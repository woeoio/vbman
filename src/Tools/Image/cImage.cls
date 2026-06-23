VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================
' cImage - 图片数据互转类 (VBMANLIB)
' 功能: Byte() <-> Base64 <-> StdPicture 互相转换
'
' 用法:
'   ' 从 Base64 显示到 Image 控件 (传统写法)
'   Set Image1.Picture = VBMAN.Image.Data(Base64Str).ReturnPicture
'
'   ' 从 Base64 显示到 Image 控件 (简写法)
'   VBMAN.Image.(Base64Str).ShowTo Image1
'
'   ' 从字节数组获取 Base64
'   Dim b64 As String
'   b64 = VBMAN.Image.Data(bytesArr).ReturnBase64
'
'   ' 从 StdPicture 获取字节数组
'   Dim b() As Byte
'   b = VBMAN.Image.Data(SomePicture).ReturnBytes
'
' 默认成员: Data - 链式调用入口
' ============================================

' --- GDI+ API 声明 ---
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
    Token As Long, _
    ByRef inputBuf As GdiplusStartupInput, _
    Optional ByVal outputBuf As Long = 0) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
    ByVal Token As Long)

Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" ( _
    ByVal pStream As Long, _
    ByRef pBitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
    ByVal pImage As Long) As Long

' --- OLE32 API 声明 ---
Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
    ByVal hGlobal As Long, _
    ByVal fDeleteOnRelease As Long, _
    ByRef ppstm As Any) As Long

Private Declare Function OleLoadPicture Lib "oleaut32" ( _
    ByVal pStream As Long, _
    ByVal lSize As Long, _
    ByVal fRunmode As Long, _
    ByRef riid As GUID, _
    ByRef ppvObj As Any) As Long

' --- Kernel32 API ---
Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

' --- 常量 ---
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40

' --- 内部状态 ---
Private pvBytes() As Byte           ' 当前字节数据
Private pvBase64 As String          ' 当前 Base64 字符串
Private pvPicture As StdPicture     ' 当前图片对象
Private pvDataType As String        ' 数据来源类型: "bytes" / "base64" / "picture"
Private pvMimeType As String        ' MIME 类型 (如 "image/jpeg", "image/png")
Private mGdiPlusToken As Long       ' GDI+ token
Private mGdiPlusStarted As Boolean  ' GDI+ 是否已初始化

' ============================================
' 内部方法
' ============================================

Private Sub InitGdiPlus()
    If Not mGdiPlusStarted Then
        Dim gsi As GdiplusStartupInput
        gsi.GdiplusVersion = 1
        GdiplusStartup mGdiPlusToken, gsi
        mGdiPlusStarted = True
    End If
End Sub

Private Sub Class_Terminate()
    If mGdiPlusStarted Then
        GdiplusShutdown mGdiPlusToken
        mGdiPlusStarted = False
    End If
    Erase pvBytes
    pvBase64 = ""
    Set pvPicture = Nothing
End Sub

' ============================================
' 公有方法
' ============================================

' Data - 默认成员，链式调用入口
' 接受: Byte() 数组 / Base64 字符串 / StdPicture 对象
' 返回: Me (cImage) 以支持链式调用
Public Function Data(ByVal vData As Variant) As cImage
Attribute Data.VB_UserMemId = 0   ' 默认成员
    Set Data = Me
    
    ' 清除旧数据
    Erase pvBytes
    pvBase64 = ""
    Set pvPicture = Nothing
    pvDataType = ""
    pvMimeType = ""
    
    ' 判断数据类型
    If IsObject(vData) Then
        ' StdPicture 对象
        On Error Resume Next
        Set pvPicture = vData
        On Error GoTo 0
        If Not pvPicture Is Nothing Then
            pvDataType = "picture"
            pvMimeType = ""  ' 从图片对象无法直接得知格式
        End If
    ElseIf VarType(vData) = (vbArray Or vbByte) Then
        ' Byte() 数组
        pvBytes = vData
        pvDataType = "bytes"
        pvMimeType = DetectMimeType(pvBytes)
    ElseIf VarType(vData) = vbString Then
        ' 可能是 Base64 字符串 (包括 data: URI 格式)
        Dim s As String
        s = CStr(vData)
        
        ' 检查 data: URI 格式
        If Left$(s, 5) = "data:" Then
            ' 解析 MIME 类型
            Dim posColon As Long, posSemi As Long
            posColon = InStr(6, s, ":")
            posSemi = InStr(6, s, ";")
            If posSemi > 0 Then
                pvMimeType = Mid$(s, 6, posSemi - 6)
            End If
            ' 提取 base64 部分
            Dim posComma As Long
            posComma = InStr(s, ",")
            If posComma > 0 Then
                s = Mid$(s, posComma + 1)
            End If
        End If
        
        ' 去除空白
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        s = Replace(s, vbTab, "")
        s = Replace(s, " ", "")
        
        pvBase64 = s
        pvDataType = "base64"
    Else
        Err.Raise vbObjectError + 1, "cImage.Data", "不支持的数据类型。请传入 Byte()、Base64字符串 或 StdPicture对象。"
    End If
End Function

' 内部方法: 确保有字节数据
Private Sub EnsureBytes()
    If pvDataType = "base64" Then
        pvBytes = ToolsBase64.Base64ToBytes(pvBase64)
        pvDataType = "bytes"
    ElseIf pvDataType = "picture" Then
        pvBytes = PictureToBytes(pvPicture)
        pvDataType = "bytes"
    End If
    ' 已经是 bytes 则什么都不做
End Sub

' 内部方法: 确保有 Base64
Private Sub EnsureBase64()
    If pvDataType = "bytes" Then
        pvBase64 = ToolsBase64.BytesToBase64(pvBytes)
    ElseIf pvDataType = "picture" Then
        EnsureBytes
        pvBase64 = ToolsBase64.BytesToBase64(pvBytes)
    End If
End Sub

' 内部方法: 确保有图片对象
Private Sub EnsurePicture()
    If pvDataType = "bytes" Then
        Set pvPicture = BytesToPicture(pvBytes)
    ElseIf pvDataType = "base64" Then
        EnsureBytes
        Set pvPicture = BytesToPicture(pvBytes)
    End If
End Sub

' --- 输出方法 ---\r

' ReturnBytes - 返回字节数组
Public Function ReturnBytes() As Byte()
    EnsureBytes
    ReturnBytes = pvBytes
End Function

' ReturnBase64 - 返回 Base64 字符串
Public Function ReturnBase64() As String
    EnsureBase64
    ReturnBase64 = pvBase64
End Function

' ReturnPicture - 返回 StdPicture 对象 (可直接赋给 Image1.Picture)
Public Function ReturnPicture() As StdPicture
    EnsurePicture
    Set ReturnPicture = pvPicture
End Function

' ShowTo - 将图片直接显示到目标控件 (链式调用，动态赋值)
' 参数: 任何拥有 Picture 属性的控件对象 (Image/PictureBox 等)
' 用法: VBMAN.Image.(base64Str).ShowTo Image1
' 返回: Me (cImage) 以支持链式调用
Public Function ShowTo(ByVal Target As Object) As cImage
    Set ShowTo = Me
    EnsurePicture
    On Error Resume Next
    Set Target.Picture = pvPicture
    On Error GoTo 0
End Function

' ReturnDataUri - 返回 Data URI 格式字符串 (如 data:image/jpeg;base64,...)
Public Function ReturnDataUri(Optional ByVal MimeType As String) As String
    EnsureBase64
    If MimeType = "" Then MimeType = pvMimeType
    If MimeType = "" Then MimeType = "image/jpeg"  ' 默认
    ReturnDataUri = "data:" & MimeType & ";base64," & pvBase64
End Function

' ReturnMimeType - 返回检测到的 MIME 类型
Public Function ReturnMimeType() As String
    If pvMimeType = "" And pvDataType = "bytes" Then
        pvMimeType = DetectMimeType(pvBytes)
    End If
    ReturnMimeType = pvMimeType
End Function

' --- 辅助方法 ---

' SetMimeType - 手动设置 MIME 类型
Public Function SetMimeType(ByVal MimeType As String) As cImage
    Set SetMimeType = Me
    pvMimeType = MimeType
End Function

' IsValid - 检查数据是否有效
Public Function IsValid() As Boolean
    IsValid = (pvDataType <> "")
End Function

' ============================================
' 内部转换函数
' ============================================



' BytesToPicture - 从字节数组创建 StdPicture 对象 (核心方法)
' 使用 GDI+ 从内存字节流加载图片
Private Function BytesToPicture(ByRef Bytes() As Byte) As StdPicture
    If UBound(Bytes) < 0 Then
        Set BytesToPicture = Nothing
        Exit Function
    End If
    
    Dim hMem As Long
    Dim pMem As Long
    Dim pStream As Long
    Dim pBitmap As Long
    Dim ret As Long
    Dim picGUID As GUID
    Dim oPic As IPictureDisp
    
    InitGdiPlus
    
    ' 定义 IPicture GUID: {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    picGUID.Data1 = &H7BF80980
    picGUID.Data2 = &HBF32
    picGUID.Data3 = &H101A
    picGUID.Data4(0) = &H8B
    picGUID.Data4(1) = &HBB
    picGUID.Data4(2) = &H0
    picGUID.Data4(3) = &HAA
    picGUID.Data4(4) = &H0
    picGUID.Data4(5) = &H30
    picGUID.Data4(6) = &HC
    picGUID.Data4(7) = &HAB
    
    ' 步骤1: 分配全局内存，写入字节数据
    hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(Bytes) + 1)
    If hMem = 0 Then
        Set BytesToPicture = Nothing
        Exit Function
    End If
    
    pMem = GlobalLock(hMem)
    If pMem = 0 Then
        GlobalFree hMem
        Set BytesToPicture = Nothing
        Exit Function
    End If
    
    CopyMemory ByVal pMem, Bytes(0), UBound(Bytes) + 1
    GlobalUnlock hMem
    
    ' 步骤2: 创建 IStream 包装全局内存
    ret = CreateStreamOnHGlobal(hMem, 1, pStream)
    If ret <> 0 Or pStream = 0 Then
        GlobalFree hMem
        Set BytesToPicture = Nothing
        Exit Function
    End If
    
    ' 步骤3: 通过 GDI+ 从 IStream 加载图片
    ret = GdipCreateBitmapFromStream(pStream, pBitmap)
    If ret = 0 And pBitmap <> 0 Then
        ' 成功加载了 GDI+ Bitmap，现在通过 OLE 创建 IPictureDisp
        ' 注意: GdipCreateBitmapFromStream 后，GDI+ bitmap 已可用
        ' 但我们需要通过 OleLoadPicture 从同一个 IStream 加载
        ' 因为 StdPicture 不直接支持 GDI+ bitmap handle
        
        ' 释放 GDI+ bitmap (我们不需要它)
        GdipDisposeImage pBitmap
        
        ' 重置 IStream 位置到开头
        Dim LARGE_INTEGER As Currency  ' 8字节
        ' 通过 OLE IStream::Seek 回到开头...
        ' 简单方法: 销毁旧 IStream，用原始数据重建新 IStream
        ' 因为 CreateStreamOnHGlobal 使用了 fDeleteOnRelease=1,
        ' 释放 IStream 时会同时释放 hGlobal
        
        ' 释放当前 IStream (会连带释放 hMem)
        ' VB6 没有直接释放 COM 对象的方法，让它自然释放
        
        ' 重新分配内存和 IStream
        Dim hMem2 As Long, pMem2 As Long, pStream2 As Long
        hMem2 = GlobalAlloc(GMEM_MOVEABLE, UBound(Bytes) + 1)
        If hMem2 = 0 Then
            Set BytesToPicture = Nothing
            Exit Function
        End If
        
        pMem2 = GlobalLock(hMem2)
        CopyMemory ByVal pMem2, Bytes(0), UBound(Bytes) + 1
        GlobalUnlock hMem2
        
        ret = CreateStreamOnHGlobal(hMem2, 1, pStream2)
        If ret = 0 And pStream2 <> 0 Then
            ' 通过 OLE 从 IStream 加载 StdPicture
            ret = OleLoadPicture(pStream2, 0, 0, picGUID, oPic)
            If ret = 0 And Not oPic Is Nothing Then
                Set BytesToPicture = oPic
            End If
        End If
    Else
        ' GDI+ 加载失败，可能不是图片格式
        Set BytesToPicture = Nothing
    End If
    
    ' 注意: IStream 会在函数结束时自然释放
    ' 由于 CreateStreamOnHGlobal 使用了 fDeleteOnRelease=1
End Function

' PictureToBytes - 从 StdPicture 对象获取字节数组
' 通过 SavePicture 保存到临时文件再读取
Private Function PictureToBytes(ByRef Pic As StdPicture) As Byte()
    On Error GoTo EH
    
    If Pic Is Nothing Then
        ReDim PictureToBytes(-1 To -1)
        Exit Function
    End If
    
    ' 获取临时文件路径
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\vbman_image_temp.jpg"
    
    ' 保存图片到临时文件
    SavePicture Pic, tempPath
    
    ' 读取文件为字节数据
    Dim fnum As Integer
    fnum = FreeFile
    Open tempPath For Binary Access Read As #fnum
    ReDim PictureToBytes(LOF(fnum) - 1)
    Get #fnum, , PictureToBytes
    Close #fnum
    
    ' 删除临时文件
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
    
    Exit Function
    
EH:
    On Error Resume Next
    Kill Environ("TEMP") & "\vbman_image_temp.jpg"
    ReDim PictureToBytes(-1 To -1)
End Function

' DetectMimeType - 通过文件魔数检测 MIME 类型
Private Function DetectMimeType(ByRef Bytes() As Byte) As String
    If UBound(Bytes) < 3 Then
        DetectMimeType = ""
        Exit Function
    End If
    
    ' JPEG: FF D8 FF
    If Bytes(0) = &HFF And Bytes(1) = &HD8 And Bytes(2) = &HFF Then
        DetectMimeType = "image/jpeg"
        Exit Function
    End If
    
    ' PNG: 89 50 4E 47
    If UBound(Bytes) >= 3 Then
        If Bytes(0) = &H89 And Bytes(1) = &H50 And Bytes(2) = &H4E And Bytes(3) = &H47 Then
            DetectMimeType = "image/png"
            Exit Function
        End If
    End If
    
    ' GIF: 47 49 46
    If Bytes(0) = &H47 And Bytes(1) = &H49 And Bytes(2) = &H46 Then
        DetectMimeType = "image/gif"
        Exit Function
    End If
    
    ' BMP: 42 4D
    If Bytes(0) = &H42 And Bytes(1) = &H4D Then
        DetectMimeType = "image/bmp"
        Exit Function
    End If
    
    ' WebP: 52 49 46 46 ... 57 45 42 50
    If UBound(Bytes) >= 11 Then
        If Bytes(0) = &H52 And Bytes(1) = &H49 And Bytes(2) = &H46 And Bytes(3) = &H46 Then
            If Bytes(8) = &H57 And Bytes(9) = &H45 And Bytes(10) = &H42 And Bytes(11) = &H50 Then
                DetectMimeType = "image/webp"
                Exit Function
            End If
        End If
    End If
    
    ' ICO: 00 00 01 00
    If UBound(Bytes) >= 3 Then
        If Bytes(0) = 0 And Bytes(1) = 0 And Bytes(2) = 1 And Bytes(3) = 0 Then
            DetectMimeType = "image/x-icon"
            Exit Function
        End If
    End If
    
    DetectMimeType = ""
End Function
