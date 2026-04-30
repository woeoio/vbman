VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' 高级文件操作类 (Open 语句实现)
'   - 使用 VB 原生 Open 语句，兼容性好
'   - 一个实例同一时间只能操作一个文件
'   - 读取内容先进入内部缓冲区，再通过 ReturnXXX 提取
'   - 接口与 cFileIO 保持一致
' ============================================================

' --- Win32 API 声明 (用于编码转换和内存操作) ---
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' --- 常量 ---
Private Const CP_UTF8                   As Long = 65001
Private Const CP_ACP                    As Long = 0

' --- 成员变量 ---
Private m_FileNo        As Integer      ' 文件号
Private m_Path          As String       ' 打开的文件路径
Private m_Buffer()      As Byte         ' 内部数据缓冲区
Private m_IsOpen        As Boolean      ' 是否已打开
Private m_LastError     As String       ' 最后一次错误信息
Private m_AccessMode    As String       ' 当前访问模式 R/W/RW

' ============================================================
' 属性
' ============================================================

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property

Public Property Get FilePath() As String
    FilePath = m_Path
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

' 文件大小 (返回 Currency 类型，单位字节，兼容 cFileIO 接口)
Public Property Get FileSize() As Currency
    If Not m_IsOpen Then ERR.Raise 53, "cFileEx.FileSize", "文件未打开"
    Dim fLen As Currency
    fLen = FileLen(m_Path)
    FileSize = fLen
End Property

' 当前缓冲区字节数
Public Property Get BufferSize() As Long
    If Not IsBufferValid Then
        BufferSize = 0
    Else
        BufferSize = UBound(m_Buffer) - LBound(m_Buffer) + 1
    End If
End Property

' ============================================================
' 文件打开/关闭
' ============================================================

' 打开文件，返回 Me 支持链式调用
' AccessMode: "R"=只读, "W"=只写, "RW"=读写(默认)
Public Function OpenFile(ByVal FilePath As String, Optional ByVal AccessMode As String = "RW") As cFileEx
    CloseFile
    
    m_AccessMode = UCase$(Trim$(AccessMode))
    
    ' 获取可用文件号
    m_FileNo = FreeFile
    
    On Error GoTo OpenError
    
    Select Case m_AccessMode
        Case "R"
            ' 只读模式 - 文件必须存在
            If Dir(FilePath, vbNormal) = "" Then
                ERR.Raise 53, "cFileEx.OpenFile", "文件不存在: " & FilePath
            End If
            Open FilePath For Binary Access Read As #m_FileNo
            
        Case "W"
            ' 只写模式 - 创建或覆盖
            Open FilePath For Output As #m_FileNo
            Close #m_FileNo
            ' 重新以 Binary 打开用于写入
            Open FilePath For Binary Access Write As #m_FileNo
            
        Case Else
            ' 读写模式 (默认)
            If Dir(FilePath, vbNormal) = "" Then
                ' 文件不存在，创建新文件
                Open FilePath For Output As #m_FileNo
                Close #m_FileNo
            End If
            Open FilePath For Binary Access Read Write As #m_FileNo
            m_AccessMode = "RW"
    End Select
    
    m_Path = FilePath
    m_IsOpen = True
    m_LastError = ""
    Set OpenFile = Me
    Exit Function
    
OpenError:
    m_LastError = "无法打开文件: " & FilePath & " - " & ERR.Description
    ERR.Raise ERR.Number, "cFileEx.OpenFile", m_LastError
End Function

Public Sub CloseFile()
    If m_IsOpen And m_FileNo > 0 Then
        On Error Resume Next
        Close #m_FileNo
        On Error GoTo 0
        m_FileNo = 0
        m_IsOpen = False
        m_Path = ""
        m_AccessMode = ""
    End If
End Sub

' ============================================================
' 缓冲区操作
' ============================================================

' 清空内部缓冲区
Public Function ClearBuffer() As cFileEx
    Erase m_Buffer
    Set ClearBuffer = Me
End Function

' 用外部字节数组填充缓冲区，返回 Me
Public Function SetBuffer(Data() As Byte) As cFileEx
    m_Buffer = Data
    Set SetBuffer = Me
End Function

' 追加字节到缓冲区
Public Function AppendBuffer(Data() As Byte) As cFileEx
    Dim oldLen As Long, newLen As Long, addLen As Long
    oldLen = UBound(m_Buffer) + 1
    addLen = UBound(Data) + 1
    If addLen <= 0 Then
        Set AppendBuffer = Me
        Exit Function
    End If
    newLen = oldLen + addLen
    ReDim Preserve m_Buffer(newLen - 1)
    CopyMemory m_Buffer(oldLen), Data(0), addLen
    Set AppendBuffer = Me
End Function

' 用字符串填充缓冲区（默认 UTF-8），返回 Me
Public Function SetBufferText(ByVal Text As String, Optional ByVal CharSet As String = "UTF-8") As cFileEx
    m_Buffer = StringToBytes(Text, CharSet)
    Set SetBufferText = Me
End Function

' ============================================================
' 数据读取（读入内部缓冲区）
' ============================================================

' 将文件指定范围读入内部缓冲区，返回 Me 支持链式调用。
'   StartPos / EndPos 为 0-based 字节偏移；传 -1 表示使用默认值。
'   StartPos = -1 : 从文件开头（第 0 字节）开始
'   EndPos   = -1 : 读到文件最后一个字节
Public Function ReadData(Optional ByVal StartPos As Long = -1, Optional ByVal EndPos As Long = -1) As cFileEx
    If Not m_IsOpen Then ERR.Raise 53, "cFileEx.ReadData", "文件未打开"
    
    Dim fLen As Long
    fLen = LOF(m_FileNo)
    If fLen <= 0 Then
        Erase m_Buffer
        Set ReadData = Me
        Exit Function
    End If
    
    Dim sPos As Long, ePos As Long
    sPos = IIf(StartPos < 0, 0, StartPos)
    If sPos < 0 Then sPos = 0
    If sPos >= fLen Then sPos = fLen - 1
    
    ePos = IIf(EndPos < 0, fLen - 1, EndPos)
    If ePos >= fLen Then ePos = fLen - 1
    If ePos < sPos Then ePos = sPos
    
    Dim rLen As Long
    rLen = ePos - sPos + 1
    
    ' Open 语句限制：一次最多读取 2GB
    If rLen > 2147483647 Then
        m_LastError = "单次读取不能超过 2GB"
        ERR.Raise 6, "cFileEx.ReadData", m_LastError
    End If
    
    ' 定位到开始位置并读取
    ReDim m_Buffer(0 To rLen - 1)
    
    On Error GoTo ReadError
    
    ' 使用 Get 语句读取二进制数据
    Get #m_FileNo, sPos + 1, m_Buffer
    
    m_LastError = ""
    Set ReadData = Me
    Exit Function
    
ReadError:
    m_LastError = "读取文件失败: " & ERR.Description
    ERR.Raise 5, "cFileEx.ReadData", m_LastError
End Function

' ============================================================
' 数据返回（从内部缓冲区提取）
' ============================================================

' 返回文本，CharSet 默认 UTF-8
Public Function ReturnText(Optional ByVal CharSet As String = "UTF-8") As String
    If Not IsBufferValid Then
        ReturnText = ""
        Exit Function
    End If
    ReturnText = BytesToString(m_Buffer, CharSet)
End Function

' 返回字节数组（副本）
Public Function ReturnBytes() As Byte()
    If Not IsBufferValid Then
        Dim emptyBuf() As Byte
        ReturnBytes = emptyBuf
        Exit Function
    End If
    ReturnBytes = m_Buffer
End Function

' 返回 Base64 编码字符串
Public Function ReturnBase64() As String
    If Not IsBufferValid Then
        ReturnBase64 = ""
        Exit Function
    End If
    ReturnBase64 = ToolsBase64.BytesToBase64(m_Buffer)
End Function

' 返回 Hex 编码字符串，可选分隔符（默认空格，每2字符一组）
Public Function ReturnHex(Optional ByVal Separator As String = " ") As String
    If Not IsBufferValid Then
        ReturnHex = ""
        Exit Function
    End If
    
    Dim i As Long, ub As Long
    Dim result As String
    Dim hexByte As String
    
    ub = UBound(m_Buffer)
    result = ""
    
    For i = LBound(m_Buffer) To ub
        ' 格式化为2位十六进制，不足补0
        hexByte = Right$("0" & Hex$(m_Buffer(i)), 2)
        result = result & hexByte
        ' 非最后一个字节时添加分隔符
        If i < ub Then
            result = result & Separator
        End If
    Next i
    
    ReturnHex = result
End Function

' 返回二进制字符串，可选分隔符（默认空格，每8字符一组）
Public Function ReturnBinString(Optional ByVal Separator As String = " ") As String
    If Not IsBufferValid Then
        ReturnBinString = ""
        Exit Function
    End If
    
    Dim i As Long, j As Long, ub As Long
    Dim result As String
    Dim binByte As String
    Dim b As Byte
    
    ub = UBound(m_Buffer)
    result = ""
    
    For i = LBound(m_Buffer) To ub
        b = m_Buffer(i)
        binByte = ""
        ' 将字节转换为8位二进制
        For j = 7 To 0 Step -1
            If (b And (2 ^ j)) <> 0 Then
                binByte = binByte & "1"
            Else
                binByte = binByte & "0"
            End If
        Next j
        result = result & binByte
        ' 非最后一个字节时添加分隔符
        If i < ub Then
            result = result & Separator
        End If
    Next i
    
    ReturnBinString = result
End Function

' 检查缓冲区是否有效（有数据）
Private Function IsBufferValid() As Boolean
    On Error Resume Next
    Dim lb As Long
    lb = LBound(m_Buffer)
    If ERR.Number <> 0 Then
        IsBufferValid = False
        ERR.Clear
    Else
        IsBufferValid = (UBound(m_Buffer) >= lb)
    End If
    On Error GoTo 0
End Function

' ============================================================
' 数据写入（从内部缓冲区写入文件）
' ============================================================

' 将内部缓冲区写入文件。
'   FilePath : 目标路径，未提供则使用 OpenFile 缓存的路径
'   IsAppend : False=覆盖(默认), True=追加到文件末尾
Public Function SaveData(Optional ByVal FilePath As String, Optional ByVal IsAppend As Boolean = False) As cFileEx
    Dim targetPath As String
    targetPath = IIf(Len(FilePath) > 0, FilePath, m_Path)
    If Len(targetPath) = 0 Then
        m_LastError = "未提供文件路径"
        ERR.Raise 5, "cFileEx.SaveData", m_LastError
    End If
    
    Dim hFile As Integer
    Dim needClose As Boolean
    
    On Error GoTo SaveError
    
    If targetPath = m_Path And m_IsOpen Then
        ' 写入当前打开的文件
        hFile = m_FileNo
        needClose = False
    Else
        ' 打开目标文件
        hFile = FreeFile
        If IsAppend And Dir(targetPath, vbNormal) <> "" Then
            ' 追加模式且文件存在
            Open targetPath For Binary Access Read Write As #hFile
        Else
            ' 覆盖模式或新建文件
            Open targetPath For Output As #hFile
            Close #hFile
            Open targetPath For Binary Access Write As #hFile
        End If
        needClose = True
    End If
    
    ' 追加模式：定位到文件末尾
    If IsAppend Then
        Seek #hFile, LOF(hFile) + 1
    ElseIf needClose Then
        ' 新建文件，从头开始写
        Seek #hFile, 1
    End If
    
    ' 写入缓冲区
    If IsBufferValid Then
        Put #hFile, , m_Buffer
    End If
    
    If needClose Then
        Close #hFile
    End If
    
    m_LastError = ""
    Set SaveData = Me
    Exit Function
    
SaveError:
    If needClose Then On Error Resume Next: Close #hFile: On Error GoTo 0
    m_LastError = "写入文件失败: " & ERR.Description
    ERR.Raise 5, "cFileEx.SaveData", m_LastError
End Function

' ============================================================
' 辅助函数
' ============================================================

' 字节数组 -> 字符串（支持 UTF-8 / ANSI / UTF-16LE）
Private Function BytesToString(ByRef b() As Byte, ByVal CharSet As String) As String
    Dim lb As Long: lb = LBound(b)
    Dim cnt As Long: cnt = UBound(b) - lb + 1
    If cnt <= 0 Then Exit Function
    
    Select Case LCase$(Trim$(CharSet))
        Case "utf-8", "utf8"
            ' 检测并跳过 BOM
            Dim startIdx As Long: startIdx = lb
            If cnt >= 3 Then
                If b(lb) = &HEF And b(lb + 1) = &HBB And b(lb + 2) = &HBF Then
                    startIdx = lb + 3
                    cnt = cnt - 3
                End If
            End If
            If cnt > 0 Then
                Dim wlen As Long
                wlen = MultiByteToWideChar(CP_UTF8, 0, b(startIdx), cnt, 0, 0)
                If wlen > 0 Then
                    BytesToString = String$(wlen, vbNullChar)
                    MultiByteToWideChar CP_UTF8, 0, b(startIdx), cnt, StrPtr(BytesToString), wlen
                End If
            End If
            
        Case "utf-16", "utf-16le", "utf16", "unicode"
            ' 检测并跳过 BOM
            startIdx = lb
            If cnt >= 2 Then
                If b(lb) = &HFF And b(lb + 1) = &HFE Then
                    startIdx = lb + 2
                    cnt = cnt - 2
                End If
            End If
            If cnt > 0 Then
                BytesToString = b
            End If
            
        Case Else ' ANSI / 系统默认编码
            wlen = MultiByteToWideChar(CP_ACP, 0, b(lb), cnt, 0, 0)
            If wlen > 0 Then
                BytesToString = String$(wlen, vbNullChar)
                MultiByteToWideChar CP_ACP, 0, b(lb), cnt, StrPtr(BytesToString), wlen
            End If
    End Select
End Function

' 字符串 -> 字节数组（支持 UTF-8 / ANSI / UTF-16LE）
Private Function StringToBytes(ByVal s As String, ByVal CharSet As String) As Byte()
    Dim b() As Byte
    If Len(s) = 0 Then
        StringToBytes = b
        Exit Function
    End If
    
    Select Case LCase$(Trim$(CharSet))
        Case "utf-8", "utf8"
            Dim blen As Long
            blen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), ByVal vbNullString, 0, vbNullString, 0)
            If blen > 0 Then
                ReDim b(blen - 1)
                WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), b(0), blen, vbNullString, 0
            End If
            
        Case "utf-16", "utf-16le", "utf16", "unicode"
            b = s
            
        Case Else ' ANSI
            blen = WideCharToMultiByte(CP_ACP, 0, StrPtr(s), Len(s), ByVal vbNullString, 0, vbNullString, 0)
            If blen > 0 Then
                ReDim b(blen - 1)
                WideCharToMultiByte CP_ACP, 0, StrPtr(s), Len(s), b(0), blen, vbNullString, 0
            End If
    End Select
    StringToBytes = b
End Function

Private Sub Class_Terminate()
    CloseFile
End Sub
