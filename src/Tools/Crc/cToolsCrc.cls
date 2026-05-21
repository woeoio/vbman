VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsCrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=============================================================================
' CRC16 计算类 (链式调用版本)
' 支持逐位计算法和查表法
'=============================================================================

' CRC16 查找表 (基于 0xA001 多项式)
Private CRC16Table(0 To 255) As Long

' CRC32 查找表 (基于 0xEDB88320 多项式) - 使用 Currency 支持无符号 32 位
Private CRC32Table(0 To 255) As Currency

' 缓存最后一次计算结果
Private m_CRC16Value As Long
Private m_CRC32Value As Currency

' 数据缓存
Private m_DataBytes() As Byte
Private m_HasData As Boolean

' 缓存类型枚举
Public Enum CRCType
    CRC16 = 0
    CRC32 = 1
End Enum

' 数据类型枚举
Public Enum CRCDataType
    DataType_ByteArray = 0   ' 字节数组
    DataType_HexString = 1   ' HEX字符串（自动处理空格）
    DataType_Text = 2        ' 文本字符串
End Enum


'=============================================================================
' 初始化 CRC16 查找表 (只需调用一次)
' 基于多项式 0xA001 (CRC16-Modbus)
'=============================================================================
Private Sub InitCRC16Table()
    Dim i As Long, j As Long
    Dim CRC As Long
    
    For i = 0 To 255
        CRC = i
        For j = 0 To 7
            If (CRC And 1) Then
                CRC = (CRC \ 2) And &H7FFF
                CRC = CRC Xor &HA001
            Else
                CRC = (CRC \ 2) And &H7FFF
            End If
        Next j
        CRC16Table(i) = CRC
    Next i
End Sub


'=============================================================================
' 设置数据 (默认成员，可省略调用)
' 参数: DataType - 数据类型 (见 CRCDataType 枚举)
'       Content - 数据内容
' 返回: cToolsCrc 实例 (支持链式调用)
' 说明: 根据 DataType 自动处理数据并缓存到内部字节数组
'=============================================================================
Public Function Data(Optional ByVal DataType As CRCDataType = DataType_ByteArray, Optional ByVal Content As Variant) As cToolsCrc
Attribute Data.VB_UserMemId = 0
    Dim HexStr As String
    Dim i As Long, j As Long
    Dim ByteCount As Long
    Dim HexValue As String
    
    ' 清除之前的数据
    Erase m_DataBytes
    m_HasData = False
    
    If IsMissing(Content) Then
        Set Data = Me
        Exit Function
    End If
    
    Select Case DataType
    Case DataType_ByteArray
        ' 字节数组
        If IsArray(Content) Then
            m_DataBytes = Content
            m_HasData = True
        End If
        
    Case DataType_HexString
        ' HEX字符串（自动处理空格）
        HexStr = CStr(Content)
        ' 移除所有空格
        HexStr = Replace(HexStr, " ", "")
        HexStr = Replace(HexStr, vbTab, "")
        HexStr = Replace(HexStr, "-", "")
        HexStr = Replace(HexStr, ":", "")
        
        If Len(HexStr) > 0 And Len(HexStr) Mod 2 = 0 Then
            ByteCount = Len(HexStr) \ 2
            ReDim m_DataBytes(0 To ByteCount - 1)
            
            For i = 0 To ByteCount - 1
                HexValue = Mid$(HexStr, i * 2 + 1, 2)
                m_DataBytes(i) = CByte("&H" & HexValue)
            Next i
            m_HasData = True
        End If
        
    Case DataType_Text
        ' 文本字符串（转换为字节数组）
        Dim Text As String
        Text = CStr(Content)
        If Len(Text) > 0 Then
            ' 转换为 ANSI 字节数组
            m_DataBytes = StrConv(Text, vbFromUnicode)
            m_HasData = True
        End If
        
    Case Else
        ' 默认为字节数组
        If IsArray(Content) Then
            m_DataBytes = Content
            m_HasData = True
        End If
    End Select
    
    Set Data = Me
End Function


'=============================================================================
' 计算 CRC16 (逐位计算法)
' 参数: StartIndex - 起始索引 (可选, 默认为 0)
'       Length - 数据长度 (可选, 默认为 -1 表示计算到数组末尾)
' 返回: cToolsCrc 实例 (支持链式调用)
' 说明: 必须使用 Data 方法预先设置数据
'=============================================================================
Public Function CalculateCRC16(Optional ByVal StartIndex As Long = 0, Optional ByVal Length As Long = -1) As cToolsCrc
    Dim CRC16High As Byte, CRC16Low As Byte
    Dim CalCRC16Hi As Byte, CalCRC16Lo As Byte
    Dim i As Long, j As Byte
    Dim defCRC16Hi As Byte, defCRC16Lo As Byte
    Dim ActualLength As Long
    
    ' 检查是否有数据
    If Not m_HasData Then
        Set CalculateCRC16 = Me
        Exit Function
    End If
    
    ' 如果 Length 为 -1，则计算到数组末尾
    If Length = -1 Then
        ActualLength = UBound(m_DataBytes) - StartIndex + 1
    Else
        ActualLength = Length
    End If
    
    CRC16Low = &HFF
    CRC16High = &HFF
    defCRC16Hi = &HA0
    defCRC16Lo = &H1
    
    For i = StartIndex To StartIndex + ActualLength - 1
        CRC16Low = CRC16Low Xor m_DataBytes(i)
        
        For j = 0 To 7
            CalCRC16Hi = CRC16High
            CalCRC16Lo = CRC16Low
            
            CRC16High = CRC16High \ 2
            CRC16Low = CRC16Low \ 2
            
            If (CalCRC16Hi And &H1) = &H1 Then
                CRC16Low = CRC16Low Or &H80
            End If
            
            If (CalCRC16Lo And &H1) = &H1 Then
                CRC16High = CRC16High Xor defCRC16Hi
                CRC16Low = CRC16Low Xor defCRC16Lo
            End If
        Next j
    Next i
    
    ' 缓存计算结果 (转换为无符号 16 位)
    m_CRC16Value = (CLng(CRC16High) * 256 + CLng(CRC16Low)) And &HFFFF&
    
    ' 返回当前实例，支持链式调用
    Set CalculateCRC16 = Me
End Function


'=============================================================================
' 计算 CRC16 (查表法 - 高性能版)
' 参数: Data() - 数据字节数组
'       StartIndex - 起始索引 (可选, 默认为 0)
'       Length - 数据长度 (可选, 默认为 -1 表示计算到数组末尾)
' 返回: cToolsCrc 实例 (支持链式调用)
' 说明: 速度比逐位计算快约 8 倍，适合大数据量处理
'=============================================================================
'Public Function CalculateCRC16Fast(ByRef Data() As Byte, Optional ByVal StartIndex As Long = 0, Optional ByVal Length As Long = -1) As cToolsCrc
'    Dim i As Long
'    Dim Index As Long
'    Dim CRC As Long
'    Dim ActualLength As Long
'
'    ' 如果 Length 为 -1，则计算到数组末尾
'    If Length = -1 Then
'        ActualLength = UBound(Data) - StartIndex + 1
'    Else
'        ActualLength = Length
'    End If
'
'    ' CRC 初始值为 0xFFFF
'    CRC = &HFFFF&
'
'    ' 查表法计算
'    For i = StartIndex To StartIndex + ActualLength - 1
'        Index = (CRC And &HFF&) Xor Data(i)
'        CRC = (CRC \ 256) Xor CRC16Table(Index)
'    Next i
'
'    ' 缓存计算结果 (转换为无符号 16 位)
'    m_CRC16Value = CRC And &HFFFF&
'
'    ' 返回当前实例，支持链式调用
'    Set CalculateCRC16Fast = Me
'End Function


'=============================================================================
' 返回 CRC 值作为 Long 类型
' 参数: CRCTypeValue - CRC类型 (可选, 默认为 CRC16)
' 返回: CRC16 返回无符号 16 位, CRC32 返回 Long 类型(若大于 2147483647 则为负数)
' 注意: 需要无符号 32 位结果时建议使用 ReturnCurrency 或 ReturnHex
'=============================================================================
Public Function ReturnLong(Optional ByVal CRCTypeValue As CRCType = CRC16) As Long
    If CRCTypeValue = CRC32 Then
        ReturnLong = CLng(m_CRC32Value)
    Else
        ReturnLong = m_CRC16Value
    End If
End Function


'=============================================================================
' 返回 CRC32 值作为 Currency 类型 (无符号 32 位)
' 参数: CRCTypeValue - CRC类型 (可选, 默认为 CRC16)
' 返回: CRC32 返回 0 到 4294967295 的正数, CRC16 返回 0 到 65535
'=============================================================================
Public Function ReturnCurrency(Optional ByVal CRCTypeValue As CRCType = CRC16) As Currency
    If CRCTypeValue = CRC32 Then
        ReturnCurrency = m_CRC32Value
    Else
        ReturnCurrency = CCur(m_CRC16Value)
    End If
End Function


'=============================================================================
' 返回 CRC 值作为字节数组
' 参数: CRCTypeValue - CRC类型 (可选, 默认为 CRC16)
' 返回: CRC16 返回 2字节(高字节在前), CRC32 返回 4字节(高字节在前)
'=============================================================================
Public Function ReturnBytes(Optional ByVal CRCTypeValue As CRCType = CRC16) As Byte()
    Dim Result() As Byte
    Dim Value As Currency
    
    If CRCTypeValue = CRC32 Then
        ReDim Result(0 To 3)
        Value = m_CRC32Value
        Result(0) = CByte((Value \ 16777216@) And 255)   ' 最高字节
        Result(1) = CByte((Value \ 65536@) And 255)      ' 次高字节
        Result(2) = CByte((Value \ 256@) And 255)        ' 次低字节
        Result(3) = CByte(Value And 255)                   ' 最低字节
    Else
        ReDim Result(0 To 1)
        Result(0) = CByte((m_CRC16Value \ 256) And 255)   ' 高字节
        Result(1) = CByte(m_CRC16Value And 255)            ' 低字节
    End If
    
    ReturnBytes = Result
End Function


'=============================================================================
' 返回 CRC 值作为十六进制字符串
' 参数: CRCTypeValue - CRC类型 (可选, 默认为 CRC16)
' 返回: CRC16 返回 4位大写(如 "B001"), CRC32 返回 8位大写(如 "A1B2C3D4")
'=============================================================================
Public Function ReturnHex(Optional ByVal CRCTypeValue As CRCType = CRC16) As String
    Dim HighWord As Long
    Dim LowWord As Long
    Dim Value As Currency
    
    If CRCTypeValue = CRC32 Then
        Value = m_CRC32Value
        ' 将 32 位值分成高低两个字
        HighWord = CLng(Value \ 65536@)
        LowWord = CLng(Value And 65535@)
        ReturnHex = Right$("0000" & Hex$(HighWord), 4) & Right$("0000" & Hex$(LowWord), 4)
    Else
        ReturnHex = Right$("0000" & Hex$(m_CRC16Value), 4)
    End If
End Function


'=============================================================================
' 返回 CRC 值作为带分隔符的十六进制字符串 (适合显示)
' 参数: Separator - 分隔符 (可选, 默认为 " ")
'       CRCTypeValue - CRC类型 (可选, 默认为 CRC16)
' 返回: CRC16 如 "B0 01", CRC32 如 "A1 B2 C3 D4"
'=============================================================================
Public Function ReturnHexString(Optional ByVal CRCTypeValue As CRCType = CRC16, Optional ByVal Separator As String = " ") As String
    Dim HexStr As String
    Dim i As Long
    Dim Result As String
    Dim ByteStr As String
    Dim Bytes() As Byte
    
    If CRCTypeValue = CRC32 Then
        ' 直接获取字节数组来构建十六进制字符串
        Bytes = Me.ReturnBytes(CRC32)
        For i = 0 To 3
            If i > 0 Then Result = Result & Separator
            Result = Result & Right$("00" & Hex$(Bytes(i)), 2)
        Next i
    Else
        HexStr = Right$("0000" & Hex$(m_CRC16Value), 4)
        ' 将4位十六进制字符串分成2个字节
        For i = 0 To 1
            If i > 0 Then Result = Result & Separator
            ByteStr = Mid$(HexStr, i * 2 + 1, 2)
            Result = Result & ByteStr
        Next i
    End If
    
    ReturnHexString = Result
End Function


'=============================================================================
' 初始化 CRC32 查找表 (只需调用一次)
' 基于多项式 0xEDB88320 (IEEE 802.3)
' 使用 Currency 类型避免符号位问题
'=============================================================================
Private Sub InitCRC32Table()
    Dim i As Long, j As Long
    Dim CRC As Currency
    
    For i = 0 To 255
        CRC = i
        For j = 0 To 7
            If (CRC And 1) Then
                CRC = Fix(CRC / 2) Xor 3988292384@  ' 0xEDB88320
            Else
                CRC = Fix(CRC / 2)
            End If
        Next j
        CRC32Table(i) = CRC
    Next i
End Sub


'=============================================================================
' 计算 CRC32 (逐位计算法)
' 参数: Data() - 数据字节数组
'       StartIndex - 起始索引 (可选, 默认为 0)
'       Length - 数据长度 (可选, 默认为 -1 表示计算到数组末尾)
' 返回: cToolsCrc 实例 (支持链式调用)
' 使用 Currency 类型避免符号位问题
'=============================================================================
'Public Function CalculateCRC32(ByRef Data() As Byte, Optional ByVal StartIndex As Long = 0, Optional ByVal Length As Long = -1) As cToolsCrc
'    Dim i As Long, j As Long
'    Dim CRC As Currency
'    Dim ActualLength As Long
'    Dim ByteVal As Byte
'
'    ' 如果 Length 为 -1，则计算到数组末尾
'    If Length = -1 Then
'        ActualLength = UBound(Data) - StartIndex + 1
'    Else
'        ActualLength = Length
'    End If
'
'    ' CRC32 初始值为 0xFFFFFFFF (4294967295)
'    CRC = 4294967295@
'
'    ' 逐位计算法
'    For i = StartIndex To StartIndex + ActualLength - 1
'        ByteVal = Data(i)
'        CRC = CRC Xor ByteVal
'
'        For j = 0 To 7
'            If (CRC And 1) Then
'                CRC = Fix(CRC / 2) Xor 3988292384@                              ' 0xEDB88320
'            Else
'                CRC = Fix(CRC / 2)
'            End If
'        Next j
'    Next i
'
'    ' 缓存计算结果 (与 0xFFFFFFFF 异或实现取反效果)
'    m_CRC32Value = CRC Xor 4294967295@
'
'    ' 返回当前实例，支持链式调用
'    Set CalculateCRC32 = Me
'End Function


'=============================================================================
' 计算 CRC32 (查表法 - 高性能版)
' 参数: Data() - 数据字节数组
'       StartIndex - 起始索引 (可选, 默认为 0)
'       Length - 数据长度 (可选, 默认为 -1 表示计算到数组末尾)
' 返回: cToolsCrc 实例 (支持链式调用)
' 说明: 速度比逐位计算快约 8 倍，适合大数据量处理
' 使用 Currency 类型避免符号位问题
'=============================================================================
'Public Function CalculateCRC32Fast(ByRef Data() As Byte, Optional ByVal StartIndex As Long = 0, Optional ByVal Length As Long = -1) As cToolsCrc
'    Dim i As Long
'    Dim Index As Long
'    Dim CRC As Currency
'    Dim ActualLength As Long
'
'    ' 如果 Length 为 -1，则计算到数组末尾
'    If Length = -1 Then
'        ActualLength = UBound(Data) - StartIndex + 1
'    Else
'        ActualLength = Length
'    End If
'
'    ' CRC32 初始值为 0xFFFFFFFF (4294967295)
'    CRC = 4294967295@
'
'    ' 查表法计算
'    For i = StartIndex To StartIndex + ActualLength - 1
'        Index = (CRC And 255) Xor Data(i)
'        CRC = Fix(CRC / 256) Xor CRC32Table(Index)
'    Next i
'
'    ' 缓存计算结果 (与 0xFFFFFFFF 异或实现取反效果)
'    m_CRC32Value = CRC Xor 4294967295@
'
'    ' 返回当前实例，支持链式调用
'    Set CalculateCRC32Fast = Me
'End Function


'=============================================================================
' 类初始化
'=============================================================================
Private Sub Class_Initialize()
    InitCRC16Table
    '    InitCRC32Table
    m_CRC16Value = 0
    m_CRC32Value = 0
    m_HasData = False
    Erase m_DataBytes
End Sub
