VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialFrameParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' 类: cSerialFrameParser
' 说明: 串口数据帧解析器 - 按协议规则从连续字节流中切分完整帧
'
' 支持的分帧模式:
'   1. 分隔符模式 (fmDelimiter)     [DATA][DELIM][DATA][DELIM]...
'   2. 起始结束标记模式 (fmStartEnd) [STX][DATA][ETX][STX][DATA][ETX]...
'   3. 长度前缀模式 (fmLengthPrefix) [Header][Length][Data...]...
'   4. 固定长度模式 (fmFixedSize)   [Fixed bytes][Fixed bytes]...
'
' 用法示例 (分隔符模式):
'   Dim parser As New cSerialFrameParser
'   parser.SetDelimiterMode vbCrLf, False   ' 按CRLF分帧, 不含分隔符
'
'   ' 在 DataReceived 事件中:
'   Private Sub sp_DataReceived(ByVal BytesCount As Long)
'       Dim buf() As Byte
'       sp.ReadData buf
'       parser.AppendData buf
'       Do While parser.HasFrame()
'           Dim frame() As Byte
'           frame = parser.GetFrame()
'           Debug.Print StrConv(frame, vbUnicode)
'       Loop
'   End Sub
'
' 用法示例 (起始结束标记模式):
'   parser.SetStartEndMode Chr$(&H2), Chr$(&H3), True  ' STX...ETX, 含标记
'
' 用法示例 (长度前缀模式):
'   ' 协议: [LEN(1字节)][DATA(n字节)]  长度值=数据长度
'   parser.SetLengthPrefixMode 0, 1, 1, False, lmPayload
'
'   ' 协议: [STX(1)][LEN(2,大端)][DATA(n)][CRC(1)]  长度值=数据长度
'   parser.SetLengthPrefixMode 1, 2, 3, True, lmPayload  ' BaseSize=3(STX+CRC)
'===============================================================================

'===============================================================================
' 公共枚举
'===============================================================================

Public Enum eFrameMode
    fmNone = 0           ' 无分帧
    fmDelimiter = 1      ' 分隔符分帧
    fmStartEnd = 2       ' 起始结束标记分帧
    fmLengthPrefix = 3   ' 长度前缀分帧
    fmFixedSize = 4      ' 固定长度分帧
End Enum

' 长度前缀模式中长度值的含义
Public Enum eLengthMeaning
    lmPayload = 0           ' 长度值 = 数据负载长度
                            ' 帧总长 = FrameBaseSize + 长度值
    lmPayloadAndTrailer = 1 ' 长度值 = 数据负载 + 帧尾
                            ' 帧总长 = FrameBaseSize + 长度值
    lmTotalFrame = 2        ' 长度值 = 整个帧长度
                            ' 帧总长 = 长度值
End Enum

'===============================================================================
' 私有成员
'===============================================================================

Private m_Mode As eFrameMode

' 分隔符模式
Private m_Delimiter() As Byte
Private m_DelimiterLen As Long
Private m_IncludeDelimiter As Boolean

' 起始结束标记模式
Private m_StartMark() As Byte
Private m_StartLen As Long
Private m_EndMark() As Byte
Private m_EndLen As Long
Private m_IncludeMarks As Boolean

' 长度前缀模式
Private m_LengthOffset As Long        ' 长度字段在帧中的偏移
Private m_LengthSize As Long          ' 长度字段字节数 (1/2/4)
Private m_LengthBigEndian As Boolean   ' 大端序
Private m_LengthMeaning As eLengthMeaning
Private m_FrameBaseSize As Long       ' 帧基础大小(长度值范围之外的字节数)

' 固定长度模式
Private m_FrameSize As Long

' 内部缓冲区
Private m_Buf() As Byte
Private m_BufLen As Long
Private m_BufCap As Long

'===============================================================================
' 类初始化/终止
'===============================================================================

Private Sub Class_Initialize()
    m_Mode = fmNone
    m_DelimiterLen = 0
    m_StartLen = 0
    m_EndLen = 0
    m_IncludeDelimiter = False
    m_IncludeMarks = True
    m_LengthOffset = 0
    m_LengthSize = 1
    m_LengthBigEndian = False
    m_LengthMeaning = lmPayload
    m_FrameBaseSize = 1
    m_FrameSize = 0
    m_BufLen = 0
    m_BufCap = 0
End Sub

Private Sub Class_Terminate()
    Erase m_Buf
    Erase m_Delimiter
    Erase m_StartMark
    Erase m_EndMark
End Sub

'===============================================================================
' 属性 - 通用
'===============================================================================

Public Property Get Mode() As eFrameMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal Value As eFrameMode)
    m_Mode = Value
End Property

' 当前缓冲区中未解析的字节数
Public Property Get BufferLength() As Long
    BufferLength = m_BufLen
End Property

'===============================================================================
' 属性 - 分隔符模式
'===============================================================================

' 分隔符 (Byte数组)
Public Property Get Delimiter() As Variant
    If m_DelimiterLen > 0 Then
        Delimiter = m_Delimiter
    Else
        Delimiter = Empty
    End If
End Property

Public Property Let Delimiter(ByVal Value As Variant)
    If IsArray(Value) Then
        m_Delimiter = Value
        On Error Resume Next
        m_DelimiterLen = UBound(m_Delimiter) - LBound(m_Delimiter) + 1
        On Error GoTo 0
        If m_DelimiterLen > 0 Then
            ReDim Preserve m_Delimiter(m_DelimiterLen - 1)
        End If
    Else
        m_DelimiterLen = 0
    End If
End Property

' GetFrame返回的帧是否包含分隔符
Public Property Get IncludeDelimiter() As Boolean
    IncludeDelimiter = m_IncludeDelimiter
End Property

Public Property Let IncludeDelimiter(ByVal Value As Boolean)
    m_IncludeDelimiter = Value
End Property

'===============================================================================
' 属性 - 起始结束标记模式
'===============================================================================

Public Property Get StartMark() As Variant
    If m_StartLen > 0 Then
        StartMark = m_StartMark
    Else
        StartMark = Empty
    End If
End Property

Public Property Let StartMark(ByVal Value As Variant)
    If IsArray(Value) Then
        m_StartMark = Value
        On Error Resume Next
        m_StartLen = UBound(m_StartMark) - LBound(m_StartMark) + 1
        On Error GoTo 0
        If m_StartLen > 0 Then
            ReDim Preserve m_StartMark(m_StartLen - 1)
        End If
    Else
        m_StartLen = 0
    End If
End Property

Public Property Get EndMark() As Variant
    If m_EndLen > 0 Then
        EndMark = m_EndMark
    Else
        EndMark = Empty
    End If
End Property

Public Property Let EndMark(ByVal Value As Variant)
    If IsArray(Value) Then
        m_EndMark = Value
        On Error Resume Next
        m_EndLen = UBound(m_EndMark) - LBound(m_EndMark) + 1
        On Error GoTo 0
        If m_EndLen > 0 Then
            ReDim Preserve m_EndMark(m_EndLen - 1)
        End If
    Else
        m_EndLen = 0
    End If
End Property

' GetFrame返回的帧是否包含起始/结束标记
Public Property Get IncludeMarks() As Boolean
    IncludeMarks = m_IncludeMarks
End Property

Public Property Let IncludeMarks(ByVal Value As Boolean)
    m_IncludeMarks = Value
End Property

'===============================================================================
' 属性 - 长度前缀模式
'===============================================================================

' 长度字段的偏移(从帧起始算)
Public Property Get LengthOffset() As Long
    LengthOffset = m_LengthOffset
End Property

Public Property Let LengthOffset(ByVal Value As Long)
    If Value < 0 Then Value = 0
    m_LengthOffset = Value
End Property

' 长度字段大小(1/2/4字节)
Public Property Get LengthSize() As Long
    LengthSize = m_LengthSize
End Property

Public Property Let LengthSize(ByVal Value As Long)
    If Value = 1 Or Value = 2 Or Value = 4 Then
        m_LengthSize = Value
    End If
End Property

' 长度字段字节序 (True=大端, False=小端)
Public Property Get LengthBigEndian() As Boolean
    LengthBigEndian = m_LengthBigEndian
End Property

Public Property Let LengthBigEndian(ByVal Value As Boolean)
    m_LengthBigEndian = Value
End Property

' 长度值含义
Public Property Get LengthMeaning() As eLengthMeaning
    LengthMeaning = m_LengthMeaning
End Property

Public Property Let LengthMeaning(ByVal Value As eLengthMeaning)
    m_LengthMeaning = Value
End Property

' 帧基础大小(长度值之外的字节数: 帧头+帧尾等)
' 帧总长度 = FrameBaseSize + 长度值 (当LengthMeaning=lmPayload时)
Public Property Get FrameBaseSize() As Long
    FrameBaseSize = m_FrameBaseSize
End Property

Public Property Let FrameBaseSize(ByVal Value As Long)
    If Value < 0 Then Value = 0
    m_FrameBaseSize = Value
End Property

'===============================================================================
' 属性 - 固定长度模式
'===============================================================================

Public Property Get FrameSize() As Long
    FrameSize = m_FrameSize
End Property

Public Property Let FrameSize(ByVal Value As Long)
    If Value < 1 Then Value = 1
    m_FrameSize = Value
End Property

'===============================================================================
' 方法 - 数据输入
'===============================================================================

' 追加数据到内部缓冲区
Public Sub AppendData(Data() As Byte)
    Dim dataLen As Long
    On Error Resume Next
    dataLen = UBound(Data) - LBound(Data) + 1
    On Error GoTo 0
    If dataLen <= 0 Then Exit Sub

    ' 确保缓冲区已初始化
    If m_BufCap = 0 Then
        m_BufCap = 4096
        ReDim m_Buf(m_BufCap - 1) As Byte
    End If

    ' 扩容
    If m_BufLen + dataLen > m_BufCap Then
        Dim newCap As Long
        newCap = m_BufCap
        Do While newCap < m_BufLen + dataLen
            newCap = newCap * 2
        Loop
        ReDim Preserve m_Buf(newCap - 1) As Byte
        m_BufCap = newCap
    End If

    ' 追加数据
    CopyMemory m_Buf(m_BufLen), Data(LBound(Data)), dataLen
    m_BufLen = m_BufLen + dataLen
End Sub

'===============================================================================
' 方法 - 帧检测和提取
'===============================================================================

' 检查是否有完整帧可用
Public Function HasFrame() As Boolean
    Select Case m_Mode
        Case fmDelimiter
            HasFrame = HasDelimiterFrame()
        Case fmStartEnd
            HasFrame = HasStartEndFrame()
        Case fmLengthPrefix
            HasFrame = HasLengthPrefixFrame()
        Case fmFixedSize
            HasFrame = (m_BufLen >= m_FrameSize And m_FrameSize > 0)
        Case Else
            HasFrame = False
    End Select
End Function

' 取出一帧(从缓冲区移除并返回)
' 返回: Byte数组，调用前应先检查HasFrame
Public Function GetFrame() As Byte()
    Dim frame() As Byte

    Select Case m_Mode
        Case fmDelimiter
            frame = GetDelimiterFrame()
        Case fmStartEnd
            frame = GetStartEndFrame()
        Case fmLengthPrefix
            frame = GetLengthPrefixFrame()
        Case fmFixedSize
            frame = GetFixedSizeFrame()
    End Select

    GetFrame = frame
End Function

' 取出所有可用帧，返回Collection(每个元素为Byte数组)
Public Function GetAllFrames() As Collection
    Dim frames As New Collection
    Do While HasFrame()
        frames.Add GetFrame()
    Loop
    Set GetAllFrames = frames
End Function

' 清空内部缓冲区
Public Sub Clear()
    m_BufLen = 0
End Sub

'===============================================================================
' 内部方法 - 分隔符模式
'===============================================================================

Private Function HasDelimiterFrame() As Boolean
    If m_DelimiterLen = 0 Or m_BufLen < m_DelimiterLen Then
        HasDelimiterFrame = False
        Exit Function
    End If
    HasDelimiterFrame = (FindPattern(0, m_Delimiter, m_DelimiterLen) >= 0)
End Function

Private Function GetDelimiterFrame() As Byte()
    Dim frame() As Byte
    Dim pos As Long
    pos = FindPattern(0, m_Delimiter, m_DelimiterLen)

    If pos < 0 Then
        GetDelimiterFrame = frame
        Exit Function
    End If

    Dim frameLen As Long
    If m_IncludeDelimiter Then
        frameLen = pos + m_DelimiterLen
    Else
        frameLen = pos
    End If

    If frameLen > 0 Then
        ReDim frame(frameLen - 1)
        CopyMemory frame(0), m_Buf(0), frameLen
    End If

    ' 从缓冲区移除帧 + 分隔符
    RemoveFromBuffer 0, pos + m_DelimiterLen

    GetDelimiterFrame = frame
End Function

'===============================================================================
' 内部方法 - 起始结束标记模式
'===============================================================================

Private Function HasStartEndFrame() As Boolean
    If m_StartLen = 0 Or m_EndLen = 0 Then
        HasStartEndFrame = False
        Exit Function
    End If
    If m_BufLen < (m_StartLen + m_EndLen) Then
        HasStartEndFrame = False
        Exit Function
    End If

    Dim startPos As Long
    startPos = FindPattern(0, m_StartMark, m_StartLen)
    If startPos < 0 Then
        HasStartEndFrame = False
        Exit Function
    End If

    ' 在StartMark之后搜索EndMark
    Dim endPos As Long
    endPos = FindPattern(startPos + m_StartLen, m_EndMark, m_EndLen)

    HasStartEndFrame = (endPos >= 0)
End Function

Private Function GetStartEndFrame() As Byte()
    Dim frame() As Byte

    Dim startPos As Long
    startPos = FindPattern(0, m_StartMark, m_StartLen)
    If startPos < 0 Then
        GetStartEndFrame = frame
        Exit Function
    End If

    Dim endPos As Long
    endPos = FindPattern(startPos + m_StartLen, m_EndMark, m_EndLen)
    If endPos < 0 Then
        GetStartEndFrame = frame
        Exit Function
    End If

    ' 计算帧范围
    Dim frameStart As Long, frameEnd As Long
    If m_IncludeMarks Then
        frameStart = startPos
        frameEnd = endPos + m_EndLen
    Else
        frameStart = startPos + m_StartLen
        frameEnd = endPos
    End If

    Dim frameLen As Long
    frameLen = frameEnd - frameStart
    If frameLen > 0 Then
        ReDim frame(frameLen - 1)
        CopyMemory frame(0), m_Buf(frameStart), frameLen
    End If

    ' 从缓冲区移除整帧(包括StartMark到EndMark)
    RemoveFromBuffer startPos, (endPos + m_EndLen) - startPos

    GetStartEndFrame = frame
End Function

'===============================================================================
' 内部方法 - 长度前缀模式
'===============================================================================

Private Function HasLengthPrefixFrame() As Boolean
    ' 需要至少 LengthOffset + LengthSize 字节才能读取长度
    Dim minBytes As Long
    minBytes = m_LengthOffset + m_LengthSize
    If m_BufLen < minBytes Then
        HasLengthPrefixFrame = False
        Exit Function
    End If

    ' 读取长度值
    Dim lengthVal As Long
    lengthVal = ReadLengthField()

    ' 计算帧总长度
    Dim totalLen As Long
    totalLen = CalcFrameTotalLength(lengthVal)

    If totalLen <= 0 Or totalLen > 1048576 Then
        ' 无效长度，可能是垃圾数据
        HasLengthPrefixFrame = False
        Exit Function
    End If

    HasLengthPrefixFrame = (m_BufLen >= totalLen)
End Function

Private Function GetLengthPrefixFrame() As Byte()
    Dim frame() As Byte

    Dim lengthVal As Long
    lengthVal = ReadLengthField()

    Dim totalLen As Long
    totalLen = CalcFrameTotalLength(lengthVal)

    If totalLen <= 0 Or totalLen > m_BufLen Then
        GetLengthPrefixFrame = frame
        Exit Function
    End If

    ReDim frame(totalLen - 1)
    CopyMemory frame(0), m_Buf(0), totalLen

    RemoveFromBuffer 0, totalLen

    GetLengthPrefixFrame = frame
End Function

' 从缓冲区读取长度字段值
Private Function ReadLengthField() As Long
    Dim val As Long
    Dim i As Long
    Dim offset As Long
    offset = m_LengthOffset

    If m_LengthBigEndian Then
        ' 大端: 高字节在前
        For i = 0 To m_LengthSize - 1
            val = val * 256 + m_Buf(offset + i)
        Next i
    Else
        ' 小端: 低字节在前
        For i = m_LengthSize - 1 To 0 Step -1
            val = val * 256 + m_Buf(offset + i)
        Next i
    End If

    ReadLengthField = val
End Function

' 根据长度值计算帧总长度
Private Function CalcFrameTotalLength(ByVal lengthVal As Long) As Long
    Select Case m_LengthMeaning
        Case lmPayload
            CalcFrameTotalLength = m_FrameBaseSize + lengthVal
        Case lmPayloadAndTrailer
            CalcFrameTotalLength = m_FrameBaseSize + lengthVal
        Case lmTotalFrame
            CalcFrameTotalLength = lengthVal
        Case Else
            CalcFrameTotalLength = 0
    End Select
End Function

'===============================================================================
' 内部方法 - 固定长度模式
'===============================================================================

Private Function GetFixedSizeFrame() As Byte()
    Dim frame() As Byte
    If m_BufLen < m_FrameSize Or m_FrameSize <= 0 Then
        GetFixedSizeFrame = frame
        Exit Function
    End If

    ReDim frame(m_FrameSize - 1)
    CopyMemory frame(0), m_Buf(0), m_FrameSize

    RemoveFromBuffer 0, m_FrameSize

    GetFixedSizeFrame = frame
End Function

'===============================================================================
' 内部方法 - 缓冲区操作
'===============================================================================

' 在m_Buf中搜索pattern (0-based)
' 返回: 找到的位置(0-based)，-1=未找到
Private Function FindPattern(ByVal startIdx As Long, pattern() As Byte, ByVal patternLen As Long) As Long
    If patternLen = 0 Then
        FindPattern = -1
        Exit Function
    End If
    If m_BufLen < patternLen Then
        FindPattern = -1
        Exit Function
    End If
    If startIdx < 0 Then startIdx = 0

    Dim i As Long, j As Long
    Dim lastPos As Long
    lastPos = m_BufLen - patternLen

    For i = startIdx To lastPos
        For j = 0 To patternLen - 1
            If m_Buf(i + j) <> pattern(j) Then Exit For
        Next j
        If j = patternLen Then
            FindPattern = i
            Exit Function
        End If
    Next i

    FindPattern = -1
End Function

' 从缓冲区中移除指定范围的字节
Private Sub RemoveFromBuffer(ByVal startIdx As Long, ByVal count As Long)
    Dim remaining As Long
    remaining = m_BufLen - (startIdx + count)
    If remaining > 0 Then
        CopyMemory m_Buf(startIdx), m_Buf(startIdx + count), remaining
    End If
    m_BufLen = m_BufLen - count
End Sub

'===============================================================================
' 辅助方法 - 快速配置
'===============================================================================

' 设置为分隔符模式
' DelimiterStr: 分隔符字符串(如 vbCrLf, Chr(10) 等)
' IncludeDelim: GetFrame返回的帧是否包含分隔符
Public Sub SetDelimiterMode(ByVal DelimiterStr As String, Optional ByVal IncludeDelim As Boolean = False)
    m_Mode = fmDelimiter
    m_Delimiter = StrConv(DelimiterStr, vbFromUnicode)
    m_DelimiterLen = LenB(StrConv(DelimiterStr, vbFromUnicode))
    If m_DelimiterLen > 0 Then
        ReDim Preserve m_Delimiter(m_DelimiterLen - 1)
    End If
    m_IncludeDelimiter = IncludeDelim
End Sub

' 设置为起始结束标记模式
' StartStr: 起始标记字符串
' EndStr: 结束标记字符串
' IncludeMarks: GetFrame返回的帧是否包含起始/结束标记
Public Sub SetStartEndMode(ByVal StartStr As String, ByVal EndStr As String, Optional ByVal IncludeMarks As Boolean = True)
    m_Mode = fmStartEnd
    m_StartMark = StrConv(StartStr, vbFromUnicode)
    m_StartLen = LenB(StrConv(StartStr, vbFromUnicode))
    If m_StartLen > 0 Then
        ReDim Preserve m_StartMark(m_StartLen - 1)
    End If
    m_EndMark = StrConv(EndStr, vbFromUnicode)
    m_EndLen = LenB(StrConv(EndStr, vbFromUnicode))
    If m_EndLen > 0 Then
        ReDim Preserve m_EndMark(m_EndLen - 1)
    End If
    m_IncludeMarks = IncludeMarks
End Sub

' 设置为长度前缀模式
' Offset: 长度字段在帧中的偏移(0=帧首就是长度字段)
' Size: 长度字段大小(1/2/4字节)
' BaseSize: 帧基础大小(长度值之外的字节数: 帧头+帧尾等固定部分)
' BigEndian: True=大端序, False=小端序
' Meaning: 长度值含义(见eLengthMeaning枚举)
'
' 示例:
'   协议 [LEN(1)][DATA(n)]           -> SetLengthPrefixMode 0, 1, 1, False, lmPayload
'   协议 [STX(1)][LEN(2,BE)][DATA(n)][CRC(1)]
'                                    -> SetLengthPrefixMode 1, 2, 3, True, lmPayload
'                                       (BaseSize=3: STX(1)+LEN(2)在LengthOffset之前... 不对)
'
' 详细说明:
'   帧总长度 = FrameBaseSize + 长度值 (当Meaning=lmPayload或lmPayloadAndTrailer)
'   帧总长度 = 长度值                 (当Meaning=lmTotalFrame)
'
'   FrameBaseSize = 帧中不属于"长度值所代表范围"的所有字节数
'   例: [STX(1)][LEN(1)][DATA(n)][ETX(1)]
'       LengthOffset=1, LengthSize=1
'       长度值=n (DATA的长度)
'       FrameBaseSize = 1(STX) + 1(LEN) + 1(ETX) = 3
'       帧总长 = 3 + n
Public Sub SetLengthPrefixMode(ByVal Offset As Long, ByVal Size As Long, _
                              ByVal BaseSize As Long, Optional ByVal BigEndian As Boolean = False, _
                              Optional ByVal Meaning As eLengthMeaning = lmPayload)
    m_Mode = fmLengthPrefix
    m_LengthOffset = Offset
    If Size = 1 Or Size = 2 Or Size = 4 Then
        m_LengthSize = Size
    Else
        m_LengthSize = 1
    End If
    m_FrameBaseSize = BaseSize
    m_LengthBigEndian = BigEndian
    m_LengthMeaning = Meaning
End Sub

' 设置为固定长度模式
Public Sub SetFixedSizeMode(ByVal Size As Long)
    m_Mode = fmFixedSize
    If Size < 1 Then Size = 1
    m_FrameSize = Size
End Sub
