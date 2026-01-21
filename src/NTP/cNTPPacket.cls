VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNTPPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cNTPPacket - NTP 协议数据包封装类
'
' Purpose: 封装 NTP 协议 v4 数据包的编码和解码
'
' Reference: RFC 5905 - Network Time Protocol Version 4
'
'=========================================================================

Option Explicit
    DefObj A-Z
    Private Const MODULE_NAME As String = "cNTPPacket"

'=========================================================================
' NTP 常量
'=========================================================================

Private Const NTP_TIMESTAMP_DELTA As Long = 2208988800  ' 70年从1900到1970的秒数
Private Const NTP_PACKET_SIZE As Long = 48              ' NTP 数据包固定大小

'=========================================================================
' NTP 数据包结构（48字节）
'=========================================================================

' 第一个字节 (Byte 0): LI VN Mode
Private m_bLI As Byte        ' Leap Indicator (2 bits)
Private m_bVN As Byte        ' Version Number (3 bits)
Private m_bMode As Byte      ' Mode (3 bits)

' 第二个字节 (Byte 1): Stratum
Private m_bStratum As Byte

' 第三个字节 (Byte 2): Poll
Private m_bPoll As Byte

' 第四个字节 (Byte 3): Precision
Private m_bPrecision As Byte

' Root Delay (Bytes 4-7): signed fixed point 16.16
Private m_lRootDelay As Long

' Root Dispersion (Bytes 8-11): unsigned fixed point 16.16
Private m_lRootDispersion As Long

' Reference Identifier (Bytes 12-15)
Private m_sReferenceID As String

' Reference Timestamp (Bytes 16-23): 64位 NTP 时间戳
Private m_dReferenceTimestamp As Double

' Origin Timestamp (Bytes 24-31): 64位 NTP 时间戳
Private m_dOriginTimestamp As Double

' Receive Timestamp (Bytes 32-39): 64位 NTP 时间戳
Private m_dReceiveTimestamp As Double

' Transmit Timestamp (Bytes 40-47): 64位 NTP 时间戳
Private m_dTransmitTimestamp As Double

'=========================================================================
' NTP 模式枚举
'=========================================================================

Public Enum NTPMode
    NTPMode_Reserved = 0
    NTPMode_SymmetricActive = 1
    NTPMode_SymmetricPassive = 2
    NTPMode_Client = 3
    NTPMode_Server = 4
    NTPMode_Broadcast = 5
    NTPMode_Control = 6
    NTPMode_Private = 7
End Enum

'=========================================================================
' NTP Stratum 层级枚举
'=========================================================================

Public Enum NTPStratum
    NTPStratum_Unspecified = 0
    NTPStratum_Primary = 1         ' 连接到原子钟
    NTPStratum_Secondary = 2       ' 二级服务器
    NTPStratum_Unsynchronized = 16 ' 未同步
    NTPStratum_Invalid = 255       ' 无效
End Enum

'=========================================================================
' 属性
'=========================================================================

' Leap Indicator (2 bits, 0-3)
Public Property Get LI() As Byte
    LI = m_bLI
End Property

Public Property Let LI(ByVal Value As Byte)
    If Value > 3 Then
        Err.Raise vbObjectError, MODULE_NAME, "LI must be 0-3"
    End If
    m_bLI = Value
End Property

' Version Number (3 bits, 0-7)
Public Property Get VN() As Byte
    VN = m_bVN
End Property

Public Property Let VN(ByVal Value As Byte)
    If Value > 7 Then
        Err.Raise vbObjectError, MODULE_NAME, "VN must be 0-7"
    End If
    m_bVN = Value
End Property

' Mode (3 bits, 0-7)
Public Property Get Mode() As NTPMode
    Mode = m_bMode
End Property

Public Property Let Mode(ByVal Value As NTPMode)
    If Value > 7 Then
        Err.Raise vbObjectError, MODULE_NAME, "Mode must be 0-7"
    End If
    m_bMode = Value
End Property

' Stratum (层级)
Public Property Get Stratum() As Byte
    Stratum = m_bStratum
End Property

Public Property Let Stratum(ByVal Value As Byte)
    m_bStratum = Value
End Property

' Poll (轮询间隔, log2秒)
Public Property Get Poll() As Byte
    Poll = m_bPoll
End Property

Public Property Let Poll(ByVal Value As Byte)
    m_bPoll = Value
End Property

' Precision (精度, log2秒)
Public Property Get Precision() As Byte
    Precision = m_bPrecision
End Property

Public Property Let Precision(ByVal Value As Byte)
    m_bPrecision = Value
End Property

' Root Delay (秒)
Public Property Get RootDelay() As Double
    RootDelay = m_lRootDelay / 65536#
End Property

Public Property Let RootDelay(ByVal Value As Double)
    m_lRootDelay = CLng(Value * 65536)
End Property

' Root Dispersion (秒)
Public Property Get RootDispersion() As Double
    RootDispersion = m_lRootDispersion / 65536#
End Property

Public Property Let RootDispersion(ByVal Value As Double)
    m_lRootDispersion = CLng(Value * 65536)
End Property

' Reference Identifier
Public Property Get ReferenceID() As String
    ReferenceID = m_sReferenceID
End Property

Public Property Let ReferenceID(ByVal Value As String)
    If Len(Value) > 4 Then
        m_sReferenceID = Left(Value, 4)
    Else
        m_sReferenceID = Value
    End If
End Property

' Reference Timestamp (本地时间)
Public Property Get ReferenceTimestamp() As Date
    ReferenceTimestamp = NtpTimeToDate(m_dReferenceTimestamp)
End Property

Public Property Let ReferenceTimestamp(ByVal Value As Date)
    m_dReferenceTimestamp = DateToNtpTime(Value)
End Property

' Origin Timestamp (本地时间)
Public Property Get OriginTimestamp() As Date
    OriginTimestamp = NtpTimeToDate(m_dOriginTimestamp)
End Property

Public Property Let OriginTimestamp(ByVal Value As Date)
    m_dOriginTimestamp = DateToNtpTime(Value)
End Property

' Receive Timestamp (本地时间)
Public Property Get ReceiveTimestamp() As Date
    ReceiveTimestamp = NtpTimeToDate(m_dReceiveTimestamp)
End Property

Public Property Let ReceiveTimestamp(ByVal Value As Date)
    m_dReceiveTimestamp = DateToNtpTime(Value)
End Property

' Transmit Timestamp (本地时间)
Public Property Get TransmitTimestamp() As Date
    TransmitTimestamp = NtpTimeToDate(m_dTransmitTimestamp)
End Property

Public Property Let TransmitTimestamp(ByVal Value As Date)
    m_dTransmitTimestamp = DateToNtpTime(Value)
End Property

'=========================================================================
' 公共方法
'=========================================================================

' 从字节数组解析 NTP 数据包
Public Function FromBytes(ByRef Data() As Byte) As Boolean
    On Error GoTo ErrorHandler
    
    If UBound(Data) + 1 < NTP_PACKET_SIZE Then
        Err.Raise vbObjectError, MODULE_NAME, "Invalid NTP packet size"
    End If
    
    ' 解析第一个字节: LI(2) VN(3) Mode(3)
    m_bLI = (Data(0) \ 64) And &H3
    m_bVN = (Data(0) \ 8) And &H7
    m_bMode = Data(0) And &H7
    
    ' 解析其他字段
    m_bStratum = Data(1)
    m_bPoll = Data(2)
    m_bPrecision = Data(3)
    
    ' Root Delay (4字节，大端序)
    m_lRootDelay = BytesToLong(Data, 4)
    
    ' Root Dispersion (4字节，大端序)
    m_lRootDispersion = BytesToLong(Data, 8)
    
    ' Reference Identifier (4字节)
    m_sReferenceID = Chr(Data(12)) & Chr(Data(13)) & Chr(Data(14)) & Chr(Data(15))
    
    ' 时间戳字段 (各8字节)
    m_dReferenceTimestamp = BytesToNtpTimestamp(Data, 16)
    m_dOriginTimestamp = BytesToNtpTimestamp(Data, 24)
    m_dReceiveTimestamp = BytesToNtpTimestamp(Data, 32)
    m_dTransmitTimestamp = BytesToNtpTimestamp(Data, 40)
    
    FromBytes = True
    Exit Function
    
ErrorHandler:
    FromBytes = False
End Function

' 将 NTP 数据包编码为字节数组
Public Function ToBytes() As Byte()
    Dim baData(NTP_PACKET_SIZE - 1) As Byte
    Dim bFirstByte As Byte
    
    ' 编码第一个字节: LI(2) VN(3) Mode(3)
    bFirstByte = (m_bLI And &H3) * 64
    bFirstByte = bFirstByte + ((m_bVN And &H7) * 8)
    bFirstByte = bFirstByte + (m_bMode And &H7)
    baData(0) = bFirstByte
    
    ' 编码其他字段
    baData(1) = m_bStratum
    baData(2) = m_bPoll
    baData(3) = m_bPrecision
    
    ' Root Delay
    LongToBytes m_lRootDelay, baData, 4
    
    ' Root Dispersion
    LongToBytes m_lRootDispersion, baData, 8
    
    ' Reference Identifier
    Dim i As Long
    For i = 0 To 3
        If i < Len(m_sReferenceID) Then
            baData(12 + i) = Asc(Mid(m_sReferenceID, i + 1, 1))
        End If
    Next i
    
    ' 时间戳字段
    NtpTimestampToBytes m_dReferenceTimestamp, baData, 16
    NtpTimestampToBytes m_dOriginTimestamp, baData, 24
    NtpTimestampToBytes m_dReceiveTimestamp, baData, 32
    NtpTimestampToBytes m_dTransmitTimestamp, baData, 40
    
    ToBytes = baData
End Function

' 创建客户端请求包
Public Function CreateClientRequest() As cNTPPacket
    Dim oPacket As New cNTPPacket
    
    oPacket.LI = 0
    oPacket.VN = 4         ' NTP v4
    oPacket.Mode = NTPMode_Client
    oPacket.Stratum = 0
    oPacket.Poll = 4       ' 2^4 = 16秒
    oPacket.Precision = &HE8  ' 2^-8 = 0.00390625秒
    
    Set CreateClientRequest = oPacket
End Function

' 创建服务器响应包
Public Function CreateServerResponse(ByVal RequestPacket As cNTPPacket) As cNTPPacket
    Dim oPacket As New cNTPPacket
    
    ' 设置基本字段
    oPacket.LI = 0
    oPacket.VN = RequestPacket.VN
    oPacket.Mode = NTPMode_Server
    oPacket.Stratum = 1  ' 一级服务器
    oPacket.Poll = RequestPacket.Poll
    oPacket.Precision = &HE8
    oPacket.RootDelay = 0
    oPacket.RootDispersion = 0.01
    oPacket.ReferenceID = "LOCL"
    
    ' 复制时间戳
    oPacket.OriginTimestamp = Now
    oPacket.ReceiveTimestamp = Now
    oPacket.TransmitTimestamp = Now
    
    Set CreateServerResponse = oPacket
End Function

' 清空数据包
Public Sub Clear()
    m_bLI = 0
    m_bVN = 4
    m_bMode = NTPMode_Client
    m_bStratum = 0
    m_bPoll = 4
    m_bPrecision = &HE8
    m_lRootDelay = 0
    m_lRootDispersion = 0
    m_sReferenceID = ""
    m_dReferenceTimestamp = 0
    m_dOriginTimestamp = 0
    m_dReceiveTimestamp = 0
    m_dTransmitTimestamp = 0
End Sub

'=========================================================================
' 私有辅助方法
'=========================================================================

' 将字节数组转换为 Long (大端序)
Private Function BytesToLong(ByRef Data() As Byte, ByVal Offset As Long) As Long
    Dim lResult As Long
    lResult = CLng(Data(Offset)) * &H1000000
    lResult = lResult + CLng(Data(Offset + 1)) * &H10000
    lResult = lResult + CLng(Data(Offset + 2)) * &H100
    lResult = lResult + CLng(Data(Offset + 3))
    BytesToLong = lResult
End Function

' 将 Long 转换为字节数组 (大端序)
Private Sub LongToBytes(ByVal Value As Long, ByRef Data() As Byte, ByVal Offset As Long)
    Data(Offset) = (Value \ &H1000000) And &HFF
    Data(Offset + 1) = (Value \ &H10000) And &HFF
    Data(Offset + 2) = (Value \ &H100) And &HFF
    Data(Offset + 3) = Value And &HFF
End Sub

' 将字节数组转换为 NTP 时间戳 (8字节)
Private Function BytesToNtpTimestamp(ByRef Data() As Byte, ByVal Offset As Long) As Double
    Dim lSeconds As Long
    Dim lFraction As Long
    
    lSeconds = BytesToLong(Data, Offset)
    lFraction = BytesToLong(Data, Offset + 4)
    
    ' NTP 时间戳 = 秒 + 分数/2^32
    BytesToNtpTimestamp = CDbl(lSeconds) + CDbl(lFraction) / 4294967296#
End Function

' 将 NTP 时间戳转换为字节数组 (8字节)
Private Sub NtpTimestampToBytes(ByVal Timestamp As Double, ByRef Data() As Byte, ByVal Offset As Long)
    Dim lSeconds As Long
    Dim lFraction As Double
    Dim dFraction As Double
    
    lSeconds = Fix(Timestamp)
    dFraction = Timestamp - lSeconds
    lFraction = Fix(dFraction * 4294967296#)
    
    LongToBytes lSeconds, Data, Offset
    LongToBytes lFraction, Data, Offset + 4
End Sub

' 将 NTP 时间戳 (自1900年的秒数) 转换为 Date
Private Function NtpTimeToDate(ByVal NtpTime As Double) As Date
    Dim dSeconds As Double
    dSeconds = NtpTime - NTP_TIMESTAMP_DELTA
    NtpTimeToDate = DateAdd("s", dSeconds, #1/1/1970#)
End Function

' 将 Date 转换为 NTP 时间戳 (自1900年的秒数)
Private Function DateToNtpTime(ByVal Dt As Date) As Double
    Dim dSeconds As Double
    dSeconds = DateDiff("s", #1/1/1970#, Dt)
    DateToNtpTime = dSeconds + NTP_TIMESTAMP_DELTA
End Function

'=========================================================================
' 类初始化/终止
'=========================================================================

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
End Sub
