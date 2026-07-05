VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' 公共枚举 - 通信事件/错误/Modem状态
' 编译为 DLL 后随类型库导出，供外部工程使用
' （原为 modSerialPortAPI.bas 中的 Public Const，标准模块常量不会导出至类型库）
'===============================================================================

' 通信事件 (SetCommMask / WaitCommEvent / PinChanged 事件)
Public Enum eCommEvent
    EV_RXCHAR = &H1&          ' 字符已接收
    EV_RXFLAG = &H2&          ' 事件字符已接收
    EV_TXEMPTY = &H4&         ' 发送缓冲区已空
    EV_CTS = &H8&             ' CTS信号变化
    EV_DSR = &H10&            ' DSR信号变化
    EV_RLSD = &H20&           ' RLSD(CD)信号变化
    EV_BREAK = &H40&          ' 检测到中断
    EV_ERR = &H80&            ' 线路状态错误
    EV_RING = &H100&          ' 振铃检测
    EV_PERR = &H200&          ' 打印机错误
    EV_RX80FULL = &H400&      ' 接收缓冲区80%满
    EV_EVENT1 = &H800&        ' 设备事件1
    EV_EVENT2 = &H1000&       ' 设备事件2
End Enum

' 通信错误 (ClearCommError / ErrorOccurred 事件 / ReadCommErrors)
Public Enum eCommError
    CE_RXOVER = &H1&          ' 接收缓冲区溢出
    CE_OVERRUN = &H2&         ' 字符覆盖（缓冲区满）
    CE_RXPARITY = &H4&        ' 校验错误
    CE_FRAME = &H8&           ' 帧错误
    CE_BREAK = &H10&          ' 中断检测
    CE_TXFULL = &H100&        ' 发送缓冲区满
    CE_PTO = &H200&           ' 并口超时
    CE_IOE = &H400&           ' I/O错误
    CE_DNS = &H800&           ' 设备未选择
    CE_OOP = &H1000&          ' 缺纸
    CE_MODE = &H8000&         ' 模式错误
End Enum

' Modem 状态 (GetCommModemStatus / ReadModemStatus)
Public Enum eModemStatus
    MS_CTS_ON = &H10&         ' CTS 清除发送
    MS_DSR_ON = &H20&         ' DSR 数据设备就绪
    MS_RING_ON = &H40&        ' RING 振铃
    MS_RLSD_ON = &H80&        ' RLSD/CD 载波检测
End Enum

'===============================================================================
' 类: cSerialPort
' 说明: 纯 Win32 API 串口通信类
'       完整封装 Windows 串口通信 API 的所有能力:
'       - 打开/关闭串口 (CreateFile/CloseHandle)
'       - 读写数据 (ReadFile/WriteFile/TransmitCommChar)
'       - 配置参数 (DCB/COMMTIMEOUTS/SetupComm)
'       - 信号线控制 (EscapeCommFunction: DTR/RTS/Break/XON/XOFF)
'       - 信号线状态查询 (GetCommModemStatus: CTS/DSR/RING/CD)
'       - 错误处理 (ClearCommError)
'       - 缓冲区管理 (PurgeComm)
'       - 事件监控 (SetCommMask/WaitCommEvent + 异步Poll)
'       - 串口属性查询 (GetCommProperties)
'       - Break信号 (SetCommBreak/ClearCommBreak)
'
' 用法示例:
'   Dim sp As New cSerialPort
'   sp.PortName = "COM3"
'   sp.Config.BaudRate = br115200
'   If sp.OpenPort() Then
'       sp.WriteText "Hello"
'       Debug.Print sp.ReadExisting()
'       sp.ClosePort
'   End If
'===============================================================================

'===============================================================================
' 事件
'===============================================================================

' 数据接收 - 当输入缓冲区有新数据时触发（需StartMonitoring）
Public Event DataReceived(ByVal BytesCount As Long)

' 通信错误
Public Event ErrorOccurred(ByVal ErrorCode As Long, ByVal ErrorMsg As String)

' 信号线变化 - CTS/DSR/RING/RLSD任一变化
Public Event PinChanged(ByVal Events As Long)

' 发送缓冲区已空
Public Event TxEmpty()

' 检测到Break信号
Public Event BreakDetected()

' 检测到振铃
Public Event RingDetected()

' 帧接收 - 当启用分帧模式(FrameInterval>0)后，每收到一个完整帧触发
' FrameData: 这一帧的原始字节数据
' 注意: 需要先设置 FrameInterval > 0 并调用 StartMonitoring
Public Event FrameReceived(FrameData() As Byte)

'===============================================================================
' 私有成员
'===============================================================================

Private m_Handle As Long                                                        ' 串口句柄
Private m_PortName As String                                                    ' 端口名
Private m_Config As cSerialConfig                                               ' 配置对象
Private m_IsOpen As Boolean                                                     ' 是否打开
Private m_TimerID As Long                                                       ' 监控定时器ID
Private m_Monitoring As Boolean                                                 ' 是否正在监控
Private m_PollInterval As Long                                                  ' 轮询间隔(ms)
Private m_LastError As Long                                                     ' 最后的错误码
Private m_ReceivedCount As Long                                                 ' 已接收总字节数
Private m_SentCount As Long                                                     ' 已发送总字节数

Private m_FrameInterval As Long                                                   ' 帧间隔超时(ms)，0=禁用分帧
Private m_MaxFrameSize As Long                                                    ' 单帧最大字节数(防溢出)
Private m_FrameBuf() As Byte                                                      ' 帧缓冲区
Private m_FrameBufLen As Long                                                      ' 帧缓冲区有效长度
Private m_FrameBufCap As Long                                                      ' 帧缓冲区容量
Private m_LastRecvTime As Long                                                     ' 最后收到数据的时间(GetTickCount)

'===============================================================================
' 类初始化/终止
'===============================================================================

Private Sub Class_Initialize()
    Set m_Config = New cSerialConfig
    m_PortName = "COM1"
    m_Handle = INVALID_HANDLE_VALUE
    m_IsOpen = False
    m_Monitoring = False
    m_TimerID = 0
    m_PollInterval = 50
    m_LastError = 0
    m_ReceivedCount = 0
    m_SentCount = 0
    m_FrameInterval = 0
    m_MaxFrameSize = 65536
    m_FrameBufLen = 0
    m_FrameBufCap = 0
    m_LastRecvTime = 0
End Sub

Private Sub Class_Terminate()
    If m_Monitoring Then StopMonitoring
    If m_IsOpen Then ClosePort
    Set m_Config = Nothing
End Sub

'===============================================================================
' 属性 - 状态
'===============================================================================

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property

Public Property Get PortName() As String
    PortName = m_PortName
End Property

Public Property Let PortName(ByVal Value As String)
    If m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口已打开，无法修改端口名"
    End If
    m_PortName = Value
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

Public Property Get LastError() As Long
    LastError = m_LastError
End Property

Public Property Get LastErrorMsg() As String
    LastErrorMsg = GetSystemErrorMessage(m_LastError)
End Property

Public Property Get ReceivedCount() As Long
    ReceivedCount = m_ReceivedCount
End Property

Public Property Get SentCount() As Long
    SentCount = m_SentCount
End Property

'===============================================================================
' 属性 - 配置
'===============================================================================

Public Property Get Config() As cSerialConfig
    Set Config = m_Config
End Property

Public Property Set Config(ByVal Value As cSerialConfig)
    If Value Is Nothing Then
        ERR.Raise 5, "cSerialPort", "配置对象不能为Nothing"
    End If
    Set m_Config = Value
    If m_IsOpen Then ApplyConfig
End Property

'===============================================================================
' 属性 - 输入缓冲区字节数
'===============================================================================

Public Property Get InBufferCount() As Long
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        InBufferCount = cs.cbInQue
    End If
End Property

'===============================================================================
' 属性 - 输出缓冲区字节数
'===============================================================================

Public Property Get OutBufferCount() As Long
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        OutBufferCount = cs.cbOutQue
    End If
End Property

'===============================================================================
' 属性 - 信号线状态
'===============================================================================

Public Property Get CtsHolding() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim stat As Long
    If GetCommModemStatus(m_Handle, stat) <> 0 Then
        CtsHolding = (stat And MS_CTS_ON) <> 0
    End If
End Property

Public Property Get DsrHolding() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim stat As Long
    If GetCommModemStatus(m_Handle, stat) <> 0 Then
        DsrHolding = (stat And MS_DSR_ON) <> 0
    End If
End Property

Public Property Get RingHolding() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim stat As Long
    If GetCommModemStatus(m_Handle, stat) <> 0 Then
        RingHolding = (stat And MS_RING_ON) <> 0
    End If
End Property

Public Property Get CdHolding() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim stat As Long
    If GetCommModemStatus(m_Handle, stat) <> 0 Then
        CdHolding = (stat And MS_RLSD_ON) <> 0
    End If
End Property

'===============================================================================
' 属性 - 通信状态 (COMSTAT)
'===============================================================================

Public Property Get IsWaitingOnCts() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        IsWaitingOnCts = ComstatCtsHold(cs.fBitFields)
    End If
End Property

Public Property Get IsWaitingOnDsr() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        IsWaitingOnDsr = ComstatDsrHold(cs.fBitFields)
    End If
End Property

Public Property Get IsWaitingOnRlsd() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        IsWaitingOnRlsd = ComstatRlsdHold(cs.fBitFields)
    End If
End Property

Public Property Get IsXoffHold() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        IsXoffHold = ComstatXoffHold(cs.fBitFields)
    End If
End Property

Public Property Get IsXoffSent() As Boolean
    If Not m_IsOpen Then Exit Property
    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        IsXoffSent = ComstatXoffSent(cs.fBitFields)
    End If
End Property

'===============================================================================
' 属性 - 监控状态
'===============================================================================

Public Property Get Monitoring() As Boolean
    Monitoring = m_Monitoring
End Property

Public Property Get PollInterval() As Long
    PollInterval = m_PollInterval
End Property

Public Property Let PollInterval(ByVal Value As Long)
    If Value < 10 Then Value = 10
    If Value > 60000 Then Value = 60000
    m_PollInterval = Value
End Property

'===============================================================================
' 属性 - 帧分帧 (方案B: 帧间隔超时分帧)
'===============================================================================

' 帧间隔超时(ms) - 0=禁用分帧(原始DataReceived模式)
' 设为>0时启用分帧: 连续收到数据视为同一帧，超过此间隔无新数据则认为一帧结束
' 典型值: 20-100ms (取决于设备协议)
' 注意: StartMonitoring的轮询间隔建议 <= FrameInterval/2 以保证响应及时
Public Property Get FrameInterval() As Long
    FrameInterval = m_FrameInterval
End Property

Public Property Let FrameInterval(ByVal Value As Long)
    If Value < 0 Then Value = 0
    If Value > 5000 Then Value = 5000
    m_FrameInterval = Value
End Property

' 单帧最大字节数 - 防止异常数据导致帧缓冲区无限增长
' 超过此大小时强制触发FrameReceived并清空缓冲区
Public Property Get MaxFrameSize() As Long
    MaxFrameSize = m_MaxFrameSize
End Property

Public Property Let MaxFrameSize(ByVal Value As Long)
    If Value < 256 Then Value = 256
    If Value > 1048576 Then Value = 1048576
    m_MaxFrameSize = Value
End Property

' 帧缓冲区当前长度(已收到的未触发帧的数据量)
Public Property Get FrameBufferLength() As Long
    FrameBufferLength = m_FrameBufLen
End Property

'===============================================================================
' 方法 - 打开串口
'===============================================================================

Public Function OpenPort(Optional ByVal Port As String, _
                         Optional ByVal Config As cSerialConfig) As Boolean
    If m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口已打开，请先关闭"
    End If

    If Len(Port) > 0 Then m_PortName = Port
    If Not Config Is Nothing Then Set m_Config = Config

    ' 构造设备名 \\.\COMx
    Dim devName As String
    If InStr(m_PortName, "\\") = 0 Then
        devName = "\\.\" & m_PortName
    Else
        devName = m_PortName
    End If

    ' 以同步模式打开串口
    m_Handle = CreateFile(devName, _
                          GENERIC_READ Or GENERIC_WRITE, _
                          0, _
                          ByVal 0&, _
                          OPEN_EXISTING, _
                          0, _
                          0)

    If m_Handle = INVALID_HANDLE_VALUE Then
        m_LastError = GetLastError()
        OpenPort = False
        Exit Function
    End If

    ' 设置缓冲区大小
    If SetupComm(m_Handle, m_Config.InBufferSize, m_Config.OutBufferSize) = 0 Then
        ' SetupComm失败不影响后续操作，使用系统默认
    End If

    ' 应用通信配置
    Dim d As DCB
    d = m_Config.BuildDCB
    If SetCommStateVB(m_Handle, d) = 0 Then
        m_LastError = GetLastError()
        CloseHandle m_Handle
        m_Handle = INVALID_HANDLE_VALUE
        ERR.Raise 5, "cSerialPort", "设置通信参数失败: " & GetSystemErrorMessage(m_LastError)
    End If

    ' 应用超时配置
    Dim ct As COMMTIMEOUTS
    ct = m_Config.BuildTimeouts
    If SetCommTimeouts(m_Handle, ct) = 0 Then
        m_LastError = GetLastError()
        CloseHandle m_Handle
        m_Handle = INVALID_HANDLE_VALUE
        ERR.Raise 5, "cSerialPort", "设置超时参数失败: " & GetSystemErrorMessage(m_LastError)
    End If

    ' 清空缓冲区
    PurgeComm m_Handle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR

    m_IsOpen = True
    m_ReceivedCount = 0
    m_SentCount = 0
    OpenPort = True
End Function

'===============================================================================
' 方法 - 关闭串口
'===============================================================================

Public Sub ClosePort()
    If Not m_IsOpen Then Exit Sub
    If m_Monitoring Then StopMonitoring
    ' 关闭前触发缓冲区中剩余的帧数据
    If m_FrameBufLen > 0 Then FlushFrame
    CloseHandle m_Handle
    m_Handle = INVALID_HANDLE_VALUE
    m_IsOpen = False
End Sub

'===============================================================================
' 方法 - 重新应用配置到已打开的端口
'===============================================================================

Public Sub ApplyConfig()
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If

    ' 设置缓冲区
    SetupComm m_Handle, m_Config.InBufferSize, m_Config.OutBufferSize

    ' 设置DCB
    Dim d As DCB
    d = m_Config.BuildDCB
    If SetCommStateVB(m_Handle, d) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "设置通信参数失败: " & GetSystemErrorMessage(m_LastError)
    End If

    ' 设置超时
    Dim ct As COMMTIMEOUTS
    ct = m_Config.BuildTimeouts
    If SetCommTimeouts(m_Handle, ct) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "设置超时参数失败: " & GetSystemErrorMessage(m_LastError)
    End If
End Sub

'===============================================================================
' 方法 - 从端口读取当前配置并更新Config对象
'===============================================================================

Public Sub RefreshConfig()
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If

    Dim d As DCB
    If GetCommStateVB(m_Handle, d) <> 0 Then
        m_Config.ApplyDCB d
    End If

    Dim ct As COMMTIMEOUTS
    If GetCommTimeouts(m_Handle, ct) <> 0 Then
        m_Config.ApplyTimeouts ct
    End If
End Sub

'===============================================================================
' 读取操作
'===============================================================================

' 读取字节数组 - 返回实际读取的字节数
' BytesToRead: 要读取的字节数，-1表示读取缓冲区中所有可用数据
Public Function ReadData(Buffer() As Byte, Optional ByVal BytesToRead As Long = -1) As Long
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"

    ' 确定要读取的字节数
    Dim bytesAvail As Long
    bytesAvail = InBufferCount

    If BytesToRead = -1 Then
        BytesToRead = bytesAvail
    End If

    If BytesToRead <= 0 Then
        ReadData = 0
        Exit Function
    End If

    ' 分配缓冲区
    ReDim Buffer(BytesToRead - 1) As Byte

    Dim bytesRead As Long
    If ReadFile(m_Handle, Buffer(0), BytesToRead, bytesRead, ByVal 0&) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "读取数据失败: " & GetSystemErrorMessage(m_LastError)
    End If

    ' 调整缓冲区大小
    If bytesRead > 0 And bytesRead < BytesToRead Then
        ReDim Preserve Buffer(bytesRead - 1) As Byte
    End If

    m_ReceivedCount = m_ReceivedCount + bytesRead
    ReadData = bytesRead
End Function

' 读取文本
Public Function ReadText(Optional ByVal BytesToRead As Long = -1) As String
    Dim buf() As Byte
    Dim nRead As Long
    nRead = ReadData(buf, BytesToRead)
    If nRead > 0 Then
        ReadText = StrConv(buf, vbUnicode)
    End If
End Function

' 读取输入缓冲区中所有现有数据（非阻塞）
Public Function ReadExisting() As String
    Dim bytesAvail As Long
    bytesAvail = InBufferCount
    If bytesAvail > 0 Then
        ReadExisting = ReadText(bytesAvail)
    End If
End Function

' 读取单个字节 - 返回0-255，超时返回-1
Public Function ReadByte() As Integer
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"

    Dim buf(0) As Byte
    Dim bytesRead As Long
    If ReadFile(m_Handle, buf(0), 1, bytesRead, ByVal 0&) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "读取字节失败: " & GetSystemErrorMessage(m_LastError)
    End If

    If bytesRead > 0 Then
        ReadByte = CInt(buf(0))
        m_ReceivedCount = m_ReceivedCount + 1
    Else
        ReadByte = -1
    End If
End Function

' 读取一行（以指定结束符为准，默认CrLf）
' 注意: 此方法可能阻塞直到收到完整行或超时
Public Function ReadLine(Optional ByVal Terminator As String = vbCrLf) As String
    Dim s As String
    Dim termLen As Long
    termLen = Len(Terminator)
    If termLen = 0 Then termLen = 2: Terminator = vbCrLf

    Dim ch As Integer
    Do
        ch = ReadByte()
        If ch < 0 Then Exit Do  ' 超时
        s = s & Chr$(ch)
        If Len(s) >= termLen Then
            If Right$(s, termLen) = Terminator Then
                ReadLine = Left$(s, Len(s) - termLen)
                Exit Function
            End If
        End If
    Loop

    ' 超时返回已收到的部分
    ReadLine = s
End Function

'===============================================================================
' 写入操作
'===============================================================================

' 写入字节数组 - 返回实际写入的字节数
Public Function WriteData(Buffer() As Byte) As Long
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"

    ' 检查数组是否有效
    If (Not Buffer) = -1 Then
        WriteData = 0
        Exit Function
    End If

    Dim bufLen As Long
    bufLen = UBound(Buffer) - LBound(Buffer) + 1
    If bufLen <= 0 Then
        WriteData = 0
        Exit Function
    End If

    Dim bytesWritten As Long
    If WriteFile(m_Handle, Buffer(LBound(Buffer)), bufLen, bytesWritten, ByVal 0&) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "写入数据失败: " & GetSystemErrorMessage(m_LastError)
    End If

    m_SentCount = m_SentCount + bytesWritten
    WriteData = bytesWritten
End Function

' 写入文本
Public Function WriteText(ByVal Text As String) As Long
    If Len(Text) = 0 Then Exit Function
    Dim buf() As Byte
    buf = StrConv(Text, vbFromUnicode)
    WriteText = WriteData(buf)
End Function

' 写入单个字节
Public Sub WriteByte(ByVal Value As Byte)
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    Dim buf(0) As Byte
    buf(0) = Value
    Dim bytesWritten As Long
    If WriteFile(m_Handle, buf(0), 1, bytesWritten, ByVal 0&) = 0 Then
        m_LastError = GetLastError()
        ERR.Raise 5, "cSerialPort", "写入字节失败: " & GetSystemErrorMessage(m_LastError)
    End If
    m_SentCount = m_SentCount + bytesWritten
End Sub

' 写入一行（自动追加换行符）
Public Function WriteLine(ByVal Text As String, Optional ByVal Terminator As String = vbCrLf) As Long
    WriteLine = WriteText(Text & Terminator)
End Function

' 优先传输字符 - 绕过发送缓冲区直接发送
Public Function TransmitImmediate(ByVal CharCode As Byte) As Boolean
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    TransmitImmediate = (TransmitCommChar(m_Handle, CharCode) <> 0)
End Function

'===============================================================================
' 信号线操作
'===============================================================================

' 设置 DTR 信号
Public Sub SetDTR(ByVal bOn As Boolean)
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    If bOn Then
        EscapeCommFunction m_Handle, modSerialPortAPI.SetDTR
    Else
        EscapeCommFunction m_Handle, modSerialPortAPI.CLRDTR
    End If
End Sub

' 设置 RTS 信号
Public Sub SetRTS(ByVal bOn As Boolean)
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    If bOn Then
        EscapeCommFunction m_Handle, modSerialPortAPI.SetRTS
    Else
        EscapeCommFunction m_Handle, modSerialPortAPI.CLRRTS
    End If
End Sub

' 设置 Break 信号
Public Sub SetBreak(ByVal bOn As Boolean)
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    If bOn Then
        If SetCommBreak(m_Handle) = 0 Then
            m_LastError = GetLastError()
        End If
    Else
        If ClearCommBreak(m_Handle) = 0 Then
            m_LastError = GetLastError()
        End If
    End If
End Sub

' 发送 XOFF 信号（暂停对方发送）
Public Sub SendXOff()
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    EscapeCommFunction m_Handle, modSerialPortAPI.SETXOFF
End Sub

' 发送 XON 信号（恢复对方发送）
Public Sub SendXOn()
    If Not m_IsOpen Then ERR.Raise 5, "cSerialPort", "端口未打开"
    EscapeCommFunction m_Handle, modSerialPortAPI.SETXON
End Sub

'===============================================================================
' 缓冲区操作
'===============================================================================

' 清空接收缓冲区
Public Sub PurgeRx()
    If Not m_IsOpen Then Exit Sub
    PurgeComm m_Handle, PURGE_RXABORT Or PURGE_RXCLEAR
End Sub

' 清空发送缓冲区
Public Sub PurgeTx()
    If Not m_IsOpen Then Exit Sub
    PurgeComm m_Handle, PURGE_TXABORT Or PURGE_TXCLEAR
End Sub

' 清空所有缓冲区（接收+发送，终止所有挂起的I/O）
Public Sub PurgeAll()
    If Not m_IsOpen Then Exit Sub
    PurgeComm m_Handle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR
End Sub

'===============================================================================
' 事件监控 - 异步轮询模式
'===============================================================================

' 开始异步监控 - 使用Timer定时轮询缓冲区
Public Function StartMonitoring(Optional ByVal IntervalMs As Long = 50) As Boolean
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If
    If m_Monitoring Then StopMonitoring

    m_PollInterval = IntervalMs
    If m_PollInterval < 10 Then m_PollInterval = 10

    m_TimerID = SetTimerVB(0, ObjPtr(Me), m_PollInterval, AddressOf SerialTimerProc)
    If m_TimerID <> 0 Then
        RegisterMonitorPort m_TimerID, Me
        m_Monitoring = True
        StartMonitoring = True
    Else
        m_LastError = GetLastError()
        StartMonitoring = False
    End If
End Function

' 停止异步监控
Public Sub StopMonitoring()
    If Not m_Monitoring Then Exit Sub
    If m_TimerID <> 0 Then
        UnregisterMonitorPort m_TimerID
        KillTimerVB 0, m_TimerID
        m_TimerID = 0
    End If
    m_Monitoring = False
    ' 停止监控前触发缓冲区中剩余的帧数据
    If m_FrameBufLen > 0 Then FlushFrame
End Sub

' Poll方法 - 由Timer回调调用，检测缓冲区和状态变化
' 必须声明为Public以便modSerialPortAPI通过Object晚绑定调用
Public Sub Poll()
    If Not m_IsOpen Then Exit Sub

    Dim cs As COMSTAT
    Dim errs As Long
    If ClearCommError(m_Handle, errs, cs) <> 0 Then
        ' 检查通信错误
        If errs <> 0 Then
            RaiseEvent ErrorOccurred(errs, GetCommErrorString(errs))
        End If

        ' 检查数据接收
        If cs.cbInQue > 0 Then
            If m_FrameInterval > 0 Then
                ' 分帧模式: 读取数据到帧缓冲区
                AppendToFrameBuf cs.cbInQue
            Else
                ' 原始模式: 直接触发DataReceived
                RaiseEvent DataReceived(cs.cbInQue)
            End If
        End If

        ' 分帧模式: 检查帧超时
        If m_FrameInterval > 0 And m_FrameBufLen > 0 Then
            If (GetTickCount() - m_LastRecvTime) >= m_FrameInterval Then
                FlushFrame
            End If
        End If
    End If
End Sub

'===============================================================================
' 内部方法 - 帧缓冲区操作 (方案B)
'===============================================================================

' 读取数据追加到帧缓冲区
Private Sub AppendToFrameBuf(ByVal BytesAvail As Long)
    ' 初始化帧缓冲区
    If m_FrameBufCap = 0 Then
        m_FrameBufCap = 4096
        ReDim m_FrameBuf(m_FrameBufCap - 1) As Byte
    End If

    ' 确保容量足够
    If m_FrameBufLen + BytesAvail > m_FrameBufCap Then
        Dim newCap As Long
        newCap = m_FrameBufCap
        Do While newCap < m_FrameBufLen + BytesAvail
            newCap = newCap * 2
        Loop
        ReDim Preserve m_FrameBuf(newCap - 1) As Byte
        m_FrameBufCap = newCap
    End If

    ' 读取数据到帧缓冲区尾部
    Dim bytesRead As Long
    If ReadFile(m_Handle, m_FrameBuf(m_FrameBufLen), BytesAvail, bytesRead, ByVal 0&) <> 0 Then
        If bytesRead > 0 Then
            m_FrameBufLen = m_FrameBufLen + bytesRead
            m_ReceivedCount = m_ReceivedCount + bytesRead
            m_LastRecvTime = GetTickCount()

            ' 防止帧缓冲区溢出: 超过最大帧大小时强制触发
            If m_FrameBufLen >= m_MaxFrameSize Then
                FlushFrame
            End If
        End If
    Else
        m_LastError = GetLastError()
    End If
End Sub

' 触发帧接收事件并清空帧缓冲区
Private Sub FlushFrame()
    If m_FrameBufLen = 0 Then Exit Sub

    ' 复制帧数据到新数组
    Dim frame() As Byte
    ReDim frame(m_FrameBufLen - 1) As Byte
    CopyMemory frame(0), m_FrameBuf(0), m_FrameBufLen

    ' 重置帧缓冲区
    m_FrameBufLen = 0

    ' 触发事件
    RaiseEvent FrameReceived(frame)
End Sub

' 手动触发当前帧缓冲区中的数据(停止监控或关闭端口前调用)
Public Sub FlushFrames()
    FlushFrame
End Sub

'===============================================================================
' 事件监控 - 同步等待模式
'===============================================================================

' 阻塞式等待通信事件
' EventMask: 要等待的事件组合（EV_RXCHAR Or EV_CTS 等）
' 返回: 实际发生的事件掩码，0表示失败
' 注意: 此方法会阻塞调用线程！在VB6 UI线程中使用将冻结界面。
'       建议使用StartMonitoring代替。仅在后台线程或控制台应用中使用。
Public Function WaitForEvent(ByVal EventMask As Long) As Long
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If

    ' 设置事件掩码
    If SetCommMaskVB(m_Handle, EventMask) = 0 Then
        m_LastError = GetLastError()
        WaitForEvent = 0
        Exit Function
    End If

    Dim evtMask As Long
    ' 同步模式等待 - 传入ByVal 0&表示不使用OVERLAPPED
    If WaitCommEvent(m_Handle, evtMask, ByVal 0&) <> 0 Then
        ' 触发相应事件
        If (evtMask And EV_RXCHAR) Then RaiseEvent DataReceived(InBufferCount)
        If (evtMask And EV_TXEMPTY) Then RaiseEvent TxEmpty
        If (evtMask And EV_BREAK) Then RaiseEvent BreakDetected
        If (evtMask And EV_RING) Then RaiseEvent RingDetected
        If (evtMask And (EV_CTS Or EV_DSR Or EV_RLSD)) Then RaiseEvent PinChanged(evtMask)
        If (evtMask And EV_ERR) Then
            Dim cs As COMSTAT
            Dim errs As Long
            ClearCommError m_Handle, errs, cs
            RaiseEvent ErrorOccurred(errs, GetCommErrorString(errs))
        End If

        WaitForEvent = evtMask
    Else
        m_LastError = GetLastError()
        WaitForEvent = 0
    End If
End Function

' 获取当前事件掩码
Public Function GetEventMask() As Long
    If Not m_IsOpen Then Exit Function
    Dim mask As Long
    If GetCommMaskVB(m_Handle, mask) <> 0 Then
        GetEventMask = mask
    End If
End Function

'===============================================================================
' 串口属性查询
'===============================================================================

' 获取串口属性(COMMPROP)
Friend Function GetProperties() As COMMPROP
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If
    Dim cp As COMMPROP
    If GetCommProperties(m_Handle, cp) = 0 Then
        m_LastError = GetLastError()
    End If
    GetProperties = cp
End Function

' 获取当前DCB
Friend Function ReadDCB() As DCB
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If
    Dim d As DCB
    If GetCommStateVB(m_Handle, d) = 0 Then
        m_LastError = GetLastError()
    End If
    ReadDCB = d
End Function

' 获取当前超时设置
Friend Function ReadTimeouts() As COMMTIMEOUTS
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If
    Dim ct As COMMTIMEOUTS
    If GetCommTimeouts(m_Handle, ct) = 0 Then
        m_LastError = GetLastError()
    End If
    ReadTimeouts = ct
End Function

' 获取通信状态
Friend Function ReadCommStat() As COMSTAT
    If Not m_IsOpen Then
        ERR.Raise 5, "cSerialPort", "端口未打开"
    End If
    Dim cs As COMSTAT
    Dim errs As Long
    ClearCommError m_Handle, errs, cs
    ReadCommStat = cs
End Function

' 获取通信错误
Public Function ReadCommErrors() As Long
    If Not m_IsOpen Then Exit Function
    Dim cs As COMSTAT
    Dim errs As Long
    ClearCommError m_Handle, errs, cs
    ReadCommErrors = errs
End Function

' 获取Modem状态字
Public Function ReadModemStatus() As Long
    If Not m_IsOpen Then Exit Function
    Dim stat As Long
    If GetCommModemStatus(m_Handle, stat) <> 0 Then
        ReadModemStatus = stat
    End If
End Function

'===============================================================================
' 统计重置
'===============================================================================

Public Sub ResetCounters()
    m_ReceivedCount = 0
    m_SentCount = 0
End Sub

'===============================================================================
' 辅助 - 诊断信息
'===============================================================================

' 获取当前串口完整状态字符串
Public Function GetStatusString() As String
    Dim s As String
    s = "Port: " & m_PortName & vbCrLf
    s = s & "Open: " & CStr(m_IsOpen) & vbCrLf
    If m_IsOpen Then
        s = s & "BaudRate: " & CStr(m_Config.BaudRate) & vbCrLf
        s = s & "DataBits: " & CStr(m_Config.DataBits) & vbCrLf
        s = s & "Parity: " & CStr(m_Config.Parity) & vbCrLf
        s = s & "StopBits: " & CStr(m_Config.StopBits) & vbCrLf
        s = s & "InBuffer: " & CStr(InBufferCount) & vbCrLf
        s = s & "OutBuffer: " & CStr(OutBufferCount) & vbCrLf
        s = s & "CTS: " & CStr(CtsHolding) & vbCrLf
        s = s & "DSR: " & CStr(DsrHolding) & vbCrLf
        s = s & "RING: " & CStr(RingHolding) & vbCrLf
        s = s & "CD: " & CStr(CdHolding) & vbCrLf
        s = s & "Received: " & CStr(m_ReceivedCount) & vbCrLf
        s = s & "Sent: " & CStr(m_SentCount) & vbCrLf
        s = s & "Monitoring: " & CStr(m_Monitoring) & vbCrLf
        If m_FrameInterval > 0 Then
            s = s & "FrameInterval: " & CStr(m_FrameInterval) & "ms" & vbCrLf
            s = s & "FrameBuffer: " & CStr(m_FrameBufLen) & "/" & CStr(m_FrameBufCap) & vbCrLf
            s = s & "MaxFrameSize: " & CStr(m_MaxFrameSize) & vbCrLf
        End If

        ' 错误状态
        Dim errs As Long
        errs = ReadCommErrors()
        If errs <> 0 Then
            s = s & "Errors: " & GetCommErrorString(errs) & vbCrLf
        End If

        ' COMSTAT状态
        Dim cs As COMSTAT
        cs = ReadCommStat()
        If cs.fBitFields <> 0 Then
            s = s & "Stat: " & GetCommStatString(cs.fBitFields) & vbCrLf
        End If
    End If
    GetStatusString = s
End Function
