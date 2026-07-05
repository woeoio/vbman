Attribute VB_Name = "modSerialPortAPI"
Option Explicit

'===============================================================================
' 模块级变量
'===============================================================================

Private m_MonitorPorts As Collection   ' 定时器回调用的端口对象映射

'===============================================================================
' 模块: modSerialPortAPI
' 说明: Win32 串口通信 API 声明、常量、类型、枚举、辅助函数
' 作者: VBManLib
'===============================================================================

'===============================================================================
' Win32 API 声明 - 文件操作
'===============================================================================

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

'===============================================================================
' Win32 API 声明 - 通信操作
'===============================================================================

Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    lpOverlapped As Any) As Long

Public Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    lpOverlapped As Any) As Long

Public Declare Function TransmitCommChar Lib "kernel32" ( _
    ByVal hCommDev As Long, _
    ByVal cChar As Byte) As Long

Public Declare Function ClearCommError Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpErrors As Long, _
    lpStat As COMSTAT) As Long

Public Declare Function GetCommStateVB Lib "kernel32" Alias "GetCommState" ( _
    ByVal hCommDev As Long, _
    lpDCB As DCB) As Long

Public Declare Function SetCommStateVB Lib "kernel32" Alias "SetCommState" ( _
    ByVal hCommDev As Long, _
    lpDCB As DCB) As Long

Public Declare Function GetCommTimeouts Lib "kernel32" ( _
    ByVal hCommDev As Long, _
    lpCommTimeouts As COMMTIMEOUTS) As Long

Public Declare Function SetCommTimeouts Lib "kernel32" ( _
    ByVal hCommDev As Long, _
    lpCommTimeouts As COMMTIMEOUTS) As Long

Public Declare Function SetupComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwInQueue As Long, _
    ByVal dwOutQueue As Long) As Long

Public Declare Function PurgeComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long

Public Declare Function GetCommModemStatus Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpModemStat As Long) As Long

Public Declare Function EscapeCommFunction Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFunc As Long) As Long

Public Declare Function SetCommBreak Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Public Declare Function ClearCommBreak Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Public Declare Function SetCommMaskVB Lib "kernel32" Alias "SetCommMask" ( _
    ByVal hFile As Long, _
    ByVal dwEventMask As Long) As Long

Public Declare Function GetCommMaskVB Lib "kernel32" Alias "GetCommMask" ( _
    ByVal hFile As Long, _
    lpEventMask As Long) As Long

Public Declare Function WaitCommEvent Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpEventMask As Long, _
    lpOverlapped As Any) As Long

Public Declare Function GetCommProperties Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpCommProp As COMMPROP) As Long

Public Declare Function BuildCommDCBVB Lib "kernel32" Alias "BuildCommDCBA" ( _
    ByVal lpDef As String, _
    lpDCB As DCB) As Long

Public Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" ( _
    ByVal lpDef As String, _
    lpDCB As DCB, _
    lpCommTimeouts As COMMTIMEOUTS) As Long

'===============================================================================
' Win32 API 声明 - 错误处理
'===============================================================================

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function FormatMessageVB Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long

'===============================================================================
' Win32 API 声明 - 同步对象
'===============================================================================

Public Declare Function CreateEventVB Lib "kernel32" Alias "CreateEventA" ( _
    lpEventAttributes As Any, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpName As String) As Long

Public Declare Function ResetEvent Lib "kernel32" ( _
    ByVal hEvent As Long) As Long

Public Declare Function SetEventVB Lib "kernel32" Alias "SetEvent" ( _
    ByVal hEvent As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

'===============================================================================
' Win32 API 声明 - 定时器
'===============================================================================

Public Declare Function SetTimerVB Lib "user32" Alias "SetTimer" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimerVB Lib "user32" Alias "KillTimer" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

'===============================================================================
' 常量 - CreateFile
'===============================================================================

Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_SHARE_READ As Long = &H1&
Public Const FILE_SHARE_WRITE As Long = &H2&
Public Const OPEN_EXISTING As Long = 3&
Public Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Public Const INVALID_HANDLE_VALUE As Long = -1

'===============================================================================
' 常量 - PurgeComm
'===============================================================================

Public Const PURGE_TXABORT As Long = &H1&
Public Const PURGE_RXABORT As Long = &H2&
Public Const PURGE_TXCLEAR As Long = &H4&
Public Const PURGE_RXCLEAR As Long = &H8&

'===============================================================================
' 常量 - EscapeCommFunction
'===============================================================================

Public Const SETXOFF As Long = 1&
Public Const SETXON As Long = 2&
Public Const SETRTS As Long = 3&
Public Const CLRRTS As Long = 4&
Public Const SETDTR As Long = 5&
Public Const CLRDTR As Long = 6&
Public Const SETBREAK As Long = 8&
Public Const CLRBREAK As Long = 9&

'===============================================================================
' 常量 - Modem 状态 (GetCommModemStatus)
'===============================================================================

Public Const MS_CTS_ON As Long = &H10&
Public Const MS_DSR_ON As Long = &H20&
Public Const MS_RING_ON As Long = &H40&
Public Const MS_RLSD_ON As Long = &H80&

'===============================================================================
' 常量 - 通信事件 (SetCommMask / WaitCommEvent)
'===============================================================================

Public Const EV_RXCHAR As Long = &H1&          ' 字符已接收
Public Const EV_RXFLAG As Long = &H2&          ' 事件字符已接收
Public Const EV_TXEMPTY As Long = &H4&         ' 发送缓冲区已空
Public Const EV_CTS As Long = &H8&             ' CTS信号变化
Public Const EV_DSR As Long = &H10&            ' DSR信号变化
Public Const EV_RLSD As Long = &H20&           ' RLSD(CD)信号变化
Public Const EV_BREAK As Long = &H40&          ' 检测到中断
Public Const EV_ERR As Long = &H80&            ' 线路状态错误
Public Const EV_RING As Long = &H100&          ' 振铃检测
Public Const EV_PERR As Long = &H200&          ' 打印机错误
Public Const EV_RX80FULL As Long = &H400&      ' 接收缓冲区80%满
Public Const EV_EVENT1 As Long = &H800&        ' 设备事件1
Public Const EV_EVENT2 As Long = &H1000&       ' 设备事件2

' 所有事件的掩码
Public Const EV_ALL As Long = EV_RXCHAR Or EV_RXFLAG Or EV_TXEMPTY Or _
    EV_CTS Or EV_DSR Or EV_RLSD Or EV_BREAK Or EV_ERR Or _
    EV_RING Or EV_PERR Or EV_RX80FULL Or EV_EVENT1 Or EV_EVENT2

'===============================================================================
' 常量 - 通信错误 (ClearCommError)
'===============================================================================

Public Const CE_RXOVER As Long = &H1&          ' 接收缓冲区溢出
Public Const CE_OVERRUN As Long = &H2&         ' 字符覆盖（缓冲区满）
Public Const CE_RXPARITY As Long = &H4&        ' 校验错误
Public Const CE_FRAME As Long = &H8&           ' 帧错误
Public Const CE_BREAK As Long = &H10&          ' 中断检测
Public Const CE_TXFULL As Long = &H100&        ' 发送缓冲区满
Public Const CE_PTO As Long = &H200&           ' 并口超时
Public Const CE_IOE As Long = &H400&           ' I/O错误
Public Const CE_DNS As Long = &H800&           ' 设备未选择
Public Const CE_OOP As Long = &H1000&          ' 缺纸
Public Const CE_MODE As Long = &H8000&         ' 模式错误

'===============================================================================
' 常量 - FormatMessage
'===============================================================================

Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Public Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&

'===============================================================================
' 常量 - WaitForSingleObject
'===============================================================================

Public Const WAIT_OBJECT_0 As Long = 0
Public Const WAIT_TIMEOUT As Long = &H102&
Public Const WAIT_FAILED As Long = -1
Public Const INFINITE As Long = -1

'===============================================================================
' 常量 - IO Pending
'===============================================================================

Public Const ERROR_IO_PENDING As Long = 997

'===============================================================================
' 常量 - DCB fBitFields 位偏移
'===============================================================================

Public Const DCB_fBinary As Long = 0              ' 必须为1
Public Const DCB_fParity As Long = 1              ' 启用校验检查
Public Const DCB_fOutxCtsFlow As Long = 2         ' CTS输出流控制
Public Const DCB_fOutxDsrFlow As Long = 3         ' DSR输出流控制
Public Const DCB_fDtrControl As Long = 4          ' DTR控制(2位)
Public Const DCB_fDsrSensitivity As Long = 6      ' DSR敏感
Public Const DCB_fTXContinueOnXoff As Long = 7    ' XOFF后继续发送
Public Const DCB_fOutX As Long = 8                ' XON/XOFF输出流控制
Public Const DCB_fInX As Long = 9                 ' XON/XOFF输入流控制
Public Const DCB_fErrorChar As Long = 10          ' 替换校验错误字符
Public Const DCB_fNull As Long = 11               ' 丢弃NULL字节
Public Const DCB_fRtsControl As Long = 12         ' RTS控制(2位)
Public Const DCB_fAbortOnError As Long = 14       ' 错误时终止读写

'===============================================================================
' 常量 - MaxBaud (COMMPROP)
'===============================================================================

Public Const BAUD_075 As Long = &H1&
Public Const BAUD_110 As Long = &H2&
Public Const BAUD_134_5 As Long = &H4&
Public Const BAUD_150 As Long = &H8&
Public Const BAUD_300 As Long = &H10&
Public Const BAUD_600 As Long = &H20&
Public Const BAUD_1200 As Long = &H40&
Public Const BAUD_1800 As Long = &H80&
Public Const BAUD_2400 As Long = &H100&
Public Const BAUD_4800 As Long = &H200&
Public Const BAUD_7200 As Long = &H400&
Public Const BAUD_9600 As Long = &H800&
Public Const BAUD_14400 As Long = &H1000&
Public Const BAUD_19200 As Long = &H2000&
Public Const BAUD_38400 As Long = &H4000&
Public Const BAUD_56K As Long = &H8000&
Public Const BAUD_57600 As Long = &H40000
Public Const BAUD_115200 As Long = &H20000
Public Const BAUD_128K As Long = &H10000
Public Const BAUD_USER As Long = &H10000000

'===============================================================================
' 枚举 - 波特率
'===============================================================================

Public Enum eBaudRate
    br110 = 110
    br300 = 300
    br600 = 600
    br1200 = 1200
    br2400 = 2400
    br4800 = 4800
    br9600 = 9600
    br14400 = 14400
    br19200 = 19200
    br38400 = 38400
    br56000 = 56000
    br57600 = 57600
    br115200 = 115200
    br128000 = 128000
    br256000 = 256000
End Enum

'===============================================================================
' 枚举 - 校验位
'===============================================================================

Public Enum eParity
    ptNone = 0       ' NOPARITY
    ptOdd = 1        ' ODDPARITY
    ptEven = 2       ' EVENPARITY
    ptMark = 3       ' MARKPARITY
    ptSpace = 4      ' SPACEPARITY
End Enum

'===============================================================================
' 枚举 - 停止位
'===============================================================================

Public Enum eStopBits
    sb1 = 0          ' ONESTOPBIT
    sb1_5 = 1        ' ONE5STOPBITS
    sb2 = 2          ' TWOSTOPBITS
End Enum

'===============================================================================
' 枚举 - DTR 控制模式
'===============================================================================

Public Enum eDtrControl
    dcDisable = 0    ' DTR_CONTROL_DISABLE
    dcEnable = 1     ' DTR_CONTROL_ENABLE
    dcHandshake = 2  ' DTR_CONTROL_HANDSHAKE
End Enum

'===============================================================================
' 枚举 - RTS 控制模式
'===============================================================================

Public Enum eRtsControl
    rcDisable = 0    ' RTS_CONTROL_DISABLE
    rcEnable = 1     ' RTS_CONTROL_ENABLE
    rcHandshake = 2  ' RTS_CONTROL_HANDSHAKE
    rcToggle = 3     ' RTS_CONTROL_TOGGLE
End Enum

'===============================================================================
' 枚举 - 流控制模式
'===============================================================================

Public Enum eFlowControl
    fcNone = 0                ' 无流控制
    fcXonXoff = 1             ' 软件流控制(XON/XOFF)
    fcRtsCts = 2              ' 硬件流控制(RTS/CTS)
    fcDtrDsr = 3              ' 硬件流控制(DTR/DSR)
    fcRtsCtsAndXonXoff = 4    ' 混合流控制
End Enum

'===============================================================================
' 类型定义 - DCB
'===============================================================================

Public Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long       ' 位域，使用辅助函数读写
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer
End Type

'===============================================================================
' 类型定义 - COMMTIMEOUTS
'===============================================================================

Public Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'===============================================================================
' 类型定义 - COMSTAT
'===============================================================================

Public Type COMSTAT
    fBitFields As Long       ' 位域: fCtsHold/fDsrHold/fRlsdHold/fXoffHold/...
    cbInQue As Long          ' 接收缓冲区字节数
    cbOutQue As Long         ' 发送缓冲区字节数
End Type

'===============================================================================
' 类型定义 - COMMPROP
'===============================================================================

Public Type COMMPROP
    wPacketLength As Integer
    wPacketVersion As Integer
    dwServiceMask As Long
    dwReserved1 As Long
    dwMaxTxQueue As Long
    dwMaxRxQueue As Long
    dwMaxBaud As Long
    dwProvSubType As Long
    dwProvCapabilities As Long
    dwSettableParams As Long
    dwSettableBaud As Long
    wSettableData As Integer
    wSettableStopParity As Integer
    dwCurrentTxQueue As Long
    dwCurrentRxQueue As Long
    dwProvSpec1 As Long
    dwProvSpec2 As Long
    wcProvChar1 As Integer
    wcProvChar2 As Integer
End Type

'===============================================================================
' 类型定义 - OVERLAPPED
'===============================================================================

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

'===============================================================================
' 位操作辅助函数 - DCB fBitFields 读写
'===============================================================================

' 获取单比特标志
Public Function GetBit(ByVal Value As Long, ByVal Bit As Long) As Boolean
    If Bit < 0 Or Bit > 31 Then Exit Function
    GetBit = (Value And (2 ^ Bit)) <> 0
End Function

' 设置单比特标志
Public Sub SetBit(ByRef Value As Long, ByVal Bit As Long, ByVal bSet As Boolean)
    If Bit < 0 Or Bit > 31 Then Exit Sub
    If bSet Then
        Value = Value Or CLng(2 ^ Bit)
    Else
        Value = Value And Not CLng(2 ^ Bit)
    End If
End Sub

' 获取多比特字段（从StartBit开始，BitCount位）
Public Function GetBits(ByVal Value As Long, ByVal StartBit As Long, ByVal BitCount As Long) As Long
    If StartBit < 0 Or BitCount < 1 Or (StartBit + BitCount) > 32 Then Exit Function
    Dim mask As Long
    mask = (2 ^ BitCount) - 1
    GetBits = (Value \ CLng(2 ^ StartBit)) And mask
End Function

' 设置多比特字段
Public Sub SetBits(ByRef Value As Long, ByVal StartBit As Long, ByVal BitCount As Long, ByVal NewValue As Long)
    If StartBit < 0 Or BitCount < 1 Or (StartBit + BitCount) > 32 Then Exit Sub
    Dim mask As Long
    mask = (2 ^ BitCount) - 1
    Value = (Value And Not (mask * CLng(2 ^ StartBit))) Or ((NewValue And mask) * CLng(2 ^ StartBit))
End Sub

'===============================================================================
' COMSTAT 位域解析
'===============================================================================

Public Function ComstatCtsHold(ByVal fBitFields As Long) As Boolean
    ComstatCtsHold = GetBit(fBitFields, 0)
End Function

Public Function ComstatDsrHold(ByVal fBitFields As Long) As Boolean
    ComstatDsrHold = GetBit(fBitFields, 1)
End Function

Public Function ComstatRlsdHold(ByVal fBitFields As Long) As Boolean
    ComstatRlsdHold = GetBit(fBitFields, 2)
End Function

Public Function ComstatXoffHold(ByVal fBitFields As Long) As Boolean
    ComstatXoffHold = GetBit(fBitFields, 3)
End Function

Public Function ComstatXoffSent(ByVal fBitFields As Long) As Boolean
    ComstatXoffSent = GetBit(fBitFields, 4)
End Function

Public Function ComstatEof(ByVal fBitFields As Long) As Boolean
    ComstatEof = GetBit(fBitFields, 5)
End Function

Public Function ComstatTxim(ByVal fBitFields As Long) As Boolean
    ComstatTxim = GetBit(fBitFields, 6)
End Function

'===============================================================================
' 错误信息辅助函数
'===============================================================================

' 获取系统错误描述
Public Function GetSystemErrorMessage(ByVal ErrorCode As Long) As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(512, 0)
    lRet = FormatMessageVB(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                           ByVal 0&, ErrorCode, 0, sBuffer, 512, ByVal 0&)
    If lRet > 0 Then
        GetSystemErrorMessage = TrimTrailing(sBuffer)
    Else
        GetSystemErrorMessage = "Unknown error: " & CStr(ErrorCode)
    End If
End Function

' 获取串口通信错误描述
Public Function GetCommErrorString(ByVal Errors As Long) As String
    Dim s As String
    If (Errors And CE_RXOVER) Then s = s & "接收缓冲区溢出; "
    If (Errors And CE_OVERRUN) Then s = s & "字符覆盖; "
    If (Errors And CE_RXPARITY) Then s = s & "校验错误; "
    If (Errors And CE_FRAME) Then s = s & "帧错误; "
    If (Errors And CE_BREAK) Then s = s & "中断检测; "
    If (Errors And CE_TXFULL) Then s = s & "发送缓冲区满; "
    If (Errors And CE_PTO) Then s = s & "并口超时; "
    If (Errors And CE_IOE) Then s = s & "I/O错误; "
    If (Errors And CE_DNS) Then s = s & "设备未选择; "
    If (Errors And CE_OOP) Then s = s & "缺纸; "
    If (Errors And CE_MODE) Then s = s & "模式错误; "
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    GetCommErrorString = s
End Function

' 获取通信事件描述
Public Function GetCommEventString(ByVal Events As Long) As String
    Dim s As String
    If (Events And EV_RXCHAR) Then s = s & "字符接收; "
    If (Events And EV_RXFLAG) Then s = s & "事件字符接收; "
    If (Events And EV_TXEMPTY) Then s = s & "发送完成; "
    If (Events And EV_BREAK) Then s = s & "中断检测; "
    If (Events And EV_CTS) Then s = s & "CTS变化; "
    If (Events And EV_DSR) Then s = s & "DSR变化; "
    If (Events And EV_ERR) Then s = s & "线路错误; "
    If (Events And EV_RING) Then s = s & "振铃; "
    If (Events And EV_RLSD) Then s = s & "RLSD变化; "
    If (Events And EV_RX80FULL) Then s = s & "接收80%满; "
    If (Events And EV_PERR) Then s = s & "打印机错误; "
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    GetCommEventString = s
End Function

' 获取COMSTAT状态描述
Public Function GetCommStatString(ByVal fBitFields As Long) As String
    Dim s As String
    If ComstatCtsHold(fBitFields) Then s = s & "等待CTS; "
    If ComstatDsrHold(fBitFields) Then s = s & "等待DSR; "
    If ComstatRlsdHold(fBitFields) Then s = s & "等待RLSD; "
    If ComstatXoffHold(fBitFields) Then s = s & "XOFF等待; "
    If ComstatXoffSent(fBitFields) Then s = s & "XOFF已发; "
    If ComstatEof(fBitFields) Then s = s & "EOF; "
    If ComstatTxim(fBitFields) Then s = s & "字符待发; "
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)
    GetCommStatString = s
End Function

'===============================================================================
' 串口枚举
'===============================================================================

' 枚举系统中可用的串口号
Public Function EnumSerialPorts() As Collection
    Dim ports As New Collection
    Dim i As Long
    Dim hPort As Long
    For i = 1 To 256
        hPort = CreateFile("\\.\COM" & CStr(i), GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, 0, 0)
        If hPort <> INVALID_HANDLE_VALUE Then
            CloseHandle hPort
            On Error Resume Next
            ports.Add "COM" & CStr(i), "COM" & CStr(i)
            On Error GoTo 0
        End If
    Next i
    Set EnumSerialPorts = ports
End Function

'===============================================================================
' 定时器回调 - 串口监控
'===============================================================================

' 注册监控端口
Public Sub RegisterMonitorPort(ByVal TimerID As Long, ByVal Port As Object)
    If m_MonitorPorts Is Nothing Then Set m_MonitorPorts = New Collection
    On Error Resume Next
    m_MonitorPorts.Remove "K" & CStr(TimerID)
    On Error GoTo 0
    m_MonitorPorts.Add Port, "K" & CStr(TimerID)
End Sub

' 注销监控端口
Public Sub UnregisterMonitorPort(ByVal TimerID As Long)
    If Not m_MonitorPorts Is Nothing Then
        On Error Resume Next
        m_MonitorPorts.Remove "K" & CStr(TimerID)
        On Error GoTo 0
    End If
End Sub

' Timer回调过程 - 由SetTimer调用
Public Sub SerialTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    If Not m_MonitorPorts Is Nothing Then
        Dim obj As Object
        Set obj = m_MonitorPorts("K" & CStr(idEvent))
        If Not obj Is Nothing Then
            obj.Poll
        End If
    End If
End Sub

'===============================================================================
' 内部辅助
'===============================================================================

Private Function TrimTrailing(ByVal s As String) As String
    Dim i As Long
    i = Len(s)
    Do While i > 0
        Select Case Mid$(s, i, 1)
            Case vbCr, vbLf, " ", vbTab
                i = i - 1
            Case Else
                Exit Do
        End Select
    Loop
    TrimTrailing = Left$(s, i)
End Function
