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

'================================================================================
' 串口类封装
' 使用 Windows 系统原生 API 实现
' 支持异步操作和事件通知
'================================================================================

'==================== Windows API 常量声明 ====================
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const INVALID_HANDLE_VALUE = -1

'==================== 串口控制常量 ====================
Private Const SETXOFF = 1
Private Const SETXON = 2
Private Const SETRTS = 3
Private Const CLRRTS = 4
Private Const SETDTR = 5
Private Const CLRDTR = 6
Private Const RESETDEV = 7
Private Const SETBREAK = 8
Private Const CLRBREAK = 9

'==================== 错误码常量 ====================
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_IO_PENDING = 997&
Private Const ERROR_OPERATION_ABORTED = 995&

'==================== DCB 结构体 ====================
Private Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long           ' 详见 DCB 位标志
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

'==================== COMMTIMEOUTS 结构体 ====================
Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'==================== OVERLAPPED 结构体 ====================
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

'==================== COMMCONFIG 结构体 ====================
Private Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type

'==================== DCB 位标志常量 ====================
Private Const fBinary = &H1
Private Const fParity = &H2
Private Const fOutxCtsFlow = &H4
Private Const fOutxDsrFlow = &H8
Private Const fDtrControl = &H30
Private Const fDsrSensitivity = &H40
Private Const fTXContinueOnXoff = &H80
Private Const fOutX = &H100
Private Const fInX = &H200
Private Const fErrorChar = &H400
Private Const fNull = &H800
Private Const fRtsControl = &H3000
Private Const fAbortOnError = &H4000
Private Const fDummy2 = &H80000000

'==================== DTR 控制常量 ====================
Private Const DTR_CONTROL_DISABLE = &H0
Private Const DTR_CONTROL_ENABLE = &H10
Private Const DTR_CONTROL_HANDSHAKE = &H20

'==================== RTS 控制常量 ====================
Private Const RTS_CONTROL_DISABLE = &H0
Private Const RTS_CONTROL_ENABLE = &H1000
Private Const RTS_CONTROL_HANDSHAKE = &H2000
Private Const RTS_CONTROL_TOGGLE = &H3000

'==================== Wait 函数常量 ====================
Private Const WAIT_OBJECT_0 = 0&
Private Const WAIT_TIMEOUT = &H102&
Private Const INFINITE = &HFFFF

'==================== 波特率常量 ====================
Private Const CBR_110 = 110
Private Const CBR_300 = 300
Private Const CBR_600 = 600
Private Const CBR_1200 = 1200
Private Const CBR_2400 = 2400
Private Const CBR_4800 = 4800
Private Const CBR_9600 = 9600
Private Const CBR_14400 = 14400
Private Const CBR_19200 = 19200
Private Const CBR_38400 = 38400
Private Const CBR_56000 = 56000
Private Const CBR_57600 = 57600
Private Const CBR_115200 = 115200
Private Const CBR_128000 = 128000
Private Const CBR_256000 = 256000

'==================== 停止位常量 ====================
Private Const ONESTOPBIT = 0
Private Const ONE5STOPBITS = 1
Private Const TWOSTOPBITS = 2

'==================== 校验位常量 ====================
Private Const NOPARITY = 0
Private Const ODDPARITY = 1
Private Const EVENPARITY = 2
Private Const MARKPARITY = 3
Private Const SPACEPARITY = 4

'==================== 事件掩码常量 ====================
Private Const EV_RXCHAR = &H1          ' 收到任何字符
Private Const EV_RXFLAG = &H2          ' 收到事件字符
Private Const EV_TXEMPTY = &H4         ' 发送缓冲区空
Private Const EV_CTS = &H8             ' CTS 信号改变
Private Const EV_DSR = &H10            ' DSR 信号改变
Private Const EV_RLSD = &H20           ' RLSD (DCD) 信号改变
Private Const EV_BREAK = &H40          ' 检测到中断
Private Const EV_ERR = &H80            ' 发生线路状态错误
Private Const EV_RING = &H100          ' 检测到振铃
Private Const EV_PERR = &H200          ' 打印机错误
Private Const EV_RX80FULL = &H400      ' 接收缓冲区 80% 满
Private Const EV_EVENT1 = &H800
Private Const EV_EVENT2 = &H1000

'==================== 错误标志常量 ====================
Private Const CE_RXOVER = &H1          ' 接收溢出
Private Const CE_OVERRUN = &H2         ' 超时错误
Private Const CE_RXPARITY = &H4        ' 校验错误
Private Const CE_FRAME = &H8           ' 帧错误
Private Const CE_BREAK = &H10          ' 中断错误
Private Const CE_TXFULL = &H100        ' 发送缓冲区满
Private Const CE_MODE = &H8000

'==================== 串口事件类型枚举 ====================
Public Enum SerialEventType
    evtDataReceived = 1
    evtCTSChanged = 2
    evtDSRChanged = 3
    evtRing = 4
    evtRLSDChanged = 5
    evtBreak = 6
    evtError = 7
    evtTxEmpty = 8
End Enum

'==================== 波特率枚举 ====================
Public Enum BaudRateEnum
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
    br230400 = 230400
    br256000 = 256000
    br460800 = 460800
    br921600 = 921600
End Enum

'==================== 停止位枚举 ====================
Public Enum StopBitsEnum
    sb1 = 1
    sb1_5 = 2
    sb2 = 3
End Enum

'==================== 校验位枚举 ====================
Public Enum ParityEnum
    peNone = 0
    peOdd = 1
    peEven = 2
    peMark = 3
    peSpace = 4
End Enum

'==================== 数据位数枚举 ====================
Public Enum DataBitsEnum
    db4 = 4
    db5 = 5
    db6 = 6
    db7 = 7
    db8 = 8
End Enum

'==================== 流控制枚举 ====================
Public Enum FlowControlEnum
    fcNone = 0
    fcXOnXOff = 1
    fcRTSCTS = 2
    fcRTSCTSXOnXOff = 3
End Enum

'==================== API 函数声明 ====================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByRef lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByRef lpOverlapped As Any) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByRef lpOverlapped As Any) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Private Declare Function PurgeComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long

Private Const PURGE_TXABORT = &H1
Private Const PURGE_RXABORT = &H2
Private Const PURGE_TXCLEAR = &H4
Private Const PURGE_RXCLEAR = &H8

Private Declare Function SetupComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwInQueue As Long, _
    ByVal dwOutQueue As Long) As Long

Private Declare Function SetCommTimeouts Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpCommTimeouts As COMMTIMEOUTS) As Long

Private Declare Function GetCommTimeouts Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpCommTimeouts As COMMTIMEOUTS) As Long

Private Declare Function SetCommState Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpDCB As DCB) As Long

Private Declare Function GetCommState Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpDCB As DCB) As Long

Private Declare Function EscapeCommFunction Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFunc As Long) As Long

Private Declare Function SetCommMask Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwEvtMask As Long) As Long

Private Declare Function GetCommMask Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpEvtMask As Long) As Long

Private Declare Function WaitCommEvent Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpEvtMask As Long, _
    ByRef lpOverlapped As OVERLAPPED) As Long

Private Declare Function GetCommModemStatus Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpModemStat As Long) As Long

Private Const MS_CTS_ON = &H10
Private Const MS_DSR_ON = &H20
Private Const MS_RING_ON = &H40
Private Const MS_RLSD_ON = &H80

Private Declare Function ClearCommError Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpErrors As Long, _
    ByRef lpStat As Any) As Long

Private Declare Function CancelIo Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( _
    ByRef lpEventAttributes As Any, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpName As String) As Long

Private Declare Function SetEvent Lib "kernel32" ( _
    ByVal hEvent As Long) As Long

Private Declare Function ResetEvent Lib "kernel32" ( _
    ByVal hEvent As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

'==================== 串口参数结构 ====================
Public Type SerialPortConfig
    PortName As String          ' 串口名称 (如 "COM1")
    BaudRate As Long            ' 波特率
    DataBits As Integer         ' 数据位
    StopBits As Integer         ' 停止位
    Parity As Integer           ' 校验位
    FlowControl As Integer      ' 流控制
End Type

'==================== 私有变量 ====================
Private m_hCommHandle As Long           ' 串口句柄
Private m_strPortName As String         ' 串口名称
Private m_lBaudRate As Long             ' 波特率
Private m_iDataBits As Integer          ' 数据位
Private m_iStopBits As Integer           ' 停止位
Private m_iParity As Integer            ' 校验位
Private m_iFlowControl As Integer       ' 流控制
Private m_bIsOpen As Boolean            ' 是否打开
Private m_lEventMask As Long            ' 事件掩码
Private m_Overlapped As OVERLAPPED      ' 异步重叠结构
Private m_bListening As Boolean         ' 是否正在监听
Private m_hStopEvent As Long            ' 停止监听事件
Private m_bUseAsyncIO As Boolean        ' 是否使用异步IO
Private m_lInQueueSize As Long          ' 输入缓冲区大小
Private m_lOutQueueSize As Long         ' 输出缓冲区大小

'==================== 事件定义 ====================
Public Event OnDataReceived(ByVal Data As String, ByVal Length As Long)
Public Event OnDataReceivedBytes(ByRef Data() As Byte, ByVal Length As Long)
Public Event OnError(ByVal ErrorCode As Long, ByVal ErrorMessage As String)
Public Event OnCTSChanged(ByVal State As Boolean)
Public Event OnDSRChanged(ByVal State As Boolean)
Public Event OnRing(ByVal State As Boolean)
Public Event OnRLSDChanged(ByVal State As Boolean)
Public Event OnBreak()
Public Event OnTxEmpty()

'================================================================================
' 类初始化
'================================================================================
Private Sub Class_Initialize()
    m_hCommHandle = INVALID_HANDLE_VALUE
    m_strPortName = ""
    m_lBaudRate = 9600
    m_iDataBits = 8
    m_iStopBits = sb1
    m_iParity = peNone
    m_iFlowControl = fcNone
    m_bIsOpen = False
    m_bListening = False
    m_bUseAsyncIO = True
    m_lInQueueSize = 4096
    m_lOutQueueSize = 4096
    m_hStopEvent = 0
End Sub

'================================================================================
' 类终止
'================================================================================
Private Sub Class_Terminate()
    Close
End Sub

'================================================================================
' 打开串口
'================================================================================
Public Function Open(ByVal PortName As String, _
                     Optional ByVal BaudRate As Long = 9600, _
                     Optional ByVal DataBits As Integer = 8, _
                     Optional ByVal Parity As ParityEnum = peNone, _
                     Optional ByVal StopBits As StopBitsEnum = sb1, _
                     Optional ByVal FlowControl As FlowControlEnum = fcNone) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 如果已打开，先关闭
    If m_bIsOpen Then
        Close
    End If
    
    ' 保存参数
    m_strPortName = PortName
    m_lBaudRate = BaudRate
    m_iDataBits = DataBits
    m_iParity = Parity
    m_iStopBits = StopBits
    m_iFlowControl = FlowControl
    
    ' 打开串口
    Dim dwFlagsAndAttributes As Long
    If m_bUseAsyncIO Then
        dwFlagsAndAttributes = FILE_FLAG_OVERLAPPED
    Else
        dwFlagsAndAttributes = 0
    End If
    
    m_hCommHandle = CreateFile( _
        "\\.\" & PortName, _
        GENERIC_READ Or GENERIC_WRITE, _
        0, _
        ByVal 0&, _
        OPEN_EXISTING, _
        dwFlagsAndAttributes, _
        0&)
    
    If m_hCommHandle = INVALID_HANDLE_VALUE Then
        Dim lErr As Long
        lErr = GetLastError()
        RaiseEvent OnError(lErr, "无法打开串口 " & PortName)
        Open = False
        Exit Function
    End If
    
    ' 配置串口缓冲区
    If SetupComm(m_hCommHandle, m_lInQueueSize, m_lOutQueueSize) = 0 Then
        CloseHandle m_hCommHandle
        m_hCommHandle = INVALID_HANDLE_VALUE
        RaiseEvent OnError(GetLastError(), "设置串口缓冲区失败")
        Open = False
        Exit Function
    End If
    
    ' 清空缓冲区
    PurgeComm m_hCommHandle, PURGE_TXCLEAR Or PURGE_RXCLEAR
    
    ' 配置串口参数
    If Not ConfigurePort(BaudRate, DataBits, Parity, StopBits, FlowControl) Then
        CloseHandle m_hCommHandle
        m_hCommHandle = INVALID_HANDLE_VALUE
        Open = False
        Exit Function
    End If
    
    ' 配置超时
    If Not SetCommTimeouts_() Then
        CloseHandle m_hCommHandle
        m_hCommHandle = INVALID_HANDLE_VALUE
        Open = False
        Exit Function
    End If
    
    ' 如果使用异步IO，创建重叠事件
    If m_bUseAsyncIO Then
        ' 创建停止事件
        m_hStopEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        If m_hStopEvent = 0 Then
            CloseHandle m_hCommHandle
            m_hCommHandle = INVALID_HANDLE_VALUE
            RaiseEvent OnError(GetLastError(), "创建停止事件失败")
            Open = False
            Exit Function
        End If
        
        ' 配置重叠事件
        m_Overlapped.hEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        If m_Overlapped.hEvent = 0 Then
            CloseHandle m_hStopEvent
            CloseHandle m_hCommHandle
            m_hCommHandle = INVALID_HANDLE_VALUE
            RaiseEvent OnError(GetLastError(), "创建重叠事件失败")
            Open = False
            Exit Function
        End If
        
        ' 设置事件掩码
        m_lEventMask = EV_RXCHAR Or EV_ERR Or EV_CTS Or EV_DSR Or EV_RING Or EV_RLSD Or EV_BREAK
        If SetCommMask(m_hCommHandle, m_lEventMask) = 0 Then
            CloseHandle m_Overlapped.hEvent
            CloseHandle m_hStopEvent
            CloseHandle m_hCommHandle
            m_hCommHandle = INVALID_HANDLE_VALUE
            RaiseEvent OnError(GetLastError(), "设置事件掩码失败")
            Open = False
            Exit Function
        End If
    End If
    
    m_bIsOpen = True
    Open = True
    Exit Function
    
ErrorHandler:
    If m_hCommHandle <> INVALID_HANDLE_VALUE Then CloseHandle m_hCommHandle
    m_hCommHandle = INVALID_HANDLE_VALUE
    m_bIsOpen = False
    Open = False
    RaiseEvent OnError(Err.Number, Err.Description)
End Function

'================================================================================
' 关闭串口
'================================================================================
Public Sub Close()
    If Not m_bIsOpen Then Exit Sub
    
    ' 停止监听
    StopListening
    
    ' 关闭事件句柄
    If m_Overlapped.hEvent <> 0 Then
        CloseHandle m_Overlapped.hEvent
        m_Overlapped.hEvent = 0
    End If
    
    If m_hStopEvent <> 0 Then
        SetEvent m_hStopEvent
        CloseHandle m_hStopEvent
        m_hStopEvent = 0
    End If
    
    ' 关闭串口
    If m_hCommHandle <> INVALID_HANDLE_VALUE Then
        CloseHandle m_hCommHandle
        m_hCommHandle = INVALID_HANDLE_VALUE
    End If
    
    m_bIsOpen = False
End Sub

'================================================================================
' 发送数据
'================================================================================
Public Function WriteString(ByVal Data As String) As Long
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        WriteString = 0
        Exit Function
    End If
    
    Dim lWritten As Long
    Dim lRet As Long
    
    If m_bUseAsyncIO Then
        ' 异步写入
        Dim ovWrite As OVERLAPPED
        ovWrite.hEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        
        lRet = WriteFile(m_hCommHandle, ByVal Data, Len(Data), lWritten, ovWrite)
        
        If lRet = 0 Then
            Dim lErr As Long
            lErr = GetLastError()
            
            If lErr = ERROR_IO_PENDING Then
                ' 等待写入完成
                Dim lWaitRet As Long
                lWaitRet = WaitForSingleObject(ovWrite.hEvent, INFINITE)
                
                If lWaitRet = WAIT_OBJECT_0 Then
                    lRet = 1
                    ' 获取实际写入的字节数
                    If GetOverlappedResult(m_hCommHandle, ovWrite, lWritten, 1) = 0 Then
                        lWritten = 0
                    End If
                Else
                    CancelIo m_hCommHandle
                    lWritten = 0
                End If
            Else
                lWritten = 0
                RaiseEvent OnError(lErr, "写入串口失败")
            End If
        End If
        
        CloseHandle ovWrite.hEvent
    Else
        ' 同步写入
        lRet = WriteFile(m_hCommHandle, ByVal Data, Len(Data), lWritten, ByVal 0&)
        If lRet = 0 Then
            lWritten = 0
            RaiseEvent OnError(GetLastError(), "写入串口失败")
        End If
    End If
    
    WriteString = lWritten
End Function

'================================================================================
' 发送字节数组
'================================================================================
Public Function WriteBytes(ByRef Data() As Byte, Optional ByVal Length As Long = -1) As Long
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        WriteBytes = 0
        Exit Function
    End If
    
    If Length < 0 Then
        Length = UBound(Data) - LBound(Data) + 1
    End If
    
    Dim lWritten As Long
    Dim lRet As Long
    
    If m_bUseAsyncIO Then
        ' 异步写入
        Dim ovWrite As OVERLAPPED
        ovWrite.hEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        
        lRet = WriteFile(m_hCommHandle, Data(LBound(Data)), Length, lWritten, ovWrite)
        
        If lRet = 0 Then
            Dim lErr As Long
            lErr = GetLastError()
            
            If lErr = ERROR_IO_PENDING Then
                ' 等待写入完成
                Dim lWaitRet As Long
                lWaitRet = WaitForSingleObject(ovWrite.hEvent, INFINITE)
                
                If lWaitRet = WAIT_OBJECT_0 Then
                    lRet = 1
                    ' 获取实际写入的字节数
                    If GetOverlappedResult(m_hCommHandle, ovWrite, lWritten, 1) = 0 Then
                        lWritten = 0
                    End If
                Else
                    CancelIo m_hCommHandle
                    lWritten = 0
                End If
            Else
                lWritten = 0
                RaiseEvent OnError(lErr, "写入串口失败")
            End If
        End If
        
        CloseHandle ovWrite.hEvent
    Else
        ' 同步写入
        lRet = WriteFile(m_hCommHandle, Data(LBound(Data)), Length, lWritten, ByVal 0&)
        If lRet = 0 Then
            lWritten = 0
            RaiseEvent OnError(GetLastError(), "写入串口失败")
        End If
    End If
    
    WriteBytes = lWritten
End Function

'================================================================================
' 写入 HEX 字符串
'================================================================================
Public Function WriteHex(ByVal HexString As String) As Long
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        WriteHex = 0
        Exit Function
    End If
    
    ' 清理 HEX 字符串 (移除空格、横线等分隔符)
    Dim sCleanHex As String
    Dim i As Long
    Dim sChar As String
    
    sCleanHex = ""
    For i = 1 To Len(HexString)
        sChar = Mid$(HexString, i, 1)
        If (sChar >= "0" And sChar <= "9") Or _
           (UCase$(sChar) >= "A" And UCase$(sChar) <= "F") Then
            sCleanHex = sCleanHex & sChar
        End If
    Next i
    
    ' 检查长度是否为偶数
    If Len(sCleanHex) Mod 2 <> 0 Then
        RaiseEvent OnError(-1, "HEX 字符串长度必须是偶数")
        WriteHex = 0
        Exit Function
    End If
    
    ' 转换为字节数组
    Dim lByteCount As Long
    lByteCount = Len(sCleanHex) \ 2
    Dim bytData() As Byte
    ReDim bytData(lByteCount - 1)
    
    For i = 0 To lByteCount - 1
        bytData(i) = CByte("&H" & Mid$(sCleanHex, i * 2 + 1, 2))
    Next i
    
    ' 发送字节数组
    WriteHex = WriteBytes(bytData)
End Function

'================================================================================
' 读取数据
'================================================================================
Public Function ReadString(Optional ByVal MaxBytes As Long = 4096) As String
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        ReadString = ""
        Exit Function
    End If
    
    Dim sBuffer As String
    Dim lRead As Long
    Dim lRet As Long
    
    sBuffer = String$(MaxBytes, 0)
    
    If m_bUseAsyncIO Then
        ' 异步读取
        Dim ovRead As OVERLAPPED
        ovRead.hEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        
        lRet = ReadFile(m_hCommHandle, ByVal sBuffer, MaxBytes, lRead, ovRead)
        
        If lRet = 0 Then
            Dim lErr As Long
            lErr = GetLastError()
            
            If lErr = ERROR_IO_PENDING Then
                ' 等待读取完成
                Dim lWaitRet As Long
                lWaitRet = WaitForSingleObject(ovRead.hEvent, INFINITE)
                
                If lWaitRet = WAIT_OBJECT_0 Then
                    lRet = 1
                    If GetOverlappedResult(m_hCommHandle, ovRead, lRead, 1) = 0 Then
                        lRead = 0
                    End If
                Else
                    CancelIo m_hCommHandle
                    lRead = 0
                End If
            Else
                lRead = 0
            End If
        End If
        
        CloseHandle ovRead.hEvent
    Else
        ' 同步读取
        lRet = ReadFile(m_hCommHandle, ByVal sBuffer, MaxBytes, lRead, ByVal 0&)
        If lRet = 0 Then
            lRead = 0
        End If
    End If
    
    If lRead > 0 Then
        ReadString = Left$(sBuffer, lRead)
    Else
        ReadString = ""
    End If
End Function

'================================================================================
' 读取字节数组
'================================================================================
Public Function ReadBytes(Optional ByVal MaxBytes As Long = 4096) As Byte()
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        ReadBytes = vbNull
        Exit Function
    End If
    
    Dim bytBuffer() As Byte
    Dim lRead As Long
    Dim lRet As Long
    
    ReDim bytBuffer(MaxBytes - 1)
    
    If m_bUseAsyncIO Then
        ' 异步读取
        Dim ovRead As OVERLAPPED
        ovRead.hEvent = CreateEvent(ByVal 0&, 1, 0, vbNullString)
        
        lRet = ReadFile(m_hCommHandle, bytBuffer(0), MaxBytes, lRead, ovRead)
        
        If lRet = 0 Then
            Dim lErr As Long
            lErr = GetLastError()
            
            If lErr = ERROR_IO_PENDING Then
                ' 等待读取完成
                Dim lWaitRet As Long
                lWaitRet = WaitForSingleObject(ovRead.hEvent, INFINITE)
                
                If lWaitRet = WAIT_OBJECT_0 Then
                    lRet = 1
                    If GetOverlappedResult(m_hCommHandle, ovRead, lRead, 1) = 0 Then
                        lRead = 0
                    End If
                Else
                    CancelIo m_hCommHandle
                    lRead = 0
                End If
            Else
                lRead = 0
            End If
        End If
        
        CloseHandle ovRead.hEvent
    Else
        ' 同步读取
        lRet = ReadFile(m_hCommHandle, bytBuffer(0), MaxBytes, lRead, ByVal 0&)
        If lRet = 0 Then
            lRead = 0
        End If
    End If
    
    If lRead > 0 Then
        ReDim Preserve bytBuffer(lRead - 1)
        ReadBytes = bytBuffer
    Else
        ReadBytes = vbNull
    End If
End Function

'================================================================================
' 读取 HEX 字符串
'================================================================================
Public Function ReadHex(Optional ByVal MaxBytes As Long = 4096, _
                        Optional ByVal Separator As String = " ") As String
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        ReadHex = ""
        Exit Function
    End If
    
    ' 读取字节数组
    Dim bytData() As Byte
    bytData = ReadBytes(MaxBytes)
    
    ' 转换为 HEX 字符串
    Dim sHex As String
    Dim i As Long
    Dim lCount As Long
    
    lCount = UBound(bytData) - LBound(bytData) + 1
    
    If lCount > 0 Then
        For i = LBound(bytData) To UBound(bytData)
            sHex = sHex & Right$("0" & Hex$(bytData(i)), 2)
            If i < UBound(bytData) Then
                sHex = sHex & Separator
            End If
        Next i
    End If
    
    ReadHex = sHex
End Function

'================================================================================
' 启动监听
'================================================================================
Public Sub StartListening()
    If Not m_bIsOpen Then
        RaiseEvent OnError(-1, "串口未打开")
        Exit Sub
    End If
    
    If m_bListening Then Exit Sub
    
    If Not m_bUseAsyncIO Then
        RaiseEvent OnError(-1, "异步IO未启用，无法监听事件")
        Exit Sub
    End If
    
    m_bListening = True
    ResetEvent m_hStopEvent
    
    ' 启动监听线程
    Call ListenThread
End Sub

'================================================================================
' 停止监听
'================================================================================
Public Sub StopListening()
    If Not m_bListening Then Exit Sub
    
    m_bListening = False
    SetEvent m_hStopEvent
    
    ' 取消所有IO操作
    If m_hCommHandle <> INVALID_HANDLE_VALUE Then
        CancelIo m_hCommHandle
    End If
End Sub

'================================================================================
' 监听线程
'================================================================================
Private Sub ListenThread()
    Dim lEvtMask As Long
    Dim lRet As Long
    Dim lWaitRet As Long
    Dim hEvents(1) As Long
    Dim ov As OVERLAPPED
    
    hEvents(0) = m_Overlapped.hEvent
    hEvents(1) = m_hStopEvent
    
    Do While m_bListening
        ' 等待串口事件
        ov.hEvent = m_Overlapped.hEvent
        ResetEvent m_Overlapped.hEvent
        
        lRet = WaitCommEvent(m_hCommHandle, lEvtMask, ov)
        
        If lRet = 0 Then
            Dim lErr As Long
            lErr = GetLastError()
            
            If lErr = ERROR_IO_PENDING Then
                ' 等待事件
                lWaitRet = WaitForMultipleObjects(2, hEvents(0), 0, INFINITE)
                
                If lWaitRet = WAIT_OBJECT_0 Then
                    ' 事件触发
                    lRet = 1
                ElseIf lWaitRet = WAIT_OBJECT_0 + 1 Then
                    ' 停止事件
                    Exit Do
                Else
                    Exit Do
                End If
            Else
                RaiseEvent OnError(lErr, "WaitCommEvent 失败")
                Exit Do
            End If
        End If
        
        If Not m_bListening Then Exit Do
        
        ' 处理事件
        ProcessEvents lEvtMask
        
        ' 允许处理其他消息
        DoEvents
    Loop
End Sub

'================================================================================
' 处理串口事件
'================================================================================
Private Sub ProcessEvents(ByVal EventMask As Long)
    ' 数据接收事件
    If (EventMask And EV_RXCHAR) <> 0 Then
        OnDataReceiveEvent
    End If
    
    ' CTS 改变
    If (EventMask And EV_CTS) <> 0 Then
        OnCTSChangedEvent
    End If
    
    ' DSR 改变
    If (EventMask And EV_DSR) <> 0 Then
        OnDSRChangedEvent
    End If
    
    ' 振铃
    If (EventMask And EV_RING) <> 0 Then
        OnRingEvent
    End If
    
    ' RLSD 改变
    If (EventMask And EV_RLSD) <> 0 Then
        OnRLSDChangedEvent
    End If
    
    ' 中断
    If (EventMask And EV_BREAK) <> 0 Then
        RaiseEvent OnBreak
    End If
    
    ' 错误
    If (EventMask And EV_ERR) <> 0 Then
        OnErrorEvent
    End If
    
    ' 发送缓冲区空
    If (EventMask And EV_TXEMPTY) <> 0 Then
        RaiseEvent OnTxEmpty
    End If
End Sub

'================================================================================
' 数据接收事件处理
'================================================================================
Private Sub OnDataReceiveEvent()
    Dim lErrors As Long
    Dim lStat As COMSTAT
    Dim lBytesToRead As Long
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim sBuffer As String
    
    ' 获取缓冲区状态
    lRet = ClearCommError(m_hCommHandle, lErrors, lStat)
    
    If lRet <> 0 And lStat.cbInQue > 0 Then
        lBytesToRead = lStat.cbInQue
        
        ' 读取数据
        ReDim bytBuffer(lBytesToRead - 1)
        lRet = ReadFile(m_hCommHandle, bytBuffer(0), lBytesToRead, lBytesToRead, ByVal 0&)
        
        If lRet <> 0 Then
            ' 转换为字符串
            sBuffer = StrConv(bytBuffer, vbUnicode)
            
            ' 触发事件
            RaiseEvent OnDataReceived(sBuffer, lBytesToRead)
            RaiseEvent OnDataReceivedBytes(bytBuffer, lBytesToRead)
        End If
    End If
End Sub

'================================================================================
' CTS 改变事件处理
'================================================================================
Private Sub OnCTSChangedEvent()
    Dim lStatus As Long
    
    If GetCommModemStatus(m_hCommHandle, lStatus) <> 0 Then
        Dim bCTS As Boolean
        bCTS = (lStatus And MS_CTS_ON) <> 0
        RaiseEvent OnCTSChanged(bCTS)
    End If
End Sub

'================================================================================
' DSR 改变事件处理
'================================================================================
Private Sub OnDSRChangedEvent()
    Dim lStatus As Long
    
    If GetCommModemStatus(m_hCommHandle, lStatus) <> 0 Then
        Dim bDSR As Boolean
        bDSR = (lStatus And MS_DSR_ON) <> 0
        RaiseEvent OnDSRChanged(bDSR)
    End If
End Sub

'================================================================================
' 振铃事件处理
'================================================================================
Private Sub OnRingEvent()
    Dim lStatus As Long
    
    If GetCommModemStatus(m_hCommHandle, lStatus) <> 0 Then
        Dim bRing As Boolean
        bRing = (lStatus And MS_RING_ON) <> 0
        RaiseEvent OnRing(bRing)
    End If
End Sub

'================================================================================
' RLSD 改变事件处理
'================================================================================
Private Sub OnRLSDChangedEvent()
    Dim lStatus As Long
    
    If GetCommModemStatus(m_hCommHandle, lStatus) <> 0 Then
        Dim bRLSD As Boolean
        bRLSD = (lStatus And MS_RLSD_ON) <> 0
        RaiseEvent OnRLSDChanged(bRLSD)
    End If
End Sub

'================================================================================
' 错误事件处理
'================================================================================
Private Sub OnErrorEvent()
    Dim lErrors As Long
    Dim lStat As COMSTAT
    Dim sErrorMsg As String
    
    ClearCommError m_hCommHandle, lErrors, lStat
    
    sErrorMsg = ""
    
    If (lErrors And CE_RXOVER) <> 0 Then sErrorMsg = sErrorMsg & "接收溢出 "
    If (lErrors And CE_OVERRUN) <> 0 Then sErrorMsg = sErrorMsg & "超时错误 "
    If (lErrors And CE_RXPARITY) <> 0 Then sErrorMsg = sErrorMsg & "校验错误 "
    If (lErrors And CE_FRAME) <> 0 Then sErrorMsg = sErrorMsg & "帧错误 "
    If (lErrors And CE_BREAK) <> 0 Then sErrorMsg = sErrorMsg & "中断错误 "
    If (lErrors And CE_TXFULL) <> 0 Then sErrorMsg = sErrorMsg & "发送缓冲区满 "
    
    If sErrorMsg <> "" Then
        RaiseEvent OnError(lErrors, sErrorMsg)
    End If
End Sub

'================================================================================
' 配置串口参数
'================================================================================
Private Function ConfigurePort(ByVal BaudRate As Long, _
                               ByVal DataBits As Integer, _
                               ByVal Parity As ParityEnum, _
                               ByVal StopBits As StopBitsEnum, _
                               ByVal FlowControl As FlowControlEnum) As Boolean
    
    Dim dcb As DCB
    Dim lRet As Long
    
    ' 获取当前配置
    dcb.DCBlength = Len(dcb)
    lRet = GetCommState(m_hCommHandle, dcb)
    
    If lRet = 0 Then
        RaiseEvent OnError(GetLastError(), "获取串口配置失败")
        ConfigurePort = False
        Exit Function
    End If
    
    ' 设置波特率
    dcb.BaudRate = BaudRate
    
    ' 设置数据位
    dcb.ByteSize = DataBits
    
    ' 设置停止位
    Select Case StopBits
        Case sb1
            dcb.StopBits = ONESTOPBIT
        Case sb1_5
            dcb.StopBits = ONE5STOPBITS
        Case sb2
            dcb.StopBits = TWOSTOPBITS
    End Select
    
    ' 设置校验位
    Select Case Parity
        Case peNone
            dcb.Parity = NOPARITY
            dcb.fBitFields = dcb.fBitFields And (Not fParity)
        Case peOdd
            dcb.Parity = ODDPARITY
            dcb.fBitFields = dcb.fBitFields Or fParity
        Case peEven
            dcb.Parity = EVENPARITY
            dcb.fBitFields = dcb.fBitFields Or fParity
        Case peMark
            dcb.Parity = MARKPARITY
            dcb.fBitFields = dcb.fBitFields Or fParity
        Case peSpace
            dcb.Parity = SPACEPARITY
            dcb.fBitFields = dcb.fBitFields Or fParity
    End Select
    
    ' 设置流控制
    Select Case FlowControl
        Case fcNone
            ' 无流控
            dcb.fBitFields = dcb.fBitFields And Not (fOutX Or fInX Or fRtsControl Or fOutxCtsFlow Or fOutxDsrFlow)
            dcb.fBitFields = dcb.fBitFields Or RTS_CONTROL_ENABLE
            
        Case fcXOnXOff
            ' 软件流控
            dcb.fBitFields = dcb.fBitFields Or fOutX Or fInX
            dcb.fBitFields = dcb.fBitFields And Not (fRtsControl Or fOutxCtsFlow Or fOutxDsrFlow)
            dcb.fBitFields = dcb.fBitFields Or RTS_CONTROL_ENABLE
            dcb.XonChar = &H11
            dcb.XoffChar = &H13
            
        Case fcRTSCTS
            ' 硬件流控
            dcb.fBitFields = dcb.fBitFields Or fRtsControl Or fOutxCtsFlow
            dcb.fBitFields = dcb.fBitFields And Not (fOutX Or fInX Or fOutxDsrFlow)
            dcb.fBitFields = dcb.fBitFields And RTS_CONTROL_HANDSHAKE
            
        Case fcRTSCTSXOnXOff
            ' 软硬件流控
            dcb.fBitFields = dcb.fBitFields Or fOutX Or fInX Or fRtsControl Or fOutxCtsFlow
            dcb.fBitFields = dcb.fBitFields And Not fOutxDsrFlow
            dcb.fBitFields = dcb.fBitFields And RTS_CONTROL_HANDSHAKE
            dcb.XonChar = &H11
            dcb.XoffChar = &H13
    End Select
    
    ' 二进制模式
    dcb.fBitFields = dcb.fBitFields Or fBinary
    
    ' 串口终止时不报告错误
    dcb.fBitFields = dcb.fBitFields And (Not fAbortOnError)
    
    ' 应用配置
    lRet = SetCommState(m_hCommHandle, dcb)
    
    If lRet = 0 Then
        RaiseEvent OnError(GetLastError(), "设置串口配置失败")
        ConfigurePort = False
        Exit Function
    End If
    
    ConfigurePort = True
End Function

'================================================================================
' 设置超时参数
'================================================================================
Private Function SetCommTimeouts_() As Boolean
    Dim ct As COMMTIMEOUTS
    
    ' 配置为立即返回模式，适合异步IO
    ct.ReadIntervalTimeout = MAXDWORD
    ct.ReadTotalTimeoutMultiplier = MAXDWORD
    ct.ReadTotalTimeoutConstant = 1      ' 最小等待时间
    ct.WriteTotalTimeoutMultiplier = 0
    ct.WriteTotalTimeoutConstant = 0     ' 无超时
    
    Dim lRet As Long
    lRet = SetCommTimeouts(m_hCommHandle, ct)
    
    If lRet = 0 Then
        RaiseEvent OnError(GetLastError(), "设置超时参数失败")
        SetCommTimeouts_ = False
        Exit Function
    End If
    
    SetCommTimeouts_ = True
End Function

'================================================================================
' 清空缓冲区
'================================================================================
Public Sub ClearBuffer()
    If Not m_bIsOpen Then Exit Sub
    
    PurgeComm m_hCommHandle, PURGE_TXCLEAR Or PURGE_RXCLEAR
End Sub

'================================================================================
' 清空发送缓冲区
'================================================================================
Public Sub ClearSendBuffer()
    If Not m_bIsOpen Then Exit Sub
    
    PurgeComm m_hCommHandle, PURGE_TXCLEAR
End Sub

'================================================================================
' 清空接收缓冲区
'================================================================================
Public Sub ClearReceiveBuffer()
    If Not m_bIsOpen Then Exit Sub
    
    PurgeComm m_hCommHandle, PURGE_RXCLEAR
End Sub

'================================================================================
' 设置 DTR
'================================================================================
Public Sub SetDTR()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, SETDTR
End Sub

'================================================================================
' 清除 DTR
'================================================================================
Public Sub ClearDTR()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, CLRDTR
End Sub

'================================================================================
' 设置 RTS
'================================================================================
Public Sub SetRTS()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, SETRTS
End Sub

'================================================================================
' 清除 RTS
'================================================================================
Public Sub ClearRTS()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, CLRRTS
End Sub

'================================================================================
' 设置中断
'================================================================================
Public Sub SetBreak()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, SETBREAK
End Sub

'================================================================================
' 清除中断
'================================================================================
Public Sub ClearBreak()
    If Not m_bIsOpen Then Exit Sub
    EscapeCommFunction m_hCommHandle, CLRBREAK
End Sub

'================================================================================
' 获取可读取的字节数
'================================================================================
Public Function BytesInQueue() As Long
    If Not m_bIsOpen Then
        BytesInQueue = 0
        Exit Function
    End If
    
    Dim lErrors As Long
    Dim lStat As COMSTAT
    
    ClearCommError m_hCommHandle, lErrors, lStat
    BytesInQueue = lStat.cbInQue
End Function

'================================================================================
' 获取等待发送的字节数
'================================================================================
Public Function BytesOutQueue() As Long
    If Not m_bIsOpen Then
        BytesOutQueue = 0
        Exit Function
    End If
    
    Dim lErrors As Long
    Dim lStat As COMSTAT
    
    ClearCommError m_hCommHandle, lErrors, lStat
    BytesOutQueue = lStat.cbOutQue
End Function

'================================================================================
' 刷新缓冲区
'================================================================================
Public Function Flush() As Boolean
    If Not m_bIsOpen Then
        Flush = False
        Exit Function
    End If
    
    Dim lRet As Long
    lRet = FlushFileBuffers(m_hCommHandle)
    
    Flush = (lRet <> 0)
End Function

'================================================================================
' 重新配置串口
'================================================================================
Public Function ReConfigure(Optional ByVal BaudRate As Long = -1, _
                            Optional ByVal DataBits As Integer = -1, _
                            Optional ByVal Parity As ParityEnum = -1, _
                            Optional ByVal StopBits As StopBitsEnum = -1, _
                            Optional ByVal FlowControl As FlowControlEnum = -1) As Boolean
    
    If Not m_bIsOpen Then
        ReConfigure = False
        Exit Function
    End If
    
    ' 更新参数
    If BaudRate >= 0 Then m_lBaudRate = BaudRate
    If DataBits >= 0 Then m_iDataBits = DataBits
    If Parity >= 0 Then m_iParity = Parity
    If StopBits >= 0 Then m_iStopBits = StopBits
    If FlowControl >= 0 Then m_iFlowControl = FlowControl
    
    ' 应用配置
    ReConfigure = ConfigurePort(m_lBaudRate, m_iDataBits, m_iParity, m_iStopBits, m_iFlowControl)
End Function

'================================================================================
' 获取 COMSTAT 类型声明
'================================================================================
Private Type COMSTAT
    fBitFields As Long       ' 详见 COMSTAT 位标志
    cbInQue As Long          ' 输入队列中的字节数
    cbOutQue As Long         ' 输出队列中的字节数
End Type

'================================================================================
' GetOverlappedResult API 函数
'================================================================================
Private Declare Function GetOverlappedResult Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpOverlapped As OVERLAPPED, _
    ByRef lpNumberOfBytesTransferred As Long, _
    ByVal bWait As Long) As Long

'================================================================================
' WaitForMultipleObjects API 函数
'================================================================================
Private Declare Function WaitForMultipleObjects Lib "kernel32" ( _
    ByVal nCount As Long, _
    ByRef lpHandles As Long, _
    ByVal bWaitAll As Long, _
    ByVal dwMilliseconds As Long) As Long

'================================================================================
' 属性: IsOpen
'================================================================================
Public Property Get IsOpen() As Boolean
    IsOpen = m_bIsOpen
End Property

'================================================================================
' 属性: PortName
'================================================================================
Public Property Get PortName() As String
    PortName = m_strPortName
End Property

'================================================================================
' 属性: BaudRate
'================================================================================
Public Property Get BaudRate() As Long
    BaudRate = m_lBaudRate
End Property

'================================================================================
' 属性: DataBits
'================================================================================
Public Property Get DataBits() As Integer
    DataBits = m_iDataBits
End Property

'================================================================================
' 属性: StopBits
'================================================================================
Public Property Get StopBits() As Integer
    StopBits = m_iStopBits
End Property

'================================================================================
' 属性: Parity
'================================================================================
Public Property Get Parity() As Integer
    Parity = m_iParity
End Property

'================================================================================
' 属性: FlowControl
'================================================================================
Public Property Get FlowControl() As Integer
    FlowControl = m_iFlowControl
End Property

'================================================================================
' 属性: UseAsyncIO
'================================================================================
Public Property Get UseAsyncIO() As Boolean
    UseAsyncIO = m_bUseAsyncIO
End Property

Public Property Let UseAsyncIO(ByVal vValue As Boolean)
    If m_bIsOpen Then
        Err.Raise -1, "cSerialPort", "串口已打开，不能修改异步IO设置"
    End If
    m_bUseAsyncIO = vValue
End Property

'================================================================================
' 属性: InQueueSize
'================================================================================
Public Property Get InQueueSize() As Long
    InQueueSize = m_lInQueueSize
End Property

Public Property Let InQueueSize(ByVal vValue As Long)
    If m_bIsOpen Then
        Err.Raise -1, "cSerialPort", "串口已打开，不能修改缓冲区大小"
    End If
    m_lInQueueSize = vValue
End Property

'================================================================================
' 属性: OutQueueSize
'================================================================================
Public Property Get OutQueueSize() As Long
    OutQueueSize = m_lOutQueueSize
End Property

Public Property Let OutQueueSize(ByVal vValue As Long)
    If m_bIsOpen Then
        Err.Raise -1, "cSerialPort", "串口已打开，不能修改缓冲区大小"
    End If
    m_lOutQueueSize = vValue
End Property

'================================================================================
' 属性: Listening
'================================================================================
Public Property Get Listening() As Boolean
    Listening = m_bListening
End Property
