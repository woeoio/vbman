VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' 类: cSerialConfig
' 说明: 串口通信配置类 - 封装 DCB 和 COMMTIMEOUTS 结构
'       提供类型安全的属性访问，以及与 Win32 结构的双向转换
'
' 注意: Public 属性统一使用 Long 而非枚举类型，
'       因为 VB6 不允许 Private 类的 Public 成员使用枚举类型。
'       调用时仍可使用 modSerialPortAPI 中定义的枚举常量:
'           cfg.BaudRate = br115200   ' br115200 自动强转为 Long
'===============================================================================

'===============================================================================
' 私有成员 - 通信参数（内部变量保留枚举类型，享有类型安全）
'===============================================================================

Private m_BaudRate As eBaudRate
Private m_DataBits As Long
Private m_StopBits As eStopBits
Private m_Parity As eParity
Private m_FlowControl As eFlowControl
Private m_DtrControl As eDtrControl
Private m_RtsControl As eRtsControl

'===============================================================================
' 私有成员 - 超时参数
'===============================================================================

Private m_ReadIntervalTimeout As Long
Private m_ReadTotalTimeoutMultiplier As Long
Private m_ReadTotalTimeoutConstant As Long
Private m_WriteTotalTimeoutMultiplier As Long
Private m_WriteTotalTimeoutConstant As Long

'===============================================================================
' 私有成员 - 缓冲区大小
'===============================================================================

Private m_InBufferSize As Long
Private m_OutBufferSize As Long

'===============================================================================
' 私有成员 - 特殊字符
'===============================================================================

Private m_XonChar As Byte
Private m_XoffChar As Byte
Private m_ErrorChar As Byte
Private m_EofChar As Byte
Private m_EvtChar As Byte

'===============================================================================
' 私有成员 - 标志
'===============================================================================

Private m_NullDiscard As Boolean       ' 丢弃NULL字节
Private m_AbortOnError As Boolean      ' 错误时终止读写
Private m_TXContinueOnXoff As Boolean  ' XOFF后继续发送
Private m_ParityReplace As Boolean     ' 校验错误时替换字符

'===============================================================================
' 类初始化
'===============================================================================

Private Sub Class_Initialize()
    ResetDefaults
End Sub

'===============================================================================
' 属性 - 通信参数（Public接口用Long，内部变量用枚举）
'===============================================================================

Public Property Get BaudRate() As Long
    BaudRate = m_BaudRate
End Property

Public Property Let BaudRate(ByVal Value As Long)
    m_BaudRate = Value
End Property

Public Property Get DataBits() As Long
    DataBits = m_DataBits
End Property

Public Property Let DataBits(ByVal Value As Long)
    If Value < 4 Or Value > 8 Then
        ERR.Raise 5, "cSerialConfig", "数据位必须在4-8之间"
    End If
    m_DataBits = Value
End Property

Public Property Get StopBits() As Long
    StopBits = m_StopBits
End Property

Public Property Let StopBits(ByVal Value As Long)
    m_StopBits = Value
End Property

Public Property Get Parity() As Long
    Parity = m_Parity
End Property

Public Property Let Parity(ByVal Value As Long)
    m_Parity = Value
End Property

Public Property Get FlowControl() As Long
    FlowControl = m_FlowControl
End Property

Public Property Let FlowControl(ByVal Value As Long)
    m_FlowControl = Value
    ' 同步更新DTR/RTS控制模式
    Select Case Value
        Case fcRtsCts
            m_RtsControl = rcHandshake
        Case fcDtrDsr
            m_DtrControl = dcHandshake
        Case fcRtsCtsAndXonXoff
            m_RtsControl = rcHandshake
        Case fcNone, fcXonXoff
            ' 不自动修改DTR/RTS
    End Select
End Property

Public Property Get DtrControl() As Long
    DtrControl = m_DtrControl
End Property

Public Property Let DtrControl(ByVal Value As Long)
    m_DtrControl = Value
End Property

Public Property Get RtsControl() As Long
    RtsControl = m_RtsControl
End Property

Public Property Let RtsControl(ByVal Value As Long)
    m_RtsControl = Value
End Property

'===============================================================================
' 属性 - 超时参数
'===============================================================================

Public Property Get ReadIntervalTimeout() As Long
    ReadIntervalTimeout = m_ReadIntervalTimeout
End Property

Public Property Let ReadIntervalTimeout(ByVal Value As Long)
    m_ReadIntervalTimeout = Value
End Property

Public Property Get ReadTotalTimeoutMultiplier() As Long
    ReadTotalTimeoutMultiplier = m_ReadTotalTimeoutMultiplier
End Property

Public Property Let ReadTotalTimeoutMultiplier(ByVal Value As Long)
    m_ReadTotalTimeoutMultiplier = Value
End Property

Public Property Get ReadTotalTimeoutConstant() As Long
    ReadTotalTimeoutConstant = m_ReadTotalTimeoutConstant
End Property

Public Property Let ReadTotalTimeoutConstant(ByVal Value As Long)
    m_ReadTotalTimeoutConstant = Value
End Property

Public Property Get WriteTotalTimeoutMultiplier() As Long
    WriteTotalTimeoutMultiplier = m_WriteTotalTimeoutMultiplier
End Property

Public Property Let WriteTotalTimeoutMultiplier(ByVal Value As Long)
    m_WriteTotalTimeoutMultiplier = Value
End Property

Public Property Get WriteTotalTimeoutConstant() As Long
    WriteTotalTimeoutConstant = m_WriteTotalTimeoutConstant
End Property

Public Property Let WriteTotalTimeoutConstant(ByVal Value As Long)
    m_WriteTotalTimeoutConstant = Value
End Property

'===============================================================================
' 属性 - 缓冲区大小
'===============================================================================

Public Property Get InBufferSize() As Long
    InBufferSize = m_InBufferSize
End Property

Public Property Let InBufferSize(ByVal Value As Long)
    If Value < 1 Then
        ERR.Raise 5, "cSerialConfig", "输入缓冲区大小必须>0"
    End If
    m_InBufferSize = Value
End Property

Public Property Get OutBufferSize() As Long
    OutBufferSize = m_OutBufferSize
End Property

Public Property Let OutBufferSize(ByVal Value As Long)
    If Value < 1 Then
        ERR.Raise 5, "cSerialConfig", "输出缓冲区大小必须>0"
    End If
    m_OutBufferSize = Value
End Property

'===============================================================================
' 属性 - 特殊字符
'===============================================================================

Public Property Get XonChar() As Byte
    XonChar = m_XonChar
End Property

Public Property Let XonChar(ByVal Value As Byte)
    m_XonChar = Value
End Property

Public Property Get XoffChar() As Byte
    XoffChar = m_XoffChar
End Property

Public Property Let XoffChar(ByVal Value As Byte)
    m_XoffChar = Value
End Property

Public Property Get ErrorChar() As Byte
    ErrorChar = m_ErrorChar
End Property

Public Property Let ErrorChar(ByVal Value As Byte)
    m_ErrorChar = Value
End Property

Public Property Get EofChar() As Byte
    EofChar = m_EofChar
End Property

Public Property Let EofChar(ByVal Value As Byte)
    m_EofChar = Value
End Property

Public Property Get EvtChar() As Byte
    EvtChar = m_EvtChar
End Property

Public Property Let EvtChar(ByVal Value As Byte)
    m_EvtChar = Value
End Property

'===============================================================================
' 属性 - 标志
'===============================================================================

Public Property Get NullDiscard() As Boolean
    NullDiscard = m_NullDiscard
End Property

Public Property Let NullDiscard(ByVal Value As Boolean)
    m_NullDiscard = Value
End Property

Public Property Get AbortOnError() As Boolean
    AbortOnError = m_AbortOnError
End Property

Public Property Let AbortOnError(ByVal Value As Boolean)
    m_AbortOnError = Value
End Property

Public Property Get TXContinueOnXoff() As Boolean
    TXContinueOnXoff = m_TXContinueOnXoff
End Property

Public Property Let TXContinueOnXoff(ByVal Value As Boolean)
    m_TXContinueOnXoff = Value
End Property

Public Property Get ParityReplace() As Boolean
    ParityReplace = m_ParityReplace
End Property

Public Property Let ParityReplace(ByVal Value As Boolean)
    m_ParityReplace = Value
End Property

'===============================================================================
' 方法 - 重置默认值
'===============================================================================

Public Sub ResetDefaults()
    ' 通信参数
    m_BaudRate = br9600
    m_DataBits = 8
    m_StopBits = sb1
    m_Parity = ptNone
    m_FlowControl = fcNone
    m_DtrControl = dcEnable
    m_RtsControl = rcEnable

    ' 超时参数（兼容MSComm的默认行为）
    m_ReadIntervalTimeout = 50
    m_ReadTotalTimeoutMultiplier = 0
    m_ReadTotalTimeoutConstant = 100
    m_WriteTotalTimeoutMultiplier = 0
    m_WriteTotalTimeoutConstant = 100

    ' 缓冲区大小
    m_InBufferSize = 4096
    m_OutBufferSize = 4096

    ' 特殊字符
    m_XonChar = &H11   ' DC1 (Ctrl+Q)
    m_XoffChar = &H13  ' DC3 (Ctrl+S)
    m_ErrorChar = 0
    m_EofChar = 0
    m_EvtChar = 0

    ' 标志
    m_NullDiscard = False
    m_AbortOnError = False
    m_TXContinueOnXoff = False
    m_ParityReplace = True
End Sub

'===============================================================================
' 方法 - 构建 DCB 结构
'===============================================================================

Friend Function BuildDCB() As DCB
    Dim d As DCB
    d.DCBlength = LenB(d)
    d.BaudRate = CLng(m_BaudRate)
    d.ByteSize = CByte(m_DataBits)
    d.Parity = CByte(m_Parity)
    d.StopBits = CByte(m_StopBits)
    d.XonChar = m_XonChar
    d.XoffChar = m_XoffChar
    d.ErrorChar = m_ErrorChar
    d.EofChar = m_EofChar
    d.EvtChar = m_EvtChar

    ' XON/XOFF 阈值
    d.XonLim = CInt(m_InBufferSize \ 4)
    d.XoffLim = CInt(m_InBufferSize \ 4)

    ' 构建 fBitFields
    Dim f As Long
    f = 0

    ' fBinary - 必须为1（Windows强制）
    SetBit f, DCB_fBinary, True

    ' fParity - 校验检查
    SetBit f, DCB_fParity, m_ParityReplace

    ' DTR控制模式（2位字段）
    SetBits f, DCB_fDtrControl, 2, CLng(m_DtrControl)

    ' RTS控制模式（2位字段）
    SetBits f, DCB_fRtsControl, 2, CLng(m_RtsControl)

    ' 流控制
    Select Case m_FlowControl
        Case fcXonXoff
            SetBit f, DCB_fOutX, True
            SetBit f, DCB_fInX, True
        Case fcRtsCts
            SetBit f, DCB_fOutxCtsFlow, True
        Case fcDtrDsr
            SetBit f, DCB_fOutxDsrFlow, True
        Case fcRtsCtsAndXonXoff
            SetBit f, DCB_fOutxCtsFlow, True
            SetBit f, DCB_fOutX, True
            SetBit f, DCB_fInX, True
    End Select

    ' 其他标志
    SetBit f, DCB_fNull, m_NullDiscard
    SetBit f, DCB_fAbortOnError, m_AbortOnError
    SetBit f, DCB_fTXContinueOnXoff, m_TXContinueOnXoff
    If m_ParityReplace And m_Parity <> ptNone Then
        SetBit f, DCB_fErrorChar, True
    End If

    d.fBitFields = f
    BuildDCB = d
End Function

'===============================================================================
' 方法 - 从 DCB 结构读取配置
'===============================================================================

Friend Sub ApplyDCB(d As DCB)
    m_BaudRate = d.BaudRate
    m_DataBits = d.ByteSize
    m_Parity = d.Parity
    m_StopBits = d.StopBits
    m_XonChar = d.XonChar
    m_XoffChar = d.XoffChar
    m_ErrorChar = d.ErrorChar
    m_EofChar = d.EofChar
    m_EvtChar = d.EvtChar

    Dim f As Long
    f = d.fBitFields

    ' 读取单比特标志
    m_ParityReplace = GetBit(f, DCB_fParity)
    m_NullDiscard = GetBit(f, DCB_fNull)
    m_AbortOnError = GetBit(f, DCB_fAbortOnError)
    m_TXContinueOnXoff = GetBit(f, DCB_fTXContinueOnXoff)

    ' 读取2位字段
    m_DtrControl = GetBits(f, DCB_fDtrControl, 2)
    m_RtsControl = GetBits(f, DCB_fRtsControl, 2)

    ' 推断流控制模式
    Dim bOutX As Boolean: bOutX = GetBit(f, DCB_fOutX)
    Dim bInX As Boolean: bInX = GetBit(f, DCB_fInX)
    Dim bCts As Boolean: bCts = GetBit(f, DCB_fOutxCtsFlow)
    Dim bDsr As Boolean: bDsr = GetBit(f, DCB_fOutxDsrFlow)

    If bCts And (bOutX Or bInX) Then
        m_FlowControl = fcRtsCtsAndXonXoff
    ElseIf bCts Then
        m_FlowControl = fcRtsCts
    ElseIf bDsr Then
        m_FlowControl = fcDtrDsr
    ElseIf bOutX Or bInX Then
        m_FlowControl = fcXonXoff
    Else
        m_FlowControl = fcNone
    End If
End Sub

'===============================================================================
' 方法 - 构建 COMMTIMEOUTS 结构
'===============================================================================

Friend Function BuildTimeouts() As COMMTIMEOUTS
    Dim ct As COMMTIMEOUTS
    ct.ReadIntervalTimeout = m_ReadIntervalTimeout
    ct.ReadTotalTimeoutMultiplier = m_ReadTotalTimeoutMultiplier
    ct.ReadTotalTimeoutConstant = m_ReadTotalTimeoutConstant
    ct.WriteTotalTimeoutMultiplier = m_WriteTotalTimeoutMultiplier
    ct.WriteTotalTimeoutConstant = m_WriteTotalTimeoutConstant
    BuildTimeouts = ct
End Function

'===============================================================================
' 方法 - 从 COMMTIMEOUTS 结构读取配置
'===============================================================================

Friend Sub ApplyTimeouts(ct As COMMTIMEOUTS)
    m_ReadIntervalTimeout = ct.ReadIntervalTimeout
    m_ReadTotalTimeoutMultiplier = ct.ReadTotalTimeoutMultiplier
    m_ReadTotalTimeoutConstant = ct.ReadTotalTimeoutConstant
    m_WriteTotalTimeoutMultiplier = ct.WriteTotalTimeoutMultiplier
    m_WriteTotalTimeoutConstant = ct.WriteTotalTimeoutConstant
End Sub

'===============================================================================
' 方法 - 克隆配置
'===============================================================================

Friend Function Clone() As cSerialConfig
    Dim cfg As New cSerialConfig
    cfg.BaudRate = m_BaudRate
    cfg.DataBits = m_DataBits
    cfg.StopBits = m_StopBits
    cfg.Parity = m_Parity
    cfg.FlowControl = m_FlowControl
    cfg.DtrControl = m_DtrControl
    cfg.RtsControl = m_RtsControl
    cfg.ReadIntervalTimeout = m_ReadIntervalTimeout
    cfg.ReadTotalTimeoutMultiplier = m_ReadTotalTimeoutMultiplier
    cfg.ReadTotalTimeoutConstant = m_ReadTotalTimeoutConstant
    cfg.WriteTotalTimeoutMultiplier = m_WriteTotalTimeoutMultiplier
    cfg.WriteTotalTimeoutConstant = m_WriteTotalTimeoutConstant
    cfg.InBufferSize = m_InBufferSize
    cfg.OutBufferSize = m_OutBufferSize
    cfg.XonChar = m_XonChar
    cfg.XoffChar = m_XoffChar
    cfg.ErrorChar = m_ErrorChar
    cfg.EofChar = m_EofChar
    cfg.EvtChar = m_EvtChar
    cfg.NullDiscard = m_NullDiscard
    cfg.AbortOnError = m_AbortOnError
    cfg.TXContinueOnXoff = m_TXContinueOnXoff
    cfg.ParityReplace = m_ParityReplace
    Set Clone = cfg
End Function

'===============================================================================
' 方法 - 快速设置模式字符串
'===============================================================================

' 示例: "baud=9600 parity=N data=8 stop=1"
' 与 BuildCommDCB API 兼容的格式
Public Sub FromModeString(ByVal ModeStr As String)
    Dim d As DCB
    d.DCBlength = LenB(d)
    If BuildCommDCBVB(ModeStr, d) = 0 Then
        ERR.Raise 5, "cSerialConfig", "无法解析模式字符串: " & ModeStr
    End If
    ApplyDCB d
End Sub

' 生成模式字符串
Public Function ToModeString() As String
    Dim sParity As String
    Select Case m_Parity
        Case ptNone: sParity = "N"
        Case ptOdd:  sParity = "O"
        Case ptEven: sParity = "E"
        Case ptMark: sParity = "M"
        Case ptSpace: sParity = "S"
    End Select
    Dim sStop As String
    Select Case m_StopBits
        Case sb1:   sStop = "1"
        Case sb1_5: sStop = "1.5"
        Case sb2:   sStop = "2"
    End Select
    ToModeString = "baud=" & CStr(m_BaudRate) & " parity=" & sParity & _
                    " data=" & CStr(m_DataBits) & " stop=" & sStop
End Function

'===============================================================================
' 方法 - 配置预设
'===============================================================================

' 非阻塞读取模式：Read立即返回，有数据就读，没数据返回空
Public Sub SetNonBlockingRead()
    m_ReadIntervalTimeout = &HFFFFFFFF   ' MAXDWORD
    m_ReadTotalTimeoutMultiplier = 0
    m_ReadTotalTimeoutConstant = 0
End Sub

' 阻塞读取模式：等待直到读取到指定字节数
Public Sub SetBlockingRead()
    m_ReadIntervalTimeout = 0
    m_ReadTotalTimeoutMultiplier = 0
    m_ReadTotalTimeoutConstant = 0
End Sub

' 带超时的读取模式：指定总超时(毫秒)
Public Sub SetReadTimeout(ByVal TimeoutMs As Long)
    m_ReadIntervalTimeout = 50
    m_ReadTotalTimeoutMultiplier = 0
    m_ReadTotalTimeoutConstant = TimeoutMs
End Sub

' 写入超时(毫秒)
Public Sub SetWriteTimeout(ByVal TimeoutMs As Long)
    m_WriteTotalTimeoutMultiplier = 0
    m_WriteTotalTimeoutConstant = TimeoutMs
End Sub
