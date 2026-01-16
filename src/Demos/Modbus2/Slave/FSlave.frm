VERSION 5.00
Begin VB.Form FSlave
   Caption         =   "Modbus 从站示例 (Slave)"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  '屏幕中央
   Begin VB.Frame fraServer
      Caption         =   "服务器设置"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtSlaveID
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtPort
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "502"
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdStop
         Caption         =   "停止"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdStart
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblSlaveID
         Caption         =   "从站ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblPort
         Caption         =   "监听端口:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblStatus
         Caption         =   "状态: 已停止"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraData
      Caption         =   "数据设置"
      Height          =   1695
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdSetInputReg
         Caption         =   "设置输入寄存器"
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdSetHoldingReg
         Caption         =   "设置保持寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdSetDiscreteInput
         Caption         =   "设置离散输入"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdSetCoil
         Caption         =   "设置线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtDataValue
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDataAddress
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDataValue
         Caption         =   "值:"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblDataAddress
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraCurrentData
      Caption         =   "当前数据状态"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   9375
      Begin VB.CommandButton cmdRefresh
         Caption         =   "刷新"
         Height          =   375
         Left            =   8280
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCurrentData
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame fraLog
      Caption         =   "通信日志"
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   9375
      Begin VB.CommandButton cmdClearLog
         Caption         =   "清空"
         Height          =   375
         Left            =   8400
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtLog
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   360
         Width           =   8175
      End
   End
End
Attribute VB_Name = "FSlave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FSlave - Modbus 从站示例程序
'
' 功能: 演示 cModbusSlave 类的基本用法
'       - TCP 模式监听主站连接
'       - 提供线圈、离散输入、保持寄存器、输入寄存器数据
'       - 响应主站的读写请求
'
' 使用方法:
'   1. 设置监听端口和从站ID
'   2. 点击"启动"按钮启动服务器
'   3. 等待主站连接并处理请求
'
' 作者: Auto
' 日期: 2026-01-16
'
'=========================================================================
Option Explicit

' Modbus 从站对象
Private WithEvents m_Slave As VBMANLIB.cModbusSlave
Attribute m_Slave.VB_VarHelpID = -1

'=========================================================================
' 窗体事件
'=========================================================================

Private Sub Form_Load()
    ' 创建 Modbus 从站对象
    Set m_Slave = New VBMANLIB.cModbusSlave

    ' 设置为 TCP 模式
    m_Slave.ProtocolType = MB_SLAVE_PROTOCOL_TCP

    ' 初始化测试数据
    InitializeTestData

    LogMessage "Modbus 从站示例程序已启动"
    LogMessage "点击启动按钮开始监听主站连接"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    ' 停止服务器
    If Not m_Slave Is Nothing Then
        If m_Slave.State = MB_SLAVE_STATE_RUNNING Then
            m_Slave.StopMe
        End If
        Set m_Slave = Nothing
    End If
End Sub

'=========================================================================
' 服务器控制
'=========================================================================

Private Sub cmdStart_Click()
    On Error GoTo ErrorHandler

    Dim lPort As Long

    ' 设置从站ID
    m_Slave.SlaveID = CByte(Val(txtSlaveID.Text))

    ' 获取端口
    lPort = CLng(Val(txtPort.Text))

    LogMessage "正在启动 TCP 服务器, 端口: " & lPort

    ' 启动服务器
    m_Slave.Start CStr(lPort)

    Exit Sub
ErrorHandler:
    LogMessage "启动失败: " & Err.Description
    MsgBox "启动失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    m_Slave.StopMe
    LogMessage "服务器已停止"
End Sub

'=========================================================================
' 数据设置
'=========================================================================

Private Sub cmdSetCoil_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim bValue As Boolean

    lAddress = CLng(Val(txtDataAddress.Text))
    bValue = (Val(txtDataValue.Text) <> 0)

    m_Slave.SetCoil lAddress, bValue
    LogMessage "设置线圈: 地址=" & lAddress & ", 值=" & IIf(bValue, "ON", "OFF")
    RefreshDataDisplay

    Exit Sub
ErrorHandler:
    LogMessage "设置线圈失败: " & Err.Description
End Sub

Private Sub cmdSetDiscreteInput_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim bValue As Boolean

    lAddress = CLng(Val(txtDataAddress.Text))
    bValue = (Val(txtDataValue.Text) <> 0)

    m_Slave.SetDiscreteInput lAddress, bValue
    LogMessage "设置离散输入: 地址=" & lAddress & ", 值=" & IIf(bValue, "ON", "OFF")
    RefreshDataDisplay

    Exit Sub
ErrorHandler:
    LogMessage "设置离散输入失败: " & Err.Description
End Sub

Private Sub cmdSetHoldingReg_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim iValue As Integer

    lAddress = CLng(Val(txtDataAddress.Text))
    iValue = CInt(Val(txtDataValue.Text))

    m_Slave.SetHoldingRegister lAddress, iValue
    LogMessage "设置保持寄存器: 地址=" & lAddress & ", 值=" & iValue
    RefreshDataDisplay

    Exit Sub
ErrorHandler:
    LogMessage "设置保持寄存器失败: " & Err.Description
End Sub

Private Sub cmdSetInputReg_Click()
    On Error GoTo ErrorHandler

    Dim lAddress As Long
    Dim iValue As Integer

    lAddress = CLng(Val(txtDataAddress.Text))
    iValue = CInt(Val(txtDataValue.Text))

    m_Slave.SetInputRegister lAddress, iValue
    LogMessage "设置输入寄存器: 地址=" & lAddress & ", 值=" & iValue
    RefreshDataDisplay

    Exit Sub
ErrorHandler:
    LogMessage "设置输入寄存器失败: " & Err.Description
End Sub

'=========================================================================
' 事件处理
'=========================================================================

Private Sub m_Slave_OnStarted()
    LogMessage "服务器已启动"
    lblStatus.Caption = "状态: 运行中"
    lblStatus.ForeColor = &HC000&
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    LogMessage "等待主站连接..."
End Sub

Private Sub m_Slave_OnStopped()
    LogMessage "服务器已停止"
    lblStatus.Caption = "状态: 已停止"
    lblStatus.ForeColor = &HFF&
    cmdStart.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub m_Slave_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
    LogMessage "客户端已连接: " & ClientID & " (" & RemoteAddress & ")"
End Sub

Private Sub m_Slave_OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
    LogMessage "客户端已断开: " & ClientID & " (" & Reason & ")"
End Sub

Private Sub m_Slave_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

Private Sub m_Slave_OnDataReceived(ByVal ClientID As String, Data() As Byte)
    ' 可选: 显示原始数据
    ' Dim sHex As String
    ' Dim i As Long
    ' sHex = "RX [" & ClientID & "]: "
    ' For i = 0 To UBound(Data)
    '     sHex = sHex & Right$("0" & Hex$(Data(i)), 2) & " "
    ' Next i
    ' LogMessage sHex
End Sub

Private Sub m_Slave_OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByVal Quantity As Long)
    Dim sFCName As String

    Select Case FunctionCode
        Case MB_SLAVE_FC_READ_COILS
            sFCName = "读线圈"
        Case MB_SLAVE_FC_READ_DISCRETE_INPUTS
            sFCName = "读离散输入"
        Case MB_SLAVE_FC_READ_HOLDING_REGISTERS
            sFCName = "读保持寄存器"
        Case MB_SLAVE_FC_READ_INPUT_REGISTERS
            sFCName = "读输入寄存器"
        Case Else
            sFCName = "未知(FC=" & Hex$(FunctionCode) & ")"
    End Select

    LogMessage "读请求 [" & ClientID & "]: " & sFCName & ", 地址=" & Address & ", 数量=" & Quantity
End Sub

Private Sub m_Slave_OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByRef Data As Variant)
    Dim sFCName As String

    Select Case FunctionCode
        Case MB_SLAVE_FC_WRITE_SINGLE_COIL
            sFCName = "写单线圈"
        Case MB_SLAVE_FC_WRITE_SINGLE_REGISTER
            sFCName = "写单寄存器"
        Case MB_SLAVE_FC_WRITE_MULTIPLE_COILS
            sFCName = "写多线圈"
        Case MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
            sFCName = "写多寄存器"
        Case Else
            sFCName = "未知(FC=" & Hex$(FunctionCode) & ")"
    End Select

    LogMessage "写请求 [" & ClientID & "]: " & sFCName & ", 地址=" & Address
    RefreshDataDisplay
End Sub

'=========================================================================
' 辅助函数
'=========================================================================

Private Sub InitializeTestData()
    Dim i As Long

    ' 初始化线圈 (0-15): 交替模式
    For i = 0 To 15
        m_Slave.SetCoil i, ((i Mod 2) = 0)
    Next i

    ' 初始化离散输入 (0-15): 每3个为1
    For i = 0 To 15
        m_Slave.SetDiscreteInput i, ((i Mod 3) = 0)
    Next i

    ' 初始化保持寄存器 (0-99): 值为 地址*10
    For i = 0 To 99
        m_Slave.SetHoldingRegister i, i * 10
    Next i

    ' 初始化输入寄存器 (0-99): 值为 1000+地址
    For i = 0 To 99
        m_Slave.SetInputRegister i, 1000 + i
    Next i

    LogMessage "测试数据已初始化:"
    LogMessage "  线圈 [0-15]: 交替模式 (1,0,1,0...)"
    LogMessage "  离散输入 [0-15]: 每3个为1"
    LogMessage "  保持寄存器 [0-99]: 值=地址*10"
    LogMessage "  输入寄存器 [0-99]: 值=1000+地址"

    RefreshDataDisplay
End Sub

Private Sub cmdRefresh_Click()
    RefreshDataDisplay
End Sub

Private Sub RefreshDataDisplay()
    On Error Resume Next

    Dim sDisplay As String
    Dim i As Long

    sDisplay = "=== 当前数据状态 ===" & vbCrLf & vbCrLf

    ' 显示线圈
    sDisplay = sDisplay & "线圈 [0-15]: "
    For i = 0 To 15
        sDisplay = sDisplay & IIf(m_Slave.GetCoil(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sDisplay = sDisplay & " "
    Next i
    sDisplay = sDisplay & vbCrLf

    ' 显示离散输入
    sDisplay = sDisplay & "离散输入 [0-15]: "
    For i = 0 To 15
        sDisplay = sDisplay & IIf(m_Slave.GetDiscreteInput(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sDisplay = sDisplay & " "
    Next i
    sDisplay = sDisplay & vbCrLf & vbCrLf

    ' 显示保持寄存器
    sDisplay = sDisplay & "保持寄存器 [0-9]: "
    For i = 0 To 9
        sDisplay = sDisplay & m_Slave.GetHoldingRegister(i)
        If i < 9 Then sDisplay = sDisplay & ", "
    Next i
    sDisplay = sDisplay & vbCrLf

    ' 显示输入寄存器
    sDisplay = sDisplay & "输入寄存器 [0-9]: "
    For i = 0 To 9
        sDisplay = sDisplay & m_Slave.GetInputRegister(i)
        If i < 9 Then sDisplay = sDisplay & ", "
    Next i

    txtCurrentData.Text = sDisplay
End Sub

Private Sub cmdClearLog_Click()
    txtLog.Text = ""
End Sub

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
