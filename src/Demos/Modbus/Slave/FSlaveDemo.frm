VERSION 5.00
Begin VB.Form FSlaveDemo
   Caption         =   "Modbus 从站演示程序 (Slave)"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   9735
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraServerControl
      Caption         =   "服务器控制"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdStop
         Caption         =   "停止服务器"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdStart
         Caption         =   "启动服务器"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Frame fraTCP
         Caption         =   "TCP 设置"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4815
         Begin VB.TextBox txtTCPPort
            Height          =   285
            Left            =   3000
            TabIndex        =   3
            Text            =   "502"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblTCPPort
            Caption         =   "监听端口:"
            Height          =   255
            Left            =   2040
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraRTU
         Caption         =   "RTU 设置"
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   4815
         Begin VB.TextBox txtStopBits
            Height          =   285
            Left            =   3120
            TabIndex        =   14
            Text            =   "1"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtParity
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Text            =   "N"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtDataBits
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Text            =   "8"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBaudRate
            Height          =   285
            Left            =   3120
            TabIndex        =   11
            Text            =   "9600"
            Top             =   -120
            Width           =   855
         End
         Begin VB.TextBox txtSerialPort
            Height          =   285
            Left            =   3120
            TabIndex        =   17
            Text            = "COM1"
            Top             =   -480
            Width           =   855
         End
         Begin VB.Label lblStopBits
            Caption         =   "停止位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   16
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblParity
            Caption         =   "校验位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   15
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblDataBits
            Caption         =   "数据位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblBaudRate
            Caption         =   "波特率:"
            Height          =   255
            Left            =   2400
            TabIndex        =   6
            Top             =   -120
            Width           =   735
         End
         Begin VB.Label lblSerialPort
            Caption         =   "串口:"
            Height          =   255
            Left            =   2400
            TabIndex        =   5
            Top             =   -480
            Width           =   615
         End
      End
      Begin VB.TextBox txtSlaveID
         Height          =   285
         Left            =   3120
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblSlaveID
         Caption         =   "从站ID:"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optProtocol
         Caption         =   "TCP"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optProtocol
         Caption         =   "RTU"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblStatus
         Caption         =   "状态: 已停止"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Frame fraDataManagement
      Caption         =   "数据管理"
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   9495
      Begin VB.CommandButton cmdSetInputRegister
         Caption         =   "设置输入寄存器"
         Height          =   375
         Left            =   7320
         TabIndex        =   35
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdSetDiscreteInput
         Caption         =   "设置离散输入"
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdSetHoldingReg
         Caption         = "设置保持寄存器"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton cmdSetCoil
         Caption         = "设置线圈"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Frame fraDiscreteInputs
         Caption         =   "离散输入"
         Height          =   735
         Left            =   5040
         TabIndex        =   29
         Top             =   960
         Width           =   4335
         Begin VB.TextBox txtDIValue
            Height          =   285
            Left            =   3000
            TabIndex        =   31
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtDIAddress
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblDIValue
            Caption         =   "值:"
            Height          =   255
            Left            =   2520
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblDIAddress
            Caption         =   "地址:"
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraInputRegs
         Caption         =   "输入寄存器"
         Height          =   735
         Left            =   7320
         TabIndex        =   26
         Top             =   960
         Width           =   2055
         Begin VB.TextBox txtIRValue
            Height          =   285
            Left            =   960
            TabIndex        =   28
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtIRAddress
            Height          =   285
            Left            =   960
            TabIndex        =   27
            Text            =   "0"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame fraHoldingRegs
         Caption         =   "保持寄存器"
         Height          =   735
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtHRValue
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtHRAddress
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Text            =   "0"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame fraCoils
         Caption         =   "线圈"
         Height          =   735
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtCoilValue
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtCoilAddress
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Text            =   "0"
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame fraLog
      Caption         =   "活动日志"
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   9495
      Begin VB.TextBox txtLog
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   240
         Width           =   9255
      End
   End
End
Attribute VB_Name = "FSlaveDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FSlaveDemo - Modbus 从站演示程序
'
' Purpose: 演示 Modbus 从站功能,支持 TCP 和 RTU 两种模式
'          - 启动服务器监听连接
'          - 管理线圈、寄存器数据
'          - 响应主站请求
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

Private WithEvents m_Slave As VBMANLIB.cModbusSlave
Attribute m_Slave.VB_VarHelpID = -1

'=========================================================================
' Form Events
'=========================================================================

Private Sub Form_Load()
    Set m_Slave = New VBMANLIB.cModbusSlave
    
    ' 默认 TCP 模式
    optProtocol(1).Value = True
    UpdateProtocolUI
    
    ' 初始化一些测试数据
    InitializeTestData
    
    LogMessage "Modbus 从站演示程序已启动"
    LogMessage "请选择协议类型并启动服务器"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_Slave Is Nothing Then
        If m_Slave.State = MB_SLAVE_STATE_RUNNING Then
            m_Slave.Stop
        End If
        Set m_Slave = Nothing
    End If
End Sub

'=========================================================================
' Protocol Selection
'=========================================================================

Private Sub optProtocol_Click(Index As Integer)
    UpdateProtocolUI
End Sub

Private Sub UpdateProtocolUI()
    If optProtocol(1).Value Then
        ' TCP 模式
        fraTCP.Visible = True
        fraRTU.Visible = False
        m_Slave.ProtocolType = MB_PROTOCOL_TCP
    Else
        ' RTU 模式
        fraTCP.Visible = False
        fraRTU.Visible = True
        m_Slave.ProtocolType = MB_PROTOCOL_RTU
    End If
End Sub

'=========================================================================
' Server Control
'=========================================================================

Private Sub cmdStart_Click()
    On Error GoTo ErrorHandler
    
    ' 设置从站ID
    m_Slave.SlaveID = CByte(Val(txtSlaveID.Text))
    
    If optProtocol(1).Value Then
        ' TCP 模式
        Dim lPort As Long
        lPort = CLng(Val(txtTCPPort.Text))
        LogMessage "正在启动 TCP 服务器, 端口: " & lPort
        m_Slave.Start CStr(lPort)
    Else
        ' RTU 模式
        m_Slave.SerialPort = txtSerialPort.Text
        m_Slave.BaudRate = CLng(Val(txtBaudRate.Text))
        m_Slave.DataBits = CLng(Val(txtDataBits.Text))
        m_Slave.Parity = txtParity.Text
        m_Slave.StopBits = CLng(Val(txtStopBits.Text))
        LogMessage "正在启动 RTU 服务器: " & m_Slave.SerialPort & " (" & m_Slave.BaudRate & ")"
        m_Slave.Start txtSerialPort.Text
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "启动失败: " & Err.Description
    MsgBox "启动失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    m_Slave.Stop
    LogMessage "服务器已停止"
End Sub

'=========================================================================
' Data Management - Coils
'=========================================================================

Private Sub cmdSetCoil_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim bValue As Boolean
    
    lAddress = CLng(Val(txtCoilAddress.Text))
    bValue = (Val(txtCoilValue.Text) <> 0)
    
    m_Slave.SetCoil lAddress, bValue
    LogMessage "设置线圈: 地址=" & lAddress & ", 值=" & IIf(bValue, "1", "0")
    
    Exit Sub
ErrorHandler:
    LogMessage "设置线圈失败: " & Err.Description
End Sub

'=========================================================================
' Data Management - Discrete Inputs
'=========================================================================

Private Sub cmdSetDiscreteInput_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim bValue As Boolean
    
    lAddress = CLng(Val(txtDIAddress.Text))
    bValue = (Val(txtDIValue.Text) <> 0)
    
    m_Slave.SetDiscreteInput lAddress, bValue
    LogMessage "设置离散输入: 地址=" & lAddress & ", 值=" & IIf(bValue, "1", "0")
    
    Exit Sub
ErrorHandler:
    LogMessage "设置离散输入失败: " & Err.Description
End Sub

'=========================================================================
' Data Management - Holding Registers
'=========================================================================

Private Sub cmdSetHoldingReg_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim iValue As Integer
    
    lAddress = CLng(Val(txtHRAddress.Text))
    iValue = CInt(Val(txtHRValue.Text))
    
    m_Slave.SetHoldingRegister lAddress, iValue
    LogMessage "设置保持寄存器: 地址=" & lAddress & ", 值=" & iValue
    
    Exit Sub
ErrorHandler:
    LogMessage "设置保持寄存器失败: " & Err.Description
End Sub

'=========================================================================
' Data Management - Input Registers
'=========================================================================

Private Sub cmdSetInputRegister_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim iValue As Integer
    
    lAddress = CLng(Val(txtIRAddress.Text))
    iValue = CInt(Val(txtIRValue.Text))
    
    m_Slave.SetInputRegister lAddress, iValue
    LogMessage "设置输入寄存器: 地址=" & lAddress & ", 值=" & iValue
    
    Exit Sub
ErrorHandler:
    LogMessage "设置输入寄存器失败: " & Err.Description
End Sub

'=========================================================================
' Modbus Events
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
    ' 调试用: 显示接收到的数据
    ' Dim sHex As String
    ' Dim i As Long
    ' If UBound(Data) >= 0 Then
    '     sHex = "接收: "
    '     For i = 0 To UBound(Data)
    '         sHex = sHex & Right$("0" & Hex$(Data(i)), 2) & " "
    '     Next i
    '     LogMessage sHex
    ' End If
End Sub

Private Sub m_Slave_OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusFunctionCode, ByVal Address As Long, ByVal Quantity As Long)
    Dim sFCName As String
    
    Select Case FunctionCode
        Case MB_FC_READ_COILS
            sFCName = "读线圈"
        Case MB_FC_READ_DISCRETE_INPUTS
            sFCName = "读离散输入"
        Case MB_FC_READ_HOLDING_REGISTERS
            sFCName = "读保持寄存器"
        Case MB_FC_READ_INPUT_REGISTERS
            sFCName = "读输入寄存器"
        Case Else
            sFCName = "未知(FC=" & Hex$(FunctionCode) & ")"
    End Select
    
    LogMessage "读请求 [" & ClientID & "]: " & sFCName & ", 地址=" & Address & ", 数量=" & Quantity
End Sub

Private Sub m_Slave_OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusFunctionCode, ByVal Address As Long, ByRef Data As Variant)
    Dim sFCName As String
    
    Select Case FunctionCode
        Case MB_FC_WRITE_SINGLE_COIL
            sFCName = "写单个线圈"
        Case MB_FC_WRITE_SINGLE_REGISTER
            sFCName = "写单个寄存器"
        Case MB_FC_WRITE_MULTIPLE_COILS
            sFCName = "写多个线圈"
        Case MB_FC_WRITE_MULTIPLE_REGISTERS
            sFCName = "写多个寄存器"
        Case Else
            sFCName = "未知(FC=" & Hex$(FunctionCode) & ")"
    End Select
    
    LogMessage "写请求 [" & ClientID & "]: " & sFCName & ", 地址=" & Address
End Sub

'=========================================================================
' Helper Functions
'=========================================================================

Private Sub InitializeTestData()
    Dim i As Long
    
    ' 初始化线圈 (0-15)
    For i = 0 To 15
        m_Slave.SetCoil i, ((i Mod 2) = 0)
    Next i
    
    ' 初始化离散输入 (0-15)
    For i = 0 To 15
        m_Slave.SetDiscreteInput i, ((i Mod 3) = 0)
    Next i
    
    ' 初始化保持寄存器 (0-99)
    For i = 0 To 99
        m_Slave.SetHoldingRegister i, i * 10
    Next i
    
    ' 初始化输入寄存器 (0-99)
    For i = 0 To 99
        m_Slave.SetInputRegister i, i + 1000
    Next i
    
    LogMessage "测试数据已初始化"
    LogMessage "  线圈: 0-15 (交替 0/1)"
    LogMessage "  离散输入: 0-15 (每3个为1)"
    LogMessage "  保持寄存器: 0-99 (0, 10, 20, ...)"
    LogMessage "  输入寄存器: 0-99 (1000, 1001, ...)"
End Sub

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
