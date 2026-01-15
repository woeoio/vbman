VERSION 5.00
Begin VB.Form FModbusTest 
   Caption         =   "Modbus 测试工具"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9000
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "操作日志"
      Height          =   3015
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   8775
      Begin VB.TextBox txtLog 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "读取操作"
      Height          =   1935
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdReadInputRegs 
         Caption         =   "读输入寄存器"
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadHoldingRegs 
         Caption         =   "读保持寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadDiscreteInputs 
         Caption         =   "读离散输入"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadCoils 
         Caption         =   "读线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtReadAddress 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtReadQuantity 
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Text            =   "10"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "数量:"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "起始地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "连接设置"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Frame FrameTCP 
         Caption         =   "TCP 设置"
         Height          =   855
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   4095
         Begin VB.TextBox txtTCPPort 
            Height          =   285
            Left            =   2040
            TabIndex        =   35
            Text            =   "502"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTCPHost 
            Height          =   285
            Left            =   2040
            TabIndex        =   36
            Text            =   "192.168.1.100"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "端口:"
            Height          =   255
            Left            =   1320
            TabIndex        =   37
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "主机:"
            Height          =   255
            Left            =   1320
            TabIndex        =   38
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "断开连接"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "连接"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Frame FrameRTU 
         Caption         =   "RTU 设置"
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4095
         Begin VB.TextBox txtStopBits 
            Height          =   285
            Left            =   3120
            TabIndex        =   9
            Text            =   "1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtParity 
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Text            =   "N"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtDataBits 
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Text            =   "8"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtBaudRate 
            Height          =   285
            Left            =   3120
            TabIndex        =   6
            Text            =   "9600"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtSerialPort 
            Height          =   285
            Left            =   3120
            TabIndex        =   5
            Text            =   "COM3"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "停止位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "校验位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   3
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "数据位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   2
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "波特率:"
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "串口:"
            Height          =   255
            Left            =   2400
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.OptionButton optProtocol 
         Caption         =   "TCP"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   39
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optProtocol 
         Caption         =   "RTU"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtSlaveID 
         Height          =   285
         Left            =   3120
         TabIndex        =   41
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "从站ID:"
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblProtocol 
         Caption         =   "协议类型:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "写入操作"
      Height          =   1935
      Left            =   4560
      TabIndex        =   22
      Top             =   2160
      Width           =   4335
      Begin VB.CommandButton cmdWriteMultipleRegs 
         Caption         =   "写多个寄存器"
         Height          =   375
         Left            =   2280
         TabIndex        =   30
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteSingleRegister 
         Caption         =   "写单个寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteMultipleCoils 
         Caption         =   "写多个线圈"
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteSingleCoil 
         Caption         =   "写单个线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtWriteValue 
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtWriteAddress 
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "值:"
         Height          =   255
         Left            =   1800
         TabIndex        =   24
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "FModbusTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FModbusTest - Modbus 测试工具表单
'
' Purpose: 提供 Modbus RTU 和 TCP 模式的完整测试界面
'
' Author: Auto
' Date: 2026-01-21
'
'=========================================================================
Option Explicit

Private WithEvents m_Modbus As cModbus
Attribute m_Modbus.VB_VarHelpID = -1

'=========================================================================
' Form Events
'=========================================================================

Private Sub Form_Load()
    Set m_Modbus = New cModbus
    
    ' 默认 TCP 模式
    optProtocol(1).Value = True
    UpdateProtocolUI
    
    LogMessage "Modbus 测试工具已启动"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_Modbus Is Nothing Then
        If m_Modbus.State = MB_STATE_CONNECTED Then
            m_Modbus.Disconnect
        End If
        Set m_Modbus = Nothing
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
        FrameTCP.Visible = True
        FrameRTU.Visible = False
        m_Modbus.ProtocolType = MB_PROTOCOL_TCP
    Else
        ' RTU 模式
        FrameTCP.Visible = False
        FrameRTU.Visible = True
        m_Modbus.ProtocolType = MB_PROTOCOL_RTU
    End If
End Sub

'=========================================================================
' Connection Management
'=========================================================================

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler
    
    ' 设置从站ID
    m_Modbus.SlaveID = CByte(Val(txtSlaveID.Text))
    m_Modbus.ResponseTimeout = 2000
    
    If optProtocol(1).Value Then
        ' TCP 模式
        m_Modbus.TCPHost = txtTCPHost.Text
        m_Modbus.TCPPort = CLng(Val(txtTCPPort.Text))
        LogMessage "正在连接到 TCP: " & m_Modbus.TCPHost & ":" & m_Modbus.TCPPort
        m_Modbus.Connect
    Else
        ' RTU 模式
        m_Modbus.SerialPort = txtSerialPort.Text
        m_Modbus.BaudRate = CLng(Val(txtBaudRate.Text))
        m_Modbus.DataBits = CLng(Val(txtDataBits.Text))
        m_Modbus.Parity = txtParity.Text
        m_Modbus.StopBits = CLng(Val(txtStopBits.Text))
        LogMessage "正在连接到串口: " & m_Modbus.SerialPort & " (" & m_Modbus.BaudRate & "," & m_Modbus.DataBits & "," & m_Modbus.Parity & "," & m_Modbus.StopBits & ")"
        m_Modbus.Connect
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "连接失败: " & Err.Description
    MsgBox "连接失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    m_Modbus.Disconnect
    LogMessage "已断开连接"
End Sub

'=========================================================================
' Read Operations
'=========================================================================

Private Sub cmdReadCoils_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baCoils() As Boolean
    Dim i As Long
    Dim sResult As String
    
    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))
    
    LogMessage "读取线圈: 地址=" & lAddress & ", 数量=" & lQuantity
    
    baCoils = m_Modbus.ReadCoils(lAddress, lQuantity)
    
    sResult = "线圈值: "
    For i = 0 To UBound(baCoils)
        sResult = sResult & IIf(baCoils(i), "1", "0")
        If i < UBound(baCoils) Then sResult = sResult & ", "
    Next i
    
    LogMessage sResult
    Exit Sub
ErrorHandler:
    LogMessage "读取线圈失败: " & Err.Description
    MsgBox "读取失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdReadDiscreteInputs_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baInputs() As Boolean
    Dim i As Long
    Dim sResult As String
    
    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))
    
    LogMessage "读取离散输入: 地址=" & lAddress & ", 数量=" & lQuantity
    
    baInputs = m_Modbus.ReadDiscreteInputs(lAddress, lQuantity)
    
    sResult = "离散输入值: "
    For i = 0 To UBound(baInputs)
        sResult = sResult & IIf(baInputs(i), "1", "0")
        If i < UBound(baInputs) Then sResult = sResult & ", "
    Next i
    
    LogMessage sResult
    Exit Sub
ErrorHandler:
    LogMessage "读取离散输入失败: " & Err.Description
    MsgBox "读取失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdReadHoldingRegs_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiRegs() As Integer
    Dim i As Long
    Dim sResult As String
    
    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))
    
    LogMessage "读取保持寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    aiRegs = m_Modbus.ReadHoldingRegisters(lAddress, lQuantity)
    
    sResult = "保持寄存器值: "
    For i = 0 To UBound(aiRegs)
        sResult = sResult & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
    Next i
    
    LogMessage sResult
    Exit Sub
ErrorHandler:
    LogMessage "读取保持寄存器失败: " & Err.Description
    MsgBox "读取失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdReadInputRegs_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiRegs() As Integer
    Dim i As Long
    Dim sResult As String
    
    lAddress = CLng(Val(txtReadAddress.Text))
    lQuantity = CLng(Val(txtReadQuantity.Text))
    
    LogMessage "读取输入寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    aiRegs = m_Modbus.ReadInputRegisters(lAddress, lQuantity)
    
    sResult = "输入寄存器值: "
    For i = 0 To UBound(aiRegs)
        sResult = sResult & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
    Next i
    
    LogMessage sResult
    Exit Sub
ErrorHandler:
    LogMessage "读取输入寄存器失败: " & Err.Description
    MsgBox "读取失败: " & Err.Description, vbCritical
End Sub

'=========================================================================
' Write Operations
'=========================================================================

Private Sub cmdWriteSingleCoil_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim bValue As Boolean
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    bValue = (Val(txtWriteValue.Text) <> 0)
    
    LogMessage "写入单个线圈: 地址=" & lAddress & ", 值=" & IIf(bValue, "1", "0")
    
    bResult = m_Modbus.WriteSingleCoil(lAddress, bValue)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写入线圈失败: " & Err.Description
    MsgBox "写入失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdWriteSingleRegister_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim iValue As Integer
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    iValue = CInt(Val(txtWriteValue.Text))
    
    LogMessage "写入单个寄存器: 地址=" & lAddress & ", 值=" & iValue
    
    bResult = m_Modbus.WriteSingleRegister(lAddress, iValue)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写入寄存器失败: " & Err.Description
    MsgBox "写入失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdWriteMultipleCoils_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baValues() As Boolean
    Dim i As Long
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    lQuantity = CLng(Val(txtWriteValue.Text))
    
    If lQuantity < 1 Or lQuantity > m_Modbus.Defaults.MAX_COILS Then
        MsgBox "数量必须在 1-" & m_Modbus.Defaults.MAX_COILS & " 之间", vbExclamation
        Exit Sub
    End If
    
    ' 创建测试数据（交替 0/1）
    ReDim baValues(lQuantity - 1) As Boolean
    For i = 0 To lQuantity - 1
        baValues(i) = ((i Mod 2) = 0)
    Next i
    
    LogMessage "写入多个线圈: 地址=" & lAddress & ", 数量=" & lQuantity
    
    bResult = m_Modbus.WriteMultipleCoils(lAddress, baValues)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写入多个线圈失败: " & Err.Description
    MsgBox "写入失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdWriteMultipleRegisters_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiValues() As Integer
    Dim i As Long
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    lQuantity = CLng(Val(txtWriteValue.Text))
    
    If lQuantity < 1 Or lQuantity > m_Modbus.Defaults.MAX_REGISTERS Then
        MsgBox "数量必须在 1-" & m_Modbus.Defaults.MAX_REGISTERS & " 之间", vbExclamation
        Exit Sub
    End If
    
    ' 创建测试数据（递增序列）
    ReDim aiValues(lQuantity - 1) As Integer
    For i = 0 To lQuantity - 1
        aiValues(i) = i + 100
    Next i
    
    LogMessage "写入多个寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    bResult = m_Modbus.WriteMultipleRegisters(lAddress, aiValues)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写入多个寄存器失败: " & Err.Description
    MsgBox "写入失败: " & Err.Description, vbCritical
End Sub

'=========================================================================
' Modbus Events
'=========================================================================

Private Sub m_Modbus_OnConnect()
    LogMessage "连接成功"
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
End Sub

Private Sub m_Modbus_OnDisconnect()
    LogMessage "连接已断开"
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub m_Modbus_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

Private Sub m_Modbus_OnDataReceived(Data() As Byte)
    ' 调试用，显示接收到的原始数据
    Dim sHex As String
    Dim i As Long
    If UBound(Data) >= 0 Then
        sHex = "接收数据: "
        For i = 0 To UBound(Data)
            sHex = sHex & Right$("0" & Hex$(Data(i)), 2) & " "
        Next i
        LogMessage sHex
    End If
End Sub

'=========================================================================
' Helper Functions
'=========================================================================

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    ' 自动滚动到底部
    txtLog.SelStart = Len(txtLog.Text)
End Sub
