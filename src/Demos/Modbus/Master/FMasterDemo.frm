VERSION 5.00
Begin VB.Form FMasterDemo
   Caption         =   "Modbus 主站演示程序 (Master)"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9735
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSlaveData
      Caption         =   "从站数据状态"
      Height          =   1935
      Left            =   5040
      TabIndex        =   30
      Top             =   5400
      Width           =   4575
      Begin VB.TextBox txtCoilStatus
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdRefreshSlaveData
         Caption         =   "刷新从站数据"
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtRegStatus
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblCoilStatus
         Caption         =   "线圈状态:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRegStatus
         Caption         =   "寄存器状态:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraLog
      Caption         =   "操作日志"
      Height          =   1815
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   4815
      Begin VB.TextBox txtLog
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraWrite
      Caption         =   "写入操作"
      Height          =   1935
      Left            =   5040
      TabIndex        =   17
      Top             =   3360
      Width           =   4575
      Begin VB.CommandButton cmdWriteMultipleRegs
         Caption         =   "写多个寄存器"
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteSingleReg
         Caption         =   "写单个寄存器"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteMultipleCoils
         Caption         =   "写多个线圈"
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdWriteSingleCoil
         Caption         =   "写单个线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtWriteValue
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Text            =   "1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtWriteAddress
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblWriteValue
         Caption         =   "值:"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblWriteAddress
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraRead
      Caption         =   "读取操作"
      Height          =   1935
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
      Begin VB.CommandButton cmdReadInputRegs
         Caption         =   "读输入寄存器"
         Height          =   375
         Left            =   2520
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
         Left            =   2520
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdReadCoils
         Caption         =   "读线圈"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtReadQuantity
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "10"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtReadAddress
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblReadQuantity
         Caption         =   "数量:"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblReadAddress
         Caption         =   "地址:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraConnection
      Caption         =   "连接设置"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Frame fraRTU
         Caption         =   "RTU 设置"
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
         Begin VB.TextBox txtStopBits
            Height          =   285
            Left            =   3120
            TabIndex        =   9
            Text            =   "1"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtParity
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Text            =   "N"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtDataBits
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Text            =   "8"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtBaudRate
            Height          =   285
            Left            =   3120
            TabIndex        =   6
            Text            =   "9600"
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtSerialPort
            Height          =   285
            Left            =   3120
            TabIndex        =   5
            Text            =   "COM1"
            Top             =   -360
            Width           =   855
         End
         Begin VB.Label lblStopBits
            Caption         =   "停止位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   10
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblParity
            Caption         =   "校验位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblDataBits
            Caption         =   "数据位:"
            Height          =   255
            Left            =   2400
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblBaudRate
            Caption         =   "波特率:"
            Height          =   255
            Left            =   2400
            TabIndex        =   2
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblSerialPort
            Caption         =   "串口:"
            Height          =   255
            Left            =   2400
            TabIndex        =   35
            Top             = -360
            Width           =   615
         End
      End
      Begin VB.Frame fraTCP
         Caption         =   "TCP 设置"
         Height          =   975
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   4575
         Begin VB.TextBox txtTCPPort
            Height          =   285
            Left            =   2760
            TabIndex        =   34
            Text            =   "502"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTCPHost
            Height          =   285
            Left            =   960
            TabIndex        =   33
            Text            =   "127.0.0.1"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblTCPPort
            Caption         =   "端口:"
            Height          =   255
            Left            =   2400
            TabIndex        =   37
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblTCPHost
            Caption         =   "主机:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdDisconnect
         Caption         =   "断开连接"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   31
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdConnect
         Caption         =   "连接"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   1935
      End
      Begin VB.OptionButton optProtocol
         Caption         =   "TCP"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   29
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optProtocol
         Caption         =   "RTU"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtSlaveID
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSlaveID
         Caption         =   "从站ID:"
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblProtocol
         Caption         =   "协议类型:"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FMasterDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FMasterDemo - Modbus 主站演示程序
'
' Purpose: 演示 Modbus 主站功能,支持 TCP 和 RTU 两种模式
'          - 连接到从站
'          - 读写线圈和寄存器
'          - 查看从站数据状态
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

Private WithEvents m_Master As VBMANLIB.cModbusMaster
Attribute m_Master.VB_VarHelpID = -1

'=========================================================================
' Form Events
'=========================================================================

Private Sub Form_Load()
    Set m_Master = New VBMANLIB.cModbusMaster
    
    ' 默认 TCP 模式
    optProtocol(1).Value = True
    UpdateProtocolUI
    
    LogMessage "Modbus 主站演示程序已启动"
    LogMessage "请选择协议类型并连接到从站"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_Master Is Nothing Then
        If m_Master.State = ModbusMasterState.MB_STATE_CONNECTED Then
            m_Master.Disconnect
        End If
        Set m_Master = Nothing
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
        m_Master.ProtocolType = MB_PROTOCOL_TCP
    Else
        ' RTU 模式
        fraTCP.Visible = False
        fraRTU.Visible = True
        m_Master.ProtocolType = MB_PROTOCOL_RTU
    End If
End Sub

'=========================================================================
' Connection Management
'=========================================================================

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler
    
    ' 设置从站ID
    m_Master.SlaveID = CByte(Val(txtSlaveID.Text))
    m_Master.ResponseTimeout = 3000
    
    If optProtocol(1).Value Then
        ' TCP 模式
        m_Master.TCPHost = txtTCPHost.Text
        m_Master.TCPPort = CLng(Val(txtTCPPort.Text))
        LogMessage "正在连接 TCP: " & m_Master.TCPHost & ":" & m_Master.TCPPort
        m_Master.Connect
    Else
        ' RTU 模式
        m_Master.SerialPort = txtSerialPort.Text
        m_Master.BaudRate = CLng(Val(txtBaudRate.Text))
        m_Master.DataBits = CLng(Val(txtDataBits.Text))
        m_Master.Parity = txtParity.Text
        m_Master.StopBits = CLng(Val(txtStopBits.Text))
        LogMessage "正在连接 RTU: " & m_Master.SerialPort & " (" & m_Master.BaudRate & ")"
        m_Master.Connect
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "连接失败: " & Err.Description
    MsgBox "连接失败: " & Err.Description, vbCritical
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    m_Master.Disconnect
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
    
    LogMessage "读线圈: 地址=" & lAddress & ", 数量=" & lQuantity
    
    baCoils = m_Master.ReadCoils(lAddress, lQuantity)
    
    sResult = ""
    For i = 0 To UBound(baCoils)
        sResult = sResult & IIf(baCoils(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i
    
    LogMessage "结果: " & sResult
    Exit Sub
ErrorHandler:
    LogMessage "读线圈失败: " & Err.Description
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
    
    LogMessage "读离散输入: 地址=" & lAddress & ", 数量=" & lQuantity
    
    baInputs = m_Master.ReadDiscreteInputs(lAddress, lQuantity)
    
    sResult = ""
    For i = 0 To UBound(baInputs)
        sResult = sResult & IIf(baInputs(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i
    
    LogMessage "结果: " & sResult
    Exit Sub
ErrorHandler:
    LogMessage "读离散输入失败: " & Err.Description
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
    
    LogMessage "读保持寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    aiRegs = m_Master.ReadHoldingRegisters(lAddress, lQuantity)
    
    sResult = ""
    For i = 0 To UBound(aiRegs)
        sResult = sResult & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i
    
    LogMessage "结果: " & vbCrLf & sResult
    Exit Sub
ErrorHandler:
    LogMessage "读保持寄存器失败: " & Err.Description
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
    
    LogMessage "读输入寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    aiRegs = m_Master.ReadInputRegisters(lAddress, lQuantity)
    
    sResult = ""
    For i = 0 To UBound(aiRegs)
        sResult = sResult & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i
    
    LogMessage "结果: " & vbCrLf & sResult
    Exit Sub
ErrorHandler:
    LogMessage "读输入寄存器失败: " & Err.Description
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
    
    LogMessage "写单个线圈: 地址=" & lAddress & ", 值=" & IIf(bValue, "1", "0")
    
    bResult = m_Master.WriteSingleCoil(lAddress, bValue)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写单个线圈失败: " & Err.Description
End Sub

Private Sub cmdWriteSingleReg_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim iValue As Integer
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    iValue = CInt(Val(txtWriteValue.Text))
    
    LogMessage "写单个寄存器: 地址=" & lAddress & ", 值=" & iValue
    
    bResult = m_Master.WriteSingleRegister(lAddress, iValue)
    
    If bResult Then
        LogMessage "写入成功"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写单个寄存器失败: " & Err.Description
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
    
    If lQuantity < 1 Or lQuantity > m_Master.Defaults.MAX_COILS Then
        MsgBox "数量必须在 1-" & m_Master.Defaults.MAX_COILS & " 之间", vbExclamation
        Exit Sub
    End If
    
    ReDim baValues(lQuantity - 1) As Boolean
    For i = 0 To lQuantity - 1
        baValues(i) = ((i Mod 2) = 0)
    Next i
    
    LogMessage "写多个线圈: 地址=" & lAddress & ", 数量=" & lQuantity
    
    bResult = m_Master.WriteMultipleCoils(lAddress, baValues)
    
    If bResult Then
        LogMessage "写入成功 (模式: 交替 0/1)"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写多个线圈失败: " & Err.Description
End Sub

Private Sub cmdWriteMultipleRegs_Click()
    On Error GoTo ErrorHandler
    
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim aiValues() As Integer
    Dim i As Long
    Dim bResult As Boolean
    
    lAddress = CLng(Val(txtWriteAddress.Text))
    lQuantity = CLng(Val(txtWriteValue.Text))
    
    If lQuantity < 1 Or lQuantity > m_Master.Defaults.MAX_REGISTERS Then
        MsgBox "数量必须在 1-" & m_Master.Defaults.MAX_REGISTERS & " 之间", vbExclamation
        Exit Sub
    End If
    
    ReDim aiValues(lQuantity - 1) As Integer
    For i = 0 To lQuantity - 1
        aiValues(i) = i + 100
    Next i
    
    LogMessage "写多个寄存器: 地址=" & lAddress & ", 数量=" & lQuantity
    
    bResult = m_Master.WriteMultipleRegisters(lAddress, aiValues)
    
    If bResult Then
        LogMessage "写入成功 (值: 100, 101, ...)"
    Else
        LogMessage "写入失败"
    End If
    
    Exit Sub
ErrorHandler:
    LogMessage "写多个寄存器失败: " & Err.Description
End Sub

'=========================================================================
' Refresh Slave Data
'=========================================================================

Private Sub cmdRefreshSlaveData_Click()
    On Error GoTo ErrorHandler
    
    ' 读取线圈状态
    Dim baCoils() As Boolean
    Dim i As Long
    Dim sCoilStr As String
    
    baCoils = m_Master.ReadCoils(0, 16)
    sCoilStr = "线圈 [0-15]: "
    For i = 0 To UBound(baCoils)
        sCoilStr = sCoilStr & IIf(baCoils(i), "1", "0")
        If (i + 1) Mod 8 = 0 And i < 15 Then sCoilStr = sCoilStr & vbCrLf & "线圈 [" & (i + 1) & "-" & (i + 8) & "]: "
    Next i
    txtCoilStatus.Text = sCoilStr
    
    ' 读取寄存器状态
    Dim aiRegs() As Integer
    Dim sRegStr As String
    
    aiRegs = m_Master.ReadHoldingRegisters(0, 10)
    sRegStr = "寄存器 [0-9]: "
    For i = 0 To UBound(aiRegs)
        sRegStr = sRegStr & aiRegs(i)
        If i < UBound(aiRegs) Then sRegStr = sRegStr & ", "
    Next i
    txtRegStatus.Text = sRegStr
    
    LogMessage "从站数据已刷新"
    
    Exit Sub
ErrorHandler:
    LogMessage "刷新从站数据失败: " & Err.Description
End Sub

'=========================================================================
' Modbus Events
'=========================================================================

Private Sub m_Master_OnConnect()
    LogMessage "连接成功"
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    LogMessage "可以开始读写操作"
End Sub

Private Sub m_Master_OnDisconnect()
    LogMessage "连接已断开"
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub m_Master_OnError(ByVal Description As String)
    LogMessage "错误: " & Description
End Sub

Private Sub m_Master_OnDataReceived(Data() As Byte)
    ' 调试用
    Dim sHex As String
    Dim i As Long
    If UBound(Data) >= 0 Then
        sHex = "接收: "
        For i = 0 To UBound(Data)
            sHex = sHex & Right$("0" & Hex$(Data(i)), 2) & " "
        Next i
        ' LogMessage sHex
    End If
End Sub

'=========================================================================
' Helper Functions
'=========================================================================

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
