VERSION 5.00
Begin VB.Form FSlaveEx 
   Caption         =   "Modbus Slave Demo (TCP + RTU)"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10320
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame fraProtocol 
      Caption         =   "Protocol Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cmbProtocol 
         Height          =   315
         ItemData        =   "FSlaveEx.frx":0000
         Left            =   1560
         List            =   "FSlaveEx.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblProtocol 
         Caption         =   "Protocol:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   280
         Width           =   1095
      End
   End
   Begin VB.Frame fraTCP 
      Caption         =   "TCP Settings"
      Height          =   1695
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   4920
      Begin VB.TextBox txtTCPSlaveID 
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   1440
      End
      Begin VB.TextBox txtTCPPort 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Text            =   "502"
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label lblTCPSlaveID 
         Caption         =   "Slave ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblTCPPort 
         Caption         =   "Port:"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraRTU 
      Caption         =   "RTU Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox cmbParity 
         Height          =   315
         ItemData        =   "FSlaveEx.frx":0018
         Left            =   3600
         List            =   "FSlaveEx.frx":001A
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbStopBits 
         Height          =   315
         ItemData        =   "FSlaveEx.frx":001C
         Left            =   1200
         List            =   "FSlaveEx.frx":001E
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmbBaudRate 
         Height          =   315
         ItemData        =   "FSlaveEx.frx":0020
         Left            =   1200
         List            =   "FSlaveEx.frx":0022
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtRTUSlaveID 
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Text            =   "1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtSerialPort 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "COM1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblParity 
         Caption         =   "Parity:"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblStopBits 
         Caption         =   "StopBits:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBaudRate 
         Caption         =   "BaudRate:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblRTUSlaveID 
         Caption         =   "Slave ID:"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSerialPort 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server Control"
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   9900
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Server"
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status: Stopped"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame fraSetData 
      Caption         =   "Set Data Values"
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   9900
      Begin VB.CommandButton cmdSetInputReg 
         Caption         =   "Set Input Reg"
         Height          =   375
         Left            =   7560
         TabIndex        =   35
         Top             =   1000
         Width           =   2175
      End
      Begin VB.CommandButton cmdSetHoldingReg 
         Caption         =   "Set Holding Reg"
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Top             =   1000
         Width           =   2295
      End
      Begin VB.CommandButton cmdSetDiscreteInput 
         Caption         =   "Set Discrete Input"
         Height          =   375
         Left            =   7560
         TabIndex        =   33
         Top             =   520
         Width           =   2175
      End
      Begin VB.CommandButton cmdSetCoil 
         Caption         =   "Set Coil"
         Height          =   375
         Left            =   5040
         TabIndex        =   32
         Top             =   520
         Width           =   2295
      End
      Begin VB.TextBox txtDataValue 
         Height          =   285
         Left            =   7560
         TabIndex        =   30
         Text            =   "1"
         Top             =   160
         Width           =   1095
      End
      Begin VB.TextBox txtDataAddress 
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Text            =   "0"
         Top             =   160
         Width           =   1095
      End
      Begin VB.Label lblDataValue 
         Caption         =   "Value:"
         Height          =   255
         Left            =   6720
         TabIndex        =   31
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblDataAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   200
         Width           =   2295
      End
   End
   Begin VB.Frame fraCurrentData 
      Caption         =   "Current Data Values"
      Height          =   1935
      Left            =   120
      TabIndex        =   36
      Top             =   5400
      Width           =   9900
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   9360
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCurrentData 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Top             =   280
         Width           =   9215
      End
   End
   Begin VB.Frame fraLog 
      Caption         =   "Communication Log"
      Height          =   1200
      Left            =   120
      TabIndex        =   25
      Top             =   7440
      Width           =   9900
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "Clear"
         Height          =   255
         Left            =   9360
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtLog 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         Top             =   280
         Width           =   9215
      End
   End
End
Attribute VB_Name = "FSlaveEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FSlaveEx - Modbus Slave Demo (TCP + RTU)
'
' Purpose: Demonstrates cModbusSlave functionality with both TCP and RTU
'          - Protocol selection (TCP or RTU)
'          - TCP mode: Listen on port for master connections
'          - RTU mode: Listen on serial port for master connections
'          - Provide read/write access to 4 data areas
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

Private WithEvents m_Slave As cModbusSlave
Attribute m_Slave.VB_VarHelpID = -1

Private Sub Form_Load()
    Set m_Slave = New cModbusSlave

    ' Initialize protocol combo
    cmbProtocol.AddItem "TCP"
    cmbProtocol.AddItem "RTU"
    cmbProtocol.ListIndex = 0

    ' Initialize RTU settings
    cmbBaudRate.AddItem "9600"
    cmbBaudRate.AddItem "19200"
    cmbBaudRate.AddItem "38400"
    cmbBaudRate.AddItem "57600"
    cmbBaudRate.ListIndex = 0

    cmbStopBits.AddItem "1"
    cmbStopBits.AddItem "2"
    cmbStopBits.ListIndex = 0

    cmbParity.AddItem "N"
    cmbParity.AddItem "E"
    cmbParity.AddItem "O"
    cmbParity.ListIndex = 0

    ' Initialize test data
    InitializeTestData

    LogMessage "Modbus Slave Demo (TCP + RTU) started"
    UpdateUI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_Slave Is Nothing Then
        If m_Slave.State = MB_SLAVE_STATE_RUNNING Then
            m_Slave.StopMe
        End If
        Set m_Slave = Nothing
    End If
End Sub

Private Sub cmbProtocol_Change()
    UpdateUI
End Sub

Private Sub UpdateUI()
    If cmbProtocol.ListIndex = 0 Then ' TCP
        fraTCP.Visible = True
        fraRTU.Visible = False
        m_Slave.ProtocolType = MB_SLAVE_PROTOCOL_TCP
    Else ' RTU
        fraTCP.Visible = False
        fraRTU.Visible = True
        m_Slave.ProtocolType = MB_SLAVE_PROTOCOL_RTU
    End If
End Sub

Private Sub cmdStart_Click()
    On Error GoTo ErrorHandler

    If cmbProtocol.ListIndex = 0 Then ' TCP
        m_Slave.SlaveID = CByte(Val(txtTCPSlaveID.Text))
        LogMessage "Starting TCP server on port " & txtTCPPort.Text & "..."
        m_Slave.Start txtTCPPort.Text
    Else ' RTU
        m_Slave.SerialPort = txtSerialPort.Text
        m_Slave.BaudRate = CLng(Val(cmbBaudRate.Text))
        m_Slave.StopBits = CLng(Val(cmbStopBits.Text))
        m_Slave.Parity = cmbParity.Text
        m_Slave.SlaveID = CByte(Val(txtRTUSlaveID.Text))
        LogMessage "Starting RTU server on " & txtSerialPort.Text & "..."
        m_Slave.Start ""
    End If

    Exit Sub
ErrorHandler:
    LogMessage "Start failed: " & Err.Description
    MsgBox "Start failed: " & Err.Description, vbCritical
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    m_Slave.Stop
    LogMessage "Server stopped"
End Sub

'=========================================================================
' Set Data Methods
'=========================================================================

Private Sub cmdSetCoil_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, bValue As Boolean

    lAddress = CLng(Val(txtDataAddress.Text))
    bValue = (Val(txtDataValue.Text) <> 0)

    m_Slave.SetCoil lAddress, bValue
    LogMessage "Set Coil: Address=" & lAddress & ", Value=" & IIf(bValue, "ON", "OFF")
    RefreshDataDisplay
    Exit Sub
ErrorHandler:
    LogMessage "Set Coil failed: " & Err.Description
End Sub

Private Sub cmdSetDiscreteInput_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, bValue As Boolean

    lAddress = CLng(Val(txtDataAddress.Text))
    bValue = (Val(txtDataValue.Text) <> 0)

    m_Slave.SetDiscreteInput lAddress, bValue
    LogMessage "Set Discrete Input: Address=" & lAddress & ", Value=" & IIf(bValue, "ON", "OFF")
    RefreshDataDisplay
    Exit Sub
ErrorHandler:
    LogMessage "Set Discrete Input failed: " & Err.Description
End Sub

Private Sub cmdSetHoldingReg_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, iValue As Integer

    lAddress = CLng(Val(txtDataAddress.Text))
    iValue = CInt(Val(txtDataValue.Text))

    m_Slave.SetHoldingRegister lAddress, iValue
    LogMessage "Set Holding Register: Address=" & lAddress & ", Value=" & iValue
    RefreshDataDisplay
    Exit Sub
ErrorHandler:
    LogMessage "Set Holding Register failed: " & Err.Description
End Sub

Private Sub cmdSetInputReg_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, iValue As Integer

    lAddress = CLng(Val(txtDataAddress.Text))
    iValue = CInt(Val(txtDataValue.Text))

    m_Slave.SetInputRegister lAddress, iValue
    LogMessage "Set Input Register: Address=" & lAddress & ", Value=" & iValue
    RefreshDataDisplay
    Exit Sub
ErrorHandler:
    LogMessage "Set Input Register failed: " & Err.Description
End Sub

'=========================================================================
' Slave Events
'=========================================================================

Private Sub m_Slave_OnStart()
    LogMessage "Server started and listening"
    lblStatus.Caption = "Status: Running"
    lblStatus.ForeColor = &HC000&
    cmdStart.Enabled = False
    cmdStop.Enabled = True
    cmbProtocol.Enabled = False
End Sub

Private Sub m_Slave_OnStop()
    LogMessage "Server stopped"
    lblStatus.Caption = "Status: Stopped"
    lblStatus.ForeColor = &HFF&
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    cmbProtocol.Enabled = True
End Sub

Private Sub m_Slave_OnError(ByVal Description As String)
    LogMessage "Error: " & Description
End Sub

Private Sub m_Slave_OnDataReceived(Data() As Byte)
    ' Optional: Log raw data
End Sub

Private Sub m_Slave_OnCoilsChanged(ByVal StartAddress As Long, ByVal Quantity As Long, Values() As Boolean)
    Dim sMsg As String, i As Long
    sMsg = "Coils changed: Address=" & StartAddress & ", Qty=" & Quantity & ", Values: "
    For i = 0 To UBound(Values)
        sMsg = sMsg & IIf(Values(i), "1", "0")
    Next i
    LogMessage sMsg
    RefreshDataDisplay
End Sub

Private Sub m_Slave_OnRegistersChanged(ByVal StartAddress As Long, ByVal Quantity As Long, Values() As Integer)
    Dim sMsg As String, i As Long
    sMsg = "Registers changed: Address=" & StartAddress & ", Qty=" & Quantity
    LogMessage sMsg
    RefreshDataDisplay
End Sub

'=========================================================================
' Utilities
'=========================================================================

Private Sub InitializeTestData()
    Dim i As Long

    ' Initialize coils (0-15): alternating pattern
    For i = 0 To 15
        m_Slave.SetCoil i, ((i Mod 2) = 0)
    Next i

    ' Initialize discrete inputs (0-15): every 3rd is 1
    For i = 0 To 15
        m_Slave.SetDiscreteInput i, ((i Mod 3) = 0)
    Next i

    ' Initialize holding registers (0-99): value = address * 10
    For i = 0 To 99
        m_Slave.SetHoldingRegister i, i * 10
    Next i

    ' Initialize input registers (0-99): value = 1000 + address
    For i = 0 To 99
        m_Slave.SetInputRegister i, 1000 + i
    Next i

    LogMessage "Test data initialized"
    RefreshDataDisplay
End Sub

Private Sub cmdRefresh_Click()
    RefreshDataDisplay
End Sub

Private Sub RefreshDataDisplay()
    On Error Resume Next

    Dim sDisplay As String, i As Long

    sDisplay = "=== Current Data Status ===" & vbCrLf & vbCrLf

    ' Show coils
    sDisplay = sDisplay & "Coils [0-15]: "
    For i = 0 To 15
        sDisplay = sDisplay & IIf(m_Slave.GetCoil(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sDisplay = sDisplay & " "
    Next i
    sDisplay = sDisplay & vbCrLf

    ' Show discrete inputs
    sDisplay = sDisplay & "Discrete Inputs [0-15]: "
    For i = 0 To 15
        sDisplay = sDisplay & IIf(m_Slave.GetDiscreteInput(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sDisplay = sDisplay & " "
    Next i
    sDisplay = sDisplay & vbCrLf & vbCrLf

    ' Show holding registers
    sDisplay = sDisplay & "Holding Registers [0-9]: "
    For i = 0 To 9
        sDisplay = sDisplay & m_Slave.GetHoldingRegister(i)
        If i < 9 Then sDisplay = sDisplay & ", "
    Next i
    sDisplay = sDisplay & vbCrLf

    ' Show input registers
    sDisplay = sDisplay & "Input Registers [0-9]: "
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
