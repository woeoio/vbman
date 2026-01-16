VERSION 5.00
Begin VB.Form FMasterEx
   Caption         =   "Modbus Master Demo (TCP + RTU)"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10320
   StartUpPosition =   2
   Begin VB.Frame fraProtocol
      Caption         =   "Protocol Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cmbProtocol
         Height          =   315
         ItemData        =   "FMasterEx.frx":0000
         Left            =   1560
         List            =   "FMasterEx.frx":000C
         Style           =   2
         TabIndex        =   0
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
      Begin VB.TextBox txtSlaveID
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Text            =   "1"
         Top             =   360
         Width           =   1440
      End
      Begin VB.TextBox txtPort
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Text            =   "502"
         Top             =   720
         Width           =   1440
      End
      Begin VB.TextBox txtHost
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblSlaveID
         Caption         =   "Slave ID:"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblPort
         Caption         =   "Port:"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblHost
         Caption         =   "Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraRTU
      Caption         =   "RTU Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Visible         =   0
      Begin VB.ComboBox cmbParity
         Height          =   315
         ItemData        =   "FMasterEx.frx":0024
         Left            =   3600
         List            =   "FMasterEx.frx":0030
         Style           =   2
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbStopBits
         Height          =   315
         ItemData        =   "FMasterEx.frx":0046
         Left            =   1200
         List            =   "FMasterEx.frx":0052
         Style           =   2
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmbBaudRate
         Height          =   315
         ItemData        =   "FMasterEx.frx":0068
         Left            =   1200
         List            =   "FMasterEx.frx":0088
         Style           =   2
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtRTUSlaveID
         Height          =   285
         Left            =   3600
         TabIndex        =   11
         Text            =   "1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtSerialPort
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "COM1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblParity
         Caption         =   "Parity:"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblStopBits
         Caption         =   "StopBits:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBaudRate
         Caption         =   "BaudRate:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblRTUSlaveID
         Caption         =   "Slave ID:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSerialPort
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraConnection
      Caption         =   "Connection"
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   9900
      Begin VB.CommandButton cmdDisconnect
         Caption         =   "Disconnect"
         Enabled         =   0
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnect
         Caption         =   "Connect"
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblStatus
         Caption         =   "Status: Disconnected"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame fraReadWrite
      Caption         =   "Read/Write Operations"
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   9900
      Begin VB.CommandButton cmdReadInputRegs
         Caption         =   "Read Input Regs"
         Height          =   375
         Left            =   7560
         TabIndex        =   40
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadHoldingRegs
         Caption         =   "Read Holding Regs"
         Height          =   375
         Left            =   5040
         TabIndex        =   39
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton cmdWriteMultipleRegs
         Caption         =   "Write Regs"
         Height          =   375
         Left            =   7560
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdWriteSingleReg
         Caption         =   "Write Reg"
         Height          =   375
         Left            =   5040
         TabIndex        =   37
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdReadDiscreteInputs
         Caption         =   "Read Discrete Inputs"
         Height          =   375
         Left            =   7560
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadCoils
         Caption         =   "Read Coils"
         Height          =   375
         Left            =   5040
         TabIndex        =   35
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdWriteMultipleCoils
         Caption         =   "Write Coils"
         Height          =   375
         Left            =   2520
         TabIndex        =   34
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdWriteSingleCoil
         Caption         =   "Write Coil"
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtOperationValue
         Height          =   285
         Left            =   7560
         TabIndex        =   31
         Text            =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtOperationQuantity
         Height          =   285
         Left            =   5280
         TabIndex        =   29
         Text            =   "10"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtOperationAddress
         Height          =   285
         Left            =   2520
         TabIndex        =   27
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblOperationValue
         Caption         =   "Value:"
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   280
         Width           =   735
      End
      Begin VB.Label lblOperationQuantity
         Caption         =   "Qty:"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   280
         Width           =   735
      End
      Begin VB.Label lblOperationAddress
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   280
         Width           =   2295
      End
   End
   Begin VB.Frame fraResult
      Caption         =   "Result"
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   6120
      Width           =   9900
      Begin VB.TextBox txtResult
         Height          =   1215
         Left            =   120
         MultiLine       =   -1
         ScrollBars      =   3
         TabIndex        =   42
         Top             =   240
         Width           =   9660
      End
   End
   Begin VB.Frame fraLog
      Caption         =   "Communication Log"
      Height          =   1305
      Left            =   120
      TabIndex        =   43
      Top             =   7800
      Width           =   9900
      Begin VB.CommandButton cmdClearLog
         Caption         =   "Clear"
         Height          =   255
         Left            =   9360
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtLog
         Height          =   960
         Left            =   120
         MultiLine       =   -1
         ScrollBars      =   3
         TabIndex        =   44
         Top             =   280
         Width           =   9215
      End
   End
End
Attribute VB_Name = "FMasterEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' FMasterEx - Modbus Master Demo (TCP + RTU)
'
' Purpose: Demonstrates cModbusMaster functionality with both TCP and RTU
'          - Protocol selection (TCP or RTU)
'          - TCP mode: Connect to slave via IP:Port
'          - RTU mode: Connect to slave via Serial Port
'          - Read/Write all Modbus function codes
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

Private WithEvents m_Master As cModbusMaster
Attribute m_Master.VB_VarHelpID = -1

Private Sub Form_Load()
    Set m_Master = New cModbusMaster

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

    LogMessage "Modbus Master Demo (TCP + RTU) started"
    UpdateUI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not m_Master Is Nothing Then
        If m_Master.State = MB_MASTER_STATE_CONNECTED Then
            m_Master.DisConnect
        End If
        Set m_Master = Nothing
    End If
End Sub

Private Sub cmbProtocol_Change()
    UpdateUI
End Sub

Private Sub UpdateUI()
    If cmbProtocol.ListIndex = 0 Then ' TCP
        fraTCP.Visible = True
        fraRTU.Visible = False
        m_Master.ProtocolType = MB_MASTER_PROTOCOL_TCP
    Else ' RTU
        fraTCP.Visible = False
        fraRTU.Visible = True
        m_Master.ProtocolType = MB_MASTER_PROTOCOL_RTU
    End If
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler

    If cmbProtocol.ListIndex = 0 Then ' TCP
        m_Master.TCPHost = txtHost.Text
        m_Master.TCPPort = CLng(Val(txtPort.Text))
        m_Master.SlaveID = CByte(Val(txtSlaveID.Text))
        LogMessage "Connecting to " & m_Master.TCPHost & ":" & m_Master.TCPPort & "..."
    Else ' RTU
        m_Master.SerialPort = txtSerialPort.Text
        m_Master.BaudRate = CLng(Val(cmbBaudRate.Text))
        m_Master.StopBits = CLng(Val(cmbStopBits.Text))
        m_Master.Parity = cmbParity.Text
        m_Master.SlaveID = CByte(Val(txtRTUSlaveID.Text))
        LogMessage "Connecting to " & m_Master.SerialPort & " (RTU)..."
    End If

    m_Master.ResponseTimeout = 3000
    m_Master.Connect

    Exit Sub
ErrorHandler:
    LogMessage "Connect failed: " & Err.Description
    MsgBox "Connect failed: " & Err.Description, vbCritical
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    m_Master.DisConnect
    LogMessage "Disconnected"
End Sub

'=========================================================================
' Read Operations
'=========================================================================

Private Sub cmdReadCoils_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, baCoils() As Boolean, i As Long, sResult As String

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    LogMessage "Reading Coils: Address=" & lAddress & ", Qty=" & lQuantity
    baCoils = m_Master.ReadCoils(lAddress, lQuantity)

    sResult = "Coils [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(baCoils)
        sResult = sResult & IIf(baCoils(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i

    txtResult.Text = sResult
    LogMessage "Read successful"
    Exit Sub
ErrorHandler:
    LogMessage "Read Coils failed: " & Err.Description
End Sub

Private Sub cmdReadDiscreteInputs_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, baInputs() As Boolean, i As Long, sResult As String

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    LogMessage "Reading Discrete Inputs: Address=" & lAddress & ", Qty=" & lQuantity
    baInputs = m_Master.ReadDiscreteInputs(lAddress, lQuantity)

    sResult = "Discrete Inputs [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(baInputs)
        sResult = sResult & IIf(baInputs(i), "1", "0")
        If (i + 1) Mod 8 = 0 Then sResult = sResult & " "
    Next i

    txtResult.Text = sResult
    LogMessage "Read successful"
    Exit Sub
ErrorHandler:
    LogMessage "Read Discrete Inputs failed: " & Err.Description
End Sub

Private Sub cmdReadHoldingRegs_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, aiRegs() As Integer, i As Long, sResult As String

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    LogMessage "Reading Holding Registers: Address=" & lAddress & ", Qty=" & lQuantity
    aiRegs = m_Master.ReadHoldingRegisters(lAddress, lQuantity)

    sResult = "Holding Registers [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(aiRegs)
        sResult = sResult & "[" & (lAddress + i) & "]=" & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i

    txtResult.Text = sResult
    LogMessage "Read successful"
    Exit Sub
ErrorHandler:
    LogMessage "Read Holding Registers failed: " & Err.Description
End Sub

Private Sub cmdReadInputRegs_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, aiRegs() As Integer, i As Long, sResult As String

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    LogMessage "Reading Input Registers: Address=" & lAddress & ", Qty=" & lQuantity
    aiRegs = m_Master.ReadInputRegisters(lAddress, lQuantity)

    sResult = "Input Registers [" & lAddress & "-" & (lAddress + lQuantity - 1) & "]:" & vbCrLf
    For i = 0 To UBound(aiRegs)
        sResult = sResult & "[" & (lAddress + i) & "]=" & aiRegs(i)
        If i < UBound(aiRegs) Then sResult = sResult & ", "
        If (i + 1) Mod 5 = 0 Then sResult = sResult & vbCrLf
    Next i

    txtResult.Text = sResult
    LogMessage "Read successful"
    Exit Sub
ErrorHandler:
    LogMessage "Read Input Registers failed: " & Err.Description
End Sub

'=========================================================================
' Write Operations
'=========================================================================

Private Sub cmdWriteSingleCoil_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, bValue As Boolean, bResult As Boolean

    lAddress = CLng(Val(txtOperationAddress.Text))
    bValue = (Val(txtOperationValue.Text) <> 0)

    LogMessage "Writing Single Coil: Address=" & lAddress & ", Value=" & IIf(bValue, "ON", "OFF")
    bResult = m_Master.WriteSingleCoil(lAddress, bValue)

    If bResult Then
        LogMessage "Write successful"
        txtResult.Text = "Write Single Coil successful" & vbCrLf & "Address: " & lAddress & vbCrLf & "Value: " & IIf(bValue, "ON (1)", "OFF (0)")
    End If
    Exit Sub
ErrorHandler:
    LogMessage "Write Single Coil failed: " & Err.Description
End Sub

Private Sub cmdWriteMultipleCoils_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, baValues() As Boolean, i As Long, bResult As Boolean

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    If lQuantity < 1 Or lQuantity > 100 Then
        MsgBox "Quantity should be between 1-100", vbExclamation
        Exit Sub
    End If

    ReDim baValues(lQuantity - 1) As Boolean
    For i = 0 To lQuantity - 1
        baValues(i) = ((i Mod 2) = 0)
    Next i

    LogMessage "Writing Multiple Coils: Address=" & lAddress & ", Qty=" & lQuantity
    bResult = m_Master.WriteMultipleCoils(lAddress, baValues)

    If bResult Then
        LogMessage "Write successful (alternating pattern)"
        txtResult.Text = "Write Multiple Coils successful" & vbCrLf & "Address: " & lAddress & vbCrLf & "Quantity: " & lQuantity
    End If
    Exit Sub
ErrorHandler:
    LogMessage "Write Multiple Coils failed: " & Err.Description
End Sub

Private Sub cmdWriteSingleReg_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, iValue As Integer, bResult As Boolean

    lAddress = CLng(Val(txtOperationAddress.Text))
    iValue = CInt(Val(txtOperationValue.Text))

    LogMessage "Writing Single Register: Address=" & lAddress & ", Value=" & iValue
    bResult = m_Master.WriteSingleRegister(lAddress, iValue)

    If bResult Then
        LogMessage "Write successful"
        txtResult.Text = "Write Single Register successful" & vbCrLf & "Address: " & lAddress & vbCrLf & "Value: " & iValue
    End If
    Exit Sub
ErrorHandler:
    LogMessage "Write Single Register failed: " & Err.Description
End Sub

Private Sub cmdWriteMultipleRegs_Click()
    On Error GoTo ErrorHandler
    Dim lAddress As Long, lQuantity As Long, aiValues() As Integer, i As Long, bResult As Boolean

    lAddress = CLng(Val(txtOperationAddress.Text))
    lQuantity = CLng(Val(txtOperationQuantity.Text))

    If lQuantity < 1 Or lQuantity > 100 Then
        MsgBox "Quantity should be between 1-100", vbExclamation
        Exit Sub
    End If

    ReDim aiValues(lQuantity - 1) As Integer
    For i = 0 To lQuantity - 1
        aiValues(i) = 100 + i
    Next i

    LogMessage "Writing Multiple Registers: Address=" & lAddress & ", Qty=" & lQuantity
    bResult = m_Master.WriteMultipleRegisters(lAddress, aiValues)

    If bResult Then
        LogMessage "Write successful (values: 100, 101, 102...)"
        txtResult.Text = "Write Multiple Registers successful" & vbCrLf & "Address: " & lAddress & vbCrLf & "Quantity: " & lQuantity
    End If
    Exit Sub
ErrorHandler:
    LogMessage "Write Multiple Registers failed: " & Err.Description
End Sub

'=========================================================================
' Master Events
'=========================================================================

Private Sub m_Master_OnConnect()
    LogMessage "Connected successfully"
    lblStatus.Caption = "Status: Connected"
    lblStatus.ForeColor = &HC000&
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmbProtocol.Enabled = False
End Sub

Private Sub m_Master_OnDisconnect()
    LogMessage "Connection closed"
    lblStatus.Caption = "Status: Disconnected"
    lblStatus.ForeColor = &HFF&
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    cmbProtocol.Enabled = True
End Sub

Private Sub m_Master_OnError(ByVal Description As String)
    LogMessage "Error: " & Description
End Sub

Private Sub m_Master_OnDataReceived(Data() As Byte)
    ' Optional: Log raw data
End Sub

'=========================================================================
' Utilities
'=========================================================================

Private Sub cmdClearLog_Click()
    txtLog.Text = ""
End Sub

Private Sub LogMessage(ByVal Message As String)
    Dim sTime As String
    sTime = Format$(Now, "hh:mm:ss")
    txtLog.Text = txtLog.Text & "[" & sTime & "] " & Message & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
