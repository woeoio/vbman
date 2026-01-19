VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModbusSlave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cModbusSlave - Modbus Slave/Server Implementation for VB6
'
' Purpose: Provides Modbus RTU and TCP slave/server functionality
'          - Read/Write Coils (0x01, 0x02, 0x05, 0x0F)
'          - Read/Write Registers (0x03, 0x04, 0x06, 0x10)
'          - Support for RTU (Serial) and TCP modes
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

'=========================================================================
' Modbus Function Codes
'=========================================================================

Public Enum ModbusSlaveFunctionCode
    ' Read Functions
    MB_SLAVE_FC_READ_COILS = &H1                                                  ' Read Coils (01)
    MB_SLAVE_FC_READ_DISCRETE_INPUTS = &H2                                        ' Read Discrete Inputs (02)
    MB_SLAVE_FC_READ_HOLDING_REGISTERS = &H3                                      ' Read Holding Registers (03)
    MB_SLAVE_FC_READ_INPUT_REGISTERS = &H4                                        ' Read Input Registers (04)

    ' Write Functions
    MB_SLAVE_FC_WRITE_SINGLE_COIL = &H5                                           ' Write Single Coil (05)
    MB_SLAVE_FC_WRITE_SINGLE_REGISTER = &H6                                       ' Write Single Register (06)
    MB_SLAVE_FC_WRITE_MULTIPLE_COILS = &HF                                        ' Write Multiple Coils (15)
    MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS = &H10                                   ' Write Multiple Registers (16)

    ' Advanced Functions (optional)
    MB_SLAVE_FC_MASK_WRITE_REGISTER = &H16                                        ' Mask Write Register (22)
    MB_SLAVE_FC_READ_WRITE_MULTIPLE_REGISTERS = &H17                              ' Read/Write Multiple (23)
End Enum

'=========================================================================
' Modbus Exception Codes
'=========================================================================

Public Enum ModbusSlaveExceptionCode
    MB_SLAVE_EX_ILLEGAL_FUNCTION = &H1                                            ' Illegal Function
    MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS = &H2                                        ' Illegal Data Address
    MB_SLAVE_EX_ILLEGAL_DATA_VALUE = &H3                                          ' Illegal Data Value
    MB_SLAVE_EX_SLAVE_DEVICE_FAILURE = &H4                                        ' Slave Device Failure
    MB_SLAVE_EX_ACKNOWLEDGE = &H5                                                 ' Acknowledge
    MB_SLAVE_EX_SLAVE_DEVICE_BUSY = &H6                                           ' Slave Device Busy
    MB_SLAVE_EX_MEMORY_PARITY_ERROR = &H8                                         ' Memory Parity Error
    MB_SLAVE_EX_GATEWAY_PATH_UNAVAILABLE = &HA                                    ' Gateway Path Unavailable
    MB_SLAVE_EX_GATEWAY_TARGET_DEVICE_FAILED = &HB                                ' Gateway Target Device Failed
End Enum

'=========================================================================
' Connection State
'=========================================================================

Public Enum ModbusSlaveState
    MB_SLAVE_STATE_STOPPED = 0                                                    ' Stopped
    MB_SLAVE_STATE_STARTING = 1                                                   ' Starting
    MB_SLAVE_STATE_RUNNING = 2                                                    ' Running
    MB_SLAVE_STATE_ERROR = 3                                                      ' Error
End Enum

'=========================================================================
' Protocol Type
'=========================================================================

Public Enum ModbusSlaveProtocolType
    MB_SLAVE_PROTOCOL_RTU = 1                                                     ' RTU over Serial
    MB_SLAVE_PROTOCOL_TCP = 2                                                     ' TCP/IP
End Enum

'=========================================================================
' Default Constants Structure
'=========================================================================

Public Type ModbusSlaveDefaults
    TCP_PORT As Long
    RTU_DEFAULT_BAUDRATE As Long
    RTU_DEFAULT_DATABITS As Long
    RTU_DEFAULT_PARITY As String
    RTU_DEFAULT_STOPBITS As Long
    RTU_DEFAULT_TIMEOUT As Long
    MAX_PDU_SIZE As Long
    MAX_REGISTERS As Long
    MAX_COILS As Long
    TCP_MBAP_SIZE As Long
End Type

'=========================================================================
' Events
'=========================================================================

Event OnStart()
Event OnStop()
Event OnError(ByVal Description As String)
Event OnDataReceived(Data() As Byte)
Event OnDataSent(Data() As Byte)
Event OnCoilsChanged(ByVal StartAddress As Long, ByVal Quantity As Long, Values() As Boolean)
Event OnRegistersChanged(ByVal StartAddress As Long, ByVal Quantity As Long, Values() As Integer)

' TCP Server mode events
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
Event OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByVal Quantity As Long)
Event OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByRef Data As Variant)

'=========================================================================
' Member Variables
'=========================================================================

' Default constants
Private m_Defaults As ModbusSlaveDefaults

' Protocol type
Private m_ProtocolType As ModbusSlaveProtocolType

' Connection state
Private m_State As ModbusSlaveState

' Transport layers
Private WithEvents m_TransportRTU As cModbusTransportRTU
Attribute m_TransportRTU.VB_VarHelpID = -1
Private WithEvents m_TransportTCP As cModbusTransportTCP
Attribute m_TransportTCP.VB_VarHelpID = -1

' RTU Mode settings
Private m_SerialPort As String
Private m_BaudRate As Long
Private m_DataBits As Long
Private m_Parity As String
Private m_StopBits As Long

' TCP Mode settings
Private m_TCPPort As Long

' Slave data
Private m_SlaveID As Byte
Private m_CoilsData() As Boolean
Private m_DiscreteInputsData() As Boolean
Private m_HoldingRegistersData() As Integer
Private m_InputRegistersData() As Integer

' Database size
Private m_CoilsSize As Long
Private m_DiscreteInputsSize As Long
Private m_HoldingRegistersSize As Long
Private m_InputRegistersSize As Long

' TCP request context (for MBAP header)
Private m_CurrentClientID As String
Private m_CurrentMBAPHeader() As Byte

'=========================================================================
' Properties
'=========================================================================

Public Property Get ProtocolType() As ModbusSlaveProtocolType
    ProtocolType = m_ProtocolType
End Property

Public Property Let ProtocolType(ByVal Value As ModbusSlaveProtocolType)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change protocol type while running"
        Exit Property
    End If
    m_ProtocolType = Value
End Property

Public Property Get State() As ModbusSlaveState
    State = m_State
End Property

Public Property Get SlaveID() As Byte
    SlaveID = m_SlaveID
End Property

Public Property Let SlaveID(ByVal Value As Byte)
    m_SlaveID = Value
End Property

Public Property Get Defaults() As ModbusSlaveDefaults
    Defaults = m_Defaults
End Property

' Coils
Public Property Get CoilsSize() As Long
    CoilsSize = m_CoilsSize
End Property

Public Property Let CoilsSize(ByVal Value As Long)
    If Value > 0 Then
        m_CoilsSize = Value
        ReDim m_CoilsData(Value - 1) As Boolean
    End If
End Property

' Discrete Inputs
Public Property Get DiscreteInputsSize() As Long
    DiscreteInputsSize = m_DiscreteInputsSize
End Property

Public Property Let DiscreteInputsSize(ByVal Value As Long)
    If Value > 0 Then
        m_DiscreteInputsSize = Value
        ReDim m_DiscreteInputsData(Value - 1) As Boolean
    End If
End Property

' Holding Registers
Public Property Get HoldingRegistersSize() As Long
    HoldingRegistersSize = m_HoldingRegistersSize
End Property

Public Property Let HoldingRegistersSize(ByVal Value As Long)
    If Value > 0 Then
        m_HoldingRegistersSize = Value
        ReDim m_HoldingRegistersData(Value - 1) As Integer
    End If
End Property

' Input Registers
Public Property Get InputRegistersSize() As Long
    InputRegistersSize = m_InputRegistersSize
End Property

Public Property Let InputRegistersSize(ByVal Value As Long)
    If Value > 0 Then
        m_InputRegistersSize = Value
        ReDim m_InputRegistersData(Value - 1) As Integer
    End If
End Property

' RTU Mode Properties
Public Property Get SerialPort() As String
    SerialPort = m_SerialPort
End Property

Public Property Let SerialPort(ByVal Value As String)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change serial port while running"
        Exit Property
    End If
    m_SerialPort = Value
End Property

Public Property Get BaudRate() As Long
    BaudRate = m_BaudRate
End Property

Public Property Let BaudRate(ByVal Value As Long)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change baud rate while running"
        Exit Property
    End If
    m_BaudRate = Value
End Property

Public Property Get DataBits() As Long
    DataBits = m_DataBits
End Property

Public Property Let DataBits(ByVal Value As Long)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change data bits while running"
        Exit Property
    End If
    m_DataBits = Value
End Property

Public Property Get Parity() As String
    Parity = m_Parity
End Property

Public Property Let Parity(ByVal Value As String)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change parity while running"
        Exit Property
    End If
    m_Parity = UCase$(Left$(Value, 1))
End Property

Public Property Get StopBits() As Long
    StopBits = m_StopBits
End Property

Public Property Let StopBits(ByVal Value As Long)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change stop bits while running"
        Exit Property
    End If
    m_StopBits = Value
End Property

' TCP Mode Properties
Public Property Get TCPPort() As Long
    TCPPort = m_TCPPort
End Property

Public Property Let TCPPort(ByVal Value As Long)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change TCP port while running"
        Exit Property
    End If
    m_TCPPort = Value
End Property

'=========================================================================
' Public Methods - Connection
'=========================================================================

Public Sub Start(Optional ByVal SerialPortOrPort As String = "")
    On Error GoTo ErrorHandler
    
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Already running"
        Exit Sub
    End If
    
    ' Validate database sizes
    If m_CoilsSize = 0 And m_DiscreteInputsSize = 0 And _
        m_HoldingRegistersSize = 0 And m_InputRegistersSize = 0 Then
        RaiseError "No data areas initialized"
        Exit Sub
    End If
    
    m_State = MB_SLAVE_STATE_STARTING
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        If LenB(SerialPortOrPort) > 0 Then m_SerialPort = SerialPortOrPort
        StartRTU
    ElseIf m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        If LenB(SerialPortOrPort) > 0 Then m_TCPPort = CLng(SerialPortOrPort)
        StartTCP
    Else
        RaiseError "Invalid protocol type"
    End If
    
    Exit Sub
ErrorHandler:
    m_State = MB_SLAVE_STATE_ERROR
    RaiseError "Start failed: " & ERR.Description
End Sub

Public Sub StopMe()
    On Error Resume Next

    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        If Not m_TransportRTU Is Nothing Then
            m_TransportRTU.Close_
        End If
    ElseIf m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        If Not m_TransportTCP Is Nothing Then
            m_TransportTCP.DisConnect
        End If
    End If

    m_State = MB_SLAVE_STATE_STOPPED
    RaiseEvent OnStop
End Sub

'=========================================================================
' Public Methods - Read Data
'=========================================================================

Public Function GetCoil(ByVal Address As Long) As Boolean
    If Address >= 0 And Address < m_CoilsSize Then
        GetCoil = m_CoilsData(Address)
    Else
        RaiseError "Coil address out of range: " & Address
    End If
End Function

Public Function GetCoils(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
    Dim baResult() As Boolean
    Dim i As Long

    If StartAddress < 0 Or StartAddress + Quantity > m_CoilsSize Then
        RaiseError "Coils address range out of bounds"
        Exit Function
    End If

    ReDim baResult(Quantity - 1) As Boolean
    For i = 0 To Quantity - 1
        baResult(i) = m_CoilsData(StartAddress + i)
    Next i

    GetCoils = baResult
End Function

Public Function GetDiscreteInput(ByVal Address As Long) As Boolean
    If Address >= 0 And Address < m_DiscreteInputsSize Then
        GetDiscreteInput = m_DiscreteInputsData(Address)
    Else
        RaiseError "Discrete input address out of range: " & Address
    End If
End Function

Public Function GetDiscreteInputs(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
    Dim baResult() As Boolean
    Dim i As Long

    If StartAddress < 0 Or StartAddress + Quantity > m_DiscreteInputsSize Then
        RaiseError "Discrete inputs address range out of bounds"
        Exit Function
    End If

    ReDim baResult(Quantity - 1) As Boolean
    For i = 0 To Quantity - 1
        baResult(i) = m_DiscreteInputsData(StartAddress + i)
    Next i

    GetDiscreteInputs = baResult
End Function

Public Function GetHoldingRegister(ByVal Address As Long) As Integer
    If Address >= 0 And Address < m_HoldingRegistersSize Then
        GetHoldingRegister = m_HoldingRegistersData(Address)
    Else
        RaiseError "Holding register address out of range: " & Address
    End If
End Function

Public Function GetHoldingRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baResult() As Integer
    Dim i As Long

    If StartAddress < 0 Or StartAddress + Quantity > m_HoldingRegistersSize Then
        RaiseError "Holding registers address range out of bounds"
        Exit Function
    End If

    ReDim baResult(Quantity - 1) As Integer
    For i = 0 To Quantity - 1
        baResult(i) = m_HoldingRegistersData(StartAddress + i)
    Next i

    GetHoldingRegisters = baResult
End Function

Public Function GetInputRegister(ByVal Address As Long) As Integer
    If Address >= 0 And Address < m_InputRegistersSize Then
        GetInputRegister = m_InputRegistersData(Address)
    Else
        RaiseError "Input register address out of range: " & Address
    End If
End Function

Public Function GetInputRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baResult() As Integer
    Dim i As Long

    If StartAddress < 0 Or StartAddress + Quantity > m_InputRegistersSize Then
        RaiseError "Input registers address range out of bounds"
        Exit Function
    End If

    ReDim baResult(Quantity - 1) As Integer
    For i = 0 To Quantity - 1
        baResult(i) = m_InputRegistersData(StartAddress + i)
    Next i

    GetInputRegisters = baResult
End Function

'=========================================================================
' Public Methods - Write Data
'=========================================================================

Public Sub SetCoil(ByVal Address As Long, ByVal Value As Boolean)
    If Address >= 0 And Address < m_CoilsSize Then
        m_CoilsData(Address) = Value
    Else
        RaiseError "Coil address out of range: " & Address
    End If
End Sub

Public Sub SetCoils(ByVal StartAddress As Long, ByRef Values() As Boolean)
    Dim i As Long
    Dim lQuantity As Long
    
    lQuantity = UBound(Values) + 1
    If StartAddress < 0 Or StartAddress + lQuantity > m_CoilsSize Then
        RaiseError "Coils address range out of bounds"
        Exit Sub
    End If
    
    For i = 0 To lQuantity - 1
        m_CoilsData(StartAddress + i) = Values(i)
    Next i
    
    RaiseEvent OnCoilsChanged(StartAddress, lQuantity, Values)
End Sub

Public Sub SetDiscreteInput(ByVal Address As Long, ByVal Value As Boolean)
    If Address >= 0 And Address < m_DiscreteInputsSize Then
        m_DiscreteInputsData(Address) = Value
    Else
        RaiseError "Discrete input address out of range: " & Address
    End If
End Sub

Public Sub SetDiscreteInputs(ByVal StartAddress As Long, ByRef Values() As Boolean)
    Dim i As Long
    Dim lQuantity As Long

    lQuantity = UBound(Values) + 1
    If StartAddress < 0 Or StartAddress + lQuantity > m_DiscreteInputsSize Then
        RaiseError "Discrete inputs address range out of bounds"
        Exit Sub
    End If

    For i = 0 To lQuantity - 1
        m_DiscreteInputsData(StartAddress + i) = Values(i)
    Next i
End Sub

Public Sub SetHoldingRegister(ByVal Address As Long, ByVal Value As Integer)
    If Address >= 0 And Address < m_HoldingRegistersSize Then
        m_HoldingRegistersData(Address) = Value
    Else
        RaiseError "Holding register address out of range: " & Address
    End If
End Sub

Public Sub SetHoldingRegisters(ByVal StartAddress As Long, ByRef Values() As Integer)
    Dim i As Long
    Dim lQuantity As Long

    lQuantity = UBound(Values) + 1
    If StartAddress < 0 Or StartAddress + lQuantity > m_HoldingRegistersSize Then
        RaiseError "Holding registers address range out of bounds"
        Exit Sub
    End If

    For i = 0 To lQuantity - 1
        m_HoldingRegistersData(StartAddress + i) = Values(i)
    Next i

    RaiseEvent OnRegistersChanged(StartAddress, lQuantity, Values)
End Sub

Public Sub SetInputRegister(ByVal Address As Long, ByVal Value As Integer)
    If Address >= 0 And Address < m_InputRegistersSize Then
        m_InputRegistersData(Address) = Value
    Else
        RaiseError "Input register address out of range: " & Address
    End If
End Sub

Public Sub SetInputRegisters(ByVal StartAddress As Long, ByRef Values() As Integer)
    Dim i As Long
    Dim lQuantity As Long

    lQuantity = UBound(Values) + 1
    If StartAddress < 0 Or StartAddress + lQuantity > m_InputRegistersSize Then
        RaiseError "Input registers address range out of bounds"
        Exit Sub
    End If

    For i = 0 To lQuantity - 1
        m_InputRegistersData(StartAddress + i) = Values(i)
    Next i
End Sub

'=========================================================================
' Private Methods - Connection
'=========================================================================

Private Sub StartRTU()
    On Error GoTo ErrorHandler

    If LenB(m_SerialPort) = 0 Then
        RaiseError "Serial port not specified"
        Exit Sub
    End If

    Set m_TransportRTU = New cModbusTransportRTU
    m_TransportRTU.SerialPort = m_SerialPort
    m_TransportRTU.BaudRate = m_BaudRate
    m_TransportRTU.DataBits = m_DataBits
    m_TransportRTU.StopBits = m_StopBits

    ' Convert parity string to enum
    Select Case m_Parity
        Case "E"
            m_TransportRTU.Parity = MB_RTU_PARITY_EVEN
        Case "O"
            m_TransportRTU.Parity = MB_RTU_PARITY_ODD
        Case Else
            m_TransportRTU.Parity = MB_RTU_PARITY_NONE
    End Select

    m_TransportRTU.Open_

    Exit Sub
ErrorHandler:
    m_State = MB_SLAVE_STATE_ERROR
    RaiseError "StartRTU failed: " & ERR.Description
End Sub

Private Sub StartTCP()
    On Error GoTo ErrorHandler
    
    If m_TCPPort = 0 Then m_TCPPort = m_Defaults.TCP_PORT
    
    Set m_TransportTCP = New cModbusTransportTCP
    m_TransportTCP.Listen m_TCPPort
    
    Exit Sub
ErrorHandler:
    m_State = MB_SLAVE_STATE_ERROR
    RaiseError "StartTCP failed: " & ERR.Description
End Sub

'=========================================================================
' Private Methods - Request Processing
'=========================================================================

Private Function ProcessRequest(ByRef Request() As Byte) As Byte()
    Dim baResponse() As Byte
    Dim baPDU() As Byte
    Dim FC As Byte
    Dim lPDUOffset As Long
    Dim lMBAPSize As Long

    On Error GoTo ErrorHandler

    ' Handle TCP vs RTU protocol differences
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ' TCP: Request includes MBAP header (7 bytes) + PDU
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        
        ' Validate minimum request length (MBAP + UnitID + FC)
        If UBound(Request) < lMBAPSize + 1 Then
            RaiseError "Invalid TCP request length"
            Exit Function
        End If
        
        ' Store MBAP header for response
        ReDim m_CurrentMBAPHeader(lMBAPSize - 1) As Byte
        Modbus_CopyMemory m_CurrentMBAPHeader(0), Request(0), lMBAPSize
        
        ' Check Unit ID (Slave ID) - it's in MBAP header at index 6
        Dim bUnitID As Byte
        bUnitID = Request(6)
        If bUnitID <> m_SlaveID And bUnitID <> 0 Then
            ' Not addressed to us, return empty response
            Exit Function
        End If
        
        ' Extract PDU (skip MBAP header, PDU starts at index 7)
        ' PDU format: Function Code + Data (no Unit ID in PDU for TCP)
        Dim lPDULen As Long
        lPDULen = UBound(Request) - lMBAPSize + 1
        ReDim baPDU(lPDULen - 1) As Byte
        Modbus_CopyMemory baPDU(0), Request(lMBAPSize), lPDULen
        
        ' Get function code from PDU (first byte after MBAP)
        If UBound(baPDU) < 0 Then
            RaiseError "Invalid PDU length"
            Exit Function
        End If
        FC = baPDU(0)
        lPDUOffset = 0
    Else
        ' RTU: Request format is SlaveID + Function Code + Data + CRC
        If UBound(Request) < 2 Then
            RaiseError "Invalid RTU request length"
            Exit Function
        End If
        
        ' Check Slave ID
        If Request(0) <> m_SlaveID And Request(0) <> 0 Then
            ' Not addressed to us
            Exit Function
        End If
        
        ' PDU starts at index 1 (skip Slave ID)
        baPDU = Request
        FC = Request(1)
        lPDUOffset = 1
    End If

    ' Process request based on function code
    Select Case FC
        Case MB_SLAVE_FC_READ_COILS
            baResponse = ProcessReadCoils(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_READ_DISCRETE_INPUTS
            baResponse = ProcessReadDiscreteInputs(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_READ_HOLDING_REGISTERS
            baResponse = ProcessReadHoldingRegisters(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_READ_INPUT_REGISTERS
            baResponse = ProcessReadInputRegisters(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_WRITE_SINGLE_COIL
            baResponse = ProcessWriteSingleCoil(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_WRITE_SINGLE_REGISTER
            baResponse = ProcessWriteSingleRegister(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_WRITE_MULTIPLE_COILS
            baResponse = ProcessWriteMultipleCoils(baPDU, lPDUOffset)
        Case MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
            baResponse = ProcessWriteMultipleRegisters(baPDU, lPDUOffset)
        Case Else
            baResponse = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_FUNCTION)
    End Select

    ProcessRequest = baResponse
    Exit Function

ErrorHandler:
    RaiseError "ProcessRequest failed: " & ERR.Description
End Function

Private Function ProcessReadCoils(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim baCoils() As Boolean
    Dim FC As Byte

    On Error GoTo ErrorHandler

    ' Validate PDU length
    If UBound(PDU) < PDUOffset + 5 Then
        ProcessReadCoils = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))

    ' Raise read request event
    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnReadRequest(m_CurrentClientID, FC, lStartAddress, lQuantity)
    End If

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_COILS Then
        ProcessReadCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_CoilsSize Then
        ProcessReadCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    baCoils = GetCoils(lStartAddress, lQuantity)
    ProcessReadCoils = BuildReadCoilsResponse(lQuantity, baCoils)
    Exit Function

ErrorHandler:
    ProcessReadCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadDiscreteInputs(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim baInputs() As Boolean
    Dim FC As Byte

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 5 Then
        ProcessReadDiscreteInputs = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))

    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnReadRequest(m_CurrentClientID, FC, lStartAddress, lQuantity)
    End If

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_COILS Then
        ProcessReadDiscreteInputs = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_DiscreteInputsSize Then
        ProcessReadDiscreteInputs = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    baInputs = GetDiscreteInputs(lStartAddress, lQuantity)
    ProcessReadDiscreteInputs = BuildReadDiscreteInputsResponse(lQuantity, baInputs)
    Exit Function

ErrorHandler:
    ProcessReadDiscreteInputs = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadHoldingRegisters(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim baRegisters() As Integer
    Dim FC As Byte

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 5 Then
        ProcessReadHoldingRegisters = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))

    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnReadRequest(m_CurrentClientID, FC, lStartAddress, lQuantity)
    End If

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessReadHoldingRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_HoldingRegistersSize Then
        ProcessReadHoldingRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    baRegisters = GetHoldingRegisters(lStartAddress, lQuantity)
    ProcessReadHoldingRegisters = BuildReadHoldingRegistersResponse(lQuantity, baRegisters)
    Exit Function

ErrorHandler:
    ProcessReadHoldingRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadInputRegisters(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim baRegisters() As Integer
    Dim FC As Byte

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 5 Then
        ProcessReadInputRegisters = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))

    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnReadRequest(m_CurrentClientID, FC, lStartAddress, lQuantity)
    End If

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessReadInputRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_InputRegistersSize Then
        ProcessReadInputRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    baRegisters = GetInputRegisters(lStartAddress, lQuantity)
    ProcessReadInputRegisters = BuildReadInputRegistersResponse(lQuantity, baRegisters)
    Exit Function

ErrorHandler:
    ProcessReadInputRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteSingleCoil(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lAddress As Long
    Dim bValue As Boolean
    Dim FC As Byte
    Dim vData As Variant

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 4 Then
        ProcessWriteSingleCoil = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    bValue = (PDU(PDUOffset + 3) = &HFF)

    If lAddress >= m_CoilsSize Then
        ProcessWriteSingleCoil = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    SetCoil lAddress, bValue
    vData = bValue
    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnWriteRequest(m_CurrentClientID, FC, lAddress, vData)
    End If

    ProcessWriteSingleCoil = BuildWriteSingleCoilResponse(lAddress, bValue)
    Exit Function

ErrorHandler:
    ProcessWriteSingleCoil = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteSingleRegister(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lAddress As Long
    Dim iValue As Integer
    Dim FC As Byte
    Dim vData As Variant

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 4 Then
        ProcessWriteSingleRegister = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    iValue = CInt(PDU(PDUOffset + 3)) * 256 + CInt(PDU(PDUOffset + 4))

    If lAddress >= m_HoldingRegistersSize Then
        ProcessWriteSingleRegister = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    SetHoldingRegister lAddress, iValue
    vData = iValue
    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnWriteRequest(m_CurrentClientID, FC, lAddress, vData)
    End If

    ProcessWriteSingleRegister = BuildWriteSingleRegisterResponse(lAddress, iValue)
    Exit Function

ErrorHandler:
    ProcessWriteSingleRegister = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteMultipleCoils(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baCoils() As Boolean
    Dim i As Long
    Dim FC As Byte
    Dim vData As Variant

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 6 Then
        ProcessWriteMultipleCoils = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))
    lByteCount = CLng(PDU(PDUOffset + 5))

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_COILS Then
        ProcessWriteMultipleCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_CoilsSize Then
        ProcessWriteMultipleCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    If UBound(PDU) < PDUOffset + 6 + lByteCount - 1 Then
        ProcessWriteMultipleCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    ReDim baCoils(lQuantity - 1) As Boolean
    For i = 0 To lQuantity - 1
        Dim lByteIndex As Long, lBitIndex As Long
        lByteIndex = i \ 8
        lBitIndex = i Mod 8
        baCoils(i) = ((PDU(PDUOffset + 6 + lByteIndex) And (2 ^ lBitIndex)) <> 0)
    Next i

    SetCoils lStartAddress, baCoils
    vData = baCoils
    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnWriteRequest(m_CurrentClientID, FC, lStartAddress, vData)
    End If

    ProcessWriteMultipleCoils = BuildWriteMultipleCoilsResponse(lStartAddress, lQuantity)
    Exit Function

ErrorHandler:
    ProcessWriteMultipleCoils = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteMultipleRegisters(ByRef PDU() As Byte, ByVal PDUOffset As Long) As Byte()
    Dim lStartAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baRegisters() As Integer
    Dim i As Long
    Dim FC As Byte
    Dim vData As Variant

    On Error GoTo ErrorHandler

    If UBound(PDU) < PDUOffset + 6 Then
        ProcessWriteMultipleRegisters = BuildExceptionResponse(PDU(PDUOffset), MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    FC = PDU(PDUOffset)
    lStartAddress = CLng(PDU(PDUOffset + 1)) * 256 + CLng(PDU(PDUOffset + 2))
    lQuantity = CLng(PDU(PDUOffset + 3)) * 256 + CLng(PDU(PDUOffset + 4))
    lByteCount = CLng(PDU(PDUOffset + 5))

    If lQuantity < 1 Or lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessWriteMultipleRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    If lStartAddress + lQuantity > m_HoldingRegistersSize Then
        ProcessWriteMultipleRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    If UBound(PDU) < PDUOffset + 6 + lByteCount - 1 Then
        ProcessWriteMultipleRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_ILLEGAL_DATA_VALUE)
        Exit Function
    End If

    ReDim baRegisters(lQuantity - 1) As Integer
    For i = 0 To lQuantity - 1
        baRegisters(i) = CInt(PDU(PDUOffset + 6 + i * 2)) * 256 + CInt(PDU(PDUOffset + 7 + i * 2))
    Next i

    SetHoldingRegisters lStartAddress, baRegisters
    vData = baRegisters
    If LenB(m_CurrentClientID) > 0 Then
        RaiseEvent OnWriteRequest(m_CurrentClientID, FC, lStartAddress, vData)
    End If

    ProcessWriteMultipleRegisters = BuildWriteMultipleRegistersResponse(lStartAddress, lQuantity)
    Exit Function

ErrorHandler:
    ProcessWriteMultipleRegisters = BuildExceptionResponse(FC, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

'=========================================================================
' Private Methods - Response Building
'=========================================================================

' Helper function to wrap PDU response with MBAP header for TCP
Private Function WrapTCPResponse(ByRef PDU() As Byte) As Byte()
    Dim baResponse() As Byte
    Dim lPDULen As Long
    Dim i As Long

    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ' TCP: Add MBAP header
        ' PDU already contains Unit ID as first byte (from response building functions)
        lPDULen = UBound(PDU) + 1
        ReDim baResponse(m_Defaults.TCP_MBAP_SIZE + lPDULen - 1) As Byte
        
        ' Copy MBAP header (from current request, preserving Transaction ID)
        If UBound(m_CurrentMBAPHeader) >= 0 Then
            For i = 0 To m_Defaults.TCP_MBAP_SIZE - 1
                baResponse(i) = m_CurrentMBAPHeader(i)
            Next i
        Else
            ' Fallback: create default MBAP header
            baResponse(0) = 0
            baResponse(1) = 0
            baResponse(2) = 0
            baResponse(3) = 0
            baResponse(6) = m_SlaveID
        End If
        
        ' Update length field (Unit ID + PDU length)
        baResponse(4) = ((lPDULen) \ 256) And &HFF
        baResponse(5) = (lPDULen) And &HFF
        
        ' Copy PDU (which includes Unit ID as first byte)
        For i = 0 To lPDULen - 1
            baResponse(m_Defaults.TCP_MBAP_SIZE + i) = PDU(i)
        Next i
    Else
        ' RTU: Return PDU as-is (with CRC)
        baResponse = PDU
    End If
    
    WrapTCPResponse = baResponse
End Function

Private Function BuildWriteSingleCoilResponse(ByVal Address As Long, ByVal Value As Boolean) As Byte()
    Dim baResponse() As Byte
    
    ReDim baResponse(5) As Byte
    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_WRITE_SINGLE_COIL
    baResponse(2) = (Address \ 256) And &HFF
    baResponse(3) = Address And &HFF
    baResponse(4) = IIf(Value, &HFF, &H0)
    baResponse(5) = 0
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildWriteSingleCoilResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildWriteSingleRegisterResponse(ByVal Address As Long, ByVal Value As Integer) As Byte()
    Dim baResponse() As Byte
    
    ReDim baResponse(5) As Byte
    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_WRITE_SINGLE_REGISTER
    baResponse(2) = (Address \ 256) And &HFF
    baResponse(3) = Address And &HFF
    baResponse(4) = (Value \ 256) And &HFF
    baResponse(5) = Value And &HFF
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildWriteSingleRegisterResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildReadCoilsResponse(ByVal Quantity As Long, ByRef Values() As Boolean) As Byte()
    Dim baResponse() As Byte
    Dim lByteCount As Long
    Dim i As Long, lByteIndex As Long, lBitIndex As Long, bByte As Byte

    lByteCount = (Quantity + 7) \ 8
    ReDim baResponse(2 + lByteCount) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_READ_COILS
    baResponse(2) = lByteCount

    For i = 0 To Quantity - 1
        lByteIndex = i \ 8
        lBitIndex = i Mod 8
        If Values(i) Then bByte = bByte Or (2 ^ lBitIndex)
        If lBitIndex = 7 Or i = Quantity - 1 Then
            baResponse(3 + lByteIndex) = bByte
            bByte = 0
        End If
    Next i

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildReadCoilsResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildReadDiscreteInputsResponse(ByVal Quantity As Long, ByRef Values() As Boolean) As Byte()
    Dim baResponse() As Byte
    Dim lByteCount As Long
    Dim i As Long, lByteIndex As Long, lBitIndex As Long, bByte As Byte

    lByteCount = (Quantity + 7) \ 8
    ReDim baResponse(2 + lByteCount) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_READ_DISCRETE_INPUTS
    baResponse(2) = lByteCount

    For i = 0 To Quantity - 1
        lByteIndex = i \ 8
        lBitIndex = i Mod 8
        If Values(i) Then bByte = bByte Or (2 ^ lBitIndex)
        If lBitIndex = 7 Or i = Quantity - 1 Then
            baResponse(3 + lByteIndex) = bByte
            bByte = 0
        End If
    Next i

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildReadDiscreteInputsResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildReadHoldingRegistersResponse(ByVal Quantity As Long, ByRef Values() As Integer) As Byte()
    Dim baResponse() As Byte
    Dim i As Long

    ReDim baResponse(2 + Quantity * 2) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_READ_HOLDING_REGISTERS
    baResponse(2) = Quantity * 2

    For i = 0 To Quantity - 1
        baResponse(3 + i * 2) = (Values(i) \ 256) And &HFF
        baResponse(4 + i * 2) = Values(i) And &HFF
    Next i

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildReadHoldingRegistersResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildReadInputRegistersResponse(ByVal Quantity As Long, ByRef Values() As Integer) As Byte()
    Dim baResponse() As Byte
    Dim i As Long

    ReDim baResponse(2 + Quantity * 2) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_READ_INPUT_REGISTERS
    baResponse(2) = Quantity * 2

    For i = 0 To Quantity - 1
        baResponse(3 + i * 2) = (Values(i) \ 256) And &HFF
        baResponse(4 + i * 2) = Values(i) And &HFF
    Next i

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildReadInputRegistersResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildWriteMultipleCoilsResponse(ByVal StartAddress As Long, ByVal Quantity As Long) As Byte()
    Dim baResponse(7) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_WRITE_MULTIPLE_COILS
    baResponse(2) = (StartAddress \ 256) And &HFF
    baResponse(3) = StartAddress And &HFF
    baResponse(4) = (Quantity \ 256) And &HFF
    baResponse(5) = Quantity And &HFF

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildWriteMultipleCoilsResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildWriteMultipleRegistersResponse(ByVal StartAddress As Long, ByVal Quantity As Long) As Byte()
    Dim baResponse(7) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
    baResponse(2) = (StartAddress \ 256) And &HFF
    baResponse(3) = StartAddress And &HFF
    baResponse(4) = (Quantity \ 256) And &HFF
    baResponse(5) = Quantity And &HFF

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildWriteMultipleRegistersResponse = WrapTCPResponse(baResponse)
End Function

Private Function BuildExceptionResponse(ByVal FunctionCode As Byte, ByVal ExceptionCode As ModbusSlaveExceptionCode) As Byte()
    Dim baResponse(3) As Byte

    baResponse(0) = m_SlaveID
    baResponse(1) = FunctionCode Or &H80
    baResponse(2) = ExceptionCode

    ' Add CRC for RTU only
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Modbus_AppendCRC baResponse
    End If
    
    BuildExceptionResponse = WrapTCPResponse(baResponse)
End Function

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusSlave] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Transport Event Handlers
'=========================================================================

Private Sub m_TransportRTU_OnOpen()
    m_State = MB_SLAVE_STATE_RUNNING
    RaiseEvent OnStart
End Sub

Private Sub m_TransportRTU_OnClose()
    m_State = MB_SLAVE_STATE_STOPPED
    RaiseEvent OnStop
End Sub

Private Sub m_TransportRTU_OnError(ByVal Description As String)
    RaiseError Description
End Sub

Private Sub m_TransportTCP_OnConnect()
    m_State = MB_SLAVE_STATE_RUNNING
    RaiseEvent OnStart
End Sub

Private Sub m_TransportTCP_OnDisconnect()
    m_State = MB_SLAVE_STATE_STOPPED
    RaiseEvent OnStop
End Sub

Private Sub m_TransportTCP_OnError(ByVal Description As String)
    RaiseError Description
End Sub

Private Sub m_TransportTCP_OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
    RaiseEvent OnClientConnect(ClientID, RemoteAddress)
End Sub

Private Sub m_TransportTCP_OnClientDisconnect(ByVal ClientID As String)
    RaiseEvent OnClientDisconnect(ClientID, "Connection closed")
End Sub

Private Sub m_TransportTCP_OnClientDataReceived(ByVal ClientID As String, Data() As Byte)
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    ' Store current client ID for event raising
    m_CurrentClientID = ClientID
    
    ' Raise data received event
    RaiseEvent OnDataReceived(Data)
    
    ' Process request
    baResponse = ProcessRequest(Data)
    
    ' Send response if available
    If UBound(baResponse) >= 0 Then
        m_TransportTCP.SendToClient ClientID, baResponse
        RaiseEvent OnDataSent(baResponse)
    End If
    
    ' Clear client ID
    m_CurrentClientID = ""
    Exit Sub
    
ErrorHandler:
    m_CurrentClientID = ""
    RaiseError "OnClientDataReceived: " & ERR.Description
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_ProtocolType = MB_SLAVE_PROTOCOL_TCP
    m_State = MB_SLAVE_STATE_STOPPED
    m_SlaveID = 1

    ' Initialize Defaults
    m_Defaults.TCP_PORT = MODBUS_TCP_PORT
    m_Defaults.RTU_DEFAULT_BAUDRATE = MODBUS_DEFAULT_BAUDRATE
    m_Defaults.RTU_DEFAULT_DATABITS = MODBUS_DEFAULT_DATABITS
    m_Defaults.RTU_DEFAULT_PARITY = "N"
    m_Defaults.RTU_DEFAULT_STOPBITS = MODBUS_DEFAULT_STOPBITS
    m_Defaults.RTU_DEFAULT_TIMEOUT = MODBUS_DEFAULT_TIMEOUT
    m_Defaults.MAX_PDU_SIZE = MODBUS_MAX_PDU_SIZE
    m_Defaults.MAX_REGISTERS = MODBUS_MAX_REGISTERS_PER_REQUEST
    m_Defaults.MAX_COILS = MODBUS_MAX_COILS_PER_REQUEST
    m_Defaults.TCP_MBAP_SIZE = MODBUS_TCP_MBAP_SIZE

    ' RTU defaults
    m_BaudRate = m_Defaults.RTU_DEFAULT_BAUDRATE
    m_DataBits = m_Defaults.RTU_DEFAULT_DATABITS
    m_Parity = m_Defaults.RTU_DEFAULT_PARITY
    m_StopBits = m_Defaults.RTU_DEFAULT_STOPBITS

    ' TCP defaults
    m_TCPPort = m_Defaults.TCP_PORT

    ' Initialize data areas
    m_CoilsSize = 1024
    m_DiscreteInputsSize = 1024
    m_HoldingRegistersSize = 1024
    m_InputRegistersSize = 1024

    ReDim m_CoilsData(m_CoilsSize - 1) As Boolean
    ReDim m_DiscreteInputsData(m_DiscreteInputsSize - 1) As Boolean
    ReDim m_HoldingRegistersData(m_HoldingRegistersSize - 1) As Integer
    ReDim m_InputRegistersData(m_InputRegistersSize - 1) As Integer
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Stop
    Set m_TransportRTU = Nothing
    Set m_TransportTCP = Nothing
End Sub
