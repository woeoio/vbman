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
'          - Listen for client connections
'          - Handle read/write requests from master
'          - Support for RTU (Serial) and TCP modes
'
' Usage Example:
'   Dim slave As New cModbusSlave
'
'   ' TCP Mode - Listen on all interfaces (default)
'   slave.Start 502
'
'   ' TCP Mode - Listen on localhost only
'   slave.BindAddress = "127.0.0.1"
'   slave.Start 502
'
'   ' TCP Mode - Listen on specific network interface
'   slave.Start 502, "192.168.1.100"
'
'   ' RTU Mode
'   slave.ProtocolType = MB_SLAVE_PROTOCOL_RTU
'   slave.Start "COM1"
'
' Author: Auto
' Date: 2026-01-15
'
'=========================================================================
Option Explicit

'=========================================================================
' Modbus Enums (imported from cModbus for consistency)
'=========================================================================

' Modbus Function Codes
Public Enum ModbusSlaveFunctionCode
' Read Functions
    MB_SLAVE_FC_READ_COILS = &H1                                                ' Read Coils (01)
    MB_SLAVE_FC_READ_DISCRETE_INPUTS = &H2                                      ' Read Discrete Inputs (02)
    MB_SLAVE_FC_READ_HOLDING_REGISTERS = &H3                                    ' Read Holding Registers (03)
    MB_SLAVE_FC_READ_INPUT_REGISTERS = &H4                                      ' Read Input Registers (04)
    
    ' Write Functions
    MB_SLAVE_FC_WRITE_SINGLE_COIL = &H5                                         ' Write Single Coil (05)
    MB_SLAVE_FC_WRITE_SINGLE_REGISTER = &H6                                     ' Write Single Register (06)
    MB_SLAVE_FC_WRITE_MULTIPLE_COILS = &HF                                      ' Write Multiple Coils (15)
    MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS = &H10                                 ' Write Multiple Registers (16)
    
    ' Mask Write Register (22)
    MB_SLAVE_FC_MASK_WRITE_REGISTER = &H16
    
    ' Read/Write Multiple Registers (23)
    MB_SLAVE_FC_READ_WRITE_MULTIPLE_REGISTERS = &H17
End Enum

' Modbus Exception Codes
Public Enum ModbusSlaveExceptionCode
    MB_SLAVE_EX_ILLEGAL_FUNCTION = &H1                                          ' Illegal Function
    MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS = &H2                                      ' Illegal Data Address
    MB_SLAVE_EX_ILLEGAL_DATA_VALUE = &H3                                        ' Illegal Data Value
    MB_SLAVE_EX_SLAVE_DEVICE_FAILURE = &H4                                      ' Slave Device Failure
    MB_SLAVE_EX_ACKNOWLEDGE = &H5                                               ' Acknowledge
    MB_SLAVE_EX_SLAVE_DEVICE_BUSY = &H6                                         ' Slave Device Busy
    MB_SLAVE_EX_MEMORY_PARITY_ERROR = &H8                                       ' Memory Parity Error
    MB_SLAVE_EX_GATEWAY_PATH_UNAVAILABLE = &HA                                  ' Gateway Path Unavailable
    MB_SLAVE_EX_GATEWAY_TARGET_DEVICE_FAILED = &HB                              ' Gateway Target Device Failed to Respond
End Enum

' Modbus Connection State
Public Enum ModbusSlaveState
    MB_SLAVE_STATE_STOPPED = 0
    MB_SLAVE_STATE_STARTING = 1
    MB_SLAVE_STATE_RUNNING = 2
    MB_SLAVE_STATE_ERROR = 3
End Enum

'=========================================================================
' Modbus Protocol Type
'=========================================================================

Public Enum ModbusSlaveProtocolType
    MB_SLAVE_PROTOCOL_RTU = 1                                                   ' RTU over Serial
    MB_SLAVE_PROTOCOL_TCP = 2                                                   ' TCP/IP
End Enum

'=========================================================================
' Events
'=========================================================================

' Server started
Event OnStarted()

' Server stopped
Event OnStopped()

' Client connected
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)

' Client disconnected
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)

' Error occurred
Event OnError(ByVal Description As String)

' Data received (for debugging)
Event OnDataReceived(ByVal ClientID As String, Data() As Byte)

' Read request received (allow custom handling)
Event OnReadRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByVal Quantity As Long)

' Write request received
Event OnWriteRequest(ByVal ClientID As String, ByVal FunctionCode As ModbusSlaveFunctionCode, ByVal Address As Long, ByRef Data As Variant)

'=========================================================================
' API Declarations for Serial Port (RTU Mode)
'=========================================================================

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function SetCommState Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpDCB As Any) As Long

Private Declare Function GetCommState Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpDCB As dcb) As Long

Private Declare Function SetCommTimeouts Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpCommTimeouts As COMMTIMEOUTS) As Long

Private Declare Function PurgeComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function ClearCommError Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpErrors As Long, _
    ByRef lpStat As COMSTAT) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Serial Port Constants
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const PURGE_TXABORT As Long = &H1
Private Const PURGE_RXABORT As Long = &H2
Private Const PURGE_TXCLEAR As Long = &H4
Private Const PURGE_RXCLEAR As Long = &H8

' DCB Structure for Serial Port Configuration
Private Type dcb
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long
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

' COMMTIMEOUTS Structure
Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

' COMSTAT Structure
Private Type COMSTAT
    fCtsHold As Long
    fDsrHold As Long
    fRlsdHold As Long
    fXoffHold As Long
    fXoffSent As Long
    fEof As Long
    fTxim As Long
    fReserved As Long
    cbInQue As Long
    cbOutQue As Long
End Type

' Modbus Defaults Structure
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
' Member Variables
'=========================================================================

' Default constants
Private m_Defaults As ModbusSlaveDefaults

' Protocol type
Private m_ProtocolType As ModbusSlaveProtocolType

' Server state
Private m_State As ModbusSlaveState

' Slave ID
Private m_SlaveID As Byte

' Data storage
Private m_Coils() As Boolean
Private m_DiscreteInputs() As Boolean
Private m_HoldingRegisters() As Integer
Private m_InputRegisters() As Integer

' Clients collection (TCP mode)
Private WithEvents m_ListenSocket As cWinsock
Attribute m_ListenSocket.VB_VarHelpID = -1
Private m_Clients As Collection
Private m_ClientCounter As Long

' TCP Mode bind address
Private m_BindAddress As String

' RTU Mode (Serial Port)
Private m_hSerialPort As Long
Private m_SerialPort As String
Private m_BaudRate As Long
Private m_DataBits As Long
Private m_Parity As String
Private m_StopBits As Long
Private m_RTUTimeout As Long
Private m_RTUBuffer As cByteBuffer

' RTU processing thread handle (simulated)
Private m_RTUProcessing As Boolean

'=========================================================================
' Properties
'=========================================================================

' Protocol type (RTU or TCP)
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

' Server state
Public Property Get State() As ModbusSlaveState
    State = m_State
End Property

' Slave/Unit ID
Public Property Get SlaveID() As Byte
    SlaveID = m_SlaveID
End Property

Public Property Let SlaveID(ByVal Value As Byte)
    m_SlaveID = Value
End Property

' Defaults constants (read-only)
Public Property Get Defaults() As ModbusSlaveDefaults
    Defaults = m_Defaults
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
Public Property Get Port() As Long
    On Error Resume Next
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        Port = m_Defaults.TCP_PORT
    End If
End Property

' Bind address for TCP listening (e.g., "0.0.0.0" for all interfaces, "127.0.0.1" for localhost only)
Public Property Get BindAddress() As String
    BindAddress = m_BindAddress
End Property

Public Property Let BindAddress(ByVal Value As String)
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Cannot change bind address while running"
        Exit Property
    End If
    m_BindAddress = Value
End Property

' Connected client count (TCP mode only)
Public Property Get ClientCount() As Long
    On Error Resume Next
    If Not m_Clients Is Nothing Then
        ClientCount = m_Clients.Count
    End If
End Property

'=========================================================================
' Public Methods - Server Control
'=========================================================================

' Start server (listen on TCP port or open serial port)
Public Sub Start(Optional ByVal PortOrSerial As String = "", Optional ByVal BindAddress As String = "")
    On Error GoTo ErrorHandler
    
    If m_State = MB_SLAVE_STATE_RUNNING Then
        RaiseError "Server already running"
        Exit Sub
    End If
    
    ' Set bind address if provided
    If LenB(BindAddress) > 0 Then
        m_BindAddress = BindAddress
    End If
    
    m_State = MB_SLAVE_STATE_STARTING
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        If LenB(PortOrSerial) > 0 Then m_SerialPort = PortOrSerial
        StartRTU
    ElseIf m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        Dim lPort As Long
        If LenB(PortOrSerial) > 0 Then
            lPort = CLng(PortOrSerial)
        Else
            lPort = m_Defaults.TCP_PORT
        End If
        StartTCP lPort
    Else
        RaiseError "Invalid protocol type"
    End If
    
    Exit Sub
ErrorHandler:
    m_State = MB_SLAVE_STATE_ERROR
    RaiseError "Start failed: " & ERR.Description
End Sub

' Stop server
Public Sub StopMe()
    On Error Resume Next

    If m_State = MB_SLAVE_STATE_STOPPED Then
        Exit Sub
    End If

    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        StopRTU
    ElseIf m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        StopTCP
    End If

    m_State = MB_SLAVE_STATE_STOPPED
    RaiseEvent OnStopped
End Sub

'=========================================================================
' Public Methods - Data Access
'=========================================================================

' Set coil value
Public Sub SetCoil(ByVal Address As Long, ByVal Value As Boolean)
    On Error GoTo ErrorHandler
    
    If Address >= m_Defaults.MAX_COILS Then
        RaiseError "Address exceeds maximum coils"
        Exit Sub
    End If
    
    EnsureCoilsAllocated Address
    m_Coils(Address) = Value
    
    Exit Sub
ErrorHandler:
    RaiseError "SetCoil failed: " & ERR.Description
End Sub

' Get coil value
Public Function GetCoil(ByVal Address As Long) As Boolean
    On Error Resume Next
    
    If Address <= UBound(m_Coils) Then
        GetCoil = m_Coils(Address)
    End If
End Function

' Set discrete input value
Public Sub SetDiscreteInput(ByVal Address As Long, ByVal Value As Boolean)
    On Error GoTo ErrorHandler
    
    If Address >= m_Defaults.MAX_COILS Then
        RaiseError "Address exceeds maximum discrete inputs"
        Exit Sub
    End If
    
    EnsureDiscreteInputsAllocated Address
    m_DiscreteInputs(Address) = Value
    
    Exit Sub
ErrorHandler:
    RaiseError "SetDiscreteInput failed: " & ERR.Description
End Sub

' Get discrete input value
Public Function GetDiscreteInput(ByVal Address As Long) As Boolean
    On Error Resume Next
    
    If Address <= UBound(m_DiscreteInputs) Then
        GetDiscreteInput = m_DiscreteInputs(Address)
    End If
End Function

' Set holding register value
Public Sub SetHoldingRegister(ByVal Address As Long, ByVal Value As Integer)
    On Error GoTo ErrorHandler
    
    If Address >= m_Defaults.MAX_REGISTERS Then
        RaiseError "Address exceeds maximum registers"
        Exit Sub
    End If
    
    EnsureHoldingRegistersAllocated Address
    m_HoldingRegisters(Address) = Value
    
    Exit Sub
ErrorHandler:
    RaiseError "SetHoldingRegister failed: " & ERR.Description
End Sub

' Get holding register value
Public Function GetHoldingRegister(ByVal Address As Long) As Integer
    On Error Resume Next
    
    If Address <= UBound(m_HoldingRegisters) Then
        GetHoldingRegister = m_HoldingRegisters(Address)
    End If
End Function

' Set input register value
Public Sub SetInputRegister(ByVal Address As Long, ByVal Value As Integer)
    On Error GoTo ErrorHandler
    
    If Address >= m_Defaults.MAX_REGISTERS Then
        RaiseError "Address exceeds maximum registers"
        Exit Sub
    End If
    
    EnsureInputRegistersAllocated Address
    m_InputRegisters(Address) = Value
    
    Exit Sub
ErrorHandler:
    RaiseError "SetInputRegister failed: " & ERR.Description
End Sub

' Get input register value
Public Function GetInputRegister(ByVal Address As Long) As Integer
    On Error Resume Next
    
    If Address <= UBound(m_InputRegisters) Then
        GetInputRegister = m_InputRegisters(Address)
    End If
End Function

' Clear all data
Public Sub ClearAllData()
    Erase m_Coils
    Erase m_DiscreteInputs
    Erase m_HoldingRegisters
    Erase m_InputRegisters
End Sub

'=========================================================================
' Private Methods - RTU Server
'=========================================================================

Private Sub StartRTU()
    Dim dcb As dcb
    Dim timeouts As COMMTIMEOUTS
    
    On Error GoTo ErrorHandler
    
    If LenB(m_SerialPort) = 0 Then
        RaiseError "Serial port not specified"
        Exit Sub
    End If
    
    ' Open serial port
    m_hSerialPort = CreateFile("\\.\" & m_SerialPort, _
                                GENERIC_READ Or GENERIC_WRITE, _
                                0, _
                                0, _
                                OPEN_EXISTING, _
                                FILE_ATTRIBUTE_NORMAL, _
                                0)
    
    If m_hSerialPort = INVALID_HANDLE_VALUE Then
        RaiseError "Failed to open serial port: " & m_SerialPort
        Exit Sub
    End If
    
    ' Configure DCB
    dcb.DCBlength = Len(dcb)
    dcb.BaudRate = m_BaudRate
    dcb.ByteSize = m_DataBits
    
    ' Parity
    Select Case m_Parity
        Case "E"
            dcb.Parity = 2  ' Even
        Case "O"
            dcb.Parity = 1  ' Odd
        Case Else
            dcb.Parity = 0  ' None
    End Select
    
    ' Stop bits
    If m_StopBits = 2 Then
        dcb.StopBits = 2
    Else
        dcb.StopBits = 0  ' 1 stop bit
    End If
    
    ' Set fBitFields
    dcb.fBitFields = &H1
    
    If SetCommState(m_hSerialPort, dcb) = 0 Then
        CloseHandle m_hSerialPort
        m_hSerialPort = INVALID_HANDLE_VALUE
        RaiseError "Failed to configure serial port"
        Exit Sub
    End If
    
    ' Set timeouts
    timeouts.ReadIntervalTimeout = 0
    timeouts.ReadTotalTimeoutMultiplier = 0
    timeouts.ReadTotalTimeoutConstant = m_RTUTimeout
    timeouts.WriteTotalTimeoutMultiplier = 0
    timeouts.WriteTotalTimeoutConstant = m_RTUTimeout
    
    SetCommTimeouts m_hSerialPort, timeouts
    
    ' Purge buffers
    PurgeComm m_hSerialPort, PURGE_TXCLEAR Or PURGE_RXCLEAR
    
    ' Clear receive buffer
    m_RTUBuffer.Clear
    
    m_State = MB_SLAVE_STATE_RUNNING
    m_RTUProcessing = True
    
    ' Start RTU processing loop (simulated with timer-like behavior)
    StartRTUProcessing
    
    RaiseEvent OnStarted
    
    Exit Sub
ErrorHandler:
    If m_hSerialPort <> INVALID_HANDLE_VALUE Then
        CloseHandle m_hSerialPort
        m_hSerialPort = INVALID_HANDLE_VALUE
    End If
    RaiseError "StartRTU failed: " & ERR.Description
End Sub

Private Sub StopRTU()
    On Error Resume Next
    
    m_RTUProcessing = False
    
    If m_hSerialPort <> INVALID_HANDLE_VALUE Then
        CloseHandle m_hSerialPort
        m_hSerialPort = INVALID_HANDLE_VALUE
    End If
    
    If Not m_RTUBuffer Is Nothing Then
        m_RTUBuffer.Clear
    End If
End Sub

Private Sub StartRTUProcessing()
    Dim baBuffer(255) As Byte
    Dim lBytesRead As Long
    Dim baFrame() As Byte
    
    On Error GoTo ErrorHandler
    
    Do While m_RTUProcessing And m_State = MB_SLAVE_STATE_RUNNING
        ' Read available data
        ReadFile m_hSerialPort, baBuffer(0), 256, lBytesRead, 0
        
        If lBytesRead > 0 Then
            Dim baTemp() As Byte
            ReDim baTemp(lBytesRead - 1) As Byte
            CopyMemory baTemp(0), baBuffer(0), lBytesRead
            m_RTUBuffer.Append baTemp
            
            ' Try to parse frames
            Do
                baFrame = ParseAndProcessRTUFrame()
                If UBound(baFrame) < 0 Then Exit Do
            Loop
        End If
        
        Sleep 10  ' Small delay to avoid CPU spinning
    Loop
    
    Exit Sub
ErrorHandler:
    RaiseError "RTUProcessing error: " & ERR.Description
End Sub

Private Function ParseAndProcessRTUFrame() As Byte()
    Dim baFrame() As Byte
    Dim lFrameLen As Long
    Dim wCRC As Long
    Dim wCalcCRC As Long
    Dim baPDU() As Byte
    Dim lPDULen As Long
    Dim bSlaveID As Byte
    Dim bFC As Byte
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    ' Minimum frame: SlaveID + FC + CRC (4 bytes)
    If m_RTUBuffer.Size < 4 Then
        Exit Function
    End If
    
    ' Peek at slave ID and function code
    bSlaveID = m_RTUBuffer.PeekByte(0)
    bFC = m_RTUBuffer.PeekByte(1)
    
    ' Check if addressed to us
    If bSlaveID <> m_SlaveID And bSlaveID <> 0 Then Exit Function
    
    ' Calculate expected frame length
    Dim lExpectedLen As Long
    lExpectedLen = CalculateFrameLength(bFC)
    
    If m_RTUBuffer.Size < lExpectedLen Then Exit Function
    
    ' Extract frame
    baFrame = m_RTUBuffer.Extract(lExpectedLen)
    
    ' Verify CRC
    wCRC = CLng(baFrame(lExpectedLen - 2)) + CLng(baFrame(lExpectedLen - 1)) * 256
    wCalcCRC = ModbusCRC16(baFrame, 0, lExpectedLen - 2)
    
    If wCRC <> wCalcCRC Then
        Exit Function
    End If
    
    ' Process request
    baResponse = ProcessRequest(baFrame)
    
    ' Send response if available
    If UBound(baResponse) >= 0 Then
        Dim lBytesWritten As Long
        WriteFile m_hSerialPort, baResponse(0), UBound(baResponse) + 1, lBytesWritten, 0
    End If
    
    ParseAndProcessRTUFrame = baResponse
    Exit Function
ErrorHandler:
    ' Error parsing, clear buffer
    m_RTUBuffer.Clear
End Function

'=========================================================================
' Private Methods - TCP Server
'=========================================================================

Private Sub StartTCP(ByVal Port As Long)
    On Error GoTo ErrorHandler

    ' Create listening socket
    Set m_ListenSocket = New cWinsock
    m_ListenSocket.Protocol = wskTCPProtocol

    ' Start listening with bind address
    ' - "0.0.0.0" or empty: Listen on all available network interfaces (default)
    ' - "127.0.0.1": Listen on localhost only
    ' - "192.168.1.100": Listen on specific network interface
    If LenB(m_BindAddress) > 0 And m_BindAddress <> "0.0.0.0" Then
        m_ListenSocket.Listen Port, m_BindAddress
    Else
        m_ListenSocket.Listen Port
    End If

    ' Initialize clients collection
    Set m_Clients = New Collection
    m_ClientCounter = 0

    m_State = MB_SLAVE_STATE_RUNNING
    RaiseEvent OnStarted

    Exit Sub
ErrorHandler:
    m_State = MB_SLAVE_STATE_ERROR
    RaiseError "StartTCP failed: " & ERR.Description
End Sub

Private Sub StopTCP()
    On Error Resume Next
    
    ' Close all client connections
    If Not m_Clients Is Nothing Then
        Dim oClient As cWinsock
        For Each oClient In m_Clients
            oClient.Close_
        Next
        Set m_Clients = Nothing
    End If
    
    ' Close listening socket
    If Not m_ListenSocket Is Nothing Then
        m_ListenSocket.Close_
        Set m_ListenSocket = Nothing
    End If
End Sub

Private Function GetClient(ByVal ClientID As String) As cWinsock
    On Error Resume Next
    Set GetClient = m_Clients(ClientID)
End Function

Private Sub RemoveClient(ByVal ClientID As String, ByVal Reason As String)
    On Error Resume Next
    
    Dim oClient As cWinsock
    Set oClient = m_Clients(ClientID)
    
    If Not oClient Is Nothing Then
        oClient.Close_
        m_Clients.Remove ClientID
        RaiseEvent OnClientDisconnect(ClientID, Reason)
    End If
End Sub

'=========================================================================
' Private Methods - Request Processing
'=========================================================================

Private Function ProcessRequest(ByRef Request() As Byte) As Byte()
    Dim bSlaveID As Byte
    Dim bFC As Byte
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim baResponse() As Byte
    Dim lMBAPSize As Long
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        bSlaveID = Request(lMBAPSize)
        bFC = Request(lMBAPSize + 1)
        
        ' Extract PDU (skip MBAP header)
        Dim baPDU() As Byte
        ReDim baPDU(UBound(Request) - lMBAPSize) As Byte
        CopyMemory baPDU(0), Request(lMBAPSize), UBound(baPDU) + 1
    Else
        bSlaveID = Request(0)
        bFC = Request(1)
    End If
    
    ' Check slave ID
    If bSlaveID <> m_SlaveID And bSlaveID <> 0 Then
        Exit Function  ' Not addressed to us
    End If
    
    ' Process by function code
    Select Case bFC
        Case MB_SLAVE_FC_READ_COILS
            baResponse = ProcessReadCoils(Request)
        Case MB_SLAVE_FC_READ_DISCRETE_INPUTS
            baResponse = ProcessReadDiscreteInputs(Request)
        Case MB_SLAVE_FC_READ_HOLDING_REGISTERS
            baResponse = ProcessReadHoldingRegisters(Request)
        Case MB_SLAVE_FC_READ_INPUT_REGISTERS
            baResponse = ProcessReadInputRegisters(Request)
        Case MB_SLAVE_FC_WRITE_SINGLE_COIL
            baResponse = ProcessWriteSingleCoil(Request)
        Case MB_SLAVE_FC_WRITE_SINGLE_REGISTER
            baResponse = ProcessWriteSingleRegister(Request)
        Case MB_SLAVE_FC_WRITE_MULTIPLE_COILS
            baResponse = ProcessWriteMultipleCoils(Request)
        Case MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
            baResponse = ProcessWriteMultipleRegisters(Request)
        Case Else
            baResponse = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_FUNCTION)
    End Select
    
    ProcessRequest = baResponse
    Exit Function
ErrorHandler:
    RaiseError "ProcessRequest failed: " & ERR.Description
End Function

Private Function ProcessReadCoils(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    Dim lBitIndex As Long
    Dim bByte As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    
    ' Check address range
    If lAddress + lQuantity > m_Defaults.MAX_COILS Then
        ProcessReadCoils = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    ' Build response
    lByteCount = (lQuantity + 7) \ 8
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 2 + lByteCount) As Byte
        ' Copy MBAP header
        CopyMemory baResponse(0), Request(0), lMBAPSize
        ' Set length
        baResponse(4) = ((lByteCount + 2) \ 256) And &HFF
        baResponse(5) = (lByteCount + 2) And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)  ' Slave ID
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_READ_COILS
        baResponse(lMBAPSize + 2) = lByteCount
    Else
        ReDim baResponse(1 + lByteCount + 1) As Byte  ' +1 for CRC
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_READ_COILS
        baResponse(2) = lByteCount
    End If
    
    ' Pack coil values
    For i = 0 To lQuantity - 1
        If (lAddress + i) <= UBound(m_Coils) And m_Coils(lAddress + i) Then
            bByte = bByte Or (2 ^ (i Mod 8))
        End If
        If (i Mod 8 = 7) Or (i = lQuantity - 1) Then
            If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
                baResponse(lMBAPSize + 3 + i \ 8) = bByte
            Else
                baResponse(3 + i \ 8) = bByte
            End If
            bByte = 0
        End If
    Next i
    
    ' Add CRC for RTU
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessReadCoils = baResponse
    Exit Function
ErrorHandler:
    ProcessReadCoils = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadDiscreteInputs(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    Dim bByte As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    
    If lAddress + lQuantity > m_Defaults.MAX_COILS Then
        ProcessReadDiscreteInputs = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    lByteCount = (lQuantity + 7) \ 8
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 2 + lByteCount) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = ((lByteCount + 2) \ 256) And &HFF
        baResponse(5) = (lByteCount + 2) And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_READ_DISCRETE_INPUTS
        baResponse(lMBAPSize + 2) = lByteCount
    Else
        ReDim baResponse(1 + lByteCount + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_READ_DISCRETE_INPUTS
        baResponse(2) = lByteCount
    End If
    
    For i = 0 To lQuantity - 1
        If (lAddress + i) <= UBound(m_DiscreteInputs) And m_DiscreteInputs(lAddress + i) Then
            bByte = bByte Or (2 ^ (i Mod 8))
        End If
        If (i Mod 8 = 7) Or (i = lQuantity - 1) Then
            If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
                baResponse(lMBAPSize + 3 + i \ 8) = bByte
            Else
                baResponse(3 + i \ 8) = bByte
            End If
            bByte = 0
        End If
    Next i
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessReadDiscreteInputs = baResponse
    Exit Function
ErrorHandler:
    ProcessReadDiscreteInputs = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadHoldingRegisters(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    
    If lAddress + lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessReadHoldingRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    lByteCount = lQuantity * 2
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 2 + lByteCount) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = ((lByteCount + 2) \ 256) And &HFF
        baResponse(5) = (lByteCount + 2) And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_READ_HOLDING_REGISTERS
        baResponse(lMBAPSize + 2) = lByteCount
    Else
        ReDim baResponse(1 + lByteCount + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_READ_HOLDING_REGISTERS
        baResponse(2) = lByteCount
    End If
    
    For i = 0 To lQuantity - 1
        Dim iValue As Integer
        If (lAddress + i) <= UBound(m_HoldingRegisters) Then
            iValue = m_HoldingRegisters(lAddress + i)
        Else
            iValue = 0
        End If
        
        If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
            baResponse(lMBAPSize + 3 + i * 2) = (iValue \ 256) And &HFF
            baResponse(lMBAPSize + 3 + i * 2 + 1) = iValue And &HFF
        Else
            baResponse(3 + i * 2) = (iValue \ 256) And &HFF
            baResponse(3 + i * 2 + 1) = iValue And &HFF
        End If
    Next i
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessReadHoldingRegisters = baResponse
    Exit Function
ErrorHandler:
    ProcessReadHoldingRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessReadInputRegisters(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    
    If lAddress + lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessReadInputRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    lByteCount = lQuantity * 2
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 2 + lByteCount) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = ((lByteCount + 2) \ 256) And &HFF
        baResponse(5) = (lByteCount + 2) And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_READ_INPUT_REGISTERS
        baResponse(lMBAPSize + 2) = lByteCount
    Else
        ReDim baResponse(1 + lByteCount + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_READ_INPUT_REGISTERS
        baResponse(2) = lByteCount
    End If
    
    For i = 0 To lQuantity - 1
        Dim iValue As Integer
        If (lAddress + i) <= UBound(m_InputRegisters) Then
            iValue = m_InputRegisters(lAddress + i)
        Else
            iValue = 0
        End If
        
        If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
            baResponse(lMBAPSize + 3 + i * 2) = (iValue \ 256) And &HFF
            baResponse(lMBAPSize + 3 + i * 2 + 1) = iValue And &HFF
        Else
            baResponse(3 + i * 2) = (iValue \ 256) And &HFF
            baResponse(3 + i * 2 + 1) = iValue And &HFF
        End If
    Next i
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessReadInputRegisters = baResponse
    Exit Function
ErrorHandler:
    ProcessReadInputRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteSingleCoil(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim bValue As Byte
    Dim bOnOff As Boolean
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    bValue = Request(lPDUOffset + 2)
    bOnOff = (bValue = &HFF And Request(lPDUOffset + 3) = 0)
    
    If lAddress >= m_Defaults.MAX_COILS Then
        ProcessWriteSingleCoil = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    EnsureCoilsAllocated lAddress
    m_Coils(lAddress) = bOnOff
    
    ' Echo request back
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 5) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = (6 \ 256) And &HFF
        baResponse(5) = 6 And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_WRITE_SINGLE_COIL
        baResponse(lMBAPSize + 2) = Request(lPDUOffset)
        baResponse(lMBAPSize + 3) = Request(lPDUOffset + 1)
        baResponse(lMBAPSize + 4) = Request(lPDUOffset + 2)
        baResponse(lMBAPSize + 5) = Request(lPDUOffset + 3)
    Else
        ReDim baResponse(1 + 4 + 1) As Byte  ' +1 for CRC
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_WRITE_SINGLE_COIL
        baResponse(2) = Request(lPDUOffset)
        baResponse(3) = Request(lPDUOffset + 1)
        baResponse(4) = Request(lPDUOffset + 2)
        baResponse(5) = Request(lPDUOffset + 3)
    End If
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessWriteSingleCoil = baResponse
    Exit Function
ErrorHandler:
    ProcessWriteSingleCoil = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteSingleRegister(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim iValue As Integer
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    iValue = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    
    If lAddress >= m_Defaults.MAX_REGISTERS Then
        ProcessWriteSingleRegister = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If
    
    EnsureHoldingRegistersAllocated lAddress
    m_HoldingRegisters(lAddress) = iValue
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 5) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = (6 \ 256) And &HFF
        baResponse(5) = 6 And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_WRITE_SINGLE_REGISTER
        baResponse(lMBAPSize + 2) = Request(lPDUOffset)
        baResponse(lMBAPSize + 3) = Request(lPDUOffset + 1)
        baResponse(lMBAPSize + 4) = Request(lPDUOffset + 2)
        baResponse(lMBAPSize + 5) = Request(lPDUOffset + 3)
    Else
        ReDim baResponse(1 + 4 + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_WRITE_SINGLE_REGISTER
        baResponse(2) = Request(lPDUOffset)
        baResponse(3) = Request(lPDUOffset + 1)
        baResponse(4) = Request(lPDUOffset + 2)
        baResponse(5) = Request(lPDUOffset + 3)
    End If
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessWriteSingleRegister = baResponse
    Exit Function
ErrorHandler:
    ProcessWriteSingleRegister = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteMultipleCoils(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    Dim lBitIndex As Long
    Dim bByte As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    ' FC15请求格式: SlaveID(0) + FC(1) + AddrHi(2) + AddrLo(3) + QtyHi(4) + QtyLo(5) + ByteCount(6) + Data(7...)
    ' 相对于PDUOffset: Addr(0,1) + Qty(2,3) + ByteCount(4) + Data(5...)
    lByteCount = Request(lPDUOffset + 4)

    If lAddress + lQuantity > m_Defaults.MAX_COILS Then
        ProcessWriteMultipleCoils = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    EnsureCoilsAllocated lAddress + lQuantity - 1

    For i = 0 To lQuantity - 1
        bByte = Request(lPDUOffset + 5 + i \ 8)
        lBitIndex = i Mod 8
        m_Coils(lAddress + i) = ((bByte And (2 ^ lBitIndex)) <> 0)
    Next i
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 5) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = (6 \ 256) And &HFF
        baResponse(5) = 6 And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_WRITE_MULTIPLE_COILS
        baResponse(lMBAPSize + 2) = Request(lPDUOffset)
        baResponse(lMBAPSize + 3) = Request(lPDUOffset + 1)
        baResponse(lMBAPSize + 4) = Request(lPDUOffset + 2)
        baResponse(lMBAPSize + 5) = Request(lPDUOffset + 3)
    Else
        ReDim baResponse(1 + 4 + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_WRITE_MULTIPLE_COILS
        baResponse(2) = Request(lPDUOffset)
        baResponse(3) = Request(lPDUOffset + 1)
        baResponse(4) = Request(lPDUOffset + 2)
        baResponse(5) = Request(lPDUOffset + 3)
    End If
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessWriteMultipleCoils = baResponse
    Exit Function
ErrorHandler:
    ProcessWriteMultipleCoils = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

Private Function ProcessWriteMultipleRegisters(ByRef Request() As Byte) As Byte()
    Dim lMBAPSize As Long
    Dim lPDUOffset As Long
    Dim lAddress As Long
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim baResponse() As Byte
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        lPDUOffset = lMBAPSize + 2
    Else
        lMBAPSize = 0
        lPDUOffset = 2
    End If
    
    lAddress = Request(lPDUOffset) * 256 + Request(lPDUOffset + 1)
    lQuantity = Request(lPDUOffset + 2) * 256 + Request(lPDUOffset + 3)
    ' FC16请求格式: SlaveID(0) + FC(1) + AddrHi(2) + AddrLo(3) + QtyHi(4) + QtyLo(5) + ByteCount(6) + Data(7...)
    ' 相对于PDUOffset: Addr(0,1) + Qty(2,3) + ByteCount(4) + Data(5...)
    lByteCount = Request(lPDUOffset + 4)

    If lAddress + lQuantity > m_Defaults.MAX_REGISTERS Then
        ProcessWriteMultipleRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_ILLEGAL_DATA_ADDRESS)
        Exit Function
    End If

    EnsureHoldingRegistersAllocated lAddress + lQuantity - 1

    For i = 0 To lQuantity - 1
        m_HoldingRegisters(lAddress + i) = Request(lPDUOffset + 5 + i * 2) * 256 + Request(lPDUOffset + 5 + i * 2 + 1)
    Next i
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        ReDim baResponse(lMBAPSize + 5) As Byte
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = (6 \ 256) And &HFF
        baResponse(5) = 6 And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        baResponse(lMBAPSize + 1) = MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
        baResponse(lMBAPSize + 2) = Request(lPDUOffset)
        baResponse(lMBAPSize + 3) = Request(lPDUOffset + 1)
        baResponse(lMBAPSize + 4) = Request(lPDUOffset + 2)
        baResponse(lMBAPSize + 5) = Request(lPDUOffset + 3)
    Else
        ReDim baResponse(1 + 4 + 1) As Byte
        baResponse(0) = m_SlaveID
        baResponse(1) = MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
        baResponse(2) = Request(lPDUOffset)
        baResponse(3) = Request(lPDUOffset + 1)
        baResponse(4) = Request(lPDUOffset + 2)
        baResponse(5) = Request(lPDUOffset + 3)
    End If
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_RTU Then
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, UBound(baResponse) - 1)
        baResponse(UBound(baResponse) - 1) = wCRC And &HFF
        baResponse(UBound(baResponse)) = (wCRC \ 256) And &HFF
    End If
    
    ProcessWriteMultipleRegisters = baResponse
    Exit Function
ErrorHandler:
    ProcessWriteMultipleRegisters = BuildExceptionResponse(Request, MB_SLAVE_EX_SLAVE_DEVICE_FAILURE)
End Function

'=========================================================================
' Private Methods - Response Building
'=========================================================================

Private Function BuildExceptionResponse(ByRef Request() As Byte, ByVal ExceptionCode As ModbusSlaveExceptionCode) As Byte()
    Dim lMBAPSize As Long
    Dim baResponse() As Byte
    Dim bFC As Byte
    
    On Error GoTo ErrorHandler
    
    If m_ProtocolType = MB_SLAVE_PROTOCOL_TCP Then
        lMBAPSize = m_Defaults.TCP_MBAP_SIZE
        ReDim baResponse(lMBAPSize + 2) As Byte
        
        CopyMemory baResponse(0), Request(0), lMBAPSize
        baResponse(4) = (3 \ 256) And &HFF
        baResponse(5) = 3 And &HFF
        baResponse(lMBAPSize) = Request(lMBAPSize)
        bFC = Request(lMBAPSize + 1)
        baResponse(lMBAPSize + 1) = bFC Or &H80
        baResponse(lMBAPSize + 2) = ExceptionCode
    Else
        ' RTU异常响应 = SlaveID(1) + FC|0x80(1) + ExCode(1) + CRC(2) = 5字节
        ReDim baResponse(4) As Byte  ' 索引0-4
        baResponse(0) = m_SlaveID
        bFC = Request(1)
        baResponse(1) = bFC Or &H80
        baResponse(2) = ExceptionCode

        ' Add CRC
        Dim wCRC As Long
        wCRC = ModbusCRC16(baResponse, 0, 3)  ' 计算前3个字节
        baResponse(3) = wCRC And &HFF
        baResponse(4) = (wCRC \ 256) And &HFF
    End If
    
    BuildExceptionResponse = baResponse
    Exit Function
ErrorHandler:
    RaiseError "BuildExceptionResponse failed: " & ERR.Description
End Function

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Function CalculateFrameLength(ByVal FC As Byte) As Long
    ' Calculate expected frame length for RTU mode
    Select Case FC
        Case MB_SLAVE_FC_READ_COILS, MB_SLAVE_FC_READ_DISCRETE_INPUTS, _
             MB_SLAVE_FC_READ_HOLDING_REGISTERS, MB_SLAVE_FC_READ_INPUT_REGISTERS
            ' 读请求: SlaveID(1) + FC(1) + Addr(2) + Qty(2) + CRC(2) = 8
            CalculateFrameLength = 8

        Case MB_SLAVE_FC_WRITE_SINGLE_COIL, MB_SLAVE_FC_WRITE_SINGLE_REGISTER
            ' 写单个: SlaveID(1) + FC(1) + Addr(2) + Value(2) + CRC(2) = 8
            CalculateFrameLength = 8

        Case MB_SLAVE_FC_WRITE_MULTIPLE_COILS, MB_SLAVE_FC_WRITE_MULTIPLE_REGISTERS
            ' 写多个: SlaveID(1) + FC(1) + Addr(2) + Qty(2) + ByteCount(1) + Data(N) + CRC(2)
            ' 需要先检查是否有足够数据读取 ByteCount
            If m_RTUBuffer.Size >= 7 Then
                Dim lByteCount As Long
                lByteCount = m_RTUBuffer.PeekByte(6)
                CalculateFrameLength = 7 + lByteCount + 2
            Else
                CalculateFrameLength = 0  ' 等待更多数据
            End If

        Case Else
            CalculateFrameLength = 8
    End Select
End Function

Private Sub EnsureCoilsAllocated(ByVal MaxAddress As Long)
    On Error Resume Next
    If UBound(m_Coils) < MaxAddress Then
        ReDim Preserve m_Coils(MaxAddress) As Boolean
    End If
End Sub

Private Sub EnsureDiscreteInputsAllocated(ByVal MaxAddress As Long)
    On Error Resume Next
    If UBound(m_DiscreteInputs) < MaxAddress Then
        ReDim Preserve m_DiscreteInputs(MaxAddress) As Boolean
    End If
End Sub

Private Sub EnsureHoldingRegistersAllocated(ByVal MaxAddress As Long)
    On Error Resume Next
    If UBound(m_HoldingRegisters) < MaxAddress Then
        ReDim Preserve m_HoldingRegisters(MaxAddress) As Integer
    End If
End Sub

Private Sub EnsureInputRegistersAllocated(ByVal MaxAddress As Long)
    On Error Resume Next
    If UBound(m_InputRegisters) < MaxAddress Then
        ReDim Preserve m_InputRegisters(MaxAddress) As Integer
    End If
End Sub

Private Function ModbusCRC16(ByRef Data() As Byte, ByVal StartIndex As Long, ByVal Length As Long) As Long
    Dim i As Long
    Dim j As Long
    Dim wCRC As Long
    Dim bByte As Byte
    
    wCRC = &HFFFF
    
    For i = StartIndex To StartIndex + Length - 1
        bByte = Data(i)
        wCRC = wCRC Xor CLng(bByte)
        
        For j = 0 To 7
            If (wCRC And &H1) <> 0 Then
                wCRC = (wCRC \ 2) Xor &HA001
            Else
                wCRC = wCRC \ 2
            End If
        Next j
    Next i
    
    ModbusCRC16 = wCRC
End Function

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusSlave] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Socket Event Handlers (TCP Mode)
'=========================================================================

Private Sub m_ListenSocket_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    On Error GoTo ErrorHandler

    ' Generate client ID
    m_ClientCounter = m_ClientCounter + 1
    Dim sClientID As String
    sClientID = "Client#" & m_ClientCounter

    ' 设置 Client.Tag 用于后续事件中识别客户端
    Client.Tag = sClientID

    ' Add to collection
    m_Clients.Add Client, sClientID

    ' Accept connection
    DisConnect = False

    RaiseEvent OnClientConnect(sClientID, Client.RemoteHostIP)

    Exit Sub
ErrorHandler:
    DisConnect = True
    RaiseError "ConnectionRequest: " & ERR.Description
End Sub

Private Sub m_ListenSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Dim sClientID As String
    Dim baResponse() As Byte

    On Error GoTo ErrorHandler

    ' 使用 Client.Tag 获取 ClientID
    sClientID = Client.Tag

    If LenB(sClientID) = 0 Then Exit Sub

    ' Receive data
    Client.GetData baData, vbByte + vbArray

    RaiseEvent OnDataReceived(sClientID, baData)

    ' Process request
    baResponse = ProcessRequest(baData)

    ' Send response
    If UBound(baResponse) >= 0 Then
        Client.SendData baResponse
    End If

    Exit Sub
ErrorHandler:
    RaiseError "DataArrival: " & ERR.Description
End Sub

Private Sub m_ListenSocket_CloseEvent(Client As cWinsock)
    Dim sClientID As String

    On Error Resume Next

    ' 使用 Client.Tag 获取 ClientID
    sClientID = Client.Tag

    If LenB(sClientID) > 0 Then
        RemoveClient sClientID, "Connection closed"
    End If
End Sub

Private Sub m_ListenSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Dim sClientID As String

    On Error Resume Next

    ' 使用 Client.Tag 获取 ClientID
    sClientID = Client.Tag

    If LenB(sClientID) > 0 Then
        RemoveClient sClientID, "Socket error: " & Description
    Else
        RaiseError "Socket error: " & Description
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_ProtocolType = MB_SLAVE_PROTOCOL_TCP
    m_State = MB_SLAVE_STATE_STOPPED
    m_SlaveID = 1
    m_ClientCounter = 0
    m_RTUProcessing = False
    m_BindAddress = "0.0.0.0"                                                   ' Default: listen on all interfaces
    
    ' Initialize Defaults constants structure
    m_Defaults.TCP_PORT = 502
    m_Defaults.RTU_DEFAULT_BAUDRATE = 9600
    m_Defaults.RTU_DEFAULT_DATABITS = 8
    m_Defaults.RTU_DEFAULT_PARITY = "N"
    m_Defaults.RTU_DEFAULT_STOPBITS = 1
    m_Defaults.RTU_DEFAULT_TIMEOUT = 1000
    m_Defaults.MAX_PDU_SIZE = 253
    m_Defaults.MAX_REGISTERS = 65536
    m_Defaults.MAX_COILS = 65536
    m_Defaults.TCP_MBAP_SIZE = 7
    
    ' RTU defaults
    m_BaudRate = m_Defaults.RTU_DEFAULT_BAUDRATE
    m_DataBits = m_Defaults.RTU_DEFAULT_DATABITS
    m_Parity = m_Defaults.RTU_DEFAULT_PARITY
    m_StopBits = m_Defaults.RTU_DEFAULT_STOPBITS
    m_RTUTimeout = m_Defaults.RTU_DEFAULT_TIMEOUT
    
    ' Initialize buffer
    Set m_RTUBuffer = New cByteBuffer
    m_hSerialPort = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    StopMe
    Set m_RTUBuffer = Nothing
    Set m_ListenSocket = Nothing
    Set m_Clients = Nothing
End Sub
