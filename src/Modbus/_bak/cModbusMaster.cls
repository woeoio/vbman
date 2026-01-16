VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModbusMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cModbusMaster - Modbus Master/Client Implementation for VB6
'
' Purpose: Provides Modbus RTU and TCP master/client functionality
'          - Read/Write Coils (0x01, 0x02, 0x05, 0x0F)
'          - Read/Write Registers (0x03, 0x04, 0x06, 0x10)
'          - Support for RTU (Serial) and TCP modes
'
' Author: Auto
' Date: 2026-01-15
'
'=========================================================================
Option Explicit
'
'=========================================================================
' Modbus Function Codes
'=========================================================================

Public Enum ModbusMasterFunctionCode
' Read Functions
    MB_MASTER_FC_READ_COILS = &H1                                                      ' Read Coils (01)
    MB_MASTER_FC_READ_DISCRETE_INPUTS = &H2                                            ' Read Discrete Inputs (02)
    MB_MASTER_FC_READ_HOLDING_REGISTERS = &H3                                          ' Read Holding Registers (03)
    MB_MASTER_FC_READ_INPUT_REGISTERS = &H4                                            ' Read Input Registers (04)

    ' Write Functions
    MB_MASTER_FC_WRITE_SINGLE_COIL = &H5                                               ' Write Single Coil (05)
    MB_MASTER_FC_WRITE_SINGLE_REGISTER = &H6                                           ' Write Single Register (06)
    MB_MASTER_FC_WRITE_MULTIPLE_COILS = &HF                                            ' Write Multiple Coils (15)
    MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS = &H10                                       ' Write Multiple Registers (16)

    ' Advanced Functions (optional)
    MB_MASTER_FC_MASK_WRITE_REGISTER = &H16                                            ' Mask Write Register (22)
    MB_MASTER_FC_READ_WRITE_MULTIPLE_REGISTERS = &H17                                  ' Read/Write Multiple (23)
End Enum

'=========================================================================
' Modbus Exception Codes
'=========================================================================

Public Enum ModbusMasterExceptionCode
    MB_MASTER_EX_ILLEGAL_FUNCTION = &H1                                                ' Illegal Function
    MB_MASTER_EX_ILLEGAL_DATA_ADDRESS = &H2                                            ' Illegal Data Address
    MB_MASTER_EX_ILLEGAL_DATA_VALUE = &H3                                              ' Illegal Data Value
    MB_MASTER_EX_SLAVE_DEVICE_FAILURE = &H4                                            ' Slave Device Failure
    MB_MASTER_EX_ACKNOWLEDGE = &H5                                                     ' Acknowledge
    MB_MASTER_EX_SLAVE_DEVICE_BUSY = &H6                                               ' Slave Device Busy
    MB_MASTER_EX_MEMORY_PARITY_ERROR = &H8                                             ' Memory Parity Error
    MB_MASTER_EX_GATEWAY_PATH_UNAVAILABLE = &HA                                        ' Gateway Path Unavailable
    MB_MASTER_EX_GATEWAY_TARGET_DEVICE_FAILED = &HB                                    ' Gateway Target Device Failed to Respond
End Enum

'=========================================================================
' Connection State
'=========================================================================

Public Enum ModbusMasterState
    MB_MASTER_STATE_DISCONNECTED = 0                                                   ' Disconnected
    MB_MASTER_STATE_CONNECTING = 1                                                     ' Connecting
    MB_MASTER_STATE_CONNECTED = 2                                                      ' Connected
    MB_MASTER_STATE_ERROR = 3                                                          ' Error
End Enum

'=========================================================================
' Protocol Type
'=========================================================================

Public Enum ModbusMasterProtocolType
    MB_MASTER_PROTOCOL_RTU = 1                                                         ' RTU over Serial
    MB_MASTER_PROTOCOL_TCP = 2                                                         ' TCP/IP
End Enum


'=========================================================================
' Default Constants Structure
'=========================================================================

Public Type ModbusMasterDefaults
    TCP_PORT As Long                                                            ' Default TCP port (502)
    RTU_DEFAULT_BAUDRATE As Long                                                ' Default baudrate (9600)
    RTU_DEFAULT_DATABITS As Long                                                ' Default data bits (8)
    RTU_DEFAULT_PARITY As String                                                ' Default parity (N)
    RTU_DEFAULT_STOPBITS As Long                                                ' Default stop bits (1)
    RTU_DEFAULT_TIMEOUT As Long                                                 ' Default timeout (1000ms)
    MAX_PDU_SIZE As Long                                                        ' Max PDU size (253)
    MAX_REGISTERS As Long                                                       ' Max registers per request (125)
    MAX_COILS As Long                                                           ' Max coils per request (2000)
    TCP_MBAP_SIZE As Long                                                       ' MBAP header size (7)
End Type

'=========================================================================
' Events
'=========================================================================

' Connection established
Event OnConnect()

' Connection closed
Event OnDisconnect()

' Error occurred
Event OnError(ByVal Description As String)

' Data received (for debugging)
Event OnDataReceived(Data() As Byte)

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

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'=========================================================================
' Serial Port Constants
'=========================================================================

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const PURGE_TXABORT As Long = &H1
Private Const PURGE_RXABORT As Long = &H2
Private Const PURGE_TXCLEAR As Long = &H4
Private Const PURGE_RXCLEAR As Long = &H8

'=========================================================================
' Structure Definitions
'=========================================================================

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

Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

'=========================================================================
' Member Variables
'=========================================================================

' Default constants
Private m_Defaults As ModbusMasterDefaults

' Protocol type
Private m_ProtocolType As ModbusMasterProtocolType

' Connection state
Private m_State As ModbusMasterState

' RTU Mode (Serial Port)
Private m_hSerialPort As Long
Private m_SerialPort As String
Private m_BaudRate As Long
Private m_DataBits As Long
Private m_Parity As String
Private m_StopBits As Long
Private m_RTUTimeout As Long

' TCP Mode
Private WithEvents m_Socket As cWinsock
Attribute m_Socket.VB_VarHelpID = -1
Private m_TCPHost As String
Private m_TCPPort As Long
Private m_TransactionID As Long

' Common
Private m_SlaveID As Byte
Private m_ResponseTimeout As Long

' Receive buffer for RTU
Private m_RTUBuffer As cByteBuffer

'=========================================================================
' Properties
'=========================================================================

' Protocol type (RTU or TCP)
Public Property Get ProtocolType() As ModbusMasterProtocolType
    ProtocolType = m_ProtocolType
End Property

Public Property Let ProtocolType(ByVal Value As ModbusMasterProtocolType)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change protocol type while connected"
        Exit Property
    End If
    m_ProtocolType = Value
End Property

' Connection state
Public Property Get State() As ModbusMasterState
    State = m_State
End Property

' Slave/Unit ID
Public Property Get SlaveID() As Byte
    SlaveID = m_SlaveID
End Property

Public Property Let SlaveID(ByVal Value As Byte)
    m_SlaveID = Value
End Property

' Response timeout (milliseconds)
Public Property Get ResponseTimeout() As Long
    ResponseTimeout = m_ResponseTimeout
End Property

Public Property Let ResponseTimeout(ByVal Value As Long)
    If Value > 0 Then m_ResponseTimeout = Value
End Property

' Defaults constants (read-only)
Public Property Get Defaults() As ModbusMasterDefaults
    Defaults = m_Defaults
End Property

' RTU Mode Properties
Public Property Get SerialPort() As String
    SerialPort = m_SerialPort
End Property

Public Property Let SerialPort(ByVal Value As String)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change serial port while connected"
        Exit Property
    End If
    m_SerialPort = Value
End Property

Public Property Get BaudRate() As Long
    BaudRate = m_BaudRate
End Property

Public Property Let BaudRate(ByVal Value As Long)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change baud rate while connected"
        Exit Property
    End If
    m_BaudRate = Value
End Property

Public Property Get DataBits() As Long
    DataBits = m_DataBits
End Property

Public Property Let DataBits(ByVal Value As Long)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change data bits while connected"
        Exit Property
    End If
    m_DataBits = Value
End Property

Public Property Get Parity() As String
    Parity = m_Parity
End Property

Public Property Let Parity(ByVal Value As String)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change parity while connected"
        Exit Property
    End If
    m_Parity = UCase$(Left$(Value, 1))
End Property

Public Property Get StopBits() As Long
    StopBits = m_StopBits
End Property

Public Property Let StopBits(ByVal Value As Long)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change stop bits while connected"
        Exit Property
    End If
    m_StopBits = Value
End Property

' TCP Mode Properties
Public Property Get TCPHost() As String
    TCPHost = m_TCPHost
End Property

Public Property Let TCPHost(ByVal Value As String)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change TCP host while connected"
        Exit Property
    End If
    m_TCPHost = Value
End Property

Public Property Get TCPPort() As Long
    TCPPort = m_TCPPort
End Property

Public Property Let TCPPort(ByVal Value As Long)
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Cannot change TCP port while connected"
        Exit Property
    End If
    m_TCPPort = Value
End Property

'=========================================================================
' Public Methods - Connection
'=========================================================================

' Connect to slave (RTU or TCP based on ProtocolType)
Public Sub Connect(Optional ByVal SerialPort As String = "", _
                   Optional ByVal TCPHost As String = "", _
                   Optional ByVal TCPPort As Long = 0)
    On Error GoTo ErrorHandler
    
    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Already connected"
        Exit Sub
    End If
    
    m_State = MB_MASTER_STATE_CONNECTING
    
    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        If LenB(SerialPort) > 0 Then m_SerialPort = SerialPort
        ConnectRTU
    ElseIf m_ProtocolType = MB_MASTER_PROTOCOL_TCP Then
        If LenB(TCPHost) > 0 Then m_TCPHost = TCPHost
        If TCPPort > 0 Then m_TCPPort = TCPPort
        ConnectTCP
    Else
        RaiseError "Invalid protocol type"
    End If
    
    Exit Sub
ErrorHandler:
    m_State = MB_MASTER_STATE_ERROR
    RaiseError "Connect failed: " & ERR.Description
End Sub

' Disconnect from slave
Public Sub DisConnect()
    On Error Resume Next
    
    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        DisconnectRTU
    ElseIf m_ProtocolType = MB_MASTER_PROTOCOL_TCP Then
        DisconnectTCP
    End If
    
    m_State = MB_MASTER_STATE_DISCONNECTED
    RaiseEvent OnDisconnect
End Sub

'=========================================================================
' Public Methods - Read Coils
'=========================================================================

' Read Coils (Function Code 01)
' StartAddress: Starting coil address (0-based)
' Quantity: Number of coils to read (1-2000)
' Returns: Array of Boolean values
Public Function ReadCoils(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Boolean
    Dim i As Long
    Dim bByte As Byte
    Dim lBitIndex As Long
    
    On Error GoTo ErrorHandler
    
    ValidateQuantity Quantity, m_Defaults.MAX_COILS
    
    baRequest = BuildReadRequest(MB_MASTER_FC_READ_COILS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        Dim lByteCount As Long
        lByteCount = baResponse(2)
        
        ReDim baResult(Quantity - 1) As Boolean
        
        For i = 0 To lByteCount - 1
            bByte = baResponse(3 + i)
            For lBitIndex = 0 To 7
                If i * 8 + lBitIndex < Quantity Then
                    baResult(i * 8 + lBitIndex) = ((bByte And (2 ^ lBitIndex)) <> 0)
                End If
            Next lBitIndex
        Next i
        
        ReadCoils = baResult
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "ReadCoils failed: " & ERR.Description
End Function

' Read Discrete Inputs (Function Code 02)
Public Function ReadDiscreteInputs(ByVal StartAddress As Long, ByVal Quantity As Long) As Boolean()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Boolean
    Dim i As Long
    Dim bByte As Byte
    Dim lBitIndex As Long
    
    On Error GoTo ErrorHandler
    
    ValidateQuantity Quantity, m_Defaults.MAX_COILS
    
    baRequest = BuildReadRequest(MB_MASTER_FC_READ_DISCRETE_INPUTS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        Dim lByteCount As Long
        lByteCount = baResponse(2)
        
        ReDim baResult(Quantity - 1) As Boolean
        
        For i = 0 To lByteCount - 1
            bByte = baResponse(3 + i)
            For lBitIndex = 0 To 7
                If i * 8 + lBitIndex < Quantity Then
                    baResult(i * 8 + lBitIndex) = ((bByte And (2 ^ lBitIndex)) <> 0)
                End If
            Next lBitIndex
        Next i
        
        ReadDiscreteInputs = baResult
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "ReadDiscreteInputs failed: " & ERR.Description
End Function

'=========================================================================
' Public Methods - Read Registers
'=========================================================================

' Read Holding Registers (Function Code 03)
' Returns: Array of Integer values (16-bit)
Public Function ReadHoldingRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Integer
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ValidateQuantity Quantity, m_Defaults.MAX_REGISTERS
    
    baRequest = BuildReadRequest(MB_MASTER_FC_READ_HOLDING_REGISTERS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        Dim lByteCount As Long
        lByteCount = baResponse(2)
        
        ReDim baResult(Quantity - 1) As Integer
        
        For i = 0 To Quantity - 1
            baResult(i) = CInt(baResponse(3 + i * 2)) * 256 + CInt(baResponse(3 + i * 2 + 1))
        Next i
        
        ReadHoldingRegisters = baResult
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "ReadHoldingRegisters failed: " & ERR.Description
End Function

' Read Input Registers (Function Code 04)
Public Function ReadInputRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Integer
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ValidateQuantity Quantity, m_Defaults.MAX_REGISTERS
    
    baRequest = BuildReadRequest(MB_MASTER_FC_READ_INPUT_REGISTERS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        Dim lByteCount As Long
        lByteCount = baResponse(2)
        
        ReDim baResult(Quantity - 1) As Integer
        
        For i = 0 To Quantity - 1
            baResult(i) = CInt(baResponse(3 + i * 2)) * 256 + CInt(baResponse(3 + i * 2 + 1))
        Next i
        
        ReadInputRegisters = baResult
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "ReadInputRegisters failed: " & ERR.Description
End Function

'=========================================================================
' Public Methods - Write Coils
'=========================================================================

' Write Single Coil (Function Code 05)
Public Function WriteSingleCoil(ByVal Address As Long, ByVal Value As Boolean) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    baRequest = BuildWriteSingleCoilRequest(Address, Value)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        If UBound(baResponse) >= 5 Then
            If baResponse(0) = baRequest(0) And _
               baResponse(1) = baRequest(1) And _
               baResponse(2) = baRequest(2) And _
               baResponse(3) = baRequest(3) And _
               baResponse(4) = baRequest(4) And _
               baResponse(5) = baRequest(5) Then
                WriteSingleCoil = True
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "WriteSingleCoil failed: " & ERR.Description
End Function

' Write Multiple Coils (Function Code 15)
Public Function WriteMultipleCoils(ByVal StartAddress As Long, ByRef Values() As Boolean) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim lQuantity As Long
    
    On Error GoTo ErrorHandler
    
    lQuantity = UBound(Values) + 1
    ValidateQuantity lQuantity, m_Defaults.MAX_COILS
    
    baRequest = BuildWriteMultipleCoilsRequest(StartAddress, Values)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        If baResponse(0) = baRequest(0) And _
           baResponse(1) = baRequest(1) And _
           baResponse(2) = baRequest(2) And _
           baResponse(3) = baRequest(3) And _
           baResponse(4) = baRequest(4) Then
            WriteMultipleCoils = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "WriteMultipleCoils failed: " & ERR.Description
End Function

'=========================================================================
' Public Methods - Write Registers
'=========================================================================

' Write Single Register (Function Code 06)
Public Function WriteSingleRegister(ByVal Address As Long, ByVal Value As Integer) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    baRequest = BuildWriteSingleRegisterRequest(Address, Value)
    baResponse = SendRequest(baRequest)

    If UBound(baResponse) >= 0 Then
        If UBound(baResponse) >= 5 Then
            If baResponse(0) = baRequest(0) And _
               baResponse(1) = baRequest(1) And _
               baResponse(2) = baRequest(2) And _
               baResponse(3) = baRequest(3) And _
               baResponse(4) = baRequest(4) And _
               baResponse(5) = baRequest(5) Then
                WriteSingleRegister = True
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "WriteSingleRegister failed: " & ERR.Description
End Function

' Write Multiple Registers (Function Code 16)
Public Function WriteMultipleRegisters(ByVal StartAddress As Long, ByRef Values() As Integer) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim lQuantity As Long
    
    On Error GoTo ErrorHandler
    
    lQuantity = UBound(Values) + 1
    ValidateQuantity lQuantity, m_Defaults.MAX_REGISTERS
    
    baRequest = BuildWriteMultipleRegistersRequest(StartAddress, Values)
    baResponse = SendRequest(baRequest)
    
    If UBound(baResponse) >= 0 Then
        If baResponse(0) = baRequest(0) And _
           baResponse(1) = baRequest(1) And _
           baResponse(2) = baRequest(2) And _
           baResponse(3) = baRequest(3) And _
           baResponse(4) = baRequest(4) Then
            WriteMultipleRegisters = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    RaiseError "WriteMultipleRegisters failed: " & ERR.Description
End Function

'=========================================================================
' Private Methods - RTU Connection
'=========================================================================

Private Sub ConnectRTU()
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
    
    ' Set fBitFields (fBinary = 1)
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
    
    m_State = MB_MASTER_STATE_CONNECTED
    RaiseEvent OnConnect
    
    Exit Sub
ErrorHandler:
    If m_hSerialPort <> INVALID_HANDLE_VALUE Then
        CloseHandle m_hSerialPort
        m_hSerialPort = INVALID_HANDLE_VALUE
    End If
    RaiseError "ConnectRTU failed: " & ERR.Description
End Sub

Private Sub DisconnectRTU()
    On Error Resume Next
    
    If m_hSerialPort <> INVALID_HANDLE_VALUE Then
        CloseHandle m_hSerialPort
        m_hSerialPort = INVALID_HANDLE_VALUE
    End If
    
    If Not m_RTUBuffer Is Nothing Then
        m_RTUBuffer.Clear
    End If
End Sub

'=========================================================================
' Private Methods - TCP Connection
'=========================================================================

Private Sub ConnectTCP()
    On Error GoTo ErrorHandler
    
    If LenB(m_TCPHost) = 0 Then
        RaiseError "TCP host not specified"
        Exit Sub
    End If
    
    If m_TCPPort = 0 Then
        m_TCPPort = m_Defaults.TCP_PORT
    End If
    
    Set m_Socket = New cWinsock
    m_Socket.Protocol = wskTCPProtocol
    m_Socket.Connect m_TCPHost, m_TCPPort
    
    Exit Sub
ErrorHandler:
    m_State = MB_MASTER_STATE_ERROR
    RaiseError "ConnectTCP failed: " & ERR.Description
End Sub

Private Sub DisconnectTCP()
    On Error Resume Next
    
    If Not m_Socket Is Nothing Then
        m_Socket.Close_
        Set m_Socket = Nothing
    End If
End Sub

'=========================================================================
' Private Methods - Request Building
'=========================================================================

Private Function BuildReadRequest(ByVal FC As ModbusMasterFunctionCode, ByVal StartAddress As Long, ByVal Quantity As Long) As Byte()
    Dim baPDU() As Byte
    ReDim baPDU(5) As Byte
    
    baPDU(0) = m_SlaveID
    baPDU(1) = FC
    baPDU(2) = (StartAddress \ 256) And &HFF
    baPDU(3) = StartAddress And &HFF
    baPDU(4) = (Quantity \ 256) And &HFF
    baPDU(5) = Quantity And &HFF
    
    BuildReadRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteSingleCoilRequest(ByVal Address As Long, ByVal Value As Boolean) As Byte()
    Dim baPDU() As Byte
    ReDim baPDU(5) As Byte
    
    baPDU(0) = m_SlaveID
    baPDU(1) = MB_MASTER_FC_WRITE_SINGLE_COIL
    baPDU(2) = (Address \ 256) And &HFF
    baPDU(3) = Address And &HFF
    If Value Then
        baPDU(4) = &HFF
        baPDU(5) = &H0
    Else
        baPDU(4) = &H0
        baPDU(5) = &H0
    End If
    
    BuildWriteSingleCoilRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteSingleRegisterRequest(ByVal Address As Long, ByVal Value As Integer) As Byte()
    Dim baPDU() As Byte
    ReDim baPDU(5) As Byte
    
    baPDU(0) = m_SlaveID
    baPDU(1) = MB_MASTER_FC_WRITE_SINGLE_REGISTER
    baPDU(2) = (Address \ 256) And &HFF
    baPDU(3) = Address And &HFF
    baPDU(4) = (Value \ 256) And &HFF
    baPDU(5) = Value And &HFF
    
    BuildWriteSingleRegisterRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteMultipleCoilsRequest(ByVal StartAddress As Long, ByRef Values() As Boolean) As Byte()
    Dim baPDU() As Byte
    Dim lQuantity As Long
    Dim lByteCount As Long
    Dim i As Long
    Dim lByteIndex As Long
    Dim lBitIndex As Long
    Dim bByte As Byte
    
    lQuantity = UBound(Values) + 1
    lByteCount = (lQuantity + 7) \ 8
    
    ReDim baPDU(6 + lByteCount) As Byte
    
    baPDU(0) = m_SlaveID
    baPDU(1) = MB_MASTER_FC_WRITE_MULTIPLE_COILS
    baPDU(2) = (StartAddress \ 256) And &HFF
    baPDU(3) = StartAddress And &HFF
    baPDU(4) = (lQuantity \ 256) And &HFF
    baPDU(5) = lQuantity And &HFF
    baPDU(6) = lByteCount
    
    ' Pack bits into bytes
    For i = 0 To lQuantity - 1
        lByteIndex = i \ 8
        lBitIndex = i Mod 8
        If Values(i) Then
            bByte = bByte Or (2 ^ lBitIndex)
        End If
        If lBitIndex = 7 Or i = lQuantity - 1 Then
            baPDU(7 + lByteIndex) = bByte
            bByte = 0
        End If
    Next i
    
    BuildWriteMultipleCoilsRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteMultipleRegistersRequest(ByVal StartAddress As Long, ByRef Values() As Integer) As Byte()
    Dim baPDU() As Byte
    Dim lQuantity As Long
    Dim i As Long
    
    lQuantity = UBound(Values) + 1
    
    ReDim baPDU(6 + lQuantity * 2) As Byte
    
    baPDU(0) = m_SlaveID
    baPDU(1) = MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS
    baPDU(2) = (StartAddress \ 256) And &HFF
    baPDU(3) = StartAddress And &HFF
    baPDU(4) = (lQuantity \ 256) And &HFF
    baPDU(5) = lQuantity And &HFF
    baPDU(6) = lQuantity * 2
    
    For i = 0 To lQuantity - 1
        baPDU(7 + i * 2) = (Values(i) \ 256) And &HFF
        baPDU(8 + i * 2) = Values(i) And &HFF
    Next i
    
    BuildWriteMultipleRegistersRequest = BuildFrame(baPDU)
End Function

Private Function BuildFrame(ByRef PDU() As Byte) As Byte()
    Dim baFrame() As Byte
    Dim lPDULen As Long
    Dim wCRC As Long
    
    lPDULen = UBound(PDU) + 1
    
    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        ' RTU: PDU + CRC16
        ReDim baFrame(lPDULen + 1) As Byte
        CopyMemory baFrame(0), PDU(0), lPDULen
        wCRC = ModbusCRC16(baFrame, 0, lPDULen)
        baFrame(lPDULen) = wCRC And &HFF
        baFrame(lPDULen + 1) = (wCRC \ 256) And &HFF
    Else
        ' TCP: MBAP Header (7 bytes) + PDU (不含SlaveID)
        m_TransactionID = m_TransactionID + 1
        If m_TransactionID > 65535 Then m_TransactionID = 1

        ' 修复: 帧大小 = MBAP(7) + PDU长度 - 1 (因为UnitID在MBAP中)
        ReDim baFrame(m_Defaults.TCP_MBAP_SIZE + lPDULen - 2) As Byte

        ' MBAP Header
        baFrame(0) = (m_TransactionID \ 256) And &HFF
        baFrame(1) = m_TransactionID And &HFF
        baFrame(2) = 0  ' Protocol ID
        baFrame(3) = 0
        baFrame(4) = ((lPDULen) \ 256) And &HFF  ' Length
        baFrame(5) = (lPDULen) And &HFF
        baFrame(6) = PDU(0)  ' Unit ID

        ' PDU (跳过SlaveID，只复制FC和数据)
        If lPDULen > 1 Then
            CopyMemory baFrame(m_Defaults.TCP_MBAP_SIZE), PDU(1), lPDULen - 1
        End If
    End If
    
    BuildFrame = baFrame
End Function

'=========================================================================
' Private Methods - Request/Response
'=========================================================================

Private Function SendRequest(ByRef Request() As Byte) As Byte()
    Dim baResponse() As Byte
    
    On Error GoTo ErrorHandler
    
    If m_State <> MB_MASTER_STATE_CONNECTED Then
        RaiseError "Not connected"
        Exit Function
    End If
    
    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        baResponse = SendRequestRTU(Request)
    Else
        baResponse = SendRequestTCP(Request)
    End If
    
    ' Check for exception response
    If UBound(baResponse) >= 0 Then
        If (baResponse(1) And &H80) <> 0 Then
            Dim lExceptionCode As Long
            lExceptionCode = baResponse(2)
            RaiseError "Modbus Exception: " & GetModbusExceptionDescription(lExceptionCode)
            Exit Function
        End If
    End If
    
    SendRequest = baResponse
    Exit Function
ErrorHandler:
    RaiseError "SendRequest failed: " & ERR.Description
End Function

Private Function SendRequestRTU(ByRef Request() As Byte) As Byte()
    Dim lBytesWritten As Long
    Dim lBytesRead As Long
    Dim baBuffer(255) As Byte
    Dim baResponse() As Byte
    Dim lStartTime As Long
    Dim lElapsed As Long
    
    On Error GoTo ErrorHandler

    ' 确保 RTUBuffer 已初始化
    If m_RTUBuffer Is Nothing Then
        Set m_RTUBuffer = New cByteBuffer
    End If

    ' Clear receive buffer
    m_RTUBuffer.Clear
    
    ' Purge buffers
    PurgeComm m_hSerialPort, PURGE_TXCLEAR Or PURGE_RXCLEAR
    
    ' Send request
    WriteFile m_hSerialPort, Request(0), UBound(Request) + 1, lBytesWritten, 0
    
    If lBytesWritten <> UBound(Request) + 1 Then
        RaiseError "Failed to write all bytes"
        Exit Function
    End If
    
    ' Wait for response
    lStartTime = GetTickCount
    
    Do
        ' Read available data
        ReadFile m_hSerialPort, baBuffer(0), 256, lBytesRead, 0
        
        If lBytesRead > 0 Then
            Dim baTemp() As Byte
            ReDim baTemp(lBytesRead - 1) As Byte
            CopyMemory baTemp(0), baBuffer(0), lBytesRead
            m_RTUBuffer.Append baTemp
        End If
        
        ' Try to parse frame
        baResponse = ParseRTUFrame()
        If UBound(baResponse) >= 0 Then
            SendRequestRTU = baResponse
            Exit Function
        End If
        
        ' Check timeout
        lElapsed = GetElapsedTime(lStartTime)
        If lElapsed > m_ResponseTimeout Then
            RaiseError "Response timeout"
            Exit Function
        End If
        
        Sleep 1
    Loop
    
    Exit Function
ErrorHandler:
    RaiseError "SendRequestRTU failed: " & ERR.Description
End Function

Private Function ParseRTUFrame() As Byte()
    Dim baFrame() As Byte
    Dim lFrameLen As Long
    Dim wCRC As Long
    Dim wCalcCRC As Long
    Dim baPDU() As Byte
    Dim lPDULen As Long
    Dim bFC As Byte
    Dim lByteCount As Long
    
    On Error GoTo ErrorHandler
    
    If m_RTUBuffer.Size < 4 Then Exit Function
    
    bFC = m_RTUBuffer.PeekByte(1)
    
    If (bFC And &H80) <> 0 Then
        lPDULen = 5
    Else
        Select Case bFC
            Case MB_MASTER_FC_READ_COILS, MB_MASTER_FC_READ_DISCRETE_INPUTS, MB_MASTER_FC_READ_HOLDING_REGISTERS, MB_MASTER_FC_READ_INPUT_REGISTERS
                If m_RTUBuffer.Size < 5 Then Exit Function
                lByteCount = m_RTUBuffer.PeekByte(2)
                lPDULen = 3 + lByteCount + 2
            Case MB_MASTER_FC_WRITE_SINGLE_COIL, MB_MASTER_FC_WRITE_SINGLE_REGISTER
                ' 写单个响应: SlaveID(1) + FC(1) + Addr(2) + Value(2) + CRC(2) = 8
                lPDULen = 8
            Case MB_MASTER_FC_WRITE_MULTIPLE_COILS, MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS
                ' 写多个响应: SlaveID(1) + FC(1) + Addr(2) + Quantity(2) + CRC(2) = 8
                lPDULen = 8
            Case Else
                lPDULen = 4
        End Select
    End If
    
    If m_RTUBuffer.Size < lPDULen Then Exit Function
    
    baFrame = m_RTUBuffer.Extract(lPDULen)
    
    wCRC = CLng(baFrame(lPDULen - 2)) + CLng(baFrame(lPDULen - 1)) * 256
    wCalcCRC = ModbusCRC16(baFrame, 0, lPDULen - 2)
    
    If wCRC <> wCalcCRC Then
        Exit Function
    End If
    
    ReDim baPDU(lPDULen - 3) As Byte
    CopyMemory baPDU(0), baFrame(0), lPDULen - 2
    
    ParseRTUFrame = baPDU
    Exit Function
ErrorHandler:
    m_RTUBuffer.Clear
End Function

Private Function SendRequestTCP(ByRef Request() As Byte) As Byte()
    Dim baResponse() As Byte
    Dim baRecvBuffer() As Byte
    Dim lStartTime As Long
    Dim lElapsed As Long
    Dim lMBAPLen As Long
    Dim lPDULen As Long
    
    On Error GoTo ErrorHandler
    
    m_Socket.SendData Request
    
    lStartTime = GetTickCount
    
    Do
        If m_Socket.BytesReceived > 0 Then
            m_Socket.GetData baRecvBuffer, vbByte + vbArray
            
            If UBound(baRecvBuffer) >= m_Defaults.TCP_MBAP_SIZE - 1 Then
                Dim lTransactionID As Long
                lTransactionID = CLng(baRecvBuffer(0)) * 256 + CLng(baRecvBuffer(1))
                
                If lTransactionID = m_TransactionID Then
                    lMBAPLen = CLng(baRecvBuffer(4)) * 256 + CLng(baRecvBuffer(5))

                    ReDim baResponse(lMBAPLen - 1) As Byte
                    baResponse(0) = baRecvBuffer(6)
                    CopyMemory baResponse(1), baRecvBuffer(m_Defaults.TCP_MBAP_SIZE), lMBAPLen - 1
                    
                    SendRequestTCP = baResponse
                    Exit Function
                End If
            End If
        End If
        
        lElapsed = GetElapsedTime(lStartTime)
        If lElapsed > m_ResponseTimeout Then
            RaiseError "Response timeout"
            Exit Function
        End If
        
        Sleep 1
    Loop
    
    Exit Function
ErrorHandler:
    RaiseError "SendRequestTCP failed: " & ERR.Description
End Function

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Function GetElapsedTime(ByVal lStartTime As Long) As Long
    Dim lNow As Long
    lNow = GetTickCount
    If lNow >= lStartTime Then
        GetElapsedTime = lNow - lStartTime
    Else
        GetElapsedTime = (2147483647 - lStartTime) + lNow + 1
    End If
End Function

Private Sub ValidateQuantity(ByVal Quantity As Long, ByVal MaxQuantity As Long)
    If Quantity < 1 Or Quantity > MaxQuantity Then
        ERR.Raise vbObjectError + 1, "cModbusMaster", "Invalid quantity: " & Quantity
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

Private Function GetModbusExceptionDescription(ByVal ExceptionCode As ModbusMasterExceptionCode) As String
    Select Case ExceptionCode
    Case MB_MASTER_EX_ILLEGAL_FUNCTION
        GetModbusExceptionDescription = "Illegal Function"
    Case MB_MASTER_EX_ILLEGAL_DATA_ADDRESS
        GetModbusExceptionDescription = "Illegal Data Address"
    Case MB_MASTER_EX_ILLEGAL_DATA_VALUE
        GetModbusExceptionDescription = "Illegal Data Value"
    Case MB_MASTER_EX_SLAVE_DEVICE_FAILURE
        GetModbusExceptionDescription = "Slave Device Failure"
    Case MB_MASTER_EX_ACKNOWLEDGE
        GetModbusExceptionDescription = "Acknowledge"
    Case MB_MASTER_EX_SLAVE_DEVICE_BUSY
        GetModbusExceptionDescription = "Slave Device Busy"
    Case MB_MASTER_EX_MEMORY_PARITY_ERROR
        GetModbusExceptionDescription = "Memory Parity Error"
    Case MB_MASTER_EX_GATEWAY_PATH_UNAVAILABLE
        GetModbusExceptionDescription = "Gateway Path Unavailable"
    Case MB_MASTER_EX_GATEWAY_TARGET_DEVICE_FAILED
        GetModbusExceptionDescription = "Gateway Target Device Failed to Respond"
    Case Else
        GetModbusExceptionDescription = "Unknown Exception (" & Hex$(ExceptionCode) & ")"
    End Select
End Function

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusMaster] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Socket Event Handlers (TCP Mode)
'=========================================================================

Private Sub m_Socket_Connect(Client As cWinsock)
    m_State = MB_MASTER_STATE_CONNECTED
    RaiseEvent OnConnect
End Sub

Private Sub m_Socket_CloseEvent(Client As cWinsock)
    m_State = MB_MASTER_STATE_DISCONNECTED
    RaiseEvent OnDisconnect
End Sub

Private Sub m_Socket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    m_State = MB_MASTER_STATE_ERROR
    RaiseError "Socket error: " & Description
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_ProtocolType = MB_MASTER_PROTOCOL_TCP
    m_State = MB_MASTER_STATE_DISCONNECTED
    m_SlaveID = 1
    m_ResponseTimeout = 1000
    m_TransactionID = 0
    
    ' Initialize Defaults
    m_Defaults.TCP_PORT = 502
    m_Defaults.RTU_DEFAULT_BAUDRATE = 9600
    m_Defaults.RTU_DEFAULT_DATABITS = 8
    m_Defaults.RTU_DEFAULT_PARITY = "N"
    m_Defaults.RTU_DEFAULT_STOPBITS = 1
    m_Defaults.RTU_DEFAULT_TIMEOUT = 1000
    m_Defaults.MAX_PDU_SIZE = 253
    m_Defaults.MAX_REGISTERS = 125
    m_Defaults.MAX_COILS = 2000
    m_Defaults.TCP_MBAP_SIZE = 7
    
    ' RTU defaults
    m_BaudRate = m_Defaults.RTU_DEFAULT_BAUDRATE
    m_DataBits = m_Defaults.RTU_DEFAULT_DATABITS
    m_Parity = m_Defaults.RTU_DEFAULT_PARITY
    m_StopBits = m_Defaults.RTU_DEFAULT_STOPBITS
    m_RTUTimeout = m_Defaults.RTU_DEFAULT_TIMEOUT
    
    ' TCP defaults
    m_TCPPort = m_Defaults.TCP_PORT
    
    ' Initialize buffer
    Set m_RTUBuffer = New cByteBuffer
    m_hSerialPort = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    DisConnect
    Set m_RTUBuffer = Nothing
    Set m_Socket = Nothing
End Sub
