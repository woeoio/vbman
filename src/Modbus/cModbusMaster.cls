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
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

'=========================================================================
' Modbus Function Codes
'=========================================================================

Public Enum ModbusMasterFunctionCode
    ' Read Functions
    MB_MASTER_FC_READ_COILS = &H1                                                 ' Read Coils (01)
    MB_MASTER_FC_READ_DISCRETE_INPUTS = &H2                                       ' Read Discrete Inputs (02)
    MB_MASTER_FC_READ_HOLDING_REGISTERS = &H3                                     ' Read Holding Registers (03)
    MB_MASTER_FC_READ_INPUT_REGISTERS = &H4                                       ' Read Input Registers (04)

    ' Write Functions
    MB_MASTER_FC_WRITE_SINGLE_COIL = &H5                                          ' Write Single Coil (05)
    MB_MASTER_FC_WRITE_SINGLE_REGISTER = &H6                                      ' Write Single Register (06)
    MB_MASTER_FC_WRITE_MULTIPLE_COILS = &HF                                       ' Write Multiple Coils (15)
    MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS = &H10                                  ' Write Multiple Registers (16)

    ' Advanced Functions (optional)
    MB_MASTER_FC_MASK_WRITE_REGISTER = &H16                                       ' Mask Write Register (22)
    MB_MASTER_FC_READ_WRITE_MULTIPLE_REGISTERS = &H17                             ' Read/Write Multiple (23)
End Enum

'=========================================================================
' Modbus Exception Codes
'=========================================================================

Public Enum ModbusMasterExceptionCode
    MB_MASTER_EX_ILLEGAL_FUNCTION = &H1                                           ' Illegal Function
    MB_MASTER_EX_ILLEGAL_DATA_ADDRESS = &H2                                       ' Illegal Data Address
    MB_MASTER_EX_ILLEGAL_DATA_VALUE = &H3                                         ' Illegal Data Value
    MB_MASTER_EX_SLAVE_DEVICE_FAILURE = &H4                                       ' Slave Device Failure
    MB_MASTER_EX_ACKNOWLEDGE = &H5                                                ' Acknowledge
    MB_MASTER_EX_SLAVE_DEVICE_BUSY = &H6                                          ' Slave Device Busy
    MB_MASTER_EX_MEMORY_PARITY_ERROR = &H8                                        ' Memory Parity Error
    MB_MASTER_EX_GATEWAY_PATH_UNAVAILABLE = &HA                                   ' Gateway Path Unavailable
    MB_MASTER_EX_GATEWAY_TARGET_DEVICE_FAILED = &HB                               ' Gateway Target Device Failed
End Enum

'=========================================================================
' Connection State
'=========================================================================

Public Enum ModbusMasterState
    MB_MASTER_STATE_DISCONNECTED = 0                                              ' Disconnected
    MB_MASTER_STATE_CONNECTING = 1                                                ' Connecting
    MB_MASTER_STATE_CONNECTED = 2                                                 ' Connected
    MB_MASTER_STATE_ERROR = 3                                                     ' Error
End Enum

'=========================================================================
' Protocol Type
'=========================================================================

Public Enum ModbusMasterProtocolType
    MB_MASTER_PROTOCOL_RTU = 1                                                    ' RTU over Serial
    MB_MASTER_PROTOCOL_TCP = 2                                                    ' TCP/IP
End Enum

'=========================================================================
' Default Constants Structure
'=========================================================================

Public Type ModbusMasterDefaults
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

Event OnConnect()
Event OnDisconnect()
Event OnError(ByVal Description As String)
Event OnDataReceived(Data() As Byte)

'=========================================================================
' Member Variables
'=========================================================================

' Default constants
Private m_Defaults As ModbusMasterDefaults

' Protocol type
Private m_ProtocolType As ModbusMasterProtocolType

' Connection state
Private m_State As ModbusMasterState

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
Private m_TCPHost As String
Private m_TCPPort As Long

' Common settings
Private m_SlaveID As Byte
Private m_ResponseTimeout As Long

'=========================================================================
' Properties
'=========================================================================

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

Public Property Get State() As ModbusMasterState
    State = m_State
End Property

Public Property Get SlaveID() As Byte
    SlaveID = m_SlaveID
End Property

Public Property Let SlaveID(ByVal Value As Byte)
    m_SlaveID = Value
End Property

Public Property Get ResponseTimeout() As Long
    ResponseTimeout = m_ResponseTimeout
End Property

Public Property Let ResponseTimeout(ByVal Value As Long)
    If Value > 0 Then
        m_ResponseTimeout = Value
        ' Update transport timeout
        If Not m_TransportRTU Is Nothing Then m_TransportRTU.Timeout = Value
        If Not m_TransportTCP Is Nothing Then m_TransportTCP.Timeout = Value
    End If
End Property

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

Public Sub Connect(Optional ByVal SerialPortOrHost As String = "", _
                   Optional ByVal TCPPort As Long = 0)
    On Error GoTo ErrorHandler

    If m_State = MB_MASTER_STATE_CONNECTED Then
        RaiseError "Already connected"
        Exit Sub
    End If

    m_State = MB_MASTER_STATE_CONNECTING

    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        If LenB(SerialPortOrHost) > 0 Then m_SerialPort = SerialPortOrHost
        ConnectRTU
    ElseIf m_ProtocolType = MB_MASTER_PROTOCOL_TCP Then
        If LenB(SerialPortOrHost) > 0 Then m_TCPHost = SerialPortOrHost
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

Public Sub DisConnect()
    On Error Resume Next

    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        If Not m_TransportRTU Is Nothing Then
            m_TransportRTU.Close_
        End If
    ElseIf m_ProtocolType = MB_MASTER_PROTOCOL_TCP Then
        If Not m_TransportTCP Is Nothing Then
            m_TransportTCP.Disconnect
        End If
    End If

    m_State = MB_MASTER_STATE_DISCONNECTED
    RaiseEvent OnDisconnect
End Sub

'=========================================================================
' Public Methods - Read Coils
'=========================================================================

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
    baResponse = SendRequest(baRequest, Modbus_CalcResponseLength(MODBUS_FC_READ_COILS, Quantity))

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
    baResponse = SendRequest(baRequest, Modbus_CalcResponseLength(MODBUS_FC_READ_DISCRETE_INPUTS, Quantity))

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

Public Function ReadHoldingRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Integer
    Dim i As Long

    On Error GoTo ErrorHandler

    ValidateQuantity Quantity, m_Defaults.MAX_REGISTERS

    baRequest = BuildReadRequest(MB_MASTER_FC_READ_HOLDING_REGISTERS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest, Modbus_CalcResponseLength(MODBUS_FC_READ_HOLDING_REGISTERS, Quantity))

    If UBound(baResponse) >= 0 Then
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

Public Function ReadInputRegisters(ByVal StartAddress As Long, ByVal Quantity As Long) As Integer()
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim baResult() As Integer
    Dim i As Long

    On Error GoTo ErrorHandler

    ValidateQuantity Quantity, m_Defaults.MAX_REGISTERS

    baRequest = BuildReadRequest(MB_MASTER_FC_READ_INPUT_REGISTERS, StartAddress, Quantity)
    baResponse = SendRequest(baRequest, Modbus_CalcResponseLength(MODBUS_FC_READ_INPUT_REGISTERS, Quantity))

    If UBound(baResponse) >= 0 Then
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

Public Function WriteSingleCoil(ByVal Address As Long, ByVal Value As Boolean) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte

    On Error GoTo ErrorHandler

    baRequest = BuildWriteSingleCoilRequest(Address, Value)
    baResponse = SendRequest(baRequest, 8)

    If UBound(baResponse) >= 0 Then
        ' Verify echo
        If baResponse(0) = m_SlaveID And _
           baResponse(1) = MB_MASTER_FC_WRITE_SINGLE_COIL Then
            WriteSingleCoil = True
        End If
    End If

    Exit Function
ErrorHandler:
    RaiseError "WriteSingleCoil failed: " & ERR.Description
End Function

Public Function WriteMultipleCoils(ByVal StartAddress As Long, ByRef Values() As Boolean) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim lQuantity As Long

    On Error GoTo ErrorHandler

    lQuantity = UBound(Values) + 1
    ValidateQuantity lQuantity, m_Defaults.MAX_COILS

    baRequest = BuildWriteMultipleCoilsRequest(StartAddress, Values)
    baResponse = SendRequest(baRequest, 8)

    If UBound(baResponse) >= 0 Then
        If baResponse(0) = m_SlaveID And _
           baResponse(1) = MB_MASTER_FC_WRITE_MULTIPLE_COILS Then
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

Public Function WriteSingleRegister(ByVal Address As Long, ByVal Value As Integer) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte

    On Error GoTo ErrorHandler

    baRequest = BuildWriteSingleRegisterRequest(Address, Value)
    baResponse = SendRequest(baRequest, 8)

    If UBound(baResponse) >= 0 Then
        If baResponse(0) = m_SlaveID And _
           baResponse(1) = MB_MASTER_FC_WRITE_SINGLE_REGISTER Then
            WriteSingleRegister = True
        End If
    End If

    Exit Function
ErrorHandler:
    RaiseError "WriteSingleRegister failed: " & ERR.Description
End Function

Public Function WriteMultipleRegisters(ByVal StartAddress As Long, ByRef Values() As Integer) As Boolean
    Dim baRequest() As Byte
    Dim baResponse() As Byte
    Dim lQuantity As Long

    On Error GoTo ErrorHandler

    lQuantity = UBound(Values) + 1
    ValidateQuantity lQuantity, m_Defaults.MAX_REGISTERS

    baRequest = BuildWriteMultipleRegistersRequest(StartAddress, Values)
    baResponse = SendRequest(baRequest, 8)

    If UBound(baResponse) >= 0 Then
        If baResponse(0) = m_SlaveID And _
           baResponse(1) = MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS Then
            WriteMultipleRegisters = True
        End If
    End If

    Exit Function
ErrorHandler:
    RaiseError "WriteMultipleRegisters failed: " & ERR.Description
End Function

'=========================================================================
' Private Methods - Connection
'=========================================================================

Private Sub ConnectRTU()
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
    m_TransportRTU.Timeout = m_ResponseTimeout

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
    m_State = MB_MASTER_STATE_ERROR
    RaiseError "ConnectRTU failed: " & ERR.Description
End Sub

Private Sub ConnectTCP()
    On Error GoTo ErrorHandler

    If LenB(m_TCPHost) = 0 Then
        RaiseError "TCP host not specified"
        Exit Sub
    End If

    If m_TCPPort = 0 Then m_TCPPort = m_Defaults.TCP_PORT

    Set m_TransportTCP = New cModbusTransportTCP
    m_TransportTCP.Timeout = m_ResponseTimeout
    m_TransportTCP.Connect m_TCPHost, m_TCPPort

    Exit Sub
ErrorHandler:
    m_State = MB_MASTER_STATE_ERROR
    RaiseError "ConnectTCP failed: " & ERR.Description
End Sub

'=========================================================================
' Private Methods - Request Building
'=========================================================================

Private Function BuildReadRequest(ByVal FC As ModbusMasterFunctionCode, ByVal StartAddress As Long, ByVal Quantity As Long) As Byte()
    Dim baPDU() As Byte

    ReDim baPDU(4) As Byte
    baPDU(0) = FC
    baPDU(1) = (StartAddress \ 256) And &HFF
    baPDU(2) = StartAddress And &HFF
    baPDU(3) = (Quantity \ 256) And &HFF
    baPDU(4) = Quantity And &HFF

    BuildReadRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteSingleCoilRequest(ByVal Address As Long, ByVal Value As Boolean) As Byte()
    Dim baPDU() As Byte

    ReDim baPDU(4) As Byte
    baPDU(0) = MB_MASTER_FC_WRITE_SINGLE_COIL
    baPDU(1) = (Address \ 256) And &HFF
    baPDU(2) = Address And &HFF
    If Value Then
        baPDU(3) = &HFF
        baPDU(4) = &H0
    Else
        baPDU(3) = &H0
        baPDU(4) = &H0
    End If

    BuildWriteSingleCoilRequest = BuildFrame(baPDU)
End Function

Private Function BuildWriteSingleRegisterRequest(ByVal Address As Long, ByVal Value As Integer) As Byte()
    Dim baPDU() As Byte

    ReDim baPDU(4) As Byte
    baPDU(0) = MB_MASTER_FC_WRITE_SINGLE_REGISTER
    baPDU(1) = (Address \ 256) And &HFF
    baPDU(2) = Address And &HFF
    baPDU(3) = (Value \ 256) And &HFF
    baPDU(4) = Value And &HFF

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

    ReDim baPDU(5 + lByteCount) As Byte

    baPDU(0) = MB_MASTER_FC_WRITE_MULTIPLE_COILS
    baPDU(1) = (StartAddress \ 256) And &HFF
    baPDU(2) = StartAddress And &HFF
    baPDU(3) = (lQuantity \ 256) And &HFF
    baPDU(4) = lQuantity And &HFF
    baPDU(5) = lByteCount

    ' Pack bits into bytes
    For i = 0 To lQuantity - 1
        lByteIndex = i \ 8
        lBitIndex = i Mod 8
        If Values(i) Then
            bByte = bByte Or (2 ^ lBitIndex)
        End If
        If lBitIndex = 7 Or i = lQuantity - 1 Then
            baPDU(6 + lByteIndex) = bByte
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

    ReDim baPDU(5 + lQuantity * 2) As Byte

    baPDU(0) = MB_MASTER_FC_WRITE_MULTIPLE_REGISTERS
    baPDU(1) = (StartAddress \ 256) And &HFF
    baPDU(2) = StartAddress And &HFF
    baPDU(3) = (lQuantity \ 256) And &HFF
    baPDU(4) = lQuantity And &HFF
    baPDU(5) = lQuantity * 2

    For i = 0 To lQuantity - 1
        baPDU(6 + i * 2) = (Values(i) \ 256) And &HFF
        baPDU(7 + i * 2) = Values(i) And &HFF
    Next i

    BuildWriteMultipleRegistersRequest = BuildFrame(baPDU)
End Function

Private Function BuildFrame(ByRef PDU() As Byte) As Byte()
    Dim baFrame() As Byte
    Dim lPDULen As Long

    lPDULen = UBound(PDU) + 1

    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        ' RTU: SlaveID + PDU + CRC16
        ReDim baFrame(1 + lPDULen + 1) As Byte
        baFrame(0) = m_SlaveID
        Modbus_CopyMemory baFrame(1), PDU(0), lPDULen
        Modbus_AppendCRC baFrame
    Else
        ' TCP: Use transport to build MBAP frame
        baFrame = m_TransportTCP.BuildMBAPFrame(PDU, m_SlaveID)
    End If

    BuildFrame = baFrame
End Function

'=========================================================================
' Private Methods - Request/Response
'=========================================================================

Private Function SendRequest(ByRef Request() As Byte, ByVal ExpectedResponseLen As Long) As Byte()
    Dim baResponse() As Byte

    On Error GoTo ErrorHandler

    If m_State <> MB_MASTER_STATE_CONNECTED Then
        RaiseError "Not connected"
        Exit Function
    End If

    If m_ProtocolType = MB_MASTER_PROTOCOL_RTU Then
        baResponse = m_TransportRTU.SendAndReceive(Request, ExpectedResponseLen)
    Else
        baResponse = m_TransportTCP.SendAndReceive(Request)
    End If

    ' Raise data received event
    If UBound(baResponse) >= 0 Then
        RaiseEvent OnDataReceived(baResponse)
    End If

    ' Check for exception response
    If UBound(baResponse) >= 1 Then
        If (baResponse(1) And &H80) <> 0 Then
            Dim lExceptionCode As Long
            lExceptionCode = baResponse(2)
            RaiseError "Modbus Exception: " & GetExceptionDescription(lExceptionCode)
            Exit Function
        End If
    End If

    SendRequest = baResponse
    Exit Function

ErrorHandler:
    RaiseError "SendRequest failed: " & ERR.Description
End Function

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Sub ValidateQuantity(ByVal Quantity As Long, ByVal MaxQuantity As Long)
    If Quantity < 1 Or Quantity > MaxQuantity Then
        ERR.Raise vbObjectError + 1, "cModbusMaster", "Invalid quantity: " & Quantity
    End If
End Sub

Private Function GetExceptionDescription(ByVal ExceptionCode As ModbusMasterExceptionCode) As String
    Select Case ExceptionCode
        Case MB_MASTER_EX_ILLEGAL_FUNCTION
            GetExceptionDescription = "Illegal Function"
        Case MB_MASTER_EX_ILLEGAL_DATA_ADDRESS
            GetExceptionDescription = "Illegal Data Address"
        Case MB_MASTER_EX_ILLEGAL_DATA_VALUE
            GetExceptionDescription = "Illegal Data Value"
        Case MB_MASTER_EX_SLAVE_DEVICE_FAILURE
            GetExceptionDescription = "Slave Device Failure"
        Case MB_MASTER_EX_ACKNOWLEDGE
            GetExceptionDescription = "Acknowledge"
        Case MB_MASTER_EX_SLAVE_DEVICE_BUSY
            GetExceptionDescription = "Slave Device Busy"
        Case MB_MASTER_EX_MEMORY_PARITY_ERROR
            GetExceptionDescription = "Memory Parity Error"
        Case MB_MASTER_EX_GATEWAY_PATH_UNAVAILABLE
            GetExceptionDescription = "Gateway Path Unavailable"
        Case MB_MASTER_EX_GATEWAY_TARGET_DEVICE_FAILED
            GetExceptionDescription = "Gateway Target Device Failed"
        Case Else
            GetExceptionDescription = "Unknown Exception (" & Hex$(ExceptionCode) & ")"
    End Select
End Function

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusMaster] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Transport Event Handlers
'=========================================================================

Private Sub m_TransportRTU_OnOpen()
    m_State = MB_MASTER_STATE_CONNECTED
    RaiseEvent OnConnect
End Sub

Private Sub m_TransportRTU_OnClose()
    m_State = MB_MASTER_STATE_DISCONNECTED
    RaiseEvent OnDisconnect
End Sub

Private Sub m_TransportRTU_OnError(ByVal Description As String)
    RaiseError Description
End Sub

Private Sub m_TransportTCP_OnConnect()
    m_State = MB_MASTER_STATE_CONNECTED
    RaiseEvent OnConnect
End Sub

Private Sub m_TransportTCP_OnDisconnect()
    m_State = MB_MASTER_STATE_DISCONNECTED
    RaiseEvent OnDisconnect
End Sub

Private Sub m_TransportTCP_OnError(ByVal Description As String)
    RaiseError Description
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_ProtocolType = MB_MASTER_PROTOCOL_TCP
    m_State = MB_MASTER_STATE_DISCONNECTED
    m_SlaveID = 1
    m_ResponseTimeout = MODBUS_DEFAULT_TIMEOUT

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
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    DisConnect
    Set m_TransportRTU = Nothing
    Set m_TransportTCP = Nothing
End Sub
