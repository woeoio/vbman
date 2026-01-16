VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModbusTransportRTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cModbusTransportRTU - Modbus RTU Transport Layer for VB6
'
' Purpose: Provides RTU (Serial) transport layer for Modbus communication
'          - Serial port management
'          - RTU frame handling (with CRC)
'          - Synchronous and asynchronous modes
'          - Non-blocking polling for slave mode
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

'=========================================================================
' Public Enums
'=========================================================================

Public Enum ModbusRTUParity
    MB_RTU_PARITY_NONE = 0
    MB_RTU_PARITY_ODD = 1
    MB_RTU_PARITY_EVEN = 2
End Enum

Public Enum ModbusRTUState
    MB_RTU_STATE_CLOSED = 0
    MB_RTU_STATE_OPEN = 1
    MB_RTU_STATE_ERROR = 2
End Enum

'=========================================================================
' Events
'=========================================================================

Event OnOpen()
Event OnClose()
Event OnError(ByVal Description As String)
Event OnDataReceived(Data() As Byte)

'=========================================================================
' Member Variables
'=========================================================================

Private m_hSerialPort As Long
Private m_SerialPort As String
Private m_BaudRate As Long
Private m_DataBits As Long
Private m_Parity As ModbusRTUParity
Private m_StopBits As Long
Private m_Timeout As Long
Private m_State As ModbusRTUState

' Receive buffer
Private m_Buffer As cByteBuffer

' Timer for polling mode (Slave)
Private WithEvents m_Timer As cTimer
Attribute m_Timer.VB_VarHelpID = -1
Private m_Polling As Boolean

'=========================================================================
' Properties
'=========================================================================

Public Property Get SerialPort() As String
    SerialPort = m_SerialPort
End Property

Public Property Let SerialPort(ByVal Value As String)
    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Cannot change serial port while open"
        Exit Property
    End If
    m_SerialPort = Value
End Property

Public Property Get BaudRate() As Long
    BaudRate = m_BaudRate
End Property

Public Property Let BaudRate(ByVal Value As Long)
    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Cannot change baud rate while open"
        Exit Property
    End If
    m_BaudRate = Value
End Property

Public Property Get DataBits() As Long
    DataBits = m_DataBits
End Property

Public Property Let DataBits(ByVal Value As Long)
    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Cannot change data bits while open"
        Exit Property
    End If
    m_DataBits = Value
End Property

Public Property Get Parity() As ModbusRTUParity
    Parity = m_Parity
End Property

Public Property Let Parity(ByVal Value As ModbusRTUParity)
    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Cannot change parity while open"
        Exit Property
    End If
    m_Parity = Value
End Property

Public Property Get StopBits() As Long
    StopBits = m_StopBits
End Property

Public Property Let StopBits(ByVal Value As Long)
    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Cannot change stop bits while open"
        Exit Property
    End If
    m_StopBits = Value
End Property

Public Property Get Timeout() As Long
    Timeout = m_Timeout
End Property

Public Property Let Timeout(ByVal Value As Long)
    If Value > 0 Then m_Timeout = Value
End Property

Public Property Get State() As ModbusRTUState
    State = m_State
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = (m_State = MB_RTU_STATE_OPEN)
End Property

'=========================================================================
' Public Methods - Connection
'=========================================================================

' Open serial port
Public Sub Open_(Optional ByVal PortName As String = "")
    On Error GoTo ErrorHandler

    If m_State = MB_RTU_STATE_OPEN Then
        RaiseError "Port already open"
        Exit Sub
    End If

    If LenB(PortName) > 0 Then m_SerialPort = PortName

    If LenB(m_SerialPort) = 0 Then
        RaiseError "Serial port not specified"
        Exit Sub
    End If

    ' Open port
    m_hSerialPort = Modbus_OpenSerialPort(m_SerialPort)

    If m_hSerialPort = MODBUS_INVALID_HANDLE Then
        m_State = MB_RTU_STATE_ERROR
        RaiseError "Failed to open serial port: " & m_SerialPort
        Exit Sub
    End If

    ' Configure port
    If Not Modbus_ConfigureSerialPort(m_hSerialPort, m_BaudRate, m_DataBits, m_Parity, m_StopBits, m_Timeout) Then
        Modbus_CloseSerialPort m_hSerialPort
        m_hSerialPort = MODBUS_INVALID_HANDLE
        m_State = MB_RTU_STATE_ERROR
        RaiseError "Failed to configure serial port"
        Exit Sub
    End If

    ' Purge buffers
    Modbus_PurgeSerialPort m_hSerialPort

    ' Clear receive buffer
    m_Buffer.Clear

    m_State = MB_RTU_STATE_OPEN
    RaiseEvent OnOpen

    Exit Sub
ErrorHandler:
    m_State = MB_RTU_STATE_ERROR
    RaiseError "Open failed: " & ERR.Description
End Sub

' Close serial port
Public Sub Close_()
    On Error Resume Next

    StopPolling

    If m_hSerialPort <> MODBUS_INVALID_HANDLE Then
        Modbus_CloseSerialPort m_hSerialPort
        m_hSerialPort = MODBUS_INVALID_HANDLE
    End If

    m_Buffer.Clear
    m_State = MB_RTU_STATE_CLOSED
    RaiseEvent OnClose
End Sub

'=========================================================================
' Public Methods - Data Transfer (Synchronous)
'=========================================================================

' Send request and wait for response (Master mode)
Public Function SendAndReceive(ByRef Request() As Byte, ByVal ExpectedResponseLen As Long) As Byte()
    Dim baResponse() As Byte
    Dim baBuffer() As Byte
    Dim lBytesRead As Long
    Dim lStartTime As Long
    Dim lElapsed As Long

    On Error GoTo ErrorHandler

    If m_State <> MB_RTU_STATE_OPEN Then
        RaiseError "Port not open"
        Exit Function
    End If

    ' Clear buffers
    m_Buffer.Clear
    Modbus_PurgeSerialPort m_hSerialPort

    ' Send request
    If Modbus_WriteSerialPort(m_hSerialPort, Request) <> UBound(Request) + 1 Then
        RaiseError "Failed to write all bytes"
        Exit Function
    End If

    ' Wait for response
    lStartTime = Modbus_GetTickCount()

    Do
        ' Read available data
        lBytesRead = Modbus_ReadSerialPort(m_hSerialPort, baBuffer, 256)

        If lBytesRead > 0 Then
            m_Buffer.Append baBuffer
        End If

        ' Check if we have enough data
        If m_Buffer.Size >= ExpectedResponseLen Then
            ' Try to parse and verify CRC
            baResponse = TryParseResponse(ExpectedResponseLen)
            If UBound(baResponse) >= 0 Then
                SendAndReceive = baResponse
                Exit Function
            End If
        End If

        ' Check timeout
        lElapsed = Modbus_GetTickCount() - lStartTime
        If lElapsed > m_Timeout Then
            RaiseError "Response timeout"
            Exit Function
        End If

        Modbus_Sleep 1
    Loop

    Exit Function
ErrorHandler:
    RaiseError "SendAndReceive failed: " & ERR.Description
End Function

' Send data without waiting for response
Public Function Send(ByRef Data() As Byte) As Boolean
    On Error GoTo ErrorHandler

    If m_State <> MB_RTU_STATE_OPEN Then
        RaiseError "Port not open"
        Send = False
        Exit Function
    End If

    Send = (Modbus_WriteSerialPort(m_hSerialPort, Data) = UBound(Data) + 1)
    Exit Function

ErrorHandler:
    RaiseError "Send failed: " & ERR.Description
    Send = False
End Function

'=========================================================================
' Public Methods - Polling Mode (for Slave)
'=========================================================================

' Start polling for incoming data
Public Sub StartPolling(Optional ByVal IntervalMs As Long = 10)
    If m_State <> MB_RTU_STATE_OPEN Then
        RaiseError "Port not open"
        Exit Sub
    End If

    If m_Polling Then Exit Sub

    Set m_Timer = New cTimer
    m_Timer.Interval = IntervalMs
    m_Timer.Enabled = True
    m_Polling = True
End Sub

' Stop polling
Public Sub StopPolling()
    m_Polling = False

    If Not m_Timer Is Nothing Then
        m_Timer.Enabled = False
        Set m_Timer = Nothing
    End If
End Sub

' Check if polling is active
Public Property Get IsPolling() As Boolean
    IsPolling = m_Polling
End Property

'=========================================================================
' Public Methods - Buffer Access
'=========================================================================

' Get current buffer size
Public Property Get BufferSize() As Long
    BufferSize = m_Buffer.Size
End Property

' Clear receive buffer
Public Sub ClearBuffer()
    m_Buffer.Clear
End Sub

' Peek byte at offset
Public Function PeekByte(ByVal Offset As Long) As Byte
    PeekByte = m_Buffer.PeekByte(Offset)
End Function

' Extract data from buffer
Public Function ExtractBuffer(ByVal Length As Long) As Byte()
    ExtractBuffer = m_Buffer.Extract(Length)
End Function

'=========================================================================
' Private Methods
'=========================================================================

Private Function TryParseResponse(ByVal ExpectedLen As Long) As Byte()
    Dim baFrame() As Byte
    Dim baResult() As Byte

    On Error GoTo ErrorHandler

    If m_Buffer.Size < ExpectedLen Then Exit Function

    ' Extract frame
    baFrame = m_Buffer.Extract(ExpectedLen)

    ' Verify CRC
    If Not Modbus_VerifyCRC(baFrame) Then
        ' CRC error, put data back? Or discard?
        ' For now, discard and return empty
        Exit Function
    End If

    ' Return frame without CRC
    ReDim baResult(ExpectedLen - 3) As Byte
    Modbus_CopyMemory baResult(0), baFrame(0), ExpectedLen - 2

    TryParseResponse = baResult
    Exit Function

ErrorHandler:
    ' Error parsing
End Function

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusTransportRTU] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Timer Event Handler
'=========================================================================

Private Sub m_Timer_Timer()
    Dim baBuffer() As Byte
    Dim lBytesRead As Long
    Dim baFrame() As Byte

    On Error Resume Next

    If m_State <> MB_RTU_STATE_OPEN Then Exit Sub
    If Not m_Polling Then Exit Sub

    ' Read available data
    lBytesRead = Modbus_ReadSerialPort(m_hSerialPort, baBuffer, 256)

    If lBytesRead > 0 Then
        m_Buffer.Append baBuffer

        ' Try to parse complete frames
        Do
            baFrame = TryParseFrame()
            If UBound(baFrame) < 0 Then Exit Do

            ' Raise event with received frame
            RaiseEvent OnDataReceived(baFrame)
        Loop
    End If
End Sub

Private Function TryParseFrame() As Byte()
    Dim bFC As Byte
    Dim lByteCount As Long
    Dim lExpectedLen As Long
    Dim baFrame() As Byte

    On Error GoTo ErrorHandler

    ' Minimum frame: SlaveID + FC + CRC (4 bytes)
    If m_Buffer.Size < 4 Then Exit Function

    ' Peek function code
    bFC = m_Buffer.PeekByte(1)

    ' Check for exception response
    If (bFC And &H80) <> 0 Then
        lExpectedLen = 5  ' SlaveID + FC + ExCode + CRC(2)
    Else
        ' Calculate expected length based on function code
        Select Case bFC
            Case MODBUS_FC_READ_COILS, MODBUS_FC_READ_DISCRETE_INPUTS, _
                 MODBUS_FC_READ_HOLDING_REGISTERS, MODBUS_FC_READ_INPUT_REGISTERS
                ' Request: SlaveID + FC + Addr(2) + Qty(2) + CRC(2) = 8
                lExpectedLen = 8

            Case MODBUS_FC_WRITE_SINGLE_COIL, MODBUS_FC_WRITE_SINGLE_REGISTER
                ' Request/Response: SlaveID + FC + Addr(2) + Value(2) + CRC(2) = 8
                lExpectedLen = 8

            Case MODBUS_FC_WRITE_MULTIPLE_COILS, MODBUS_FC_WRITE_MULTIPLE_REGISTERS
                ' Need to peek byte count
                If m_Buffer.Size < 7 Then Exit Function
                lByteCount = m_Buffer.PeekByte(6)
                lExpectedLen = 9 + lByteCount  ' SlaveID + FC + Addr(2) + Qty(2) + ByteCount + Data + CRC(2)

            Case Else
                lExpectedLen = 8
        End Select
    End If

    If m_Buffer.Size < lExpectedLen Then Exit Function

    ' Extract frame
    baFrame = m_Buffer.Extract(lExpectedLen)

    ' Verify CRC
    If Not Modbus_VerifyCRC(baFrame) Then
        ' CRC error, discard first byte and try again
        ' This helps resync after corruption
        Exit Function
    End If

    TryParseFrame = baFrame
    Exit Function

ErrorHandler:
    ' Error parsing, clear buffer to resync
    m_Buffer.Clear
End Function

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_hSerialPort = MODBUS_INVALID_HANDLE
    m_State = MB_RTU_STATE_CLOSED
    m_BaudRate = MODBUS_DEFAULT_BAUDRATE
    m_DataBits = MODBUS_DEFAULT_DATABITS
    m_Parity = MB_RTU_PARITY_NONE
    m_StopBits = MODBUS_DEFAULT_STOPBITS
    m_Timeout = MODBUS_DEFAULT_TIMEOUT
    m_Polling = False

    Set m_Buffer = New cByteBuffer
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Close_
    Set m_Buffer = Nothing
    Set m_Timer = Nothing
End Sub
