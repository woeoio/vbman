Attribute VB_Name = "mModbus"
'=========================================================================
'
' mModbus - Modbus Common Module for VB6
'
' Purpose: Provides common utilities for Modbus implementation
'          - CRC16 calculation
'          - Serial port operations
'          - Common constants
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

'=========================================================================
' API Declarations for Serial Port
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
    ByRef lpDCB As MODBUS_DCB) As Long

Private Declare Function SetCommTimeouts Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpCommTimeouts As MODBUS_COMMTIMEOUTS) As Long

Private Declare Function PurgeComm Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function ClearCommError Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpErrors As Long, _
    ByRef lpStat As MODBUS_COMSTAT) As Long

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

Private Const PURGE_TXABORT As Long = &H1
Private Const PURGE_RXABORT As Long = &H2
Private Const PURGE_TXCLEAR As Long = &H4
Private Const PURGE_RXCLEAR As Long = &H8

'=========================================================================
' Public Constants
'=========================================================================

Public Const MODBUS_INVALID_HANDLE As Long = -1
Public Const MODBUS_DEFAULT_TIMEOUT As Long = 1000
Public Const MODBUS_DEFAULT_BAUDRATE As Long = 9600
Public Const MODBUS_DEFAULT_DATABITS As Long = 8
Public Const MODBUS_DEFAULT_STOPBITS As Long = 1
Public Const MODBUS_TCP_PORT As Long = 502
Public Const MODBUS_TCP_MBAP_SIZE As Long = 7
Public Const MODBUS_MAX_PDU_SIZE As Long = 253
Public Const MODBUS_MAX_REGISTERS_PER_REQUEST As Long = 125
Public Const MODBUS_MAX_COILS_PER_REQUEST As Long = 2000
Public Const MODBUS_MAX_REGISTER_ADDRESS As Long = 65535
Public Const MODBUS_MAX_COIL_ADDRESS As Long = 65535

'=========================================================================
' Modbus Function Codes (Internal use)
'=========================================================================

Public Const MODBUS_FC_READ_COILS As Byte = &H1
Public Const MODBUS_FC_READ_DISCRETE_INPUTS As Byte = &H2
Public Const MODBUS_FC_READ_HOLDING_REGISTERS As Byte = &H3
Public Const MODBUS_FC_READ_INPUT_REGISTERS As Byte = &H4
Public Const MODBUS_FC_WRITE_SINGLE_COIL As Byte = &H5
Public Const MODBUS_FC_WRITE_SINGLE_REGISTER As Byte = &H6
Public Const MODBUS_FC_WRITE_MULTIPLE_COILS As Byte = &HF
Public Const MODBUS_FC_WRITE_MULTIPLE_REGISTERS As Byte = &H10

'=========================================================================
' Structure Definitions
'=========================================================================

Public Type MODBUS_DCB
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

Public Type MODBUS_COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

Public Type MODBUS_COMSTAT
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

'=========================================================================
' Public Functions - CRC Calculation
'=========================================================================

' Calculate Modbus CRC16
Public Function Modbus_CRC16(ByRef Data() As Byte, ByVal StartIndex As Long, ByVal Length As Long) As Long
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

    Modbus_CRC16 = wCRC
End Function

' Append CRC16 to data frame
Public Sub Modbus_AppendCRC(ByRef Data() As Byte)
    Dim wCRC As Long
    Dim lLen As Long

    lLen = UBound(Data) + 1
    wCRC = Modbus_CRC16(Data, 0, lLen - 2)
    Data(lLen - 2) = wCRC And &HFF
    Data(lLen - 1) = (wCRC \ 256) And &HFF
End Sub

' Verify CRC16 of received frame
Public Function Modbus_VerifyCRC(ByRef Data() As Byte) As Boolean
    Dim wCRC As Long
    Dim wRecvCRC As Long
    Dim lLen As Long

    lLen = UBound(Data) + 1
    If lLen < 4 Then
        Modbus_VerifyCRC = False
        Exit Function
    End If

    wCRC = Modbus_CRC16(Data, 0, lLen - 2)
    wRecvCRC = CLng(Data(lLen - 2)) + CLng(Data(lLen - 1)) * 256

    Modbus_VerifyCRC = (wCRC = wRecvCRC)
End Function

'=========================================================================
' Public Functions - Serial Port Operations
'=========================================================================

' Open serial port
Public Function Modbus_OpenSerialPort(ByVal PortName As String) As Long
    Dim hPort As Long

    hPort = CreateFile("\\.\" & PortName, _
                       GENERIC_READ Or GENERIC_WRITE, _
                       0, _
                       0, _
                       OPEN_EXISTING, _
                       FILE_ATTRIBUTE_NORMAL, _
                       0)

    Modbus_OpenSerialPort = hPort
End Function

' Close serial port
Public Sub Modbus_CloseSerialPort(ByVal hPort As Long)
    If hPort <> MODBUS_INVALID_HANDLE Then
        CloseHandle hPort
    End If
End Sub

' Configure serial port
Public Function Modbus_ConfigureSerialPort(ByVal hPort As Long, _
                                           ByVal BaudRate As Long, _
                                           ByVal DataBits As Long, _
                                           ByVal Parity As Long, _
                                           ByVal StopBits As Long, _
                                           ByVal Timeout As Long) As Boolean
    Dim dcb As MODBUS_DCB
    Dim timeouts As MODBUS_COMMTIMEOUTS

    On Error GoTo ErrorHandler

    If hPort = MODBUS_INVALID_HANDLE Then
        Modbus_ConfigureSerialPort = False
        Exit Function
    End If

    ' Configure DCB
    dcb.DCBlength = Len(dcb)
    dcb.BaudRate = BaudRate
    dcb.ByteSize = DataBits
    dcb.Parity = Parity

    ' Stop bits: 0 = 1 stop bit, 2 = 2 stop bits
    If StopBits = 2 Then
        dcb.StopBits = 2
    Else
        dcb.StopBits = 0
    End If

    ' Set fBitFields (fBinary = 1)
    dcb.fBitFields = &H1

    If SetCommState(hPort, dcb) = 0 Then
        Modbus_ConfigureSerialPort = False
        Exit Function
    End If

    ' Set timeouts
    timeouts.ReadIntervalTimeout = 0
    timeouts.ReadTotalTimeoutMultiplier = 0
    timeouts.ReadTotalTimeoutConstant = Timeout
    timeouts.WriteTotalTimeoutMultiplier = 0
    timeouts.WriteTotalTimeoutConstant = Timeout

    SetCommTimeouts hPort, timeouts

    Modbus_ConfigureSerialPort = True
    Exit Function

ErrorHandler:
    Modbus_ConfigureSerialPort = False
End Function

' Write data to serial port
Public Function Modbus_WriteSerialPort(ByVal hPort As Long, ByRef Data() As Byte) As Long
    Dim lBytesWritten As Long

    If hPort = MODBUS_INVALID_HANDLE Then
        Modbus_WriteSerialPort = 0
        Exit Function
    End If

    WriteFile hPort, Data(0), UBound(Data) + 1, lBytesWritten, 0
    Modbus_WriteSerialPort = lBytesWritten
End Function

' Read data from serial port
Public Function Modbus_ReadSerialPort(ByVal hPort As Long, ByRef Buffer() As Byte, ByVal MaxBytes As Long) As Long
    Dim lBytesRead As Long

    If hPort = MODBUS_INVALID_HANDLE Then
        Modbus_ReadSerialPort = 0
        Exit Function
    End If

    ReDim Buffer(MaxBytes - 1) As Byte
    ReadFile hPort, Buffer(0), MaxBytes, lBytesRead, 0

    If lBytesRead > 0 And lBytesRead < MaxBytes Then
        ReDim Preserve Buffer(lBytesRead - 1) As Byte
    End If

    Modbus_ReadSerialPort = lBytesRead
End Function

' Purge serial port buffers
Public Sub Modbus_PurgeSerialPort(ByVal hPort As Long)
    If hPort <> MODBUS_INVALID_HANDLE Then
        PurgeComm hPort, PURGE_TXCLEAR Or PURGE_RXCLEAR
    End If
End Sub

' Get bytes available in receive buffer
Public Function Modbus_GetBytesAvailable(ByVal hPort As Long) As Long
    Dim lErrors As Long
    Dim stat As MODBUS_COMSTAT

    If hPort = MODBUS_INVALID_HANDLE Then
        Modbus_GetBytesAvailable = 0
        Exit Function
    End If

    ClearCommError hPort, lErrors, stat
    Modbus_GetBytesAvailable = stat.cbInQue
End Function

'=========================================================================
' Public Functions - Timing
'=========================================================================

' Get current tick count
Public Function Modbus_GetTickCount() As Long
    Modbus_GetTickCount = GetTickCount()
End Function

' Sleep for specified milliseconds
Public Sub Modbus_Sleep(ByVal Milliseconds As Long)
    Sleep Milliseconds
End Sub

'=========================================================================
' Public Functions - Memory
'=========================================================================

' Copy memory
Public Sub Modbus_CopyMemory(ByRef Destination As Variant, ByRef Source As Variant, ByVal Length As Long)
    CopyMemory Destination, Source, Length
End Sub

'=========================================================================
' Public Functions - Frame Length Calculation
'=========================================================================

' Calculate expected RTU response frame length based on function code
Public Function Modbus_CalcResponseLength(ByVal FunctionCode As Byte, ByVal Quantity As Long) As Long
    Dim lLen As Long

    Select Case FunctionCode
        Case MODBUS_FC_READ_COILS, MODBUS_FC_READ_DISCRETE_INPUTS
            ' SlaveID(1) + FC(1) + ByteCount(1) + Data(n) + CRC(2)
            lLen = 5 + ((Quantity + 7) \ 8)

        Case MODBUS_FC_READ_HOLDING_REGISTERS, MODBUS_FC_READ_INPUT_REGISTERS
            ' SlaveID(1) + FC(1) + ByteCount(1) + Data(n*2) + CRC(2)
            lLen = 5 + (Quantity * 2)

        Case MODBUS_FC_WRITE_SINGLE_COIL, MODBUS_FC_WRITE_SINGLE_REGISTER
            ' SlaveID(1) + FC(1) + Addr(2) + Value(2) + CRC(2)
            lLen = 8

        Case MODBUS_FC_WRITE_MULTIPLE_COILS, MODBUS_FC_WRITE_MULTIPLE_REGISTERS
            ' SlaveID(1) + FC(1) + Addr(2) + Quantity(2) + CRC(2)
            lLen = 8

        Case Else
            lLen = 8
    End Select

    Modbus_CalcResponseLength = lLen
End Function

' Calculate expected RTU request frame length based on function code
Public Function Modbus_CalcRequestLength(ByVal FunctionCode As Byte, ByVal ByteCount As Long) As Long
    Dim lLen As Long

    Select Case FunctionCode
        Case MODBUS_FC_READ_COILS, MODBUS_FC_READ_DISCRETE_INPUTS, _
             MODBUS_FC_READ_HOLDING_REGISTERS, MODBUS_FC_READ_INPUT_REGISTERS
            ' SlaveID(1) + FC(1) + Addr(2) + Quantity(2) + CRC(2)
            lLen = 8

        Case MODBUS_FC_WRITE_SINGLE_COIL, MODBUS_FC_WRITE_SINGLE_REGISTER
            ' SlaveID(1) + FC(1) + Addr(2) + Value(2) + CRC(2)
            lLen = 8

        Case MODBUS_FC_WRITE_MULTIPLE_COILS, MODBUS_FC_WRITE_MULTIPLE_REGISTERS
            ' SlaveID(1) + FC(1) + Addr(2) + Quantity(2) + ByteCount(1) + Data(n) + CRC(2)
            lLen = 9 + ByteCount

        Case Else
            lLen = 8
    End Select

    Modbus_CalcRequestLength = lLen
End Function
