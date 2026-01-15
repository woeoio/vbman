Attribute VB_Name = "mModbus"
'=========================================================================
'
' mModbus - Modbus Protocol Constants
'
' Purpose: Provides Modbus protocol constants (must be in .bas file for DLL compilation)
'
' Author: Auto
' Date: 2026-01-21
'
'=========================================================================
Option Explicit

'=========================================================================
' Modbus Constants
'=========================================================================

' Modbus TCP Default Port
Public Const MB_TCP_PORT As Long = 502

' Modbus RTU Default Settings
Public Const MB_RTU_DEFAULT_BAUDRATE As Long = 9600
Public Const MB_RTU_DEFAULT_DATABITS As Long = 8
Public Const MB_RTU_DEFAULT_PARITY As String = "N"                              ' N=None, E=Even, O=Odd
Public Const MB_RTU_DEFAULT_STOPBITS As Long = 1
Public Const MB_RTU_DEFAULT_TIMEOUT As Long = 1000                              ' milliseconds

' Modbus RTU Frame Timing (3.5 character times)
' For 9600 baud: 3.5 * 11 bits / 9600 = ~4ms
Public Const MB_RTU_FRAME_DELAY_MS As Long = 4

' Maximum Modbus PDU size
Public Const MB_MAX_PDU_SIZE As Long = 253
Public Const MB_MAX_REGISTERS As Long = 125
Public Const MB_MAX_COILS As Long = 2000

' Modbus TCP MBAP Header size
Public Const MB_TCP_MBAP_SIZE As Long = 7
