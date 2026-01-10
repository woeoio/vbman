VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cByteBuffer - High Performance Byte Buffer for VB6
'
' Purpose: Provides efficient byte buffer management with pre-allocation
'          and minimal memory reallocation operations.
'
' Features:
'   - Pre-allocated buffer with automatic growth
'   - Efficient append and consume operations
'   - Peek without consuming
'   - Thread-safe design considerations
'
' Author: Claude (Anthropic)
' Date: 2026-01-10
'
'=========================================================================
Option Explicit

'=========================================================================
' API Declarations
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants
'=========================================================================

Private Const INITIAL_CAPACITY As Long = 4096
Private Const GROWTH_FACTOR As Double = 1.5

'=========================================================================
' Member Variables
'=========================================================================

Private m_Data() As Byte
Private m_Capacity As Long
Private m_Size As Long

'=========================================================================
' Properties
'=========================================================================

' Current data size in buffer
Public Property Get Size() As Long
    Size = m_Size
End Property

' Current buffer capacity
Public Property Get Capacity() As Long
    Capacity = m_Capacity
End Property

' Check if buffer is empty
Public Property Get IsEmpty() As Boolean
    IsEmpty = (m_Size = 0)
End Property

'=========================================================================
' Public Methods
'=========================================================================

' Append data to buffer
Public Sub Append(ByRef Data() As Byte)
    Dim lDataLen As Long

    On Error Resume Next
    lDataLen = UBound(Data) + 1
    If ERR.Number <> 0 Then
        ERR.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If lDataLen <= 0 Then Exit Sub

    ' Ensure capacity
    EnsureCapacity m_Size + lDataLen

    ' Copy data
    CopyMemory m_Data(m_Size), Data(0), lDataLen
    m_Size = m_Size + lDataLen
End Sub

' Append single byte
Public Sub AppendByte(ByVal Value As Byte)
    EnsureCapacity m_Size + 1
    m_Data(m_Size) = Value
    m_Size = m_Size + 1
End Sub

' Peek data without consuming (read-only view)
Public Function Peek(ByVal Offset As Long, ByVal Length As Long, ByRef OutData() As Byte) As Boolean
    If Offset < 0 Or Length <= 0 Or Offset + Length > m_Size Then
        Peek = False
        Exit Function
    End If

    ReDim OutData(Length - 1) As Byte
    CopyMemory OutData(0), m_Data(Offset), Length
    Peek = True
End Function

' Peek single byte
Public Function PeekByte(ByVal Offset As Long) As Byte
    If Offset >= 0 And Offset < m_Size Then
        PeekByte = m_Data(Offset)
    End If
End Function

' Consume (remove) data from front of buffer
Public Sub Consume(ByVal Length As Long)
    If Length <= 0 Then Exit Sub

    If Length >= m_Size Then
        ' Clear all
        m_Size = 0
    Else
        ' Shift remaining data to front
        CopyMemory m_Data(0), m_Data(Length), m_Size - Length
        m_Size = m_Size - Length
    End If
End Sub

' Extract and consume data
Public Function Extract(ByVal Length As Long) As Byte()
    Dim baResult() As Byte

    If Length <= 0 Or Length > m_Size Then
        Extract = baResult
        Exit Function
    End If

    ReDim baResult(Length - 1) As Byte
    CopyMemory baResult(0), m_Data(0), Length
    Consume Length

    Extract = baResult
End Function

' Get all data as array (copy)
Public Function ToArray() As Byte()
    Dim baResult() As Byte

    If m_Size > 0 Then
        ReDim baResult(m_Size - 1) As Byte
        CopyMemory baResult(0), m_Data(0), m_Size
    End If

    ToArray = baResult
End Function

' Clear buffer
Public Sub Clear()
    m_Size = 0
End Sub

' Reset buffer to initial state
Public Sub Reset()
    m_Size = 0
    m_Capacity = INITIAL_CAPACITY
    ReDim m_Data(m_Capacity - 1) As Byte
End Sub

' Direct access to internal buffer (for performance-critical operations)
' WARNING: Do not modify beyond m_Size
Public Function GetBufferPtr() As Long
    If m_Capacity > 0 Then
        GetBufferPtr = VarPtr(m_Data(0))
    End If
End Function

'=========================================================================
' Private Methods
'=========================================================================

' Ensure buffer has enough capacity
Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim lNewCapacity As Long

    If RequiredCapacity <= m_Capacity Then Exit Sub

    ' Calculate new capacity
    lNewCapacity = m_Capacity
    Do While lNewCapacity < RequiredCapacity
        lNewCapacity = CLng(lNewCapacity * GROWTH_FACTOR)
        If lNewCapacity < RequiredCapacity Then
            lNewCapacity = RequiredCapacity
        End If
    Loop

    ' Resize buffer
    ReDim Preserve m_Data(lNewCapacity - 1) As Byte
    m_Capacity = lNewCapacity
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_Capacity = INITIAL_CAPACITY
    m_Size = 0
    ReDim m_Data(m_Capacity - 1) As Byte
End Sub

Private Sub Class_Terminate()
    Erase m_Data
End Sub
