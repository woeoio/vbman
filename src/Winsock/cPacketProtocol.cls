VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPacketProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cPacketProtocol - Packet Protocol Handler for cWinsock
' by woeoio@qq.com
'
' Purpose: Solve TCP packet fragmentation and reassembly problems.
'          Provides built-in delimiter, fixed-length, and length-header
'          protocols with automatic buffering.
'
'=========================================================================
Option Explicit
    DefObj A-Z
Private Const MODULE_NAME As String = "cPacketProtocol"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'--- Protocol Type Enum ---
Public Enum PacketProtocolType
    ppNone = 0                ' No protocol (raw data)
    ppDelimiter = 1           ' Delimiter-based protocol
    ppFixedLength = 2         ' Fixed-length protocol
    ppLengthHeader = 3        ' Length-header protocol
End Enum

'--- Endian Enum ---
Public Enum EndianEnum
    eeLittleEndian = 0        ' Little-endian (x86 default)
    eeBigEndian = 1           ' Big-endian (network byte order)
End Enum

'--- Member Variables ---
Private m_eProtocol As PacketProtocolType
Private m_baRecvBuffer() As Byte       ' Internal receive buffer

' Delimiter protocol params
Private m_sDelimiter As String         ' Delimiter string (e.g. vbCrLf, vbNullChar)
Private m_baDelimiter() As Byte        ' Delimiter as byte array (for binary matching)

' Fixed-length protocol params
Private m_lFixedLength As Long         ' Fixed message length

' Length-header protocol params
Private m_nHeaderBytes As Long         ' Header size: 2 or 4 bytes
Private m_eEndian As EndianEnum        ' Byte order for header

'=========================================================================
' Properties
'=========================================================================

Public Property Get ProtocolType() As PacketProtocolType
    ProtocolType = m_eProtocol
End Property

Public Property Let ProtocolType(ByVal Value As PacketProtocolType)
    m_eProtocol = Value
    Clear
End Property

Public Property Get Delimiter() As String
    Delimiter = m_sDelimiter
End Property

Public Property Let Delimiter(ByVal Value As String)
    m_sDelimiter = Value
    m_baDelimiter = StrConv(Value, vbFromUnicode)
End Property

Public Property Get FixedLength() As Long
    FixedLength = m_lFixedLength
End Property

Public Property Let FixedLength(ByVal Value As Long)
    m_lFixedLength = Value
End Property

Public Property Get HeaderBytes() As Long
    HeaderBytes = m_nHeaderBytes
End Property

Public Property Let HeaderBytes(ByVal Value As Long)
    If Value <> 2 And Value <> 4 Then
        ERR.Raise vbObjectError, MODULE_NAME, "HeaderBytes must be 2 or 4"
    End If
    m_nHeaderBytes = Value
End Property

Public Property Get Endian() As EndianEnum
    Endian = m_eEndian
End Property

Public Property Let Endian(ByVal Value As EndianEnum)
    m_eEndian = Value
End Property

' Check if there is pending (incomplete) data in buffer
Public Property Get HasPendingData() As Boolean
    On Error Resume Next
    HasPendingData = (UBound(m_baRecvBuffer) >= 0)
End Property

' Get pending buffer size
Public Property Get PendingBytes() As Long
    On Error Resume Next
    PendingBytes = UBound(m_baRecvBuffer) + 1
End Property

'=========================================================================
' Public Methods
'=========================================================================

' Encode (packet framing): add protocol markers to data before sending
Public Function Encode(ByRef baData() As Byte) As Byte()
    Dim baResult() As Byte
    Dim lDataLen As Long
    Dim lOffset As Long
    
    On Error GoTo EH
    
    lDataLen = UBound(baData) + 1
    
    Select Case m_eProtocol
    Case ppNone
        Encode = baData
        Exit Function
        
    Case ppDelimiter
        ' Append delimiter bytes
        Dim lDelimLen As Long
        lDelimLen = UBound(m_baDelimiter) + 1
        ReDim baResult(0 To lDataLen + lDelimLen - 1) As Byte
        CopyMemory baResult(0), baData(0), lDataLen
        CopyMemory baResult(lDataLen), m_baDelimiter(0), lDelimLen
        
    Case ppFixedLength
        ' Pad or truncate to fixed length
        If lDataLen >= m_lFixedLength Then
            ReDim baResult(0 To m_lFixedLength - 1) As Byte
            CopyMemory baResult(0), baData(0), m_lFixedLength
        Else
            ReDim baResult(0 To m_lFixedLength - 1) As Byte
            CopyMemory baResult(0), baData(0), lDataLen
            ' Remaining bytes are zero (padding)
        End If
        
    Case ppLengthHeader
        ' Prepend length header
        ReDim baResult(0 To m_nHeaderBytes + lDataLen - 1) As Byte
        If m_nHeaderBytes = 2 Then
            Dim nLen As Integer
            nLen = CInt(lDataLen)
            If m_eEndian = eeLittleEndian Then
                CopyMemory baResult(0), nLen, 2
            Else
                baResult(0) = (nLen And &HFF00&) \ &H100&
                baResult(1) = nLen And &HFF&
            End If
        Else ' 4 bytes
            If m_eEndian = eeLittleEndian Then
                CopyMemory baResult(0), lDataLen, 4
            Else
                baResult(0) = (lDataLen And &HFF000000) \ &H1000000
                baResult(1) = (lDataLen And &HFF0000) \ &H10000
                baResult(2) = (lDataLen And &HFF00&) \ &H100&
                baResult(3) = lDataLen And &HFF&
            End If
        End If
        CopyMemory baResult(m_nHeaderBytes), baData(0), lDataLen
        
    Case Else
        Encode = baData
        Exit Function
    End Select
    
    Encode = baResult
    Exit Function
EH:
    Encode = baData
End Function

' Decode (packet deframing): parse received data into complete messages
' Returns a Collection of Byte() arrays (each is a complete message)
' Returns Nothing if no complete message available
Public Function Decode(ByRef baData() As Byte) As Collection
    Dim colResult As New Collection
    Dim baWork() As Byte
    Dim lWorkLen As Long
    
    On Error GoTo EH
    
    ' Merge new data into internal buffer
    Dim bHasData As Boolean
    bHasData = False
    On Error Resume Next
    bHasData = (UBound(baData) >= 0)
    On Error GoTo EH
    
    If bHasData Then
        If UBound(m_baRecvBuffer) >= 0 Then
            Dim lOldLen As Long
            lOldLen = UBound(m_baRecvBuffer) + 1
            lWorkLen = lOldLen + UBound(baData) + 1
            ReDim Preserve m_baRecvBuffer(0 To lWorkLen - 1) As Byte
            CopyMemory m_baRecvBuffer(lOldLen), baData(0), UBound(baData) + 1
        Else
            m_baRecvBuffer = baData
        End If
    End If
    
    ' Extract complete messages
    Select Case m_eProtocol
    Case ppDelimiter
        ExtractDelimiterMessages colResult
        
    Case ppFixedLength
        ExtractFixedLengthMessages colResult
        
    Case ppLengthHeader
        ExtractLengthHeaderMessages colResult
        
    Case Else
        ' ppNone: treat entire buffer as one message
        If UBound(m_baRecvBuffer) >= 0 Then
            colResult.Add m_baRecvBuffer
            m_baRecvBuffer = vbNullString
        End If
    End Select
    
    If colResult.count > 0 Then
        Set Decode = colResult
    Else
        Set Decode = Nothing
    End If
    
    Exit Function
EH:
    Set Decode = Nothing
End Function

' Clear internal buffer
Public Sub Clear()
    m_baRecvBuffer = vbNullString
End Sub

'=========================================================================
' Private Methods - Message Extraction
'=========================================================================

Private Sub ExtractDelimiterMessages(colResult As Collection)
    Dim lDelimLen As Long
    Dim lBufLen As Long
    Dim lPos As Long
    Dim lSearchStart As Long
    Dim baMessage() As Byte
    
    On Error Resume Next
    
    lDelimLen = UBound(m_baDelimiter) + 1
    If lDelimLen = 0 Then Exit Sub
    
    lBufLen = UBound(m_baRecvBuffer) + 1
    If lBufLen = 0 Then Exit Sub
    
    Do
        lPos = FindDelimiter(m_baRecvBuffer, lSearchStart, m_baDelimiter)
        If lPos < 0 Then Exit Do
        
        ' Extract message (without delimiter)
        If lPos > 0 Then
            ReDim baMessage(0 To lPos - 1) As Byte
            CopyMemory baMessage(0), m_baRecvBuffer(0), lPos
        Else
            ' Empty message (just delimiter)
            ReDim baMessage(0 To -1) As Byte
        End If
        
        colResult.Add baMessage
        
        ' Remove processed data (message + delimiter) from buffer
        Dim lRemaining As Long
        lRemaining = lBufLen - lPos - lDelimLen
        If lRemaining > 0 Then
            Dim baTemp() As Byte
            ReDim baTemp(0 To lRemaining - 1) As Byte
            CopyMemory baTemp(0), m_baRecvBuffer(lPos + lDelimLen), lRemaining
            m_baRecvBuffer = baTemp
            lBufLen = lRemaining
            lSearchStart = 0
        Else
            m_baRecvBuffer = vbNullString
            Exit Do
        End If
    Loop
End Sub

Private Sub ExtractFixedLengthMessages(colResult As Collection)
    Dim lBufLen As Long
    Dim baMessage() As Byte
    Dim lRemaining As Long
    Dim baTemp() As Byte
    
    On Error Resume Next
    
    If m_lFixedLength <= 0 Then Exit Sub
    
    lBufLen = UBound(m_baRecvBuffer) + 1
    
    Do While lBufLen >= m_lFixedLength
        ' Extract fixed-length message
        ReDim baMessage(0 To m_lFixedLength - 1) As Byte
        CopyMemory baMessage(0), m_baRecvBuffer(0), m_lFixedLength
        colResult.Add baMessage
        
        ' Remove from buffer
        lRemaining = lBufLen - m_lFixedLength
        If lRemaining > 0 Then
            ReDim baTemp(0 To lRemaining - 1) As Byte
            CopyMemory baTemp(0), m_baRecvBuffer(m_lFixedLength), lRemaining
            m_baRecvBuffer = baTemp
            lBufLen = lRemaining
        Else
            m_baRecvBuffer = vbNullString
            Exit Do
        End If
    Loop
End Sub

Private Sub ExtractLengthHeaderMessages(colResult As Collection)
    Dim lBufLen As Long
    Dim lMsgLen As Long
    Dim baMessage() As Byte
    Dim lRemaining As Long
    Dim baTemp() As Byte
    
    On Error Resume Next
    
    lBufLen = UBound(m_baRecvBuffer) + 1
    
    Do
        ' Need at least header bytes
        If lBufLen < m_nHeaderBytes Then Exit Do
        
        ' Parse length from header
        lMsgLen = ParseLengthHeader()
        
        ' Validate length
        If lMsgLen < 0 Or lMsgLen > 10485760 Then  ' Max 10MB sanity check
            ' Invalid header, discard buffer
            m_baRecvBuffer = vbNullString
            Exit Do
        End If
        
        ' Check if complete message is in buffer
        If lBufLen < m_nHeaderBytes + lMsgLen Then Exit Do
        
        ' Extract message (without header)
        If lMsgLen > 0 Then
            ReDim baMessage(0 To lMsgLen - 1) As Byte
            CopyMemory baMessage(0), m_baRecvBuffer(m_nHeaderBytes), lMsgLen
        Else
            ReDim baMessage(0 To -1) As Byte
        End If
        colResult.Add baMessage
        
        ' Remove from buffer
        lRemaining = lBufLen - m_nHeaderBytes - lMsgLen
        If lRemaining > 0 Then
            ReDim baTemp(0 To lRemaining - 1) As Byte
            CopyMemory baTemp(0), m_baRecvBuffer(m_nHeaderBytes + lMsgLen), lRemaining
            m_baRecvBuffer = baTemp
            lBufLen = lRemaining
        Else
            m_baRecvBuffer = vbNullString
            Exit Do
        End If
    Loop
End Sub

'=========================================================================
' Private Helpers
'=========================================================================

' Find delimiter byte sequence in buffer, starting from lStart
' Returns position (0-based) or -1 if not found
Private Function FindDelimiter(ByRef baBuffer() As Byte, ByVal lStart As Long, ByRef baDelim() As Byte) As Long
    Dim lBufLen As Long
    Dim lDelimLen As Long
    Dim i As Long
    Dim j As Long
    Dim bMatch As Boolean
    
    FindDelimiter = -1
    
    On Error Resume Next
    
    lBufLen = UBound(baBuffer) + 1
    lDelimLen = UBound(baDelim) + 1
    
    If lDelimLen = 0 Or lBufLen < lDelimLen Then Exit Function
    
    For i = lStart To lBufLen - lDelimLen
        bMatch = True
        For j = 0 To lDelimLen - 1
            If baBuffer(i + j) <> baDelim(j) Then
                bMatch = False
                Exit For
            End If
        Next j
        If bMatch Then
            FindDelimiter = i
            Exit Function
        End If
    Next i
End Function

' Parse length value from the beginning of the buffer
Private Function ParseLengthHeader() As Long
    Dim nVal2 As Integer
    Dim lVal4 As Long
    
    On Error GoTo EH
    
    If m_nHeaderBytes = 2 Then
        If m_eEndian = eeLittleEndian Then
            CopyMemory nVal2, m_baRecvBuffer(0), 2
            ParseLengthHeader = nVal2 And &HFFFF&
        Else
            ParseLengthHeader = CLng(m_baRecvBuffer(0)) * 256& + CLng(m_baRecvBuffer(1))
        End If
    Else ' 4 bytes
        If m_eEndian = eeLittleEndian Then
            CopyMemory lVal4, m_baRecvBuffer(0), 4
            ' Ensure positive value
            If lVal4 < 0 Then
                ParseLengthHeader = lVal4 And &H7FFFFFFF
            Else
                ParseLengthHeader = lVal4
            End If
        Else
            ParseLengthHeader = CLng(m_baRecvBuffer(0)) * &H1000000 + _
                                CLng(m_baRecvBuffer(1)) * &H10000 + _
                                CLng(m_baRecvBuffer(2)) * &H100& + _
                                CLng(m_baRecvBuffer(3))
        End If
    End If
    
    Exit Function
EH:
    ParseLengthHeader = -1
End Function

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_eProtocol = ppNone
    m_sDelimiter = vbCrLf
    m_baDelimiter = StrConv(vbCrLf, vbFromUnicode)
    m_lFixedLength = 0
    m_nHeaderBytes = 4
    m_eEndian = eeLittleEndian
    m_baRecvBuffer = vbNullString
End Sub

Private Sub Class_Terminate()
    Clear
End Sub
