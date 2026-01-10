VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebSocketFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cWebSocketFrame - WebSocket Frame Parser for VB6
'
' Purpose: Provides frame parsing and building for WebSocket protocol
'          This class is READ-ONLY for parsing - it does not modify input data
'
' Features:
'   - Parse frame header without modifying source data
'   - Extract and unmask payload separately
'   - Build frames for sending
'   - Support for fragmented messages
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
' Frame Header Info (parsed result)
'=========================================================================

' Frame header information (public readable properties)
Private m_FIN As Boolean
Private m_RSV1 As Boolean
Private m_RSV2 As Boolean
Private m_RSV3 As Boolean
Private m_OpCode As WsOpCode
Private m_HasMask As Boolean
Private m_MaskingKey(3) As Byte
Private m_PayloadLength As Long
Private m_HeaderLength As Long
Private m_TotalFrameLength As Long
Private m_IsValid As Boolean
Private m_ErrorMessage As String

'=========================================================================
' Properties - Frame Header Info
'=========================================================================

Public Property Get FIN() As Boolean
    FIN = m_FIN
End Property

Public Property Get RSV1() As Boolean
    RSV1 = m_RSV1
End Property

Public Property Get RSV2() As Boolean
    RSV2 = m_RSV2
End Property

Public Property Get RSV3() As Boolean
    RSV3 = m_RSV3
End Property

Public Property Get OpCode() As WsOpCode
    OpCode = m_OpCode
End Property

Public Property Get HasMask() As Boolean
    HasMask = m_HasMask
End Property

Public Property Get PayloadLength() As Long
    PayloadLength = m_PayloadLength
End Property

Public Property Get HeaderLength() As Long
    HeaderLength = m_HeaderLength
End Property

Public Property Get TotalFrameLength() As Long
    TotalFrameLength = m_TotalFrameLength
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_ErrorMessage
End Property

' Check if this is a control frame
Public Property Get IsControlFrame() As Boolean
    IsControlFrame = (m_OpCode >= WS_OPCODE_CLOSE)
End Property

' Check if this is a data frame
Public Property Get IsDataFrame() As Boolean
    IsDataFrame = (m_OpCode <= WS_OPCODE_BINARY)
End Property

'=========================================================================
' Public Methods - Parsing
'=========================================================================

' Parse frame header from buffer (READ-ONLY, does not modify buffer)
' Returns True if header was successfully parsed
' Note: This only parses the header, call ExtractPayload to get payload
Public Function ParseHeader(ByRef Buffer As cByteBuffer) As Boolean
    Dim bByte0 As Byte
    Dim bByte1 As Byte
    Dim lLenType As Long
    Dim lOffset As Long
    
    ' Reset state
    ResetState
    
    ' Need at least 2 bytes for minimal header
    If Buffer.Size < 2 Then
        m_ErrorMessage = "Insufficient data for frame header"
        ParseHeader = False
        Exit Function
    End If
    
    ' Read first two bytes
    bByte0 = Buffer.PeekByte(0)
    bByte1 = Buffer.PeekByte(1)
    
    ' Parse first byte: FIN, RSV1-3, Opcode
    m_FIN = ((bByte0 And &H80) = &H80)
    m_RSV1 = ((bByte0 And &H40) = &H40)
    m_RSV2 = ((bByte0 And &H20) = &H20)
    m_RSV3 = ((bByte0 And &H10) = &H10)
    m_OpCode = bByte0 And &HF
    
    ' Parse second byte: Mask flag, Payload length
    m_HasMask = ((bByte1 And &H80) = &H80)
    lLenType = bByte1 And &H7F
    
    lOffset = 2
    
    ' Parse extended payload length
    If lLenType = 126 Then
        ' 16-bit length
        If Buffer.Size < 4 Then
            m_ErrorMessage = "Insufficient data for 16-bit length"
            ParseHeader = False
            Exit Function
        End If
        m_PayloadLength = CLng(Buffer.PeekByte(2)) * 256 + CLng(Buffer.PeekByte(3))
        lOffset = 4
        
    ElseIf lLenType = 127 Then
        ' 64-bit length (we only support 32-bit in VB6)
        If Buffer.Size < 10 Then
            m_ErrorMessage = "Insufficient data for 64-bit length"
            ParseHeader = False
            Exit Function
        End If
        ' Check high 4 bytes are zero (we can't handle > 2GB)
        If Buffer.PeekByte(2) <> 0 Or Buffer.PeekByte(3) <> 0 Or _
            Buffer.PeekByte(4) <> 0 Or Buffer.PeekByte(5) <> 0 Then
            m_ErrorMessage = "Payload too large (>2GB not supported)"
            ParseHeader = False
            Exit Function
        End If
        ' Read lower 4 bytes (big-endian)
        m_PayloadLength = CLng(Buffer.PeekByte(6)) * &H1000000 + _
        CLng(Buffer.PeekByte(7)) * &H10000 + _
        CLng(Buffer.PeekByte(8)) * &H100& + _
        CLng(Buffer.PeekByte(9))
        lOffset = 10
    Else
        ' 7-bit length
        m_PayloadLength = lLenType
    End If
    
    ' Parse masking key if present
    If m_HasMask Then
        If Buffer.Size < lOffset + 4 Then
            m_ErrorMessage = "Insufficient data for masking key"
            ParseHeader = False
            Exit Function
        End If
        m_MaskingKey(0) = Buffer.PeekByte(lOffset)
        m_MaskingKey(1) = Buffer.PeekByte(lOffset + 1)
        m_MaskingKey(2) = Buffer.PeekByte(lOffset + 2)
        m_MaskingKey(3) = Buffer.PeekByte(lOffset + 3)
        lOffset = lOffset + 4
    End If
    
    m_HeaderLength = lOffset
    m_TotalFrameLength = m_HeaderLength + m_PayloadLength
    m_IsValid = True
    
    ParseHeader = True
End Function

' Check if complete frame is available in buffer
Public Function IsCompleteFrame(ByRef Buffer As cByteBuffer) As Boolean
    If Not m_IsValid Then
        IsCompleteFrame = False
        Exit Function
    End If
    IsCompleteFrame = (Buffer.Size >= m_TotalFrameLength)
End Function

' Extract and unmask payload from buffer
' This consumes the entire frame from the buffer
Public Function ExtractPayload(ByRef Buffer As cByteBuffer) As Byte()
    Dim baPayload() As Byte
    Dim baFrame() As Byte
    Dim i As Long

    If Not m_IsValid Then
        ExtractPayload = baPayload
        Exit Function
    End If

    If Buffer.Size < m_TotalFrameLength Then
        ExtractPayload = baPayload
        Exit Function
    End If

    ' Extract entire frame
    baFrame = Buffer.Extract(m_TotalFrameLength)

    ' Extract payload portion
    If m_PayloadLength > 0 Then
        ReDim baPayload(m_PayloadLength - 1) As Byte
        CopyMemory baPayload(0), baFrame(m_HeaderLength), m_PayloadLength

        ' Unmask if needed
        If m_HasMask Then
            For i = 0 To m_PayloadLength - 1
                baPayload(i) = baPayload(i) Xor m_MaskingKey(i Mod 4)
            Next i
        End If
    End If

    ExtractPayload = baPayload
End Function

' Skip this frame in buffer (consume without extracting payload)
Public Sub SkipFrame(ByRef Buffer As cByteBuffer)
    If m_IsValid And Buffer.Size >= m_TotalFrameLength Then
        Buffer.Consume m_TotalFrameLength
    End If
End Sub

'=========================================================================
' Public Methods - Frame Building
'=========================================================================

' Build a WebSocket frame (for sending)
' UseMask should be True for client->server, False for server->client
Public Function BuildFrame(ByRef Payload() As Byte, ByVal OpCode As WsOpCode, _
                          Optional ByVal UseMask As Boolean = False, _
                          Optional ByVal IsFinal As Boolean = True) As Byte()
    Dim baFrame() As Byte
    Dim lPayloadLen As Long
    Dim lHeaderSize As Long
    Dim lOffset As Long
    Dim baMaskKey(3) As Byte
    Dim i As Long

    ' Get payload length
    On Error Resume Next
    lPayloadLen = UBound(Payload) + 1
    If ERR.Number <> 0 Then
        ERR.Clear
        lPayloadLen = 0
    End If
    On Error GoTo 0

    ' Calculate header size
    lHeaderSize = 2
    If lPayloadLen < 126 Then
        ' 7-bit length
    ElseIf lPayloadLen <= 65535 Then
        lHeaderSize = lHeaderSize + 2  ' 16-bit length
    Else
        lHeaderSize = lHeaderSize + 8  ' 64-bit length
    End If

    If UseMask Then
        lHeaderSize = lHeaderSize + 4  ' Masking key
    End If

    ' Allocate frame
    ReDim baFrame(lHeaderSize + lPayloadLen - 1) As Byte
    lOffset = 0

    ' Byte 0: FIN + RSV + Opcode
    If IsFinal Then
        baFrame(0) = &H80 Or (OpCode And &HF)
    Else
        baFrame(0) = OpCode And &HF
    End If
    lOffset = 1

    ' Byte 1: Mask bit + Payload length
    If UseMask Then
        baFrame(1) = &H80
    Else
        baFrame(1) = 0
    End If

    If lPayloadLen < 126 Then
        baFrame(1) = baFrame(1) Or lPayloadLen
        lOffset = 2
    ElseIf lPayloadLen <= 65535 Then
        baFrame(1) = baFrame(1) Or 126
        baFrame(2) = (lPayloadLen And &HFF00&) \ &H100&
        baFrame(3) = lPayloadLen And &HFF&
        lOffset = 4
    Else
        baFrame(1) = baFrame(1) Or 127
        ' 64-bit length (big-endian, only using lower 32 bits)
        baFrame(2) = 0
        baFrame(3) = 0
        baFrame(4) = 0
        baFrame(5) = 0
        baFrame(6) = (lPayloadLen And &HFF000000) \ &H1000000
        baFrame(7) = (lPayloadLen And &HFF0000) \ &H10000
        baFrame(8) = (lPayloadLen And &HFF00&) \ &H100&
        baFrame(9) = lPayloadLen And &HFF&
        lOffset = 10
    End If

    ' Add masking key if needed
    If UseMask Then
        Randomize Timer
        baMaskKey(0) = CByte(Int(Rnd * 256))
        baMaskKey(1) = CByte(Int(Rnd * 256))
        baMaskKey(2) = CByte(Int(Rnd * 256))
        baMaskKey(3) = CByte(Int(Rnd * 256))

        baFrame(lOffset) = baMaskKey(0)
        baFrame(lOffset + 1) = baMaskKey(1)
        baFrame(lOffset + 2) = baMaskKey(2)
        baFrame(lOffset + 3) = baMaskKey(3)
        lOffset = lOffset + 4

        ' Copy and mask payload
        For i = 0 To lPayloadLen - 1
            baFrame(lOffset + i) = Payload(i) Xor baMaskKey(i Mod 4)
        Next i
    Else
        ' Copy payload without masking
        If lPayloadLen > 0 Then
            CopyMemory baFrame(lOffset), Payload(0), lPayloadLen
        End If
    End If

    BuildFrame = baFrame
End Function

' Build a text frame
Public Function BuildTextFrame(ByVal Text As String, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
    Dim baPayload() As Byte
    baPayload = StringToUTF8(Text)
    BuildTextFrame = BuildFrame(baPayload, WS_OPCODE_TEXT, UseMask, True)
End Function

' Build a binary frame
Public Function BuildBinaryFrame(ByRef Data() As Byte, _
                                Optional ByVal UseMask As Boolean = False) As Byte()
    BuildBinaryFrame = BuildFrame(Data, WS_OPCODE_BINARY, UseMask, True)
End Function

' Build a close frame
Public Function BuildCloseFrame(Optional ByVal StatusCode As WsCloseCode = WS_CLOSE_NORMAL, _
                               Optional ByVal Reason As String = "", _
                               Optional ByVal UseMask As Boolean = False) As Byte()
    Dim baPayload() As Byte
    Dim baReason() As Byte
    Dim lReasonLen As Long

    If LenB(Reason) > 0 Then
        baReason = StringToUTF8(Reason)
        On Error Resume Next
        lReasonLen = UBound(baReason) + 1
        If ERR.Number <> 0 Then lReasonLen = 0
        On Error GoTo 0
    End If

    ReDim baPayload(1 + lReasonLen) As Byte
    ' Status code (big-endian)
    baPayload(0) = (StatusCode And &HFF00&) \ &H100&
    baPayload(1) = StatusCode And &HFF&

    If lReasonLen > 0 Then
        CopyMemory baPayload(2), baReason(0), lReasonLen
    End If

    BuildCloseFrame = BuildFrame(baPayload, WS_OPCODE_CLOSE, UseMask, True)
End Function

' Build a ping frame
Public Function BuildPingFrame(ByRef Payload() As Byte, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
    Dim baEmpty() As Byte
    On Error Resume Next
    If UBound(Payload) >= 0 Then
        BuildPingFrame = BuildFrame(Payload, WS_OPCODE_PING, UseMask, True)
    Else
        BuildPingFrame = BuildFrame(baEmpty, WS_OPCODE_PING, UseMask, True)
    End If
End Function

' Build a pong frame
Public Function BuildPongFrame(ByRef Payload() As Byte, _
                              Optional ByVal UseMask As Boolean = False) As Byte()
    Dim baEmpty() As Byte
    On Error Resume Next
    If UBound(Payload) >= 0 Then
        BuildPongFrame = BuildFrame(Payload, WS_OPCODE_PONG, UseMask, True)
    Else
        BuildPongFrame = BuildFrame(baEmpty, WS_OPCODE_PONG, UseMask, True)
    End If
End Function

'=========================================================================
' Private Methods
'=========================================================================

Private Sub ResetState()
    m_FIN = False
    m_RSV1 = False
    m_RSV2 = False
    m_RSV3 = False
    m_OpCode = WS_OPCODE_CONTINUATION
    m_HasMask = False
    m_MaskingKey(0) = 0
    m_MaskingKey(1) = 0
    m_MaskingKey(2) = 0
    m_MaskingKey(3) = 0
    m_PayloadLength = 0
    m_HeaderLength = 0
    m_TotalFrameLength = 0
    m_IsValid = False
    m_ErrorMessage = ""
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    ResetState
End Sub
