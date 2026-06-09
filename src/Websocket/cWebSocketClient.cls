VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebSocketClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cWebSocketClient - WebSocket Client for VB6
'
' Purpose: Provides WebSocket client functionality
'          - Connect to WebSocket servers
'          - Send/receive text and binary messages
'          - Automatic Ping/Pong handling
'          - Message fragmentation support
'
' Author: Claude (Anthropic)
' Date: 2026-01-10
'
'=========================================================================
Option Explicit

'=========================================================================
' Helper Function for Tick Count
'=========================================================================

Private Declare Function GetTickCount Lib "kernel32" () As Long
'=========================================================================
' Events
'=========================================================================

' Connection established
Event OnOpen()

' Connection closed
Event OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)

' Text message received
Event OnTextMessage(ByVal Message As String)

' Binary message received
Event OnBinaryMessage(Data() As Byte)

' Error occurred
Event OnError(ByVal Description As String)

' Pong received (for latency measurement)
Event OnPong(Data() As Byte)

'=========================================================================
' Member Variables
'=========================================================================

Private WithEvents m_Socket As cWinsock
Attribute m_Socket.VB_VarHelpID = -1

Private m_State As WsState
Private m_RecvBuffer As cByteBuffer
Private m_FrameParser As cWebSocketFrame

' URL components
Private m_URL As String
Private m_Host As String
Private m_Port As Long
Private m_Path As String
Private m_Query As String
Private m_IsSecure As Boolean

' Handshake
Private m_SecretKey As String
Private m_SubProtocol As String

' Fragmented message assembly
Private m_FragmentBuffer As cByteBuffer
Private m_FragmentOpCode As WsOpCode
Private m_IsFragmented As Boolean

' Auto ping
Private m_AutoPing As Boolean
Private m_PingInterval As Long
Private m_LastPingTime As Long

'=========================================================================
' Properties
'=========================================================================

' Current connection state
Public Property Get State() As WsState
    State = m_State
End Property

' Connected URL
Public Property Get url() As String
    url = m_URL
End Property

' Server host
Public Property Get Host() As String
    Host = m_Host
End Property

' Server port
Public Property Get Port() As Long
    Port = m_Port
End Property

' Enable auto ping
Public Property Get AutoPing() As Boolean
    AutoPing = m_AutoPing
End Property

Public Property Let AutoPing(ByVal Value As Boolean)
    m_AutoPing = Value
End Property

' Ping interval in milliseconds
Public Property Get PingInterval() As Long
    PingInterval = m_PingInterval
End Property

Public Property Let PingInterval(ByVal Value As Long)
    If Value > 0 Then m_PingInterval = Value
End Property

'=========================================================================
' Public Methods - Connection
'=========================================================================

' Connect to WebSocket server
' URL format: ws[s]://host[:port][/path][?query]
Public Sub Connect(ByVal WebSocketURL As String, Optional ByVal SubProtocol As String = "")
    On Error GoTo ErrorHandler
    
    ' Validate state
    If m_State <> WS_STATE_CLOSED Then
        RaiseError "Already connected or connecting"
        Exit Sub
    End If
    
    ' Parse URL
    If Not ParseURL(WebSocketURL) Then
        RaiseError "Invalid WebSocket URL: " & WebSocketURL
        Exit Sub
    End If
    
    ' Check for secure WebSocket (not implemented)
    ' wss:// 现在已支持，通过 cWinsock.Tls 启用加密
    
    m_SubProtocol = SubProtocol
    m_SecretKey = GenerateWebSocketKey()
    
    ' Initialize buffers
    m_RecvBuffer.Reset
    m_FragmentBuffer.Reset
    m_IsFragmented = False
    
    ' Set state
    m_State = WS_STATE_CONNECTING
    
    ' Create and connect socket
    Set m_Socket = New cWinsock
    m_Socket.Protocol = wskTCPProtocol
    
    ' wss:// 自动启用 TLS
    If m_IsSecure Then
        m_Socket.Tls
    End If
    
    m_Socket.Connect m_Host, m_Port
    
    Exit Sub
ErrorHandler:
    m_State = WS_STATE_CLOSED
    RaiseError "Connect failed: " & ERR.Description
End Sub

' Close connection
Public Sub CloseConnection(Optional ByVal Code As WsCloseCode = WS_CLOSE_NORMAL, _
                          Optional ByVal Reason As String = "")
    On Error Resume Next

    If m_State = WS_STATE_CLOSED Then Exit Sub

    ' Send close frame if connected
    If m_State = WS_STATE_OPEN Then
        m_State = WS_STATE_CLOSING
        Dim baFrame() As Byte
        baFrame = m_FrameParser.BuildCloseFrame(Code, Reason, True)
        m_Socket.SendData baFrame
    End If

    ' Close socket
    If Not m_Socket Is Nothing Then
        m_Socket.Close_
    End If

    m_State = WS_STATE_CLOSED

    ' Raise close event
    If LenB(Reason) = 0 Then
        Reason = GetCloseCodeDescription(Code)
    End If
    RaiseEvent OnClose(Code, Reason)
End Sub

'=========================================================================
' Public Methods - Sending
'=========================================================================

' Send text message
Public Sub SendText(ByVal Message As String)
    If m_State <> WS_STATE_OPEN Then
        RaiseError "WebSocket is not connected"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildTextFrame(Message, True)  ' Client must mask
    m_Socket.SendData baFrame

    Exit Sub
ErrorHandler:
    RaiseError "SendText failed: " & ERR.Description
End Sub

' Send binary message
Public Sub SendBinary(Data() As Byte)
    If m_State <> WS_STATE_OPEN Then
        RaiseError "WebSocket is not connected"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildBinaryFrame(Data, True)  ' Client must mask
    m_Socket.SendData baFrame

    Exit Sub
ErrorHandler:
    RaiseError "SendBinary failed: " & ERR.Description
End Sub

' Send ping
Public Sub SendPing(Optional Payload As Variant)
    If m_State <> WS_STATE_OPEN Then Exit Sub

    On Error GoTo ErrorHandler

    Dim baPayload() As Byte
    Dim baFrame() As Byte

    If Not IsMissing(Payload) Then
        If varType(Payload) = vbString Then
            baPayload = StringToUTF8(CStr(Payload))
        ElseIf varType(Payload) = vbByte + vbArray Then
            baPayload = Payload
        End If
    End If

    baFrame = m_FrameParser.BuildPingFrame(baPayload, True)
    m_Socket.SendData baFrame

    Exit Sub
ErrorHandler:
    RaiseError "SendPing failed: " & ERR.Description
End Sub

'=========================================================================
' Private Methods - URL Parsing
'=========================================================================

Private Function ParseURL(ByVal url As String) As Boolean
    Dim lPos As Long
    Dim sScheme As String
    Dim sRest As String
    Dim sHostPart As String

    url = Trim$(url)
    If LenB(url) = 0 Then Exit Function

    ' Parse scheme
    lPos = InStr(1, url, "://", vbTextCompare)
    If lPos = 0 Then Exit Function

    sScheme = LCase$(Left$(url, lPos - 1))

    If sScheme = "ws" Then
        m_IsSecure = False
        m_Port = 80
    ElseIf sScheme = "wss" Then
        m_IsSecure = True
        m_Port = 443
    Else
        Exit Function
    End If

    ' Parse rest
    sRest = Mid$(url, lPos + 3)

    ' Find path separator
    lPos = InStr(sRest, "/")
    If lPos = 0 Then lPos = InStr(sRest, "?")

    If lPos > 0 Then
        sHostPart = Left$(sRest, lPos - 1)
        m_Path = Mid$(sRest, lPos)

        ' Separate path and query
        lPos = InStr(m_Path, "?")
        If lPos > 0 Then
            m_Query = Mid$(m_Path, lPos + 1)
            m_Path = Left$(m_Path, lPos - 1)
        Else
            m_Query = ""
        End If
    Else
        sHostPart = sRest
        m_Path = "/"
        m_Query = ""
    End If

    ' Parse host and port
    lPos = InStr(sHostPart, ":")
    If lPos > 0 Then
        m_Host = Left$(sHostPart, lPos - 1)
        m_Port = CLng(Mid$(sHostPart, lPos + 1))
    Else
        m_Host = sHostPart
    End If

    m_URL = url
    ParseURL = True
End Function

'=========================================================================
' Private Methods - Handshake
'=========================================================================

Private Sub SendHandshake()
    Dim sHandshake As String
    Dim sHostHeader As String

    ' Build host header
    sHostHeader = m_Host
    If (Not m_IsSecure And m_Port <> 80) Or (m_IsSecure And m_Port <> 443) Then
        sHostHeader = sHostHeader & ":" & m_Port
    End If

    ' Build request path
    sHandshake = "GET " & m_Path
    If LenB(m_Query) > 0 Then
        sHandshake = sHandshake & "?" & m_Query
    End If
    sHandshake = sHandshake & " HTTP/1.1" & vbCrLf

    ' Headers
    sHandshake = sHandshake & "Host: " & sHostHeader & vbCrLf
    sHandshake = sHandshake & "Upgrade: websocket" & vbCrLf
    sHandshake = sHandshake & "Connection: Upgrade" & vbCrLf
    sHandshake = sHandshake & "Sec-WebSocket-Key: " & m_SecretKey & vbCrLf
    sHandshake = sHandshake & "Sec-WebSocket-Version: " & WS_VERSION & vbCrLf

    If LenB(m_SubProtocol) > 0 Then
        sHandshake = sHandshake & "Sec-WebSocket-Protocol: " & m_SubProtocol & vbCrLf
    End If

    sHandshake = sHandshake & vbCrLf

    m_Socket.SendData sHandshake, wcpUtf8
End Sub

Private Function ValidateHandshakeResponse(ByVal Response As String) As Boolean
    Dim sAccept As String
    Dim sExpected As String

    ' Check for 101 status
    If InStr(1, Response, "HTTP/1.1 101", vbTextCompare) <> 1 Then
        Exit Function
    End If

    ' Check Upgrade header
    If InStr(1, Response, "Upgrade: websocket", vbTextCompare) = 0 Then
        Exit Function
    End If

    ' Validate accept key
    sAccept = GetHeaderValue(Response, "Sec-WebSocket-Accept")
    If LenB(sAccept) = 0 Then Exit Function

    sExpected = ComputeAcceptKey(m_SecretKey)
    If sAccept <> sExpected Then Exit Function

    ValidateHandshakeResponse = True
End Function

'=========================================================================
' Private Methods - Frame Processing
'=========================================================================

Private Sub ProcessReceivedData()
    On Error GoTo ErrorHandler

    Do While m_RecvBuffer.Size >= 2
        ' Parse frame header
        If Not m_FrameParser.ParseHeader(m_RecvBuffer) Then
            Exit Do  ' Need more data
        End If

        ' Check if complete frame available
        If Not m_FrameParser.IsCompleteFrame(m_RecvBuffer) Then
            Exit Do  ' Need more data
        End If

        ' Extract payload (consumes frame from buffer)
        Dim baPayload() As Byte
        baPayload = m_FrameParser.ExtractPayload(m_RecvBuffer)

        ' Process frame by opcode
        Select Case m_FrameParser.OpCode
            Case WS_OPCODE_TEXT, WS_OPCODE_BINARY
                ProcessDataFrame baPayload, m_FrameParser.OpCode, m_FrameParser.FIN

            Case WS_OPCODE_CONTINUATION
                ProcessContinuationFrame baPayload, m_FrameParser.FIN

            Case WS_OPCODE_CLOSE
                ProcessCloseFrame baPayload

            Case WS_OPCODE_PING
                ProcessPingFrame baPayload

            Case WS_OPCODE_PONG
                ProcessPongFrame baPayload
        End Select
    Loop

    Exit Sub
ErrorHandler:
    RaiseError "ProcessReceivedData: " & ERR.Description
    m_RecvBuffer.Clear
End Sub

Private Sub ProcessDataFrame(ByRef Payload() As Byte, ByVal OpCode As WsOpCode, ByVal IsFinal As Boolean)
    If IsFinal And Not m_IsFragmented Then
        ' Complete standalone message
        DeliverMessage Payload, OpCode
    Else
        ' Start of fragmented message
        m_IsFragmented = True
        m_FragmentOpCode = OpCode
        m_FragmentBuffer.Clear

        On Error Resume Next
        If UBound(Payload) >= 0 Then
            m_FragmentBuffer.Append Payload
        End If
        On Error GoTo 0

        If IsFinal Then
            ' Single fragment (unusual but valid)
            DeliverFragmentedMessage
        End If
    End If
End Sub

Private Sub ProcessContinuationFrame(ByRef Payload() As Byte, ByVal IsFinal As Boolean)
    If Not m_IsFragmented Then
        RaiseError "Unexpected continuation frame"
        Exit Sub
    End If

    ' Append to fragment buffer
    On Error Resume Next
    If UBound(Payload) >= 0 Then
        m_FragmentBuffer.Append Payload
    End If
    On Error GoTo 0

    If IsFinal Then
        DeliverFragmentedMessage
    End If
End Sub

Private Sub DeliverFragmentedMessage()
    Dim baData() As Byte
    baData = m_FragmentBuffer.ToArray

    DeliverMessage baData, m_FragmentOpCode

    m_FragmentBuffer.Clear
    m_IsFragmented = False
End Sub

Private Sub DeliverMessage(ByRef Data() As Byte, ByVal OpCode As WsOpCode)
    On Error Resume Next

    If OpCode = WS_OPCODE_TEXT Then
        Dim sText As String
        If UBound(Data) >= 0 Then
            sText = UTF8ToString(Data)
        Else
            sText = ""
        End If
        RaiseEvent OnTextMessage(sText)
    Else
        RaiseEvent OnBinaryMessage(Data)
    End If
End Sub

Private Sub ProcessCloseFrame(ByRef Payload() As Byte)
    Dim lCode As WsCloseCode
    Dim sReason As String

    On Error Resume Next

    ' Parse close code and reason
    If UBound(Payload) >= 1 Then
        lCode = CLng(Payload(0)) * 256 + CLng(Payload(1))

        If UBound(Payload) >= 2 Then
            Dim baReason() As Byte
            ReDim baReason(UBound(Payload) - 2) As Byte
            Dim i As Long
            For i = 2 To UBound(Payload)
                baReason(i - 2) = Payload(i)
            Next i
            sReason = UTF8ToString(baReason)
        End If
    Else
        lCode = WS_CLOSE_NO_STATUS
    End If

    On Error GoTo 0

    ' Close connection
    CloseConnection lCode, sReason
End Sub

Private Sub ProcessPingFrame(ByRef Payload() As Byte)
    ' Auto-reply with pong
    On Error Resume Next

    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildPongFrame(Payload, True)
    m_Socket.SendData baFrame
End Sub

Private Sub ProcessPongFrame(ByRef Payload() As Byte)
    RaiseEvent OnPong(Payload)
End Sub

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cWebSocketClient] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Socket Event Handlers
'=========================================================================

Private Sub m_Socket_Connect(Client As cWinsock)
    ' Socket connected, send WebSocket handshake
    SendHandshake
End Sub

Private Sub m_Socket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    On Error GoTo ErrorHandler

    Dim baData() As Byte
    Dim sData As String

    If m_State = WS_STATE_CONNECTING Then
        ' Expecting handshake response
        Client.GetData sData, vbString, , wcpUtf8

        If ValidateHandshakeResponse(sData) Then
            m_State = WS_STATE_OPEN
            m_LastPingTime = GetTickCount()
            RaiseEvent OnOpen
        Else
            m_State = WS_STATE_CLOSED
            RaiseError "Handshake failed"
            Client.Close_
        End If

    ElseIf m_State = WS_STATE_OPEN Or m_State = WS_STATE_CLOSING Then
        ' Receive WebSocket data
        Client.GetData baData, vbByte + vbArray

        On Error Resume Next
        If UBound(baData) >= 0 Then
            m_RecvBuffer.Append baData
        End If
        On Error GoTo ErrorHandler

        ProcessReceivedData
    End If

    Exit Sub
ErrorHandler:
    RaiseError "DataArrival: " & ERR.Description
End Sub

Private Sub m_Socket_CloseEvent(Client As cWinsock)
    If m_State <> WS_STATE_CLOSED Then
        m_State = WS_STATE_CLOSED
        RaiseEvent OnClose(WS_CLOSE_ABNORMAL, "Connection closed unexpectedly")
    End If
End Sub

Private Sub m_Socket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    m_State = WS_STATE_CLOSED
    RaiseError "Socket error: " & Description
End Sub



'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_State = WS_STATE_CLOSED
    m_AutoPing = False
    m_PingInterval = 30000
    
    Set m_RecvBuffer = New cByteBuffer
    Set m_FragmentBuffer = New cByteBuffer
    Set m_FrameParser = New cWebSocketFrame
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CloseConnection WS_CLOSE_GOING_AWAY
    Set m_Socket = Nothing
    Set m_RecvBuffer = Nothing
    Set m_FragmentBuffer = Nothing
    Set m_FrameParser = Nothing
End Sub
