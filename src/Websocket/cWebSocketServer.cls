VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebSocketServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cWebSocketServer - WebSocket Server for VB6
'
' Purpose: Provides WebSocket server functionality
'          - Listen for client connections
'          - Handle multiple clients
'          - Broadcast messages
'          - Per-client messaging
'
' Author: Claude (Anthropic)
' Date: 2026-01-10
'
'=========================================================================
Option Explicit

'=========================================================================
' Events
'=========================================================================

' Client connected (after successful handshake)
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)

' Client disconnected
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)

' Text message received from client
Event OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)

' Binary message received from client
Event OnClientBinaryMessage(ByVal ClientID As String, Data() As Byte)

' Error occurred
Event OnError(ByVal Description As String)

' Server started
Event OnStart(ByVal Port As Long)

' Server stopped
Event OnStop()

'=========================================================================
' Member Variables
'=========================================================================

Private WithEvents m_ListenSocket As cWinsock
Attribute m_ListenSocket.VB_VarHelpID = -1

Private m_Clients As Collection                                                 ' Key: ClientID, Value: cWebSocketServerClient
Private m_ClientCounter As Long
Private m_ListenPort As Long
Private m_IsListening As Boolean

' Frame builder (shared for efficiency)
Private m_FrameBuilder As cWebSocketFrame

'=========================================================================
' Properties
'=========================================================================

' Listening port
Public Property Get Port() As Long
    Port = m_ListenPort
End Property

' Is server listening
Public Property Get IsListening() As Boolean
    IsListening = m_IsListening
End Property

' Connected client count
Public Property Get ClientCount() As Long
    On Error Resume Next
    If Not m_Clients Is Nothing Then
        ClientCount = m_Clients.Count
    End If
End Property

' Get client IDs as array
Public Property Get ClientIDs() As Variant
    Dim vResult() As String
    Dim oClient As cWebSocketServerClient
    Dim i As Long

    On Error Resume Next
    If m_Clients Is Nothing Or m_Clients.Count = 0 Then
        ClientIDs = vResult
        Exit Property
    End If

    ReDim vResult(m_Clients.Count - 1)
    i = 0
    For Each oClient In m_Clients
        vResult(i) = oClient.ID
        i = i + 1
    Next

    ClientIDs = vResult
End Property

'=========================================================================
' Public Methods - Server Control
'=========================================================================

' Start listening on specified port
Public Sub Listen(Optional ByVal Port As Long = 8080)
    On Error GoTo ErrorHandler
    
    ' Stop if already listening
    If m_IsListening Then
        StopServer
    End If
    
    m_ListenPort = Port
    
    ' Create listening socket
    Set m_ListenSocket = New cWinsock
    m_ListenSocket.Protocol = wskTCPProtocol
    m_ListenSocket.Listen Port
    
    ' Initialize clients collection
    Set m_Clients = New Collection
    m_ClientCounter = 0
    
    m_IsListening = True
    RaiseEvent OnStart(Port)
    
    Exit Sub
ErrorHandler:
    m_IsListening = False
    RaiseError "Listen failed: " & ERR.Description
End Sub

' Stop server
Public Sub StopServer()
    On Error Resume Next

    ' Close all client connections
    If Not m_Clients Is Nothing Then
        Dim oClient As cWebSocketServerClient
        For Each oClient In m_Clients
            oClient.SendClose WS_CLOSE_GOING_AWAY, "Server shutting down"
            oClient.CloseSocket
        Next
        Set m_Clients = Nothing
    End If

    ' Close listening socket
    If Not m_ListenSocket Is Nothing Then
        m_ListenSocket.Close_
        Set m_ListenSocket = Nothing
    End If

    m_IsListening = False
    RaiseEvent OnStop
End Sub

'=========================================================================
' Public Methods - Messaging
'=========================================================================

' Send text message to specific client
Public Sub SendText(ByVal ClientID As String, ByVal Message As String)
    On Error GoTo ErrorHandler

    Dim oClient As cWebSocketServerClient
    Set oClient = GetClient(ClientID)

    If Not oClient Is Nothing Then
        If oClient.State = WS_CLIENT_CONNECTED Then
            oClient.SendText Message
        End If
    End If

    Exit Sub
ErrorHandler:
    RaiseError "SendText failed: " & ERR.Description
End Sub

' Send binary message to specific client
Public Sub SendBinary(ByVal ClientID As String, Data() As Byte)
    On Error GoTo ErrorHandler

    Dim oClient As cWebSocketServerClient
    Set oClient = GetClient(ClientID)

    If Not oClient Is Nothing Then
        If oClient.State = WS_CLIENT_CONNECTED Then
            oClient.SendBinary Data
        End If
    End If

    Exit Sub
ErrorHandler:
    RaiseError "SendBinary failed: " & ERR.Description
End Sub

' Broadcast text message to all connected clients
Public Sub BroadcastText(ByVal Message As String, Optional ByVal ExcludeClientID As String = "")
    On Error Resume Next

    Dim oClient As cWebSocketServerClient
    Dim baFrame() As Byte

    ' Build frame once
    baFrame = m_FrameBuilder.BuildTextFrame(Message, False)

    For Each oClient In m_Clients
        If oClient.State = WS_CLIENT_CONNECTED Then
            If LenB(ExcludeClientID) = 0 Or oClient.ID <> ExcludeClientID Then
                oClient.SendFrame baFrame
            End If
        End If
    Next
End Sub

' Broadcast binary message to all connected clients
Public Sub BroadcastBinary(Data() As Byte, Optional ByVal ExcludeClientID As String = "")
    On Error Resume Next

    Dim oClient As cWebSocketServerClient
    Dim baFrame() As Byte

    ' Build frame once
    baFrame = m_FrameBuilder.BuildBinaryFrame(Data, False)

    For Each oClient In m_Clients
        If oClient.State = WS_CLIENT_CONNECTED Then
            If LenB(ExcludeClientID) = 0 Or oClient.ID <> ExcludeClientID Then
                oClient.SendFrame baFrame
            End If
        End If
    Next
End Sub

' Disconnect specific client
Public Sub DisconnectClient(ByVal ClientID As String, _
                           Optional ByVal Code As WsCloseCode = WS_CLOSE_NORMAL, _
                           Optional ByVal Reason As String = "")
    On Error Resume Next

    Dim oClient As cWebSocketServerClient
    Set oClient = GetClient(ClientID)

    If Not oClient Is Nothing Then
        oClient.SendClose Code, Reason
        oClient.CloseSocket
    End If
End Sub

'=========================================================================
' Private Methods - Client Management
'=========================================================================

Private Function GetClient(ByVal ClientID As String) As cWebSocketServerClient
    On Error Resume Next
    Set GetClient = m_Clients(ClientID)
End Function

Private Function FindClientBySocket(ByVal Socket As cWinsock) As cWebSocketServerClient
    Dim oClient As cWebSocketServerClient

    On Error Resume Next
    For Each oClient In m_Clients
        If oClient.Socket Is Socket Then
            Set FindClientBySocket = oClient
            Exit Function
        End If
    Next
End Function

Private Sub RemoveClient(ByVal ClientID As String, ByVal Reason As String)
    On Error Resume Next

    Dim oClient As cWebSocketServerClient
    Set oClient = m_Clients(ClientID)

    If Not oClient Is Nothing Then
        oClient.CloseSocket
        m_Clients.Remove ClientID
        RaiseEvent OnClientDisconnect(ClientID, Reason)
    End If
End Sub

'=========================================================================
' Private Methods - Handshake
'=========================================================================

Private Function HandleHandshake(ByVal Client As cWebSocketServerClient, ByVal Request As String) As Boolean
    Dim sKey As String
    Dim sAccept As String
    Dim sResponse As String

    On Error GoTo ErrorHandler

    ' Check for WebSocket upgrade request
    If InStr(1, Request, "Upgrade: websocket", vbTextCompare) = 0 Then
        HandleHandshake = False
        Exit Function
    End If

    ' Get Sec-WebSocket-Key
    sKey = GetHeaderValue(Request, "Sec-WebSocket-Key")
    If LenB(sKey) = 0 Then
        HandleHandshake = False
        Exit Function
    End If

    ' Compute accept key
    sAccept = ComputeAcceptKey(sKey)
    If LenB(sAccept) = 0 Then
        HandleHandshake = False
        Exit Function
    End If

    ' Build response
    sResponse = "HTTP/1.1 101 Switching Protocols" & vbCrLf
    sResponse = sResponse & "Upgrade: websocket" & vbCrLf
    sResponse = sResponse & "Connection: Upgrade" & vbCrLf
    sResponse = sResponse & "Sec-WebSocket-Accept: " & sAccept & vbCrLf
    sResponse = sResponse & vbCrLf

    ' Send response
    Client.Socket.SendData sResponse, ScpUtf8

    HandleHandshake = True
    Exit Function

ErrorHandler:
    HandleHandshake = False
End Function

'=========================================================================
' Private Methods - Frame Processing
'=========================================================================

Private Sub ProcessClientData(ByVal Client As cWebSocketServerClient)
    Dim oParser As cWebSocketFrame
    Dim baPayload() As Byte

    On Error GoTo ErrorHandler

    Set oParser = Client.FrameParser

    Do While Client.RecvBuffer.Size >= 2
        ' Parse frame header
        If Not oParser.ParseHeader(Client.RecvBuffer) Then
            Exit Do  ' Need more data
        End If

        ' Check if complete frame available
        If Not oParser.IsCompleteFrame(Client.RecvBuffer) Then
            Exit Do  ' Need more data
        End If

        ' Extract payload
        baPayload = oParser.ExtractPayload(Client.RecvBuffer)

        ' Process by opcode
        Select Case oParser.OpCode
            Case WS_OPCODE_TEXT, WS_OPCODE_BINARY
                ProcessDataFrame Client, baPayload, oParser.OpCode, oParser.FIN

            Case WS_OPCODE_CONTINUATION
                ProcessContinuationFrame Client, baPayload, oParser.FIN

            Case WS_OPCODE_CLOSE
                ProcessCloseFrame Client, baPayload

            Case WS_OPCODE_PING
                ProcessPingFrame Client, baPayload

            Case WS_OPCODE_PONG
                ' Ignore pong
        End Select
    Loop

    Exit Sub
ErrorHandler:
    RaiseError "ProcessClientData: " & ERR.Description
End Sub

Private Sub ProcessDataFrame(ByVal Client As cWebSocketServerClient, _
                            ByRef Payload() As Byte, _
                            ByVal OpCode As WsOpCode, _
                            ByVal IsFinal As Boolean)
    If IsFinal And Not Client.IsFragmented Then
        ' Complete standalone message
        DeliverClientMessage Client, Payload, OpCode
    Else
        ' Start fragmented message
        Client.IsFragmented = True
        Client.FragmentOpCode = OpCode
        Client.FragmentBuffer.Clear

        On Error Resume Next
        If UBound(Payload) >= 0 Then
            Client.FragmentBuffer.Append Payload
        End If
        On Error GoTo 0

        If IsFinal Then
            DeliverFragmentedMessage Client
        End If
    End If
End Sub

Private Sub ProcessContinuationFrame(ByVal Client As cWebSocketServerClient, _
                                    ByRef Payload() As Byte, _
                                    ByVal IsFinal As Boolean)
    If Not Client.IsFragmented Then
        RaiseError "Unexpected continuation frame from " & Client.ID
        Exit Sub
    End If

    On Error Resume Next
    If UBound(Payload) >= 0 Then
        Client.FragmentBuffer.Append Payload
    End If
    On Error GoTo 0

    If IsFinal Then
        DeliverFragmentedMessage Client
    End If
End Sub

Private Sub DeliverFragmentedMessage(ByVal Client As cWebSocketServerClient)
    Dim baData() As Byte
    baData = Client.GetFragmentedData

    DeliverClientMessage Client, baData, Client.FragmentOpCode

    Client.ClearFragmentBuffer
End Sub

Private Sub DeliverClientMessage(ByVal Client As cWebSocketServerClient, _
                                ByRef Data() As Byte, _
                                ByVal OpCode As WsOpCode)
    On Error Resume Next

    If OpCode = WS_OPCODE_TEXT Then
        Dim sText As String
        If UBound(Data) >= 0 Then
            sText = UTF8ToString(Data)
        Else
            sText = ""
        End If
        RaiseEvent OnClientTextMessage(Client.ID, sText)
    Else
        RaiseEvent OnClientBinaryMessage(Client.ID, Data)
    End If
End Sub

Private Sub ProcessCloseFrame(ByVal Client As cWebSocketServerClient, ByRef Payload() As Byte)
    Dim sReason As String

    On Error Resume Next
    sReason = "Client initiated close"

    ' Parse reason if available
    If UBound(Payload) >= 3 Then
        Dim baReason() As Byte
        ReDim baReason(UBound(Payload) - 2) As Byte
        Dim i As Long
        For i = 2 To UBound(Payload)
            baReason(i - 2) = Payload(i)
        Next i
        sReason = UTF8ToString(baReason)
    End If

    On Error GoTo 0

    RemoveClient Client.ID, sReason
End Sub

Private Sub ProcessPingFrame(ByVal Client As cWebSocketServerClient, ByRef Payload() As Byte)
    ' Auto-reply with pong
    Client.SendPong Payload
End Sub

'=========================================================================
' Private Methods - Helpers
'=========================================================================

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cWebSocketServer] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Socket Event Handlers
'=========================================================================

Private Sub m_ListenSocket_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    On Error GoTo ErrorHandler

    ' Create new client
    Dim oClient As cWebSocketServerClient
    Set oClient = New cWebSocketServerClient

    ' Generate client ID
    m_ClientCounter = m_ClientCounter + 1
    oClient.ID = "Client#" & m_ClientCounter

    ' Set client properties
    Set oClient.Socket = Client
    oClient.RemoteAddress = Client.RemoteHostIP
    oClient.RemotePort = Client.RemotePort
    oClient.State = WS_CLIENT_PENDING

    ' Add to collection
    m_Clients.Add oClient, oClient.ID

    ' Accept connection
    DisConnect = False

    Exit Sub
ErrorHandler:
    DisConnect = True
    RaiseError "ConnectionRequest: " & ERR.Description
End Sub

Private Sub m_ListenSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim oClient As cWebSocketServerClient
    Dim baData() As Byte
    Dim sData As String

    On Error GoTo ErrorHandler

    ' Find client
    Set oClient = FindClientBySocket(Client)
    If oClient Is Nothing Then Exit Sub

    If oClient.State = WS_CLIENT_PENDING Then
        ' Expecting handshake
        Client.GetData sData, vbString, , ScpUtf8

        If HandleHandshake(oClient, sData) Then
            oClient.State = WS_CLIENT_CONNECTED
            RaiseEvent OnClientConnect(oClient.ID, oClient.RemoteAddress, oClient.RemotePort)
        Else
            ' Handshake failed
            Client.Close_
            m_Clients.Remove oClient.ID
            RaiseError "Handshake failed for " & oClient.ID
        End If

    ElseIf oClient.State = WS_CLIENT_CONNECTED Then
        ' Receive WebSocket data
        Client.GetData baData, vbByte + vbArray

        On Error Resume Next
        If UBound(baData) >= 0 Then
            oClient.AppendRecvData baData
        End If
        On Error GoTo ErrorHandler

        ProcessClientData oClient
    End If

    Exit Sub
ErrorHandler:
    RaiseError "DataArrival: " & ERR.Description
End Sub

Private Sub m_ListenSocket_CloseEvent(Client As cWinsock)
    Dim oClient As cWebSocketServerClient

    On Error Resume Next

    Set oClient = FindClientBySocket(Client)
    If Not oClient Is Nothing Then
        RemoveClient oClient.ID, "Connection closed"
    End If
End Sub

Private Sub m_ListenSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Dim oClient As cWebSocketServerClient

    On Error Resume Next

    Set oClient = FindClientBySocket(Client)
    If Not oClient Is Nothing Then
        RemoveClient oClient.ID, "Socket error: " & Description
    Else
        RaiseError "Socket error: " & Description
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_IsListening = False
    m_ClientCounter = 0
    Set m_FrameBuilder = New cWebSocketFrame
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    StopServer
    Set m_FrameBuilder = Nothing
End Sub
