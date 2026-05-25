VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cWinsock - Simplified Winsock Wrapper for VB6 by woeoio@qq.com
' Based on VbAsyncSocket by wqweto@gmail.com
'
' Purpose: Provides a clean, event-driven programming model similar to
'          the classic VB6 Winsock control, but with direct object
'          references instead of control array indices.
'
' Key Features:
'   - Event DataArrival(Client As cWinsock, bytesTotal As Long)
'     Directly operate on the client socket object
'   - Automatic client connection management for server mode
'   - Simple Connect/Listen/Send API
'   - Supports TCP and UDP protocols
'
'=========================================================================
Option Explicit
    DefObj A-Z
Private Const MODULE_NAME As String = "cWinsock"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'=========================================================================
' Public Events
'=========================================================================

' Client connected to server (TCP only)
Event Connect(Client As cWinsock)

' Client disconnected (TCP only)
Event CloseEvent(Client As cWinsock)

' New connection request on listening socket
' DisConnect: Set to True to reject the connection and clean up resources
Event ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)

' Data received - directly get the client socket object
Event DataArrival(Client As cWinsock, ByVal bytesTotal As Long)

' Send progress notification
Event SendProgress(Client As cWinsock, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

' Send completed
Event SendComplete(Client As cWinsock)

' Error occurred
Event Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)

' Client count changed (for server mode)
Event ClientCountChange(Client As cWinsock, ByVal ClientCount As Long, ByVal Action As WinsockClientAction)

' Complete message received (only fires when PacketProtocol is set)
' Unlike DataArrival which fires on any raw data, this fires only when
' a complete protocol-framed message is available
Event MessageArrival(Client As cWinsock, ByVal bytesTotal As Long)

' Client heartbeat timeout (server-side, fires when client idle exceeds HeartbeatTimeout)
Event ClientTimeout(Client As cWinsock)

' Heartbeat sent (client-side)
Event HeartbeatSent(Client As cWinsock)

Public Enum WinsockClientAction
    wskServerStoped = -1
    wskServerStarted = 0
    wskClientConnected = 1
    wskClientDisconnected = 2
End Enum

'=========================================================================
' Public Enums

Public Enum EnumScpCodePage
    wcpAcp = 0
    wcpOem = 1
    wcpUtf8 = 65001
    wcpUnicode = -1                                                             '--- handled internally -> no conversion, keep text wide-char
    wcpNone = ucsScpUnicode                                                     '--- and some more aliases
    wcpWideChar = ucsScpUnicode
End Enum

'=========================================================================

Public Enum WinsockProtocol
    wskTCPProtocol = 1
    wskUDPProtocol = 2
End Enum

Public Enum WinsockState
    wckClosed = 0
    wckOpen = 1
    wckListening = 2
    wckConnectionPending = 3
    wckResolvingHost = 4
    wckHostResolved = 5
    wckConnecting = 6
    wckConnected = 7
    wckClosing = 8
    wckError = 9
End Enum

Public Enum WinSocketEventMaskEnum
    wskSfdRead = 2 ^ 0
    wskSfdWrite = 2 ^ 1
    wskSfdOob = 2 ^ 2
    wskSfdAccept = 2 ^ 3
    wskSfdConnect = 2 ^ 4
    wskSfdClose = 2 ^ 5
    wskSfdAll = 2 ^ 6 - 1
    [_wskSfdResolve] = 2 ^ 15
    [_wskSfdForceRead] = 2 ^ 14
End Enum

'=========================================================================
' Private Constants
'=========================================================================

Private Const DEF_LOCALPORT         As Long = 0
Private Const DEF_PROTOCOL          As Long = wskTCPProtocol
Private Const DEF_REMOTEPORT        As Long = 0

'=========================================================================
' Member Variables
'=========================================================================

Private WithEvents m_oSocket        As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_eState                    As WinsockState
Private m_eProtocol                 As WinsockProtocol
Private m_lLocalPort                As Long
Private m_sRemoteHost               As String
Private m_lRemotePort               As Long
Private m_sRemoteHostIP             As String
Private m_sTag                      As String
Private m_vUserData                 As Variant
Private m_baRecvBuffer()            As Byte
Private m_bIsServer                 As Boolean
Private m_bIsAcceptedClient         As Boolean
Private m_oParentServer             As cWinsock
Private m_cClients                  As Collection
Private m_lClientCounter            As Long
Private m_bCloseEventRaised         As Boolean                                  ' Prevent duplicate CloseEvent
Private m_oPacketProtocol           As cPacketProtocol                          ' Packet protocol handler
Private WithEvents m_oHeartbeat     As cHeartbeat                               ' Heartbeat manager

Private m_Users As New Dictionary
Private m_Groups As New Dictionary                                              ' GroupName -> Dictionary(UserName -> True)

'For Client of tcp server
Public CurrentUser As Variant
Public CurrentUserToken As String
Public CurrentUserInfo As New cJson
Public CurrentGroups As New Dictionary                                          ' GroupName -> True
'=========================================================================
' Properties
'=========================================================================

' Current Users Count
Public Property Get CountUsers() As Long
    CountUsers = m_Users.count
End Property

' Current m_Groups Count
Public Property Get CountGroups() As Long
    CountGroups = m_Groups.count
End Property

' Current socket state
Public Property Get State() As WinsockState
    State = m_eState
End Property

' Protocol type (TCP/UDP)
Public Property Get Protocol() As WinsockProtocol
    Protocol = m_eProtocol
End Property

Public Property Let Protocol(ByVal Value As WinsockProtocol)
    If m_eState <> wckClosed Then
        ERR.Raise vbObjectError, MODULE_NAME, "Cannot change protocol while socket is open"
    End If
    m_eProtocol = Value
End Property

Public Property Let RecvBuffer(ByRef Value() As Byte)
    m_baRecvBuffer = Value
End Property

' Local port to bind
Public Property Get LocalPort() As Long
    Dim lPort As Long
    If m_eState <> wckClosed And Not m_oSocket Is Nothing Then
        m_oSocket.GetSockName vbNullString, lPort
        LocalPort = lPort
    Else
        LocalPort = m_lLocalPort
    End If
End Property

Public Property Let LocalPort(ByVal Value As Long)
    If m_eState <> wckClosed Then
        ERR.Raise vbObjectError, MODULE_NAME, "Cannot change local port while socket is open"
    End If
    m_lLocalPort = Value
End Property

' Remote host address
Public Property Get RemoteHost() As String
    RemoteHost = m_sRemoteHost
End Property

Public Property Let RemoteHost(ByVal Value As String)
    m_sRemoteHost = Value
End Property

' Remote port
Public Property Get RemotePort() As Long
    RemotePort = m_lRemotePort
End Property

Public Property Let RemotePort(ByVal Value As Long)
    m_lRemotePort = Value
End Property

' Resolved remote IP address
Public Property Get RemoteHostIP() As String
    ' For UDP server mode clients (no socket but has parent server), return stored IP
    If m_oSocket Is Nothing And Not m_oParentServer Is Nothing And m_eProtocol = wskUDPProtocol Then
        RemoteHostIP = m_sRemoteHostIP
        Exit Property
    End If
    If m_eState >= wckHostResolved And Not m_oSocket Is Nothing Then
        Dim sAddr As String
        Dim lPort As Long
        m_oSocket.GetPeerName sAddr, lPort
        If LenB(sAddr) <> 0 Then
            RemoteHostIP = sAddr
            Exit Property
        End If
    End If
    RemoteHostIP = m_sRemoteHostIP
End Property

' Local host name
Public Property Get LocalHostName() As String
    If Not m_oSocket Is Nothing Then
        m_oSocket.GetLocalHost LocalHostName, vbNullString
    End If
End Property

' Local IP address
Public Property Get LocalIP() As String
    If Not m_oSocket Is Nothing Then
        Dim sAddr As String
        Dim lPort As Long
        m_oSocket.GetSockName sAddr, lPort
        LocalIP = sAddr
    End If
End Property

' User-defined tag for identification
Public Property Get Tag() As String
    Tag = m_sTag
End Property

Public Property Let Tag(ByVal Value As String)
    m_sTag = Value
End Property

' User-defined data storage
Public Property Get UserData() As Variant
    If IsObject(m_vUserData) Then
        Set UserData = m_vUserData
    Else
        UserData = m_vUserData
    End If
End Property

Public Property Let UserData(ByVal Value As Variant)
    m_vUserData = Value
End Property

Public Property Set UserData(ByVal Value As Variant)
    Set m_vUserData = Value
End Property

' Socket handle
Public Property Get SocketHandle() As Long
    If Not m_oSocket Is Nothing Then
        SocketHandle = m_oSocket.SocketHandle
    End If
End Property

' Available bytes in receive buffer
Public Property Get BytesReceived() As Long
    If Not m_oSocket Is Nothing Then
        BytesReceived = m_oSocket.AvailableBytes
    End If
    If UBound(m_baRecvBuffer) >= 0 Then
        BytesReceived = BytesReceived + UBound(m_baRecvBuffer) + 1
    End If
End Property

' Is this a server (listening) socket
Public Property Get IsServer() As Boolean
    IsServer = m_bIsServer
End Property

' Is this an accepted client socket
Public Property Get IsAcceptedClient() As Boolean
    IsAcceptedClient = m_bIsAcceptedClient
End Property

' Get the parent server socket (for accepted clients)
Public Property Get ParentServer() As cWinsock
    Set ParentServer = m_oParentServer
End Property

Friend Property Set ParentServer(ByVal Value As cWinsock)
    Set m_oParentServer = Value
End Property

' Get all connected clients (for server socket)
Public Property Get Clients() As Collection
    Set Clients = m_cClients
End Property

' Get client count (for server socket)
Public Property Get ClientCount() As Long
    If Not m_cClients Is Nothing Then
        ClientCount = m_cClients.count
    End If
End Property

' Internal socket reference
Friend Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

'--- Packet Protocol Properties ---

' Get/Set the packet protocol handler object
' Each cWinsock instance has its own independent protocol handler
Public Property Get PacketHandler() As cPacketProtocol
    Set PacketHandler = m_oPacketProtocol
End Property

Public Property Set PacketHandler(ByVal Value As cPacketProtocol)
    Set m_oPacketProtocol = Value
End Property

' Quick setup: set packet protocol type (creates handler if needed)
Public Property Get PacketProtocol() As PacketProtocolType
    If m_oPacketProtocol Is Nothing Then
        PacketProtocol = ppNone
    Else
        PacketProtocol = m_oPacketProtocol.ProtocolType
    End If
End Property

Public Property Let PacketProtocol(ByVal Value As PacketProtocolType)
    If Value = ppNone Then
        Set m_oPacketProtocol = Nothing
    Else
        If m_oPacketProtocol Is Nothing Then
            Set m_oPacketProtocol = New cPacketProtocol
        End If
        m_oPacketProtocol.ProtocolType = Value
    End If
End Property

' Delimiter for ppDelimiter protocol
Public Property Get Delimiter() As String
    If m_oPacketProtocol Is Nothing Then
        Delimiter = vbCrLf
    Else
        Delimiter = m_oPacketProtocol.Delimiter
    End If
End Property

Public Property Let Delimiter(ByVal Value As String)
    If m_oPacketProtocol Is Nothing Then
        Set m_oPacketProtocol = New cPacketProtocol
        m_oPacketProtocol.ProtocolType = ppDelimiter
    End If
    m_oPacketProtocol.Delimiter = Value
End Property

' Fixed length for ppFixedLength protocol
Public Property Get FixedLength() As Long
    If m_oPacketProtocol Is Nothing Then
        FixedLength = 0
    Else
        FixedLength = m_oPacketProtocol.FixedLength
    End If
End Property

Public Property Let FixedLength(ByVal Value As Long)
    If m_oPacketProtocol Is Nothing Then
        Set m_oPacketProtocol = New cPacketProtocol
        m_oPacketProtocol.ProtocolType = ppFixedLength
    End If
    m_oPacketProtocol.FixedLength = Value
End Property

' Header bytes for ppLengthHeader protocol (2 or 4)
Public Property Get HeaderBytes() As Long
    If m_oPacketProtocol Is Nothing Then
        HeaderBytes = 4
    Else
        HeaderBytes = m_oPacketProtocol.HeaderBytes
    End If
End Property

Public Property Let HeaderBytes(ByVal Value As Long)
    If m_oPacketProtocol Is Nothing Then
        Set m_oPacketProtocol = New cPacketProtocol
        m_oPacketProtocol.ProtocolType = ppLengthHeader
    End If
    m_oPacketProtocol.HeaderBytes = Value
End Property

' Endian for ppLengthHeader protocol
Public Property Get HeaderEndian() As EndianEnum
    If m_oPacketProtocol Is Nothing Then
        HeaderEndian = eeLittleEndian
    Else
        HeaderEndian = m_oPacketProtocol.Endian
    End If
End Property

Public Property Let HeaderEndian(ByVal Value As EndianEnum)
    If m_oPacketProtocol Is Nothing Then
        Set m_oPacketProtocol = New cPacketProtocol
        m_oPacketProtocol.ProtocolType = ppLengthHeader
    End If
    m_oPacketProtocol.Endian = Value
End Property

'--- Heartbeat Properties (auto-driven by internal cTimer) ---

' Get the heartbeat handler object (for advanced configuration)
Public Property Get Heartbeat() As cHeartbeat
    Set Heartbeat = m_oHeartbeat
End Property

' Enable/disable auto heartbeat (uses internal cTimer, no external polling needed)
Public Property Let AutoHeartbeat(ByVal Value As Boolean)
    If m_oHeartbeat Is Nothing Then
        Set m_oHeartbeat = New cHeartbeat
    End If
    m_oHeartbeat.Enabled = Value
End Property

Public Property Get AutoHeartbeat() As Boolean
    If m_oHeartbeat Is Nothing Then
        AutoHeartbeat = False
    Else
        AutoHeartbeat = m_oHeartbeat.Enabled
    End If
End Property

' Server: heartbeat timeout in seconds (default 120)
' Clients idle beyond this are auto-disconnected
Public Property Let HeartbeatTimeout(ByVal Value As Long)
    If m_oHeartbeat Is Nothing Then
        Set m_oHeartbeat = New cHeartbeat
    End If
    m_oHeartbeat.ServerTimeout = Value
End Property

Public Property Get HeartbeatTimeout() As Long
    If m_oHeartbeat Is Nothing Then
        HeartbeatTimeout = 120
    Else
        HeartbeatTimeout = m_oHeartbeat.ServerTimeout
    End If
End Property

' Client: heartbeat interval in seconds (default 50)
Public Property Let HeartbeatInterval(ByVal Value As Long)
    If m_oHeartbeat Is Nothing Then
        Set m_oHeartbeat = New cHeartbeat
    End If
    m_oHeartbeat.ClientInterval = Value
End Property

Public Property Get HeartbeatInterval() As Long
    If m_oHeartbeat Is Nothing Then
        HeartbeatInterval = 50
    Else
        HeartbeatInterval = m_oHeartbeat.ClientInterval
    End If
End Property

' Heartbeat packet content (byte array, default single zero byte)
Public Property Let HeartbeatData(ByRef Value() As Byte)
    If m_oHeartbeat Is Nothing Then
        Set m_oHeartbeat = New cHeartbeat
    End If
    m_oHeartbeat.HeartbeatData = Value
End Property

Public Property Get HeartbeatData() As Byte()
    If m_oHeartbeat Is Nothing Then
        Dim baDefault(0 To 0) As Byte
        baDefault(0) = 0
        HeartbeatData = baDefault
    Else
        HeartbeatData = m_oHeartbeat.HeartbeatData
    End If
End Property

' Last activity time (seconds since last send/receive)
Public Property Get IdleSeconds() As Long
    If m_oHeartbeat Is Nothing Then
        IdleSeconds = 0
    Else
        IdleSeconds = m_oHeartbeat.IdleSeconds
    End If
End Property

'=========================================================================
' Public Methods
'=========================================================================

Public Sub BindUser(ByVal User As Variant, Client As cWinsock, Optional ByVal Token As String, Optional Info As cJson)
    Set m_Users(User) = Client
    Client.CurrentUser = User
    If Token <> "" Then Client.CurrentUserToken = Token
    If Not Info Is Nothing Then Set Client.CurrentUserInfo = Info: Set Info = Nothing
End Sub

' Check if a user is bound
Public Function ExistsUser(ByVal User As Variant) As Boolean
    ExistsUser = m_Users.Exists(User)
End Function

' Unbind a user
Public Sub UnbindUser(ByVal User As Variant)
    On Error Resume Next
    m_Users.Remove User
    CurrentUser = vbNull
    CurrentUserToken = ""
    CurrentUserInfo.Clear
End Sub

Public Function GetClientByUser(ByVal User As Variant) As cWinsock
    Set GetClientByUser = m_Users(User)
End Function
'Public Function GetClientByIndex(ByVal Index As Long) As cWinsock
'    '    Set GetClientByUser = m_Users(User)
'End Function

' Send data to a bound user
Public Sub SendToUser(ByVal User As Variant, Data As Variant, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    On Error GoTo EH
    
    ' Check if user exists
    If Not m_Users.Exists(User) Then
        ERR.Raise vbObjectError, MODULE_NAME, "User not bound: " & CStr(User)
        Exit Sub
    End If
    
    Dim oClient As cWinsock
    Set oClient = m_Users(User)
    
    oClient.SendData Data, CodePage
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

'=========================================================================
' Group Management
'=========================================================================

' Bind client to a group (client must be bound to a user first)
Public Sub BindGroup(ByVal GroupName As String, Client As cWinsock)
    If LenB(CStr(Client.CurrentUser)) = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "Client must be bound to a user first"
        Exit Sub
    End If
    
    Dim dMembers As Dictionary
    Dim sUserName As String
    sUserName = CStr(Client.CurrentUser)
    
    ' Create group if not exists
    If Not m_Groups.Exists(GroupName) Then
        Set dMembers = New Dictionary
        dMembers.CompareMode = TextCompare
        m_Groups.Add GroupName, dMembers
    Else
        Set dMembers = m_Groups(GroupName)
    End If
    
    ' Add user to group if not already in
    If Not dMembers.Exists(sUserName) Then
        dMembers.Add sUserName, True
    End If
    
    ' Sync client's CurrentGroups
    If Not Client.CurrentGroups.Exists(GroupName) Then
        Client.CurrentGroups.Add GroupName, True
    End If
End Sub

' Check if a group exists
Public Function ExistsGroup(ByVal GroupName As String) As Boolean
    ExistsGroup = m_Groups.Exists(GroupName)
End Function

' Unbind a user from a group
Public Sub UnbindGroup(ByVal GroupName As String, ByVal User As Variant)
    On Error Resume Next
    If m_Groups.Exists(GroupName) Then
        m_Groups(GroupName).Remove CStr(User)
        ' Remove empty group
        If m_Groups(GroupName).count = 0 Then
            m_Groups.Remove GroupName
        End If
    End If
    
    ' Sync client's CurrentGroups
    If m_Users.Exists(User) Then
        Dim oClient As cWinsock
        Set oClient = m_Users(User)
        If oClient.CurrentGroups.Exists(GroupName) Then
            oClient.CurrentGroups.Remove GroupName
        End If
    End If
End Sub

' Unbind a user from all groups (called internally when user disconnects)
Friend Sub UnbindUserFromAllGroups(ByVal User As Variant)
    Dim sGroup As Variant
    Dim oClient As cWinsock
    
    ' Get client instance to sync CurrentGroups
    If m_Users.Exists(User) Then
        Set oClient = m_Users(User)
    End If
    
    For Each sGroup In m_Groups.Keys
        On Error Resume Next
        m_Groups(sGroup).Remove CStr(User)
        On Error GoTo 0
        
        ' Sync client's CurrentGroups
        If Not oClient Is Nothing Then
            If oClient.CurrentGroups.Exists(sGroup) Then
                oClient.CurrentGroups.Remove sGroup
            End If
        End If
        
        ' Remove empty groups
        If m_Groups.Exists(sGroup) And m_Groups(sGroup).count = 0 Then
            m_Groups.Remove sGroup
        End If
    Next
End Sub

' Get members of a group (returns array of user names)
Public Function GetGroupMembers(ByVal GroupName As String) As String()
    Dim aResult() As String
    Dim dMembers As Dictionary
    Dim vMember As Variant
    Dim lIdx As Long
    
    If m_Groups.Exists(GroupName) Then
        Set dMembers = m_Groups(GroupName)
        If dMembers.count > 0 Then
            ReDim aResult(0 To dMembers.count - 1)
            lIdx = 0
            For Each vMember In dMembers.Keys
                aResult(lIdx) = CStr(vMember)
                lIdx = lIdx + 1
            Next
            GetGroupMembers = aResult
            Exit Function
        End If
    End If
    
    GetGroupMembers = aResult
End Function

' Send data to all members in a group
Public Sub SendToGroup(ByVal GroupName As String, Data As Variant, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    Dim dMembers As Dictionary
    Dim vMember As Variant
    Dim oClient As cWinsock
    
    If Not m_Groups.Exists(GroupName) Then Exit Sub
    
    Set dMembers = m_Groups(GroupName)
    For Each vMember In dMembers.Keys
        If m_Users.Exists(vMember) Then
            Set oClient = m_Users(vMember)
            oClient.SendData Data, CodePage
        End If
    Next
End Sub

' Send data to all connected clients
Public Sub SendToAll(Data As Variant, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    Dim oClient As cWinsock
    Dim vKey As Variant
    
    If m_cClients Is Nothing Then Exit Sub
    
    For Each vKey In m_cClients
        Set oClient = vKey
        oClient.SendData Data, CodePage
    Next
End Sub

' Get all groups a user belongs to
Public Function GetUserGroups(ByVal User As Variant) As String()
    Dim aResult() As String
    Dim sGroup As Variant
    Dim lIdx As Long
    
    ' Fast path: use client's CurrentGroups if user is online
    If m_Users.Exists(User) Then
        Dim oClient As cWinsock
        Set oClient = m_Users(User)
        If oClient.CurrentGroups.count > 0 Then
            ReDim aResult(0 To oClient.CurrentGroups.count - 1)
            lIdx = 0
            For Each sGroup In oClient.CurrentGroups.Keys
                aResult(lIdx) = CStr(sGroup)
                lIdx = lIdx + 1
            Next
            GetUserGroups = aResult
            Exit Function
        End If
    End If
    
    ' Fallback: scan all groups (for offline users)
    For Each sGroup In m_Groups.Keys
        If m_Groups(sGroup).Exists(CStr(User)) Then
            ReDim Preserve aResult(0 To lIdx)
            aResult(lIdx) = CStr(sGroup)
            lIdx = lIdx + 1
        End If
    Next
    
    GetUserGroups = aResult
End Function

' Delete a group (disband)
Public Sub DeleteGroup(ByVal GroupName As String)
    If Not m_Groups.Exists(GroupName) Then Exit Sub
    
    Dim dMembers As Dictionary
    Dim vUser As Variant
    Dim oClient As cWinsock
    
    Set dMembers = m_Groups(GroupName)
    
    ' Sync remove group from each member's CurrentGroups
    For Each vUser In dMembers.Keys
        If m_Users.Exists(vUser) Then
            Set oClient = m_Users(vUser)
            If oClient.CurrentGroups.Exists(GroupName) Then
                oClient.CurrentGroups.Remove GroupName
            End If
        End If
    Next
    
    ' Remove the group
    m_Groups.Remove GroupName
End Sub

' Close a specific user's client connection
Public Sub CloseUser(ByVal User As Variant)
    On Error Resume Next
    If m_Users.Exists(User) Then
        Dim oClient As cWinsock
        Set oClient = m_Users(User)
        RemoveClient oClient
        oClient.Close_
    End If
End Sub


' Connect to remote host
Public Sub Connect(Optional RemoteHost As String, Optional ByVal RemotePort As Long)
    Dim eSocketType As UcsAsyncSocketTypeEnum
    
    On Error GoTo EH
    
    ' Close existing connection
    Close_
    
    ' Update remote host/port if provided
    If LenB(RemoteHost) <> 0 Then
        m_sRemoteHost = RemoteHost
    End If
    If RemotePort <> 0 Then
        m_lRemotePort = RemotePort
    End If
    
    ' Validate parameters
    If LenB(m_sRemoteHost) = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "RemoteHost not specified"
    End If
    If m_lRemotePort = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "RemotePort not specified"
    End If
    
    ' Determine socket type
    eSocketType = IIf(m_eProtocol = wskTCPProtocol, ucsSckStream, ucsSckDatagram)
    
    ' Create socket
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=eSocketType) Then
        pvRaiseError m_oSocket.LastError
        Exit Sub
    End If
    
    m_eState = sckResolvingHost
    
    ' Connect
    If Not m_oSocket.Connect(m_sRemoteHost, m_lRemotePort) Then
        pvRaiseError m_oSocket.LastError
        Exit Sub
    End If
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Start listening for incoming connections
Public Sub Listen(Optional ByVal Port As Long, Optional ByVal LocalAddress As String = "")
    Dim eSocketType As UcsAsyncSocketTypeEnum
    
    On Error GoTo EH
    
    ' Close existing connection
    Close_
    
    ' Update port if provided
    If Port <> 0 Then
        m_lLocalPort = Port
    End If
    
    ' Determine socket type
    eSocketType = IIf(m_eProtocol = wskTCPProtocol, ucsSckStream, ucsSckDatagram)
    
    ' Create socket
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketPort:=m_lLocalPort, SocketType:=eSocketType, SocketAddress:=LocalAddress) Then
        pvRaiseError m_oSocket.LastError
        Exit Sub
    End If
    
    ' For TCP, start listening
    If m_eProtocol = wskTCPProtocol Then
        If Not m_oSocket.Listen() Then
            pvRaiseError m_oSocket.LastError
            Exit Sub
        End If
    End If
    
    m_eState = sckListening
    m_bIsServer = True
    Set m_cClients = New Collection
    
    ' Raise ClientCountChange event for new connection
    RaiseEvent ClientCountChange(Nothing, 0, wskServerStarted)
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Send string data
' CodePage: Text encoding for string data (default: ACP/GBK)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
' Note: When PacketProtocol is set, string data is auto-encoded and framed
Public Sub SendData(Data As Variant, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    On Error GoTo EH
    
    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If
    
    '--- Auto packet framing (Encode) ---
    ' If protocol is set and input is a string, convert to bytes first then Encode
    If Not m_oPacketProtocol Is Nothing Then
        If m_oPacketProtocol.ProtocolType <> ppNone Then
            Dim baRaw() As Byte
            
            Select Case varType(Data)
            Case vbString
                ' Convert string to bytes using specified CodePage
                Dim oConvSocket As cAsyncSocket
                If Not m_oSocket Is Nothing Then
                    Set oConvSocket = m_oSocket
                ElseIf Not m_oParentServer Is Nothing And Not m_oParentServer.Socket Is Nothing Then
                    Set oConvSocket = m_oParentServer.Socket
                Else
                    Set oConvSocket = New cAsyncSocket
                End If
                baRaw = oConvSocket.ToTextArray(CStr(Data), CodePage)
            Case vbByte + vbArray
                baRaw = Data
            Case Else
                ERR.Raise vbObjectError, MODULE_NAME, "Unsupported data type for packet protocol: " & TypeName(Data)
            End Select
            
            ' Encode (add framing markers)
            Dim baEncoded() As Byte
            baEncoded = m_oPacketProtocol.Encode(baRaw)
            
            ' Send encoded byte array
            If m_eProtocol = wskUDPProtocol And m_bIsServer Then
                If LenB(m_sRemoteHostIP) <> 0 Then
                    m_oSocket.SendArray baEncoded, m_sRemoteHostIP, m_lRemotePort
                ElseIf LenB(m_sRemoteHost) <> 0 Then
                    m_oSocket.SendArray baEncoded, m_sRemoteHost, m_lRemotePort
                Else
                    ERR.Raise vbObjectError, MODULE_NAME, "RemoteHost or RemoteHostIP not specified"
                End If
            Else
                m_oSocket.SendArray baEncoded
            End If
            
            Exit Sub
        End If
    End If
    
    '--- Normal send (no protocol framing) ---
    Select Case varType(Data)
    Case vbString
        If m_eProtocol = wskUDPProtocol And m_bIsServer Then
            ' For UDP server, need to specify remote address
            ' Use m_sRemoteHost if m_sRemoteHostIP is empty (allows domain name resolution)
            If LenB(m_sRemoteHostIP) <> 0 Then
                m_oSocket.SendText CStr(Data), m_sRemoteHostIP, m_lRemotePort, CodePage
            ElseIf LenB(m_sRemoteHost) <> 0 Then
                m_oSocket.SendText CStr(Data), m_sRemoteHost, m_lRemotePort, CodePage
            Else
                ERR.Raise vbObjectError, MODULE_NAME, "RemoteHost or RemoteHostIP not specified"
            End If
        Else
            m_oSocket.SendText CStr(Data), , , CodePage
        End If
    Case vbByte + vbArray

        Dim byteData() As Byte
        byteData = Data

        If m_eProtocol = wskUDPProtocol And m_bIsServer Then
            ' Use m_sRemoteHost if m_sRemoteHostIP is empty (allows domain name resolution)
            If LenB(m_sRemoteHostIP) <> 0 Then
                m_oSocket.SendArray byteData, m_sRemoteHostIP, m_lRemotePort
            ElseIf LenB(m_sRemoteHost) <> 0 Then
                m_oSocket.SendArray byteData, m_sRemoteHost, m_lRemotePort
            Else
                ERR.Raise vbObjectError, MODULE_NAME, "RemoteHost or RemoteHostIP not specified"
            End If
        Else
            m_oSocket.SendArray byteData
        End If
    Case Else
        ERR.Raise vbObjectError, MODULE_NAME, "Unsupported data type: " & TypeName(Data)
    End Select
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

'--- Heartbeat Auto-Tick (driven by internal cTimer) ---
' This runs automatically every 10 seconds when AutoHeartbeat is enabled.
' No external PollHeartbeat() call needed - cHeartbeat uses cTimer internally.

Private Sub m_oHeartbeat_Tick()
    On Error Resume Next
    
    If m_oHeartbeat Is Nothing Then Exit Sub
    If Not m_oHeartbeat.Enabled Then Exit Sub
    
    '--- Server mode: check client timeouts ---
    If m_bIsServer Then
        If m_oHeartbeat.ServerTimeout > 0 Then
            Dim colTimedOut As Collection
            Set colTimedOut = m_oHeartbeat.CheckClientTimeouts(Me)
            
            If Not colTimedOut Is Nothing Then
                Dim oClient As cWinsock
                For Each oClient In colTimedOut
                    ' Raise timeout event
                    RaiseEvent ClientTimeout(oClient)
                    ' Auto-disconnect
                    RemoveClient oClient
                    oClient.Close_
                Next
            End If
        End If
    End If
    
    '--- Client mode: send heartbeat ---
    If Not m_bIsServer Then
        If m_oHeartbeat.ClientInterval > 0 Then
            If m_oHeartbeat.ShouldSendHeartbeat() Then
                ' Send heartbeat data
                Dim baHB() As Byte
                baHB = m_oHeartbeat.HeartbeatData
                On Error Resume Next
                If UBound(baHB) >= 0 Then
                    m_oSocket.SendArray baHB
                    m_oHeartbeat.RecordHeartbeatSent
                    RaiseEvent HeartbeatSent(Me)
                End If
            End If
        End If
    End If
End Sub

' Get data as text string (default ACP/GBK encoding)
Public Function GetDataText(Optional ByVal maxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = wcpAcp) As String
    Dim vData As Variant
    GetData vData, vbString, maxLen, CodePage
    GetDataText = vData
End Function

' Get data as UTF-8 text string
Public Function GetDataTextUTF8(Optional ByVal maxLen As Long = -1) As String
    GetDataTextUTF8 = GetDataText(maxLen, wcpUtf8)
End Function

' Get data as Unicode text string
Public Function GetDataTextUnicode(Optional ByVal maxLen As Long = -1) As String
    GetDataTextUnicode = GetDataText(maxLen, wcpUnicode)
End Function

' Get data as hex string (space-separated, e.g. "48 65 6C 6C 6F")
Public Function GetDataHex(Optional ByVal maxLen As Long = -1) As String
    Dim baData() As Byte
    Dim i As Long
    Dim sHex As String
    
    GetData baData, vbByte + vbArray, maxLen
    
    On Error Resume Next
    For i = 0 To UBound(baData)
        sHex = sHex & Right$("0" & Hex$(baData(i)), 2) & " "
    Next
    On Error GoTo 0
    
    If Len(sHex) > 0 Then
        GetDataHex = Left$(sHex, Len(sHex) - 1)
    End If
End Function

' Get data as byte array
Public Function GetDataByteArray(Optional ByVal maxLen As Long = -1) As Byte()
    On Error Resume Next
    Dim baData() As Byte
    GetData baData, vbByte + vbArray, maxLen
    GetDataByteArray = baData
End Function

'--- Legacy alias ---
Public Function GetDataToString(Optional ByVal maxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = wcpAcp) As String
    GetDataToString = GetDataText(maxLen, CodePage)
End Function
' Receive data from buffer
' CodePage: Text encoding for string data (default: ACP/GBK for compatibility)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
Public Sub GetData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal maxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    Dim baBuffer() As Byte
    Dim lIdx As Long
    Dim oSocket As cAsyncSocket
    
    On Error GoTo EH
    
    If VarType_ = 0 Then
        VarType_ = varType(Data)
    End If
    
    ' First check internal buffer
    If UBound(m_baRecvBuffer) >= 0 Then
        If maxLen < 0 Or maxLen >= UBound(m_baRecvBuffer) + 1 Then
            baBuffer = m_baRecvBuffer
            m_baRecvBuffer = vbNullString
        Else
            ReDim baBuffer(0 To maxLen - 1) As Byte
            CopyMemory baBuffer(0), m_baRecvBuffer(0), maxLen
            lIdx = UBound(m_baRecvBuffer) + 1 - maxLen
            If lIdx > 0 Then
                Dim baTemp() As Byte
                ReDim baTemp(0 To lIdx - 1) As Byte
                CopyMemory baTemp(0), m_baRecvBuffer(maxLen), lIdx
                m_baRecvBuffer = baTemp
            Else
                m_baRecvBuffer = vbNullString
            End If
        End If
    ElseIf Not m_oSocket Is Nothing Then
        ' Read from socket
        If m_oSocket.ReceiveArray(baBuffer) Then
            ' If MaxLen is specified and data exceeds it, truncate and save remainder
            If maxLen >= 0 And UBound(baBuffer) >= 0 And UBound(baBuffer) + 1 > maxLen Then
                ' Save remainder to internal buffer
                Dim baRemainder() As Byte
                ReDim baRemainder(0 To UBound(baBuffer) - maxLen) As Byte
                CopyMemory baRemainder(0), baBuffer(maxLen), UBound(baBuffer) - maxLen + 1
                m_baRecvBuffer = baRemainder
                
                ' Truncate baBuffer to MaxLen
                ReDim Preserve baBuffer(0 To maxLen - 1) As Byte
            End If
        Else
            baBuffer = vbNullString
        End If
    ElseIf varType(m_vUserData) = (vbByte + vbArray) Then
        ' For UDP server mode clients, data is stored in UserData
        baBuffer = m_vUserData
        ' Clear UserData after reading
        m_vUserData = Empty
        ' If MaxLen is specified, truncate
        If maxLen >= 0 And UBound(baBuffer) >= 0 And UBound(baBuffer) + 1 > maxLen Then
            Dim baTruncated() As Byte
            ReDim baTruncated(0 To maxLen - 1) As Byte
            CopyMemory baTruncated(0), baBuffer(0), maxLen
            baBuffer = baTruncated
        End If
    Else
        baBuffer = vbNullString
    End If
    
    ' Convert to requested type
    Select Case VarType_
    Case vbString
        If UBound(baBuffer) >= 0 Then
            ' Use parent server's socket for conversion if available, otherwise create a temporary one
            If Not m_oSocket Is Nothing Then
                Set oSocket = m_oSocket
            ElseIf Not m_oParentServer Is Nothing And Not m_oParentServer.Socket Is Nothing Then
                Set oSocket = m_oParentServer.Socket
            Else
                Set oSocket = New cAsyncSocket
            End If
            ' Use specified encoding (default ACP for backward compatibility with UDP)
            Data = oSocket.FromTextArray(baBuffer, CodePage)
        Else
            Data = vbNullString
        End If
    Case vbByte + vbArray
        Data = baBuffer
    Case Else
        Data = baBuffer
    End Select
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Peek data without removing from buffer
' CodePage: Text encoding for string data (default: ACP/GBK for compatibility)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
Public Sub PeekData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal maxLen As Long = -1, Optional ByVal CodePage As EnumScpCodePage = wcpAcp)
    Dim baBuffer() As Byte

    On Error GoTo EH

    If VarType_ = 0 Then
        VarType_ = varType(Data)
    End If

    ' Copy from internal buffer
    If UBound(m_baRecvBuffer) >= 0 Then
        If maxLen < 0 Or maxLen >= UBound(m_baRecvBuffer) + 1 Then
            baBuffer = m_baRecvBuffer
        Else
            ReDim baBuffer(0 To maxLen - 1) As Byte
            CopyMemory baBuffer(0), m_baRecvBuffer(0), maxLen
        End If
    End If

    ' Convert to requested type
    Select Case VarType_
    Case vbString
        If UBound(baBuffer) >= 0 Then
            ' Use specified encoding (default ACP for backward compatibility with UDP)
            Dim oSocket As cAsyncSocket
            If Not m_oSocket Is Nothing Then
                Set oSocket = m_oSocket
            ElseIf Not m_oParentServer Is Nothing And Not m_oParentServer.Socket Is Nothing Then
                Set oSocket = m_oParentServer.Socket
            Else
                Set oSocket = New cAsyncSocket
            End If
            Data = oSocket.FromTextArray(baBuffer, CodePage)
        Else
            Data = vbNullString
        End If
    Case vbByte + vbArray
        Data = baBuffer
    Case Else
        Data = baBuffer
    End Select

    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Close socket
Public Sub Close_()
    On Error Resume Next
    
    ' Close all client connections for server socket
    If m_bIsServer And Not m_cClients Is Nothing Then
        Dim oClient As cWinsock
        Do While m_cClients.count > 0
            Set oClient = m_cClients(1)
            RemoveClient oClient
            oClient.Close_
        Loop
        Set m_cClients = Nothing
    End If
    
    ' Close socket
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If
    
    m_eState = sckClosed
    m_bIsServer = False
    m_bIsAcceptedClient = False
    m_bCloseEventRaised = False
    Set m_oParentServer = Nothing
    m_baRecvBuffer = vbNullString
    
    ' Clear packet protocol buffer
    If Not m_oPacketProtocol Is Nothing Then
        m_oPacketProtocol.Clear
    End If
    
    ' Stop heartbeat
    If Not m_oHeartbeat Is Nothing Then
        m_oHeartbeat.[Stop]
    End If

' Bind to local address/port (for UDP)
Public Sub Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String)
    On Error GoTo EH
    
    Close_
    
    If LocalPort <> 0 Then
        m_lLocalPort = LocalPort
    End If
    
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketPort:=m_lLocalPort, SocketType:=ucsSckDatagram) Then
        pvRaiseError m_oSocket.LastError
        Exit Sub
    End If
    
    m_eState = sckOpen
    m_eProtocol = wskUDPProtocol
    m_bIsServer = True
    Set m_cClients = New Collection
    
    ' Raise ClientCountChange event for new connection
    RaiseEvent ClientCountChange(Nothing, 0, wskServerStarted)
    
    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Get error description
Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
    If Not m_oSocket Is Nothing Then
        GetErrorDescription = m_oSocket.GetErrorDescription(ErrorCode)
    Else
        Dim oTempSocket As cAsyncSocket
        Set oTempSocket = New cAsyncSocket
        GetErrorDescription = oTempSocket.GetErrorDescription(ErrorCode)
    End If
End Function

'=========================================================================
' Friend Methods (for internal use)
'=========================================================================

' Accept a new client connection (called internally)
Friend Sub AcceptFrom(oListenSocket As cAsyncSocket, oParent As cWinsock)
    On Error GoTo EH

    Set m_oSocket = New cAsyncSocket
    If Not oListenSocket.Accept(m_oSocket) Then
        pvRaiseError oListenSocket.LastError
        Exit Sub
    End If

    m_eState = sckConnected
    m_eProtocol = wskTCPProtocol
    m_bIsAcceptedClient = True
    Set m_oParentServer = oParent

    ' Get peer info
    Dim sAddr As String
    Dim lPort As Long
    m_oSocket.GetPeerName sAddr, lPort
    m_sRemoteHostIP = sAddr
    m_sRemoteHost = sAddr
    m_lRemotePort = lPort

    Exit Sub
EH:
    pvRaiseError ERR.Number, ERR.Description
End Sub

' Set UDP client info (for server)
Friend Sub SetUdpClientInfo(sAddress As String, ByVal lPort As Long)
    m_sRemoteHostIP = sAddress
    m_sRemoteHost = sAddress
    m_lRemotePort = lPort
End Sub

'=========================================================================
' Socket Events
'=========================================================================

Private Sub m_oSocket_OnResolve(IpAddress As String)
    m_sRemoteHostIP = IpAddress
    m_eState = sckHostResolved
End Sub

Private Sub m_oSocket_OnConnect()
    m_eState = sckConnected
    
    ' Auto-start heartbeat if configured
    If Not m_oHeartbeat Is Nothing Then
        If m_oHeartbeat.Enabled Then
            m_oHeartbeat.ResetActivity
        End If
    End If
    
    RaiseEvent Connect(Me)
End Sub

Private Sub m_oSocket_OnClose()
    Dim oParent As cWinsock

    On Error Resume Next

    ' Prevent duplicate CloseEvent
    If m_bCloseEventRaised Then Exit Sub
    m_bCloseEventRaised = True

    m_eState = sckClosing

    ' Notify parent server to remove this client
    If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
        Set oParent = m_oParentServer
    End If

    ' If this is an accepted client, raise event through parent server
    ' so the host can receive it via the server's event handler
    If Not oParent Is Nothing Then
        oParent.RaiseCloseEvent Me
    Else
        RaiseEvent CloseEvent(Me)
    End If

    ' Remove from parent's client list
    If Not oParent Is Nothing Then
        oParent.RemoveClient Me
    End If

    m_eState = sckClosed
End Sub

Private Sub m_oSocket_OnAccept()
    Dim oClient As cWinsock
    
    On Error GoTo EH
    
    ' Create new client socket
    Set oClient = New cWinsock
    oClient.AcceptFrom m_oSocket, Me
    
    If oClient.State = wckConnected Then
        ' Add to clients collection
        m_lClientCounter = m_lClientCounter + 1
        oClient.Tag = "#" & m_lClientCounter
        m_cClients.Add oClient, oClient.Tag
        
        ' Inherit packet protocol configuration from server
        ' Each client gets its own independent protocol instance
        If Not m_oPacketProtocol Is Nothing Then
            If m_oPacketProtocol.ProtocolType <> ppNone Then
                oClient.PacketProtocol = m_oPacketProtocol.ProtocolType
                Select Case m_oPacketProtocol.ProtocolType
                Case ppDelimiter
                    oClient.Delimiter = m_oPacketProtocol.Delimiter
                Case ppFixedLength
                    oClient.FixedLength = m_oPacketProtocol.FixedLength
                Case ppLengthHeader
                    oClient.HeaderBytes = m_oPacketProtocol.HeaderBytes
                    oClient.HeaderEndian = m_oPacketProtocol.Endian
                End Select
            End If
        End If
        
        ' Inherit heartbeat configuration from server
        ' New client gets its own heartbeat instance for idle tracking
        If Not m_oHeartbeat Is Nothing Then
            If m_oHeartbeat.Enabled Then
                oClient.AutoHeartbeat = True
                oClient.HeartbeatTimeout = m_oHeartbeat.ServerTimeout
            End If
        End If
        
        ' Raise event to notify user - allow rejection
        Dim bDisConnect As Boolean
        bDisConnect = False
        RaiseEvent ConnectionRequest(oClient, bDisConnect)
        
        ' If user wants to reject, close and remove client
        If bDisConnect Then
            oClient.Close_
            RemoveClient oClient
            Exit Sub
        End If
        
        ' Raise ClientCountChange event for new connection
        RaiseEvent ClientCountChange(oClient, m_cClients.count, wskClientConnected)
    End If
    
    Exit Sub
EH:
    Debug.Print "OnAccept Error: " & ERR.Description
End Sub

Private Sub m_oSocket_OnReceive()
    Dim baBuffer() As Byte
    Dim lBytes As Long
    Dim lIdx As Long
    
    On Error GoTo EH
    
    ' Reset heartbeat activity timer
    If Not m_oHeartbeat Is Nothing Then m_oHeartbeat.ResetActivity
    
    lBytes = m_oSocket.AvailableBytes
    
    ' For UDP server, handle differently
    If m_bIsServer And m_eProtocol = wskUDPProtocol Then
        Dim sAddress As String
        Dim lPort As Long
        Dim sKey As String
        Dim oClient As cWinsock
        
        If m_oSocket.ReceiveArray(baBuffer, sAddress, lPort) Then
            sKey = sAddress & ":" & lPort
            
            ' Find or create client
            On Error Resume Next
            Set oClient = m_cClients(sKey)
            On Error GoTo EH
            
            If oClient Is Nothing Then
                Set oClient = New cWinsock
                oClient.SetUdpClientInfo sAddress, lPort
                ' UDP client doesn't need its own socket - it's just an identifier
                ' All data is received through the server's socket
                oClient.Protocol = wskUDPProtocol
                Set oClient.ParentServer = Me
                oClient.Tag = sKey
                m_cClients.Add oClient, sKey
                
                ' Raise event to notify user - allow rejection
                Dim bDisConnect As Boolean
                bDisConnect = False
                RaiseEvent ConnectionRequest(oClient, bDisConnect)
                
                ' If user wants to reject, remove client and discard data
                If bDisConnect Then
                    RemoveClient oClient
                    Exit Sub
                End If
                
                ' Raise ClientCountChange event for new UDP client
                RaiseEvent ClientCountChange(oClient, m_cClients.count, wskClientConnected)
            End If
            
            ' Store data and raise event
            oClient.UserData = baBuffer
            RaiseEvent DataArrival(oClient, UBound(baBuffer) + 1)
        End If
    Else
        ' TCP or client socket
        If m_oSocket.ReceiveArray(baBuffer) Then
            If UBound(baBuffer) >= 0 Then
                '--- Packet protocol: auto-deframe ---
                If Not m_oPacketProtocol Is Nothing Then
                    If m_oPacketProtocol.ProtocolType <> ppNone Then
                        ' Decode through protocol handler
                        Dim colMessages As Collection
                        Dim vMsg As Variant
                        Dim baMsg() As Byte
                        Dim lMsgIdx As Long
                        
                        Set colMessages = m_oPacketProtocol.Decode(baBuffer)
                        
                        If Not colMessages Is Nothing Then
                            ' Process each complete message
                            For Each vMsg In colMessages
                                baMsg = vMsg
                                lMsgIdx = 0
                                
                                ' Replace recv buffer with this message's data
                                ' (GetData will read from m_baRecvBuffer)
                                m_baRecvBuffer = baMsg
                                
                                ' Calculate message size
                                On Error Resume Next
                                lBytes = UBound(baMsg) + 1
                                If ERR.Number <> 0 Then lBytes = 0
                                On Error GoTo EH
                                
                                ' Raise MessageArrival event
                                If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
                                    m_oParentServer.RaiseMessageArrivalEvent Me, lBytes
                                Else
                                    RaiseEvent MessageArrival(Me, lBytes)
                                End If
                                
                                ' Also raise DataArrival for backward compatibility
                                If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
                                    m_oParentServer.RaiseDataArrivalEvent Me, lBytes
                                Else
                                    RaiseEvent DataArrival(Me, lBytes)
                                End If
                            Next
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                '--- Normal (no protocol): append to receive buffer ---
                ' Append to receive buffer
                If UBound(m_baRecvBuffer) >= 0 Then
                    lIdx = UBound(m_baRecvBuffer) + 1
                    ReDim Preserve m_baRecvBuffer(0 To lIdx + UBound(baBuffer)) As Byte
                    CopyMemory m_baRecvBuffer(lIdx), baBuffer(0), UBound(baBuffer) + 1
                Else
                    m_baRecvBuffer = baBuffer
                End If
                
                lBytes = UBound(m_baRecvBuffer) + 1
                
                ' If this is an accepted client, raise event through parent server
                ' so the host can receive it via the server's event handler
                If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
                    m_oParentServer.RaiseDataArrivalEvent Me, lBytes
                Else
                    ' Direct client or standalone socket
                    RaiseEvent DataArrival(Me, lBytes)
                End If
            End If
        End If
    End If
    
    Exit Sub
EH:
    Debug.Print "OnReceive Error: " & ERR.Description
End Sub

Private Sub m_oSocket_OnSend()
    ' Reset heartbeat activity timer
    If Not m_oHeartbeat Is Nothing Then m_oHeartbeat.ResetActivity
    
    ' If this is an accepted client, raise event through parent server
    If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
        m_oParentServer.RaiseSendCompleteEvent Me
    Else
        RaiseEvent SendComplete(Me)
    End If
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    m_eState = sckError
    pvRaiseError ErrorCode
End Sub

'=========================================================================
' Friend Methods
'=========================================================================

Friend Sub RemoveClient(oClient As cWinsock)
    On Error Resume Next
    If Not m_cClients Is Nothing Then
        ' Only process if client is still in collection to avoid duplicate events
        If m_cClients.Exists(oClient.Tag) Then
            m_cClients.Remove oClient.Tag
            ' Unbind user if client was bound (including from groups)
            If LenB(CStr(oClient.CurrentUser)) <> 0 Then
                UnbindUserFromAllGroups oClient.CurrentUser
                m_Users.Remove oClient.CurrentUser
            End If
            ' Raise ClientCountChange event for disconnection
            RaiseEvent ClientCountChange(oClient, m_cClients.count, wskClientDisconnected)
        End If
    End If
End Sub

' Raise DataArrival event for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseDataArrivalEvent(oClient As cWinsock, ByVal bytesTotal As Long)
    RaiseEvent DataArrival(oClient, bytesTotal)
End Sub

' Raise CloseEvent for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseCloseEvent(oClient As cWinsock)
    RaiseEvent CloseEvent(oClient)
End Sub

' Raise SendComplete event for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseSendCompleteEvent(oClient As cWinsock)
    RaiseEvent SendComplete(oClient)
End Sub

' Raise Error event for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseErrorEvent(oClient As cWinsock, ByVal Number As Long, Description As String)
    RaiseEvent Error(oClient, Number, Description, Number)
End Sub

' Raise MessageArrival event for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseMessageArrivalEvent(oClient As cWinsock, ByVal bytesTotal As Long)
    RaiseEvent MessageArrival(oClient, bytesTotal)
End Sub

'=========================================================================
' Private Methods
'=========================================================================

Private Sub pvRaiseError(ByVal ErrorCode As Long, Optional Description As String)
    Dim sDesc As String

    m_eState = sckError

    If LenB(Description) <> 0 Then
        sDesc = Description
    Else
        sDesc = GetErrorDescription(ErrorCode)
    End If

    ' If this is an accepted client, raise event through parent server
    If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
        m_oParentServer.RaiseErrorEvent Me, ErrorCode, sDesc
    Else
        RaiseEvent Error(Me, ErrorCode, sDesc, ErrorCode)
    End If
End Sub



'=========================================================================
' Sync Methods (blocking operations with timeout)
'=========================================================================

' Synchronous connection with timeout
Public Function SyncConnect(HostAddress As String, ByVal HostPort As Long, Optional ByVal Timeout As Long) As Boolean
    Dim eSocketType As UcsAsyncSocketTypeEnum

    On Error GoTo EH

    ' Close existing connection
    Close_

    ' Update remote host/port
    m_sRemoteHost = HostAddress
    m_lRemotePort = HostPort

    ' Validate parameters
    If LenB(m_sRemoteHost) = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "RemoteHost not specified"
    End If
    If m_lRemotePort = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "RemotePort not specified"
    End If

    ' Determine socket type
    eSocketType = IIf(m_eProtocol = wskTCPProtocol, ucsSckStream, ucsSckDatagram)

    ' Create socket
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=eSocketType) Then
        pvRaiseError m_oSocket.LastError
        Exit Function
    End If

    m_eState = sckResolvingHost

    ' Sync connect
    SyncConnect = m_oSocket.SyncConnect(m_sRemoteHost, m_lRemotePort, Timeout)

    ' Update state on success
    If SyncConnect Then
        m_eState = sckConnected
        ' Get peer info
        Dim sAddr As String
        Dim lPort As Long
        m_oSocket.GetPeerName sAddr, lPort
        m_sRemoteHostIP = sAddr
    End If

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous receive text with timeout
Public Function SyncReceiveText( _
    Optional ByVal NeedLen As Long = 1, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional HostPort As Long, _
    Optional ByVal CodePage As EnumScpCodePage = wcpUtf8) As String

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncReceiveText = m_oSocket.SyncReceiveText(NeedLen, Timeout, HostAddress, HostPort, CodePage)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous receive byte array with timeout
Public Function SyncReceiveArray( _
    Buffer() As Byte, _
    Optional ByVal NeedLen As Long, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional HostPort As Long, _
    Optional ByVal flags As Long = 0) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncReceiveArray = m_oSocket.SyncReceiveArray(Buffer, NeedLen, Timeout, HostAddress, HostPort, flags)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous receive to buffer pointer with timeout
Public Function SyncReceive( _
    ByVal BufPtr As Long, _
    ByVal bufLen As Long, _
    ByRef Received As Long, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional HostPort As Long, _
    Optional ByVal flags As Long = 0) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncReceive = m_oSocket.SyncReceive(BufPtr, bufLen, Received, Timeout, HostAddress, HostPort, flags)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous send text with timeout
Public Function SyncSendText( _
    Text As String, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional ByVal HostPort As Long, _
    Optional ByVal CodePage As EnumScpCodePage = wcpUtf8) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncSendText = m_oSocket.SyncSendText(Text, Timeout, HostAddress, HostPort, CodePage)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous send byte array with timeout
Public Function SyncSendArray( _
    Buffer() As Byte, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional ByVal HostPort As Long, _
    Optional ByVal flags As Long = 0) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncSendArray = m_oSocket.SyncSendArray(Buffer, Timeout, HostAddress, HostPort, flags)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Synchronous send from buffer pointer with timeout
Public Function SyncSend( _
    ByVal BufPtr As Long, _
    ByVal bufLen As Long, _
    Optional ByVal Timeout As Long, _
    Optional HostAddress As String, _
    Optional ByVal HostPort As Long, _
    Optional ByVal flags As Long = 0) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncSend = m_oSocket.SyncSend(BufPtr, bufLen, Timeout, HostAddress, HostPort, flags)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Wait for specific socket event with timeout
Public Function SyncWaitForEvent( _
    ByVal Timeout As Long, _
    Optional ByVal EventMask As WinSocketEventMaskEnum = wskSfdAll) As Boolean

    On Error GoTo EH

    If m_oSocket Is Nothing Then
        ERR.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If

    SyncWaitForEvent = m_oSocket.SyncWaitForEvent(Timeout, EventMask)

    Exit Function
EH:
    pvRaiseError ERR.Number, ERR.Description
End Function

' Cancel pending synchronous wait operation
Public Sub SyncCancelWait()
    On Error Resume Next

    If Not m_oSocket Is Nothing Then
        m_oSocket.SyncCancelWait
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_eProtocol = DEF_PROTOCOL
    m_lLocalPort = DEF_LOCALPORT
    m_lRemotePort = DEF_REMOTEPORT
    m_baRecvBuffer = vbNullString
    m_Users.CompareMode = TextCompare
    m_Groups.CompareMode = TextCompare
    CurrentGroups.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    ' Unbind user if this client was bound
    If LenB(CStr(CurrentUser)) <> 0 Then
        If Not m_oParentServer Is Nothing Then
            m_oParentServer.UnbindUserFromAllGroups CurrentUser
            m_oParentServer.UnbindUser CurrentUser
        End If
    End If
    
    ' Clear CurrentGroups
    Set CurrentGroups = New Dictionary
    CurrentGroups.CompareMode = TextCompare
    
    ' Clear packet protocol
    Set m_oPacketProtocol = Nothing
    
    ' Clear heartbeat
    Set m_oHeartbeat = Nothing
    
    Close_
    ' Raise ClientCountChange event for new connection
    RaiseEvent ClientCountChange(Nothing, 0, wskServerStoped)
End Sub
