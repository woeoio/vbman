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
' cWinsock - Simplified Winsock Wrapper for VB6
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

'=========================================================================
' Public Enums
'=========================================================================

Public Enum WinsockProtocol
    sckTCPProtocol = 1
    sckUDPProtocol = 2
End Enum

Public Enum WinsockState
    sckClosed = 0
    sckOpen = 1
    sckListening = 2
    sckConnectionPending = 3
    sckResolvingHost = 4
    sckHostResolved = 5
    sckConnecting = 6
    sckConnected = 7
    sckClosing = 8
    sckError = 9
End Enum

'=========================================================================
' Private Constants
'=========================================================================

Private Const DEF_LOCALPORT         As Long = 0
Private Const DEF_PROTOCOL          As Long = sckTCPProtocol
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

'=========================================================================
' Properties
'=========================================================================

' Current socket state
Public Property Get State() As WinsockState
    State = m_eState
End Property

' Protocol type (TCP/UDP)
Public Property Get Protocol() As WinsockProtocol
    Protocol = m_eProtocol
End Property

Public Property Let Protocol(ByVal Value As WinsockProtocol)
    If m_eState <> sckClosed Then
        Err.Raise vbObjectError, MODULE_NAME, "Cannot change protocol while socket is open"
    End If
    m_eProtocol = Value
End Property

Public Property Let RecvBuffer(ByRef Value() As Byte)
    m_baRecvBuffer = Value
End Property

' Local port to bind
Public Property Get LocalPort() As Long
    Dim lPort As Long
    If m_eState <> sckClosed And Not m_oSocket Is Nothing Then
        m_oSocket.GetSockName vbNullString, lPort
        LocalPort = lPort
    Else
        LocalPort = m_lLocalPort
    End If
End Property

Public Property Let LocalPort(ByVal Value As Long)
    If m_eState <> sckClosed Then
        Err.Raise vbObjectError, MODULE_NAME, "Cannot change local port while socket is open"
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
    If m_oSocket Is Nothing And Not m_oParentServer Is Nothing And m_eProtocol = sckUDPProtocol Then
        RemoteHostIP = m_sRemoteHostIP
        Exit Property
    End If
    If m_eState >= sckHostResolved And Not m_oSocket Is Nothing Then
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
        ClientCount = m_cClients.Count
    End If
End Property

' Internal socket reference
Friend Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

'=========================================================================
' Public Methods
'=========================================================================

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
        Err.Raise vbObjectError, MODULE_NAME, "RemoteHost not specified"
    End If
    If m_lRemotePort = 0 Then
        Err.Raise vbObjectError, MODULE_NAME, "RemotePort not specified"
    End If

    ' Determine socket type
    eSocketType = IIf(m_eProtocol = sckTCPProtocol, ucsSckStream, ucsSckDatagram)

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
    pvRaiseError Err.Number, Err.Description
End Sub

' Start listening for incoming connections
Public Sub Listen(Optional ByVal Port As Long)
    Dim eSocketType As UcsAsyncSocketTypeEnum

    On Error GoTo EH

    ' Close existing connection
    Close_

    ' Update port if provided
    If Port <> 0 Then
        m_lLocalPort = Port
    End If

    ' Determine socket type
    eSocketType = IIf(m_eProtocol = sckTCPProtocol, ucsSckStream, ucsSckDatagram)

    ' Create socket
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketPort:=m_lLocalPort, SocketType:=eSocketType) Then
        pvRaiseError m_oSocket.LastError
        Exit Sub
    End If

    ' For TCP, start listening
    If m_eProtocol = sckTCPProtocol Then
        If Not m_oSocket.Listen() Then
            pvRaiseError m_oSocket.LastError
            Exit Sub
        End If
    End If

    m_eState = sckListening
    m_bIsServer = True
    Set m_cClients = New Collection

    Exit Sub
EH:
    pvRaiseError Err.Number, Err.Description
End Sub

' Send string data
' CodePage: Text encoding for string data (default: ACP/GBK)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
Public Sub SendData(Data As Variant, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpAcp)
    On Error GoTo EH
    
    If m_oSocket Is Nothing Then
        Err.Raise vbObjectError, MODULE_NAME, "Socket not connected"
    End If
    
    Select Case VarType(Data)
    Case vbString
        If m_eProtocol = sckUDPProtocol And m_bIsServer Then
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

        If m_eProtocol = sckUDPProtocol And m_bIsServer Then
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
        Err.Raise vbObjectError, MODULE_NAME, "Unsupported data type: " & TypeName(Data)
    End Select
    
    Exit Sub
EH:
    pvRaiseError Err.Number, Err.Description
End Sub

' Receive data from buffer
' CodePage: Text encoding for string data (default: ACP/GBK for compatibility)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
Public Sub GetData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpAcp)
    Dim baBuffer() As Byte
    Dim lIdx As Long
    Dim oSocket As cAsyncSocket

    On Error GoTo EH

    If VarType_ = 0 Then
        VarType_ = VarType(Data)
    End If

    ' First check internal buffer
    If UBound(m_baRecvBuffer) >= 0 Then
        If MaxLen < 0 Or MaxLen >= UBound(m_baRecvBuffer) + 1 Then
            baBuffer = m_baRecvBuffer
            m_baRecvBuffer = vbNullString
        Else
            ReDim baBuffer(0 To MaxLen - 1) As Byte
            CopyMemory baBuffer(0), m_baRecvBuffer(0), MaxLen
            lIdx = UBound(m_baRecvBuffer) + 1 - MaxLen
            If lIdx > 0 Then
                Dim baTemp() As Byte
                ReDim baTemp(0 To lIdx - 1) As Byte
                CopyMemory baTemp(0), m_baRecvBuffer(MaxLen), lIdx
                m_baRecvBuffer = baTemp
            Else
                m_baRecvBuffer = vbNullString
            End If
        End If
    ElseIf Not m_oSocket Is Nothing Then
        ' Read from socket
        If m_oSocket.ReceiveArray(baBuffer) Then
            ' If MaxLen is specified and data exceeds it, truncate and save remainder
            If MaxLen >= 0 And UBound(baBuffer) >= 0 And UBound(baBuffer) + 1 > MaxLen Then
                ' Save remainder to internal buffer
                Dim baRemainder() As Byte
                ReDim baRemainder(0 To UBound(baBuffer) - MaxLen) As Byte
                CopyMemory baRemainder(0), baBuffer(MaxLen), UBound(baBuffer) - MaxLen + 1
                m_baRecvBuffer = baRemainder
                
                ' Truncate baBuffer to MaxLen
                ReDim Preserve baBuffer(0 To MaxLen - 1) As Byte
            End If
        Else
            baBuffer = vbNullString
        End If
    ElseIf VarType(m_vUserData) = (vbByte + vbArray) Then
        ' For UDP server mode clients, data is stored in UserData
        baBuffer = m_vUserData
        ' Clear UserData after reading
        m_vUserData = Empty
        ' If MaxLen is specified, truncate
        If MaxLen >= 0 And UBound(baBuffer) >= 0 And UBound(baBuffer) + 1 > MaxLen Then
            Dim baTruncated() As Byte
            ReDim baTruncated(0 To MaxLen - 1) As Byte
            CopyMemory baTruncated(0), baBuffer(0), MaxLen
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
    pvRaiseError Err.Number, Err.Description
End Sub

' Peek data without removing from buffer
' CodePage: Text encoding for string data (default: ACP/GBK for compatibility)
'           Use ucsScpAcp for GBK/ACP encoding, ucsScpUtf8 for UTF-8 encoding
Public Sub PeekData(Data As Variant, Optional ByVal VarType_ As Long, Optional ByVal MaxLen As Long = -1, Optional ByVal CodePage As UcsAsyncSocketCodePageEnum = ucsScpAcp)
    Dim baBuffer() As Byte

    On Error GoTo EH

    If VarType_ = 0 Then
        VarType_ = VarType(Data)
    End If

    ' Copy from internal buffer
    If UBound(m_baRecvBuffer) >= 0 Then
        If MaxLen < 0 Or MaxLen >= UBound(m_baRecvBuffer) + 1 Then
            baBuffer = m_baRecvBuffer
        Else
            ReDim baBuffer(0 To MaxLen - 1) As Byte
            CopyMemory baBuffer(0), m_baRecvBuffer(0), MaxLen
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
    pvRaiseError Err.Number, Err.Description
End Sub

' Close socket
Public Sub Close_()
    On Error Resume Next

    ' Close all client connections for server socket
    If m_bIsServer And Not m_cClients Is Nothing Then
        Dim oClient As cWinsock
        For Each oClient In m_cClients
            oClient.Close_
        Next
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
    Set m_oParentServer = Nothing
    m_baRecvBuffer = vbNullString
End Sub

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
    m_eProtocol = sckUDPProtocol
    m_bIsServer = True
    Set m_cClients = New Collection

    Exit Sub
EH:
    pvRaiseError Err.Number, Err.Description
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
    m_eProtocol = sckTCPProtocol
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
    pvRaiseError Err.Number, Err.Description
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
    RaiseEvent Connect(Me)
End Sub

Private Sub m_oSocket_OnClose()
    Dim oParent As cWinsock

    On Error Resume Next

    m_eState = sckClosing

    ' Notify parent server to remove this client
    If m_bIsAcceptedClient And Not m_oParentServer Is Nothing Then
        Set oParent = m_oParentServer
    End If

    RaiseEvent CloseEvent(Me)

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
    
    If oClient.State = sckConnected Then
        ' Add to clients collection
        m_lClientCounter = m_lClientCounter + 1
        oClient.Tag = "#" & m_lClientCounter
        m_cClients.Add oClient, oClient.Tag
        
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
    End If
    
    Exit Sub
EH:
    Debug.Print "OnAccept Error: " & Err.Description
End Sub

Private Sub m_oSocket_OnReceive()
    Dim baBuffer() As Byte
    Dim lBytes As Long
    Dim lIdx As Long
    
    On Error GoTo EH
    
    lBytes = m_oSocket.AvailableBytes
    
    ' For UDP server, handle differently
    If m_bIsServer And m_eProtocol = sckUDPProtocol Then
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
                oClient.Protocol = sckUDPProtocol
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
            End If
            
            ' Store data and raise event
            oClient.UserData = baBuffer
            RaiseEvent DataArrival(oClient, UBound(baBuffer) + 1)
        End If
    Else
        ' TCP or client socket
        If m_oSocket.ReceiveArray(baBuffer) Then
            If UBound(baBuffer) >= 0 Then
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
    Debug.Print "OnReceive Error: " & Err.Description
End Sub

Private Sub m_oSocket_OnSend()
    RaiseEvent SendComplete(Me)
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
        m_cClients.Remove oClient.Tag
    End If
End Sub

' Raise DataArrival event for a client (used by accepted clients to notify through parent server)
Friend Sub RaiseDataArrivalEvent(oClient As cWinsock, ByVal bytesTotal As Long)
    RaiseEvent DataArrival(oClient, bytesTotal)
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
    
    RaiseEvent Error(Me, ErrorCode, sDesc, ErrorCode)
End Sub



'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_eProtocol = DEF_PROTOCOL
    m_lLocalPort = DEF_LOCALPORT
    m_lRemotePort = DEF_REMOTEPORT
    m_baRecvBuffer = vbNullString
End Sub

Private Sub Class_Terminate()
    Close_
End Sub
