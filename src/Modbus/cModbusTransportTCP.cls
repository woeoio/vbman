VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModbusTransportTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' cModbusTransportTCP - Modbus TCP Transport Layer for VB6
'
' Purpose: Provides TCP/IP transport layer for Modbus communication
'          - Client mode: Connect to Modbus TCP server
'          - Server mode: Accept connections from Modbus TCP clients
'          - MBAP header handling
'          - Multiple client management
'
' Author: Auto
' Date: 2026-01-16
'
'=========================================================================
Option Explicit

'=========================================================================
' Public Enums
'=========================================================================

Public Enum ModbusTCPState
    MB_TCP_STATE_CLOSED = 0
    MB_TCP_STATE_CONNECTING = 1
    MB_TCP_STATE_CONNECTED = 2
    MB_TCP_STATE_LISTENING = 3
    MB_TCP_STATE_ERROR = 4
End Enum

Public Enum ModbusTCPMode
    MB_TCP_MODE_CLIENT = 1
    MB_TCP_MODE_SERVER = 2
End Enum

'=========================================================================
' Events
'=========================================================================

' Client mode events
Event OnConnect()
Event OnDisconnect()
Event OnError(ByVal Description As String)
Event OnDataReceived(Data() As Byte)

' Server mode events
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String)
Event OnClientDisconnect(ByVal ClientID As String)
Event OnClientDataReceived(ByVal ClientID As String, Data() As Byte)

'=========================================================================
' Member Variables
'=========================================================================

Private m_Mode As ModbusTCPMode
Private m_State As ModbusTCPState
Private m_Host As String
Private m_Port As Long
Private m_BindAddress As String
Private m_Timeout As Long
Private m_TransactionID As Long

' Client mode socket
Private WithEvents m_Socket As cWinsock
Attribute m_Socket.VB_VarHelpID = -1

' Server mode
Private WithEvents m_ListenSocket As cWinsock
Attribute m_ListenSocket.VB_VarHelpID = -1
Private m_Clients As Collection
Private m_ClientIDMap As Collection
Private m_ClientCounter As Long

'=========================================================================
' Properties
'=========================================================================

Public Property Get Mode() As ModbusTCPMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal Value As ModbusTCPMode)
    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Cannot change mode while connected"
        Exit Property
    End If
    m_Mode = Value
End Property

Public Property Get State() As ModbusTCPState
    State = m_State
End Property

Public Property Get Host() As String
    Host = m_Host
End Property

Public Property Let Host(ByVal Value As String)
    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Cannot change host while connected"
        Exit Property
    End If
    m_Host = Value
End Property

Public Property Get Port() As Long
    Port = m_Port
End Property

Public Property Let Port(ByVal Value As Long)
    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Cannot change port while connected"
        Exit Property
    End If
    m_Port = Value
End Property

Public Property Get BindAddress() As String
    BindAddress = m_BindAddress
End Property

Public Property Let BindAddress(ByVal Value As String)
    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Cannot change bind address while connected"
        Exit Property
    End If
    m_BindAddress = Value
End Property

Public Property Get Timeout() As Long
    Timeout = m_Timeout
End Property

Public Property Let Timeout(ByVal Value As Long)
    If Value > 0 Then m_Timeout = Value
End Property

Public Property Get ClientCount() As Long
    On Error Resume Next
    If Not m_Clients Is Nothing Then
        ClientCount = m_Clients.Count
    End If
End Property

Public Property Get IsConnected() As Boolean
    IsConnected = (m_State = MB_TCP_STATE_CONNECTED)
End Property

Public Property Get IsListening() As Boolean
    IsListening = (m_State = MB_TCP_STATE_LISTENING)
End Property

'=========================================================================
' Public Methods - Client Mode
'=========================================================================

' Connect to server
Public Sub Connect(Optional ByVal HostAddr As String = "", Optional ByVal PortNum As Long = 0)
    On Error GoTo ErrorHandler

    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Already connected or connecting"
        Exit Sub
    End If

    If LenB(HostAddr) > 0 Then m_Host = HostAddr
    If PortNum > 0 Then m_Port = PortNum

    If LenB(m_Host) = 0 Then
        RaiseError "Host not specified"
        Exit Sub
    End If

    If m_Port = 0 Then m_Port = MODBUS_TCP_PORT

    m_Mode = MB_TCP_MODE_CLIENT
    m_State = MB_TCP_STATE_CONNECTING

    Set m_Socket = New cWinsock
    m_Socket.Protocol = wskTCPProtocol
    m_Socket.Connect m_Host, m_Port

    Exit Sub
ErrorHandler:
    m_State = MB_TCP_STATE_ERROR
    RaiseError "Connect failed: " & ERR.Description
End Sub

' Disconnect from server
Public Sub Disconnect()
    On Error Resume Next

    If Not m_Socket Is Nothing Then
        m_Socket.Close_
        Set m_Socket = Nothing
    End If

    m_State = MB_TCP_STATE_CLOSED
    RaiseEvent OnDisconnect
End Sub

' Send request and wait for response (synchronous, for Master)
Public Function SendAndReceive(ByRef Request() As Byte) As Byte()
    Dim baResponse() As Byte
    Dim baRecvBuffer() As Byte
    Dim lStartTime As Long
    Dim lElapsed As Long
    Dim lMBAPLen As Long

    On Error GoTo ErrorHandler

    If m_State <> MB_TCP_STATE_CONNECTED Then
        RaiseError "Not connected"
        Exit Function
    End If

    ' Send request
    m_Socket.SendData Request

    ' Wait for response
    lStartTime = Modbus_GetTickCount()

    Do
        If m_Socket.BytesReceived > 0 Then
            m_Socket.GetData baRecvBuffer, vbByte + vbArray

            If UBound(baRecvBuffer) >= MODBUS_TCP_MBAP_SIZE - 1 Then
                ' Verify transaction ID
                Dim lTransactionID As Long
                lTransactionID = CLng(baRecvBuffer(0)) * 256 + CLng(baRecvBuffer(1))

                If lTransactionID = m_TransactionID Then
                    ' Get length from MBAP
                    lMBAPLen = CLng(baRecvBuffer(4)) * 256 + CLng(baRecvBuffer(5))

                    ' Build response (UnitID + PDU)
                    ReDim baResponse(lMBAPLen) As Byte
                    baResponse(0) = baRecvBuffer(6)  ' Unit ID
                    Modbus_CopyMemory baResponse(1), baRecvBuffer(MODBUS_TCP_MBAP_SIZE), lMBAPLen - 1

                    SendAndReceive = baResponse
                    Exit Function
                End If
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

' Send data without waiting (for responses)
Public Function Send(ByRef Data() As Byte) As Boolean
    On Error GoTo ErrorHandler

    If m_State <> MB_TCP_STATE_CONNECTED Then
        RaiseError "Not connected"
        Send = False
        Exit Function
    End If

    m_Socket.SendData Data
    Send = True
    Exit Function

ErrorHandler:
    RaiseError "Send failed: " & ERR.Description
    Send = False
End Function

'=========================================================================
' Public Methods - Server Mode
'=========================================================================

' Start listening
Public Sub Listen(Optional ByVal PortNum As Long = 0, Optional ByVal BindAddr As String = "")
    On Error GoTo ErrorHandler

    If m_State <> MB_TCP_STATE_CLOSED Then
        RaiseError "Already listening or connected"
        Exit Sub
    End If

    If PortNum > 0 Then m_Port = PortNum
    If LenB(BindAddr) > 0 Then m_BindAddress = BindAddr

    If m_Port = 0 Then m_Port = MODBUS_TCP_PORT

    m_Mode = MB_TCP_MODE_SERVER

    ' Initialize collections
    Set m_Clients = New Collection
    Set m_ClientIDMap = New Collection
    m_ClientCounter = 0

    ' Create listening socket
    Set m_ListenSocket = New cWinsock
    m_ListenSocket.Protocol = wskTCPProtocol

    ' Listen with bind address if specified
    If LenB(m_BindAddress) > 0 And m_BindAddress <> "0.0.0.0" Then
        m_ListenSocket.Listen m_Port, m_BindAddress
    Else
        m_ListenSocket.Listen m_Port
    End If

    m_State = MB_TCP_STATE_LISTENING

    Exit Sub
ErrorHandler:
    m_State = MB_TCP_STATE_ERROR
    RaiseError "Listen failed: " & ERR.Description
End Sub

' Stop listening
Public Sub StopListening()
    On Error Resume Next

    ' Close all client connections
    If Not m_Clients Is Nothing Then
        Dim oClient As cWinsock
        For Each oClient In m_Clients
            oClient.Close_
        Next
        Set m_Clients = Nothing
    End If

    Set m_ClientIDMap = Nothing

    ' Close listening socket
    If Not m_ListenSocket Is Nothing Then
        m_ListenSocket.Close_
        Set m_ListenSocket = Nothing
    End If

    m_State = MB_TCP_STATE_CLOSED
End Sub

' Send data to specific client
Public Function SendToClient(ByVal ClientID As String, ByRef Data() As Byte) As Boolean
    Dim oClient As cWinsock

    On Error GoTo ErrorHandler

    Set oClient = GetClientByID(ClientID)
    If oClient Is Nothing Then
        RaiseError "Client not found: " & ClientID
        SendToClient = False
        Exit Function
    End If

    oClient.SendData Data
    SendToClient = True
    Exit Function

ErrorHandler:
    RaiseError "SendToClient failed: " & ERR.Description
    SendToClient = False
End Function

' Disconnect specific client
Public Sub DisconnectClient(ByVal ClientID As String)
    RemoveClient ClientID, "Disconnected by server"
End Sub

'=========================================================================
' Public Methods - MBAP Frame Building
'=========================================================================

' Build MBAP frame from PDU
Public Function BuildMBAPFrame(ByRef PDU() As Byte, ByVal UnitID As Byte) As Byte()
    Dim baFrame() As Byte
    Dim lPDULen As Long

    lPDULen = UBound(PDU) + 1

    ' Increment transaction ID
    m_TransactionID = m_TransactionID + 1
    If m_TransactionID > 65535 Then m_TransactionID = 1

    ' MBAP Header (7 bytes) + PDU
    ReDim baFrame(MODBUS_TCP_MBAP_SIZE + lPDULen - 1) As Byte

    ' Transaction ID (2 bytes)
    baFrame(0) = (m_TransactionID \ 256) And &HFF
    baFrame(1) = m_TransactionID And &HFF

    ' Protocol ID (2 bytes) - always 0 for Modbus
    baFrame(2) = 0
    baFrame(3) = 0

    ' Length (2 bytes) - Unit ID + PDU length
    baFrame(4) = ((lPDULen + 1) \ 256) And &HFF
    baFrame(5) = (lPDULen + 1) And &HFF

    ' Unit ID (1 byte)
    baFrame(6) = UnitID

    ' Copy PDU
    Modbus_CopyMemory baFrame(MODBUS_TCP_MBAP_SIZE), PDU(0), lPDULen

    BuildMBAPFrame = baFrame
End Function

' Get current transaction ID
Public Property Get TransactionID() As Long
    TransactionID = m_TransactionID
End Property

'=========================================================================
' Private Methods - Client Management
'=========================================================================

Private Function GetClientByID(ByVal ClientID As String) As cWinsock
    On Error Resume Next
    Set GetClientByID = m_Clients(ClientID)
End Function

Private Function GetClientIDBySocket(ByVal Socket As cWinsock) As String
    Dim sKey As String
    Dim oClient As cWinsock
    Dim vKey As Variant

    On Error Resume Next

    ' Use object hash as lookup key
    sKey = "S" & ObjPtr(Socket)

    GetClientIDBySocket = m_ClientIDMap(sKey)
End Function

Private Sub AddClient(ByVal Client As cWinsock, ByVal ClientID As String)
    Dim sKey As String

    On Error Resume Next

    m_Clients.Add Client, ClientID

    ' Add reverse mapping
    sKey = "S" & ObjPtr(Client)
    m_ClientIDMap.Add ClientID, sKey
End Sub

Private Sub RemoveClient(ByVal ClientID As String, ByVal Reason As String)
    Dim oClient As cWinsock
    Dim sKey As String

    On Error Resume Next

    Set oClient = m_Clients(ClientID)

    If Not oClient Is Nothing Then
        ' Remove reverse mapping
        sKey = "S" & ObjPtr(oClient)
        m_ClientIDMap.Remove sKey

        ' Close and remove client
        oClient.Close_
        m_Clients.Remove ClientID

        RaiseEvent OnClientDisconnect(ClientID)
    End If
End Sub

Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cModbusTransportTCP] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' Socket Event Handlers - Client Mode
'=========================================================================

Private Sub m_Socket_Connect(Client As cWinsock)
    m_State = MB_TCP_STATE_CONNECTED
    RaiseEvent OnConnect
End Sub

Private Sub m_Socket_CloseEvent(Client As cWinsock)
    m_State = MB_TCP_STATE_CLOSED
    RaiseEvent OnDisconnect
End Sub

Private Sub m_Socket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte

    On Error Resume Next

    Client.GetData baData, vbByte + vbArray
    RaiseEvent OnDataReceived(baData)
End Sub

Private Sub m_Socket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    m_State = MB_TCP_STATE_ERROR
    RaiseError "Socket error: " & Description
End Sub

'=========================================================================
' Socket Event Handlers - Server Mode
'=========================================================================

Private Sub m_ListenSocket_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    Dim sClientID As String

    On Error GoTo ErrorHandler

    ' Generate client ID
    m_ClientCounter = m_ClientCounter + 1
    sClientID = "Client#" & m_ClientCounter

    ' Store client reference with ID
    AddClient Client, sClientID

    ' Accept connection
    DisConnect = False

    RaiseEvent OnClientConnect(sClientID, Client.RemoteHostIP)
    Exit Sub

ErrorHandler:
    DisConnect = True
    RaiseError "ConnectionRequest: " & ERR.Description
End Sub

Private Sub m_ListenSocket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Dim sClientID As String

    On Error GoTo ErrorHandler

    ' Find client ID
    sClientID = GetClientIDBySocket(Client)
    If LenB(sClientID) = 0 Then Exit Sub

    ' Receive data
    Client.GetData baData, vbByte + vbArray

    RaiseEvent OnClientDataReceived(sClientID, baData)
    Exit Sub

ErrorHandler:
    RaiseError "DataArrival: " & ERR.Description
End Sub

Private Sub m_ListenSocket_CloseEvent(Client As cWinsock)
    Dim sClientID As String

    On Error Resume Next

    sClientID = GetClientIDBySocket(Client)
    If LenB(sClientID) > 0 Then
        RemoveClient sClientID, "Connection closed"
    End If
End Sub

Private Sub m_ListenSocket_Error(Client As cWinsock, ByVal Number As Long, Description As String, ByVal Scode As Long)
    Dim sClientID As String

    On Error Resume Next

    sClientID = GetClientIDBySocket(Client)
    If LenB(sClientID) > 0 Then
        RemoveClient sClientID, "Socket error: " & Description
    Else
        RaiseError "Socket error: " & Description
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_Mode = MB_TCP_MODE_CLIENT
    m_State = MB_TCP_STATE_CLOSED
    m_Port = MODBUS_TCP_PORT
    m_BindAddress = "0.0.0.0"
    m_Timeout = MODBUS_DEFAULT_TIMEOUT
    m_TransactionID = 0
    m_ClientCounter = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next

    If m_Mode = MB_TCP_MODE_CLIENT Then
        Disconnect
    Else
        StopListening
    End If

    Set m_Socket = Nothing
    Set m_ListenSocket = Nothing
    Set m_Clients = Nothing
    Set m_ClientIDMap = Nothing
End Sub
