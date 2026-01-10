VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebSocketServerClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' cWebSocketServerClient - WebSocket Server Client Connection for VB6
'
' Purpose: Represents a single client connection on the server side
'          Manages per-client state, buffers, and fragmented messages
'
' Author: Claude (Anthropic)
' Date: 2026-01-10
'
'=========================================================================
Option Explicit

'=========================================================================
' Client States
'=========================================================================

Public Enum WsClientState
    WS_CLIENT_PENDING = 0       ' Waiting for handshake
    WS_CLIENT_CONNECTED = 1     ' Handshake complete, ready for messages
    WS_CLIENT_CLOSING = 2       ' Close initiated
End Enum

'=========================================================================
' Member Variables
'=========================================================================

Private m_Socket As cWinsock
Private m_ID As String
Private m_State As WsClientState
Private m_RemoteAddress As String
Private m_RemotePort As Long

' Buffers
Private m_RecvBuffer As cByteBuffer
Private m_FragmentBuffer As cByteBuffer
Private m_FragmentOpCode As WsOpCode
Private m_IsFragmented As Boolean

' Frame parser (each client has its own)
Private m_FrameParser As cWebSocketFrame

'=========================================================================
' Properties
'=========================================================================

' Socket object
Public Property Get Socket() As cWinsock
    Set Socket = m_Socket
End Property

Public Property Set Socket(ByVal Value As cWinsock)
    Set m_Socket = Value
End Property

' Client ID
Public Property Get ID() As String
    ID = m_ID
End Property

Public Property Let ID(ByVal Value As String)
    m_ID = Value
End Property

' Client state
Public Property Get State() As WsClientState
    State = m_State
End Property

Public Property Let State(ByVal Value As WsClientState)
    m_State = Value
End Property

' Remote address
Public Property Get RemoteAddress() As String
    RemoteAddress = m_RemoteAddress
End Property

Public Property Let RemoteAddress(ByVal Value As String)
    m_RemoteAddress = Value
End Property

' Remote port
Public Property Get RemotePort() As Long
    RemotePort = m_RemotePort
End Property

Public Property Let RemotePort(ByVal Value As Long)
    m_RemotePort = Value
End Property

' Receive buffer
Public Property Get RecvBuffer() As cByteBuffer
    Set RecvBuffer = m_RecvBuffer
End Property

' Fragment buffer
Public Property Get FragmentBuffer() As cByteBuffer
    Set FragmentBuffer = m_FragmentBuffer
End Property

' Fragment opcode
Public Property Get FragmentOpCode() As WsOpCode
    FragmentOpCode = m_FragmentOpCode
End Property

Public Property Let FragmentOpCode(ByVal Value As WsOpCode)
    m_FragmentOpCode = Value
End Property

' Is fragmented message in progress
Public Property Get IsFragmented() As Boolean
    IsFragmented = m_IsFragmented
End Property

Public Property Let IsFragmented(ByVal Value As Boolean)
    m_IsFragmented = Value
End Property

' Frame parser
Public Property Get FrameParser() As cWebSocketFrame
    Set FrameParser = m_FrameParser
End Property

'=========================================================================
' Public Methods
'=========================================================================

' Append data to receive buffer
Public Sub AppendRecvData(ByRef Data() As Byte)
    On Error Resume Next
    If UBound(Data) >= 0 Then
        m_RecvBuffer.Append Data
    End If
End Sub

' Clear fragment buffer
Public Sub ClearFragmentBuffer()
    m_FragmentBuffer.Clear
    m_IsFragmented = False
End Sub

' Get assembled fragmented message
Public Function GetFragmentedData() As Byte()
    GetFragmentedData = m_FragmentBuffer.ToArray
End Function

' Send frame to this client (server doesn't mask)
Public Sub SendFrame(ByRef FrameData() As Byte)
    On Error Resume Next
    If Not m_Socket Is Nothing Then
        m_Socket.SendData FrameData
    End If
End Sub

' Send text message
Public Sub SendText(ByVal Message As String)
    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildTextFrame(Message, False)  ' Server doesn't mask
    SendFrame baFrame
End Sub

' Send binary message
Public Sub SendBinary(ByRef Data() As Byte)
    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildBinaryFrame(Data, False)
    SendFrame baFrame
End Sub

' Send close frame
Public Sub SendClose(Optional ByVal Code As WsCloseCode = WS_CLOSE_NORMAL, _
                    Optional ByVal Reason As String = "")
    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildCloseFrame(Code, Reason, False)
    SendFrame baFrame
    m_State = WS_CLIENT_CLOSING
End Sub

' Send pong frame
Public Sub SendPong(ByRef Payload() As Byte)
    Dim baFrame() As Byte
    baFrame = m_FrameParser.BuildPongFrame(Payload, False)
    SendFrame baFrame
End Sub

' Close socket
Public Sub CloseSocket()
    On Error Resume Next
    If Not m_Socket Is Nothing Then
        m_Socket.Close_
    End If
End Sub

'=========================================================================
' Class Events
'=========================================================================

Private Sub Class_Initialize()
    m_State = WS_CLIENT_PENDING
    m_IsFragmented = False

    Set m_RecvBuffer = New cByteBuffer
    Set m_FragmentBuffer = New cByteBuffer
    Set m_FrameParser = New cWebSocketFrame
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_Socket = Nothing
    Set m_RecvBuffer = Nothing
    Set m_FragmentBuffer = Nothing
    Set m_FrameParser = Nothing
End Sub
