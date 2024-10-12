VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim IsRunning  As Boolean

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection

Public Event OnAccept(ClientInfo As cHttpServerClientInfo, Disconnect As Boolean)

Public Router As New cHttpServerRouter
Public RouteBefore As New cHttpServerRouteBefore
Public RouteAfter As New cHttpServerRouterAfter
Public Database As New cDataBase

Dim HttpApp As cHttpServerApp
Dim HttpSvr As cHttpServerSvr

'支持MVC部署方式, 即 脚本文件 在 App 目录,静态文件在 Webroot

Public Property Get Version() As String
    Version = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function Start(Optional Port As Long = 80, Optional WebRoot As String = "", Optional IP As String = "0.0.0.0") As Boolean
    If IsRunning = True Then Exit Function
    Set m_oServer = New cTlsReMaster
    With m_oServer
        .Protocol = RemasterProtocolConstants.sckTCPProtocol
        .Bind Port, IP
        .Listen
    End With
    '    Shell "cmd /c start http://localhost:8088/"
    Set HttpApp = New cHttpServerApp
    Set HttpSvr = New cHttpServerSvr
    With HttpSvr
        .IP = IP
        .Port = Port
        .WebRoot = WebRoot
    End With
    IsRunning = True
    Start = True
End Function


Private Sub Class_Initialize()
    '
End Sub

Private Sub m_oServer_ConnectionRequest(ByVal requestID As Long, Disconnect As Boolean)
    
    ' requestID 就是客户端的句柄 hSocket
    '    Debug.Print "m_oServer_ConnectionRequest, requestID=" & requestID & ", RemoteHostIP=" & m_oServer.RemoteHostIP & ", RemotePort=" & m_oServer.RemotePort, Timer
    'todo 实现 Disconnect,
    Dim CI As New cHttpServerClientInfo, IsDiss As Boolean
    With CI
        .IP = m_oServer.RemoteHostIP
        .Port = m_oServer.RemotePort
        .ConnectAt = Now()
    End With
    RaiseEvent OnAccept(CI, IsDiss)
    If IsDiss = True Then
        Disconnect = True
        Exit Sub
    End If
    Dim oCallback           As New cClientCallback
    m_cConnPool.Add oCallback
    With oCallback
        .Index = m_cConnPool.Count
        Set .Parent = Me
        Set .ClientInfo = CI
        Set .Socket = New cTlsReMaster
        Set .Context = New cHttpServerContext
        Set .Context.ClientInfo = CI
        Set .Context.Server = HttpSvr
        .Socket.Accept requestID
    End With
    '    Debug.Print "oCallback.Socket.Protocol=" & oCallback.Socket.Protocol
End Sub

Friend Sub OnDataArrival(Index As Long)
    
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    
    Call Router.Entry(oCallback)
    
End Sub

Friend Sub OnCloseSck(Index As Long)
    m_cConnPool.Remove Index
    If m_cConnPool.Count >= Index Then
        m_cConnPool.Add Nothing, Before:=Index
    End If
    Do While m_cConnPool.Count > 0
        If Not m_cConnPool.Item(m_cConnPool.Count) Is Nothing Then
            Exit Do
        End If
        m_cConnPool.Remove m_cConnPool.Count
    Loop
End Sub



