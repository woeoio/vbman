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
'作者：邓伟，215879458@qq.com

Option Explicit

Dim IsRunning  As Boolean

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection

Public Event OnAccept(ClientInfo As cHttpServerClientInfo, Disconnect As Boolean)
Public Event OnLogs(ByVal Level As String, ByVal Content As String)

Public SSE As New cSSE
Public ScriptEngine As New cScriptEngine

Public Router As New cHttpServerRouter
Public RouteBefore As New cHttpServerRouteBefore
Public RouteAfter As New cHttpServerRouterAfter
Public Database As New cDataBase

Public CrossDomain As New cHttpCrossDomain

Public LastError As String

Dim HttpApp As cHttpServerApp
Dim HttpSvr As cHttpServerSvr

'支持MVC部署方式, 即 脚本文件 在 App 目录,静态文件在 Webroot


Public Function Start(Optional Port As Long = 80, Optional WebRoot As String = "", Optional IP As String = "0.0.0.0") As Boolean
    If IsRunning = True Then LastError = "服务器正在运行中": Exit Function
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

Public Function StopMe() As Boolean
    On Error GoTo EH
    m_oServer.CloseSck
    IsRunning = False
    'clear
    Set m_cConnPool = Nothing
    Set HttpApp = Nothing
    Set HttpSvr = Nothing
    Set m_oServer = Nothing
    
    Set SSE = Nothing
    
    Set Router = Nothing
    Set RouteBefore = Nothing
    Set RouteAfter = Nothing
    Set Database = Nothing
    
    Set CrossDomain = Nothing
    
    StopMe = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
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
        Set .Socket = New cTlsReMaster
        Set .Context = New cHttpServerContext
        .Index = m_cConnPool.Count
        Set .Parent = Me
        'todo 这里可以优化为对象池子, 并发请求的时候处理更快
        Set .Context.Server = HttpSvr
        .Socket.Accept requestID
        '
        CI.hSocket = .Socket.SocketHandle
        Set .ClientInfo = CI
        Set .Context.ClientInfo = CI
        .hSocket = .Socket.SocketHandle
    End With
    '    Debug.Print "oCallback.Socket.Protocol=" & oCallback.Socket.Protocol
End Sub

Friend Sub OnDataArrival(Index As Long)
    
    On Error GoTo EH
    '写入对象到上下文
    Dim Request As New cHttpServerRequest
    Dim response As New cHttpServerResponse
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    '传入Client给输出对象
    Set response.fClient = oCallback
    Dim vRequest As Variant
    Dim sRequestHead            As String
    Dim sRequestBody            As String
    Dim aReq() As Byte
    Dim aReqHead() As Byte
    Dim aReqBody() As Byte
    oCallback.Socket.GetData aReq, vbArray + vbByte
    '分割请求头和请求体
    vRequest = ToolsArray.SplitByteArrayByKeyword(aReq, vbCrLf & vbCrLf, 1)
    aReqHead = vRequest(0)
    '请求头转为字符串
    sRequestHead = ToolsUtf8.Decode(aReqHead)
    
    '解析请求头
    If Request.ParseHeader(sRequestHead) = False Then ERR.Raise 500, "REQUEST", Request.LastError
    '开始 sse 流程
    If Request.HeaderAccept = "text/event-stream" Then
        oCallback.IsSSE = True
        Set oCallback.Context.Request = Request
        Call SSE.Entry(oCallback)
        Exit Sub
    End If
    '是否 options 查询, 且没有中间件处理流程, 则这里简单处理下, 避免重入业务流程
    If Request.Method = ReqOptions Then
        If CrossDomain.Enable = True Then
            response.Text ""
        Else
            response.State403
        End If
        Exit Sub
    End If
    '开始 http 流程
    Dim Session As New cHttpServerSession
    Dim Cookies As New cHttpServerCookies
    With oCallback.Context
        Set .Request = Request
        Set .response = response
        Set .Session = Session
        Set .Cookies = Cookies
        Set .Db = oCallback.Parent.Database
        Set .SSE.flsSSEInst = SSE
    End With
    '解析请求体, 这里可以按需要决定是否解析,比如头部指定为文件传输, 可能就要异步或者稍后解析
    aReqBody = vRequest(1)
    If Request.ParseBody(aReqBody) = False Then ERR.Raise 500, "REQUEST", Request.LastError
    
    '开始路由
    Call Router.Entry(oCallback)
    Exit Sub
EH:
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    RaiseEvent OnLogs(ToolsLogs.LEVEL_ERROR, ErrInfo)
    'todo 如果开发者定义了全局错误处理，就执行，否则执行默认错误处理函数
    response.State500 ErrInfo
End Sub

Friend Sub OnCloseSck(Index As Long)
    'todo 这里可能有内存泄漏
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    If oCallback.IsSSE = True Then
        SSE.CloseClient oCallback.hSocket
    End If
    Set oCallback = Nothing
    '    Set m_cConnPool.Item(Index) = Nothing
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



