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
'作者：张伟，215879458@qq.com

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Session 存储类型枚举
Public Enum EnumSessionStorageType
    SessionStorageMemory = 0      ' 内存存储（非持久化）
    SessionStorageFileSystem = 1   ' 文件系统存储
    SessionStorageDatabase = 2     ' 数据库存储
End Enum

Dim IsRunning  As Boolean

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1

' 内置定时清理器：定期清理空闲连接和过期 Session
Private WithEvents m_oCleanupTimer As cTimer
Attribute m_oCleanupTimer.VB_VarHelpID = -1

' BUG-1 修复：用 Dictionary 替代 Collection，以 hSocket 为 Key 彻底消除索引漂移
' As New→显式：改 As New 为显式创建，确保 Set = Nothing 时真正释放
Private m_cConnPool             As Scripting.Dictionary

' TLS 模式枚举
Private Enum EnumTlsMode
    TlsNone          ' 未启用 TLS（默认）
    TlsModeCertFile  ' 证书文件
    TlsModeSubject   ' Windows 证书存储
    TlsModeMemory    ' 内存证书集合
End Enum

' TLS 状态
Private m_TlsMode As EnumTlsMode

' TlsCertFile 参数
Private m_TlsCertFile As String
Private m_TlsPassword As String

' TlsCertSubject 参数
Private m_TlsCertSubject As String

' TlsCertMemory 参数
Private m_TlsCertificates As Collection
Private m_TlsPrivateKey As Collection

' 其他 TLS 参数
Private m_TlsAlpnProtocols As String

' WebRoot
Private m_WebRoot As String
' As New→显式：Session 字典显式创建
Dim m_Sessions                  As Scripting.Dictionary

Public Event OnAccept(ClientInfo As cHttpServerClientInfo, Disconnect As Boolean)
Public Event OnLogs(ByVal Level As String, ByVal Content As String)

' As New→显式：所有共享对象改为显式创建
Public SSE As cSSE
Public ScriptEngine As cScriptEngine

Public Router As cHttpServerRouter
Public RouteBefore As cHttpServerRouteBefore
Public RouteAfter As cHttpServerRouterAfter
Public Database As cDataBase

Public CrossDomain As cHttpCrossDomain

' 性能统计对象
Public Statistics As cHttpServerStatistics

Public LastError As String

' Session 持久化参数
Public SessionStorageType As EnumSessionStorageType ' Session 存储类型
Public SessionStoragePath As String ' 文件系统路径或数据库名
Public SessionCookieName As String ' Session Cookie 名称
Public SessionAutoCleanup As Boolean ' 是否自动清理过期 Session

' BUG-7 修复：最大连接数限制（默认 1000）
Public MaxConnections As Long
' BUG-3 修复：最大请求体大小（默认 10MB）
Public MaxRequestSize As Long
' BUG-2 修复：空闲连接超时秒数（默认 120 秒）
Public IdleTimeoutSeconds As Long
' 内置清理定时器间隔（毫秒，默认 30000 = 30秒）
Private m_CleanupTimerInterval As Long

Dim HttpApp As cHttpServerApp
Dim HttpSvr As cHttpServerSvr

'支持MVC方式，如 脚本文件 在 App 目录，静态文件在 WebRoot

'=========================================================================
' TLS 链式配置
'=========================================================================

' 证书文件模式（最常用）
' CertFile: PFX/P12 或 PEM 格式证书路径，PEM 支持 | 号分隔多文件
Public Function TlsCertFile( _
    ByVal CertFile As String, _
    Optional ByVal Password As String, _
    Optional ByVal AlpnProtocols As String = "http/1.1") As cHttpServer

    m_TlsMode = TlsModeCertFile
    m_TlsCertFile = CertFile
    m_TlsPassword = Password
    m_TlsAlpnProtocols = AlpnProtocols
    Set TlsCertFile = Me
End Function

' Windows 证书存储模式
' CertSubject: 证书主题（如 "www.example.com"）
Public Function TlsCertSubject( _
    ByVal CertSubject As String, _
    Optional ByVal AlpnProtocols As String = "http/1.1") As cHttpServer

    m_TlsMode = TlsModeSubject
    m_TlsCertSubject = CertSubject
    m_TlsAlpnProtocols = AlpnProtocols
    Set TlsCertSubject = Me
End Function

' 内存证书集合模式（高级接口）
Public Function TlsCertMemory( _
    ByVal Certificates As Collection, _
    ByVal PrivateKey As Collection, _
    Optional ByVal AlpnProtocols As String = "http/1.1") As cHttpServer

    m_TlsMode = TlsModeMemory
    Set m_TlsCertificates = Certificates
    Set m_TlsPrivateKey = PrivateKey
    m_TlsAlpnProtocols = AlpnProtocols
    Set TlsCertMemory = Me
End Function

'=========================================================================
' WebRoot 链式配置
'=========================================================================

Public Function WebRoot(ByVal Path As String) As cHttpServer
    m_WebRoot = Path
    Set WebRoot = Me
End Function

'=========================================================================
' Start / StopMe
'=========================================================================

Public Function Start(Optional Port As Long = 80, Optional IP As String = "0.0.0.0") As Boolean
    If IsRunning = True Then LastError = "服务器已经在运行中": Exit Function
    Set m_oServer = New cTlsReMaster
    With m_oServer
        If m_TlsMode = TlsNone Then
            ' --- 纯 HTTP ---
            .Protocol = RemasterProtocolConstants.sckTCPProtocol
            .Bind Port, IP
            .Listen
        Else
            ' --- HTTPS ---
            .Protocol = RemasterProtocolConstants.sckTLSProtocol
            .Bind Port, IP
            Select Case m_TlsMode
                Case TlsModeCertFile
                    .Listen m_TlsCertFile, m_TlsPassword, , , , m_TlsAlpnProtocols
                Case TlsModeSubject
                    .Listen , , m_TlsCertSubject, , , m_TlsAlpnProtocols
                Case TlsModeMemory
                    .Listen , , , m_TlsCertificates, m_TlsPrivateKey, m_TlsAlpnProtocols
            End Select
        End If
    End With
    Set HttpApp = New cHttpServerApp
    Set HttpSvr = New cHttpServerSvr
    With HttpSvr
        .IP = IP
        .Port = Port
        .WebRoot = m_WebRoot
    End With
    IsRunning = True
    Statistics.StartTime = Now
    
    ' 启动内置定时清理器
    Set m_oCleanupTimer = New cTimer
    m_oCleanupTimer.Interval = m_CleanupTimerInterval
    m_oCleanupTimer.Enable
    
    Start = True
End Function

Public Function StopMe() As Boolean
    On Error GoTo EH
    If Not m_oServer Is Nothing Then
        m_oServer.CloseSck
    End If
    IsRunning = False
    
    ' 停止内置定时清理器
    If Not m_oCleanupTimer Is Nothing Then
        m_oCleanupTimer.Disable
        Set m_oCleanupTimer = Nothing
    End If
    
    ' 逐一释放所有连接，断开循环引用链
    Dim key As Variant
    Dim oCallback As cClientCallback
    If Not m_cConnPool Is Nothing Then
        For Each key In m_cConnPool.Keys
            Set oCallback = m_cConnPool.Item(key)
            If Not oCallback Is Nothing Then
                oCallback.Release
            End If
        Next
        m_cConnPool.RemoveAll
    End If
    Set HttpApp = Nothing
    Set HttpSvr = Nothing
    Set m_oServer = Nothing
    
    ' BUG-12 修复：StopMe 不再销毁共享对象（Router/SSE/Database/CrossDomain 等）
    ' 这些对象由用户管理，服务器重启后可继续使用
    
    ' 清理内存中 Session
    If Not m_Sessions Is Nothing Then
        ' 先释放每个 Session 内部资源
        Dim sk As Variant
        For Each sk In m_Sessions.Keys
            Dim sess As cHttpServerSession
            Set sess = m_Sessions(sk)
            If Not sess Is Nothing Then sess.Release
        Next
        m_Sessions.RemoveAll
    End If
    
    ' 清理 TLS 状态
    m_TlsMode = TlsNone
    m_TlsCertFile = ""
    m_TlsPassword = ""
    m_TlsCertSubject = ""
    Set m_TlsCertificates = Nothing
    Set m_TlsPrivateKey = Nothing
    m_TlsAlpnProtocols = ""
    m_WebRoot = ""
    
    StopMe = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
End Function

Private Sub Class_Initialize()
    ' 显式创建所有子对象（替代 As New 隐式创建）
    Set m_cConnPool = New Scripting.Dictionary
    Set m_Sessions = New Scripting.Dictionary
    m_Sessions.CompareMode = TextCompare
    
    Set SSE = New cSSE
    Set ScriptEngine = New cScriptEngine
    Set Router = New cHttpServerRouter
    Set RouteBefore = New cHttpServerRouteBefore
    Set RouteAfter = New cHttpServerRouterAfter
    Set Database = New cDataBase
    Set CrossDomain = New cHttpCrossDomain
    Set Statistics = New cHttpServerStatistics
    
    ' 默认配置
    SessionStorageType = SessionStorageMemory ' 默认内存存储
    SessionCookieName = "VBMAN_SESSIONID"
    SessionAutoCleanup = True
    ' BUG-7 修复：默认最大连接数
    MaxConnections = 1000
    ' BUG-3 修复：默认最大请求体 10MB
    MaxRequestSize = 10485760
    ' BUG-2 修复：默认空闲超时 120 秒
    IdleTimeoutSeconds = 120
    ' 内置清理定时器默认 30 秒
    m_CleanupTimerInterval = 30000
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    StopMe
    ' 显式释放所有共享对象
    Set SSE = Nothing
    Set ScriptEngine = Nothing
    Set Router = Nothing
    Set RouteBefore = Nothing
    Set RouteAfter = Nothing
    Set Database = Nothing
    Set CrossDomain = Nothing
    Set Statistics = Nothing
    If Not m_cConnPool Is Nothing Then
        m_cConnPool.RemoveAll
        Set m_cConnPool = Nothing
    End If
    If Not m_Sessions Is Nothing Then
        m_Sessions.RemoveAll
        Set m_Sessions = Nothing
    End If
End Sub

' 内置清理定时器事件：定期清理空闲连接和过期 Session
Private Sub m_oCleanupTimer_Timer()
    On Error Resume Next
    ' 定时清理空闲连接
    If IdleTimeoutSeconds > 0 Then CleanupIdleConnections
    ' 定时清理过期 Session
    If SessionAutoCleanup Then CleanupExpiredSessions
End Sub

' 获取/设置内置清理定时器间隔（毫秒）
Public Property Get CleanupTimerInterval() As Long
    CleanupTimerInterval = m_CleanupTimerInterval
End Property

Public Property Let CleanupTimerInterval(ByVal value As Long)
    m_CleanupTimerInterval = value
    ' 如果定时器已在运行，更新间隔并重启
    If Not m_oCleanupTimer Is Nothing Then
        m_oCleanupTimer.Interval = value
        m_oCleanupTimer.Reset
    End If
End Property

Private Sub m_oServer_ConnectionRequest(ByVal requestID As Long, Disconnect As Boolean)
    
    ' requestID 就是客户端的句柄 hSocket
    Dim CI As cHttpServerClientInfo, IsDiss As Boolean
    Set CI = New cHttpServerClientInfo
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
    
    ' BUG-7 修复：检查最大连接数
    If m_cConnPool.Count >= MaxConnections Then
        Disconnect = True
        RaiseEvent OnLogs("WARN", "连接数已达上限 " & MaxConnections & "，拒绝新连接")
        Statistics.RejectedConnections = Statistics.RejectedConnections + 1
        Exit Sub
    End If
    
    Dim oCallback           As cClientCallback
    Set oCallback = New cClientCallback
    Set oCallback.Socket = New cTlsReMaster
    Set oCallback.Context = New cHttpServerContext
    
    ' BUG-1 修复：先 Accept 获取 SocketHandle，再用 hSocket 作为 Dictionary Key
    oCallback.Socket.Accept requestID
    Dim sockHandle As Long
    sockHandle = oCallback.Socket.SocketHandle
    
    With oCallback
        .hSocket = sockHandle
        Set .Parent = Me
        .LastActivity = Now  ' BUG-2 修复：记录连接创建时间
        Set .Context.Server = HttpSvr
        CI.hSocket = sockHandle
        Set .ClientInfo = CI
        Set .Context.ClientInfo = CI
    End With
    
    ' BUG-1 修复：用 hSocket 做 Key，不依赖 Collection 位置索引
    m_cConnPool.Add sockHandle, oCallback
    
    ' 统计：连接接受 + 峰值更新
    Statistics.TotalConnectionsAccepted = Statistics.TotalConnectionsAccepted + 1
    Statistics.UpdatePeakConnections m_cConnPool.Count
End Sub

' BUG-1 修复：参数从 Index 改为 hSocket，用 Dictionary 按 Key 查找
Friend Sub OnDataArrival(ByVal hSocket As Long)
    
    On Error GoTo EH
    Dim oCallback As cClientCallback
    
    ' BUG-1 修复：用 hSocket 直接查找，不存在则跳过
    If Not m_cConnPool.Exists(hSocket) Then Exit Sub
    Set oCallback = m_cConnPool.Item(hSocket)
    
    ' BUG-2 修复：更新活动时间
    oCallback.LastActivity = Now
    
    ' 1. 读取数据并追加到缓冲区（解决粘包问题）
    Dim newData() As Byte
    oCallback.Socket.GetData newData, vbArray + vbByte
    
    ' 统计：记录接收字节数
    Statistics.TotalBytesReceived = Statistics.TotalBytesReceived + UBound(newData) + 1
    
    ' BUG-3 修复：检查缓冲区大小是否超过限制
    If MaxRequestSize > 0 Then
        If oCallback.GetRecvBufferLen + UBound(newData) + 1 > MaxRequestSize Then
            RaiseEvent OnLogs("ERROR", "请求体超过最大限制 " & MaxRequestSize & " bytes, hSocket=" & hSocket)
            oCallback.Release
            m_cConnPool.Remove hSocket
            Exit Sub
        End If
    End If
    
    ' 追加数据到缓冲区
    oCallback.AppendToBuffer newData
    
    ' 2. 循环处理缓冲区中的所有完整请求（解决粘包问题）
    Do While ContainsCompleteRequest(oCallback)
        ' 3. 提取一个完整的 HTTP 请求
        Dim completeRequest As Variant
        completeRequest = ExtractOneRequest(oCallback)
        
        ' 4. 处理该请求
        ProcessHttpRequest hSocket, completeRequest
        
        ' 5. 从缓冲区移除已处理的数据
        RemoveProcessedData oCallback, completeRequest
    Loop
    
    Exit Sub
EH:
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    RaiseEvent OnLogs(ToolsLogs.LEVEL_ERROR, "OnDataArrival: " & ErrInfo)
    ' 发生错误时清空缓冲区，防止持续积压导致崩溃
    oCallback.GetRecvBufferLen = 0
End Sub

' 判断是否包含完整的 HTTP 请求（解决粘包问题）
Private Function ContainsCompleteRequest(oCallback As cClientCallback) As Boolean
    Dim bufferLen As Long
    bufferLen = oCallback.GetRecvBufferLen
    
    ' 如果缓冲区为空，直接返回 False
    If bufferLen = 0 Then
        ContainsCompleteRequest = False
        Exit Function
    End If
    
    ' 获取缓冲区数据
    Dim buffer() As Byte
    buffer = oCallback.GetRecvBuffer
    
    ' 查找请求头结束标记 CRLFCRLF
    Dim headerEnd As Long
    headerEnd = FindCRLFCRLF(buffer, bufferLen)
    
    ' BUG-6 修复：FindCRLFCRLF 返回 -1 表示未找到
    If headerEnd = -1 Then
        ' 未找到完整的请求头
        ContainsCompleteRequest = False
        Exit Function
    End If
    
    ' 将请求头部分转换为字符串
    Dim headerBytes() As Byte
    ReDim headerBytes(0 To headerEnd - 1)
    CopyMemory headerBytes(0), buffer(0), headerEnd
    
    Dim headerStr As String
    headerStr = ToolsUtf8.Decode(headerBytes)
    
    ' 解析 Content-Length
    Dim contentLength As Long
    contentLength = ParseContentLengthFromHeader(headerStr)
    
    ' BUG-3 修复：检查 Content-Length 是否超过最大限制
    If MaxRequestSize > 0 And contentLength > MaxRequestSize Then
        ' 超限请求，返回 True 让上层处理并断开
        ContainsCompleteRequest = True
        Exit Function
    End If
    
    ' 计算已接收到体内容量
    Dim bodyStart As Long
    bodyStart = headerEnd + 4  ' \r\n\r\n 的长度（4个字节）
    
    ' 如果是 GET、HEAD、OPTIONS 等没有 body 的请求方法，且 contentLength 为 0
    If contentLength = 0 Then
        ' 判断是否是那些没有 body 的请求方法
        If IsMethodWithoutBody(headerStr) Then
            ContainsCompleteRequest = True
            Exit Function
        End If
    End If
    
    ' 判断缓冲区是否包含完整的请求头 + 体部分
    Dim totalRequired As Long
    totalRequired = bodyStart + contentLength
    
    ContainsCompleteRequest = (bufferLen >= totalRequired)
End Function

' BUG-6 修复：找不到 CRLFCRLF 时返回 -1（而非 0），消除歧义
Private Function FindCRLFCRLF(buffer() As Byte, bufLen As Long) As Long
    Dim i As Long
    For i = 0 To bufLen - 4
        If buffer(i) = 13 And buffer(i + 1) = 10 And _
           buffer(i + 2) = 13 And buffer(i + 3) = 10 Then
            FindCRLFCRLF = i
            Exit Function
        End If
    Next
    FindCRLFCRLF = -1  ' BUG-6 修复：未找到返回 -1 而非 0
End Function

' 提取一个完整的请求
Private Function ExtractOneRequest(oCallback As cClientCallback) As Variant
    Dim buffer() As Byte
    buffer = oCallback.GetRecvBuffer
    
    Dim bufferLen As Long
    bufferLen = oCallback.GetRecvBufferLen
    
    ' 返回 Variant 数组：Array(headerBytes, bodyBytes, totalLength)
    Dim result(2) As Variant
    
    ' 查找请求头结束位置
    Dim headerEnd As Long
    headerEnd = FindCRLFCRLF(buffer, bufferLen)
    
    ' BUG-6 修复：headerEnd 为 -1 时不应进入此函数（调用方已保证）
    If headerEnd = -1 Then
        result(2) = 0
        ExtractOneRequest = result
        Exit Function
    End If
    
    ' 获取请求头
    Dim headerBytes() As Byte
    ReDim headerBytes(0 To headerEnd - 1)
    CopyMemory headerBytes(0), buffer(0), headerEnd
    result(0) = headerBytes
    
    ' 将请求头转换为字符串以解析 Content-Length
    Dim headerStr As String
    headerStr = ToolsUtf8.Decode(headerBytes)
    Dim contentLength As Long
    contentLength = ParseContentLengthFromHeader(headerStr)
    
    ' 获取请求体
    Dim bodyStart As Long
    bodyStart = headerEnd + 4
    
    If contentLength > 0 Then
        Dim bodyBytes() As Byte
        ReDim bodyBytes(0 To contentLength - 1)
        CopyMemory bodyBytes(0), buffer(bodyStart), contentLength
        result(1) = bodyBytes
    Else
        ' 没有 body
        result(1) = vbNullString
    End If
    
    ' 计算总长度
    result(2) = bodyStart + contentLength
    
    ExtractOneRequest = result
End Function

' 从缓冲区移除已处理的数据
Private Sub RemoveProcessedData(oCallback As cClientCallback, extractedRequest As Variant)
    Dim processedLen As Long
    processedLen = extractedRequest(2)
    
    oCallback.RemoveFromBuffer processedLen
End Sub

' 从请求头字符串中解析 Content-Length
Private Function ParseContentLengthFromHeader(headerStr As String) As Long
    On Error GoTo EH
    Dim lines As Variant
    lines = Split(headerStr, vbCrLf)
    
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = lines(i)
        
        ' 匹配 Content-Length 头（不区分大小写）
        If InStr(1, line, "Content-Length:", vbTextCompare) > 0 Then
            Dim parts As Variant
            parts = Split(line, ":", 2)
            If UBound(parts) >= 1 Then
                ParseContentLengthFromHeader = CLng(Trim(parts(1)))
                Exit Function
            End If
        End If
    Next
    
    ' 如果没有找到 Content-Length，返回 0
    ParseContentLengthFromHeader = 0
    Exit Function
EH:
    ParseContentLengthFromHeader = 0
End Function

' 判断请求方法是否没有 Body
Private Function IsMethodWithoutBody(headerStr As String) As Boolean
    On Error GoTo EH
    ' 获取第一行
    Dim firstLine As String
    Dim lines As Variant
    lines = Split(headerStr, vbCrLf)
    If UBound(lines) >= 0 Then
        firstLine = lines(0)
        
        ' 判断请求方法
        If InStr(1, firstLine, "GET ", vbTextCompare) = 1 Or _
           InStr(1, firstLine, "HEAD ", vbTextCompare) = 1 Or _
           InStr(1, firstLine, "OPTIONS ", vbTextCompare) = 1 Or _
           InStr(1, firstLine, "DELETE ", vbTextCompare) = 1 Then
            IsMethodWithoutBody = True
            Exit Function
        End If
    End If
    IsMethodWithoutBody = False
    Exit Function
EH:
    IsMethodWithoutBody = False
End Function

' 处理一个完整的 HTTP 请求
' BUG-1 修复：参数从 Index As Long 改为 hSocket As Long
Private Sub ProcessHttpRequest(ByVal hSocket As Long, completeRequest As Variant)
    On Error GoTo EH
    
    If Not m_cConnPool.Exists(hSocket) Then Exit Sub
    Dim oCallback As cClientCallback
    Set oCallback = m_cConnPool.Item(hSocket)
    
    ' As New→显式：本地变量用显式创建，不用 As New
    Dim Request As cHttpServerRequest
    Dim response As cHttpServerResponse
    Set Request = New cHttpServerRequest
    Set response = New cHttpServerResponse
    
    ' 设置Client回调引用
    Set response.fClient = oCallback
    
    Dim aReqHead() As Byte
    Dim aReqBody() As Byte
    Dim sRequestHead As String
    Dim sessionFile As String
    Dim sessionTable As String
    
    aReqHead = completeRequest(0)
    aReqBody = completeRequest(1)
    
    ' 请求头转为字符串
    sRequestHead = ToolsUtf8.Decode(aReqHead)
    
    ' 解析请求头
    If Request.ParseHeader(sRequestHead) = False Then
        ERR.Raise 500, "REQUEST", Request.LastError
    End If
    
    ' BUG-3 修复：请求体大小超限检查
    If MaxRequestSize > 0 And Request.HeaderContentLength > MaxRequestSize Then
        response.State413 "Request entity too large"
        Statistics.RequestSizeRejected = Statistics.RequestSizeRejected + 1
        ' 请求处理完毕，显式清理本次请求对象
        On Error Resume Next
        oCallback.Context.ReleaseRequest
        Request.ClearRequest
        Set response = Nothing: Set Request = Nothing
        On Error GoTo EH
        Exit Sub
    End If
    
    ' 初始 sse 连接
    If Request.HeaderAccept = "text/event-stream" Then
        oCallback.IsSSE = True
        Set oCallback.Context.Request = Request
        Call SSE.Entry(oCallback)
        Statistics.SSEConnectionsAccepted = Statistics.SSEConnectionsAccepted + 1
        Exit Sub
    End If
    
    ' 是否 options 查询
    If Request.Method = ReqOptions Then
        If CrossDomain.Enable = True Then
            response.Text ""
        Else
            response.State403
        End If
        ' 请求处理完毕，显式清理本次请求对象
        On Error Resume Next
        oCallback.Context.ReleaseRequest
        Request.ClearRequest
        Set response = Nothing: Set Request = Nothing
        On Error GoTo EH
        Exit Sub
    End If
    
    ' 初始 http 会话
    Dim Session As cHttpServerSession
    Dim Cookies As cHttpServerCookies
    Set Cookies = New cHttpServerCookies
    
    ' 解析请求头中的 Cookie，分离请求 Cookie（只解析到 m_RequestCookies，响应 Cookie 通过 Cookie() 设置）
    If Request.Header.Exists("Cookie") = True Then
        Cookies.Decode Request.Header("Cookie")
    End If
    
    ' 从请求 Cookie 获取 SessionID，查找 Session
    Dim sessionID As String
    Dim sessionCookieName As String
    sessionCookieName = IIf(Me.SessionCookieName = "", "VBMAN_SESSIONID", Me.SessionCookieName)
    
    ' 使用 GetValue 获取请求端 Cookie（不影响响应端 Cookie）
    sessionID = Cookies.GetValue(sessionCookieName)
    
    If sessionID <> "" Then
        ' 根据不同存储类型加载 Session
        If SessionStorageType = SessionStorageMemory Then
            ' 内存模式：直接查找存储中的 Session 对象
            If m_Sessions.Exists(sessionID) Then
                Set Session = m_Sessions(sessionID)
                If Session.IsExpired Then
                    m_Sessions.Remove sessionID
                    Set Session = Nothing
                Else
                    Session.Touch
                End If
            End If
        ElseIf SessionStorageType = SessionStorageFileSystem Then
            ' 文件系统
            Set Session = New cHttpServerSession
            Session.sessionID = sessionID
            If SessionStoragePath = "" Then
                sessionFile = Environ("TEMP") & "\vbman\sessions\" & sessionID & ".json"
            Else
                sessionFile = SessionStoragePath & "\" & sessionID & ".json"
            End If
            If Session.LoadFromFile(sessionFile) = False Or Session.IsExpired Then
                ' 加载失败或已过期，删除临时文件
                If Session.IsExpired Then
                    On Error Resume Next
                    Kill sessionFile
                    On Error GoTo EH
                End If
                Set Session = Nothing
            End If
        ElseIf SessionStorageType = SessionStorageDatabase Then
            ' 数据库
            Set Session = New cHttpServerSession
            sessionTable = IIf(SessionStoragePath = "", "sessions", SessionStoragePath)
            If Session.LoadFromDatabase(Database, sessionID, sessionTable) = False Or Session.IsExpired Then
                ' 加载失败或已过期，删除过期记录
                If Session.IsExpired Then
                    On Error Resume Next
                    Database.Sql("DELETE FROM " & sessionTable & " WHERE SessionID=?") _
                        .Param("SessionID", sessionID, adVarWChar) _
                        .ExecParam
                    On Error GoTo EH
                End If
                Set Session = Nothing
            End If
        End If
    End If
    
    ' 如果 Session 未加载（或用户会话过期），创建新 Session
    If Session Is Nothing Then
        Set Session = New cHttpServerSession
    End If
    
    With oCallback.Context
        Set .Request = Request
        Set .response = response
        Set .Session = Session
        Set .Cookies = Cookies
        Set .Db = oCallback.Parent.Database
        Set .SSE.flsSSEInst = SSE
        Set .Statistics = Statistics    ' 引用服务器级统计对象
    End With
    
    ' 解析请求体
    If Request.ParseBody(aReqBody) = False Then
        ERR.Raise 500, "REQUEST", Request.LastError
    End If
    
    ' 开始路由
    Call Router.Entry(oCallback)
    
    ' 统计：请求完成（按方法和状态码分类）
    Statistics.TotalRequests = Statistics.TotalRequests + 1
    Statistics.IncrementByMethod Request.Method
    If Not response Is Nothing Then
        Statistics.IncrementByStatusCode response.StatusCode
    End If
    
    ' 关键修复：只有 Session.HasID（实际有数据被写入）时才写入 Cookie 和存储
    ' 之前 sessionID Property Get 会自动生成 GUID，导致每个请求的空 Session 都被存入 m_Sessions
    ' 造成 900 QPS × 20 分钟 = 1,080,000 个空 Session 永驻内存（约 1-2GB）
    If Session.HasID Then
        Statistics.TotalSessionsCreated = Statistics.TotalSessionsCreated + 1
        ' 将 SessionID 写入响应 Cookie
        With Cookies.Cookie(sessionCookieName)
            .Value = Session.sessionID
            .Path = "/"
            .HttpOnly = True
            .SameSite = "Lax"
            If Session.TimeOut > 0 Then
                .Expires = DateAdd("n", Session.TimeOut, Now)
            End If
        End With
        
        ' 根据不同存储方式保存 Session
        If SessionStorageType = SessionStorageMemory Then
            ' 内存模式：更新或添加到字典（下次请求时复用，不需要额外保存）
            Set m_Sessions(Session.sessionID) = Session
        ElseIf Session.Dirty Then
            ' 文件/数据库模式：仅在数据变更时持久化，避免不必要的 I/O
            If SessionStorageType = SessionStorageFileSystem Then
                ' 文件系统
                If SessionStoragePath = "" Then
                    sessionFile = Environ("TEMP") & "\vbman\sessions\" & Session.sessionID & ".json"
                Else
                    sessionFile = SessionStoragePath & "\" & Session.sessionID & ".json"
                End If
                Session.SaveToFile sessionFile
            ElseIf SessionStorageType = SessionStorageDatabase Then
                ' 数据库
                sessionTable = IIf(SessionStoragePath = "", "sessions", SessionStoragePath)
                Session.SaveToDatabase Database, sessionTable
            End If
        End If
    End If
    
    ' 请求处理完毕，显式清理本次请求对象，切断循环引用链
    ' 这样即使在高并发下，VB6 COM 引用计数也能及时回收
    On Error Resume Next
    oCallback.Context.ReleaseRequest
    Request.ClearRequest
    ' 显式释放本地变量引用（包括 Session 和 Cookies，不再等 Sub 退出才释放）
    Set Session = Nothing
    Set Cookies = Nothing
    Set response = Nothing: Set Request = Nothing
    On Error GoTo 0
    
    Exit Sub
EH:
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    Statistics.RequestErrors = Statistics.RequestErrors + 1
    Statistics.TotalRequests = Statistics.TotalRequests + 1
    RaiseEvent OnLogs(ToolsLogs.LEVEL_ERROR, "ProcessHttpRequest: " & ErrInfo)
    On Error Resume Next
    If Not response Is Nothing Then
        response.State500 ErrInfo
    End If
    ' 错误路径也要清理请求对象
    oCallback.Context.ReleaseRequest
    Set Session = Nothing
    Set Cookies = Nothing
    Set response = Nothing: Set Request = Nothing
End Sub

' BUG-1 修复：参数从 Index 改为 hSocket，用 Dictionary 按 Key 移除
Friend Sub OnCloseSck(ByVal hSocket As Long)
    On Error GoTo EH
    If Not m_cConnPool.Exists(hSocket) Then Exit Sub
    
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(hSocket)
    If Not oCallback Is Nothing Then
        If oCallback.IsSSE = True Then
            SSE.CloseClient oCallback.hSocket
        End If
        ' 释放连接，断开循环引用链
        oCallback.Release
    End If
    
    m_cConnPool.Remove hSocket
    Exit Sub
EH:
    ' 出错时尝试清理连接池条目
    On Error Resume Next
    m_cConnPool.Remove hSocket
End Sub

' BUG-2 修复：清理空闲超时连接
Public Sub CleanupIdleConnections()
    On Error GoTo EH
    If IdleTimeoutSeconds <= 0 Then Exit Sub
    If m_cConnPool.Count = 0 Then Exit Sub
    
    Dim key As Variant
    Dim oCallback As cClientCallback
    Dim cutoffTime As Date
    cutoffTime = DateAdd("s", -IdleTimeoutSeconds, Now)
    Dim keysToClose As Collection
    Set keysToClose = New Collection
    
    ' 先收集需要关闭的连接 Key（不能在遍历中修改 Dictionary）
    For Each key In m_cConnPool.Keys
        Set oCallback = m_cConnPool.Item(key)
        If Not oCallback Is Nothing Then
            If oCallback.LastActivity < cutoffTime Then
                keysToClose.Add key
            End If
        End If
    Next
    
    ' 统计：空闲连接清理数
    Statistics.IdleConnectionsCleaned = Statistics.IdleConnectionsCleaned + keysToClose.Count
    
    ' 关闭空闲连接
    Dim i As Long
    For i = 1 To keysToClose.Count
        Dim hSocketToClose As Long
        hSocketToClose = keysToClose(i)
        If m_cConnPool.Exists(hSocketToClose) Then
            Set oCallback = m_cConnPool.Item(hSocketToClose)
            If Not oCallback Is Nothing Then
                RaiseEvent OnLogs("INFO", "关闭空闲连接 hSocket=" & hSocketToClose & _
                    ", 空闲时间=" & DateDiff("s", oCallback.LastActivity, Now) & "s")
                If oCallback.IsSSE = True Then
                    SSE.CloseClient hSocketToClose
                End If
                oCallback.Release
            End If
            m_cConnPool.Remove hSocketToClose
        End If
    Next
    Exit Sub
EH:
    ' 忽略清理过程中的错误
End Sub

' 清理过期的 Session
Public Sub CleanupExpiredSessions()
    On Error GoTo EH
    Dim k As Variant
    Dim Session As cHttpServerSession
    Dim keysToRemove As Collection
    Set keysToRemove = New Collection
    
    If SessionStorageType = SessionStorageMemory Then
        ' 内存模式：遍历字典，删除过期 Session
        For Each k In m_Sessions.Keys
            Set Session = m_Sessions(k)
            If Session.IsExpired Then
                keysToRemove.Add k
            End If
        Next
        Dim i As Long
        For i = 1 To keysToRemove.Count
            ' 先释放 Session 内部资源
            Set Session = m_Sessions(keysToRemove(i))
            If Not Session Is Nothing Then Session.Release
            m_Sessions.Remove keysToRemove(i)
        Next
    ElseIf SessionStorageType = SessionStorageFileSystem Then
        ' 文件系统模式，删除过期的 Session 文件
        Dim sessionPath As String
        If SessionStoragePath = "" Then
            sessionPath = Environ("TEMP") & "\vbman\sessions\"
        Else
            sessionPath = SessionStoragePath & "\"
        End If
        If Dir(sessionPath, vbDirectory) <> "" Then
            Dim f As String
            f = Dir(sessionPath & "*.json")
            Do While f <> ""
                Dim sess As cHttpServerSession
                Set sess = New cHttpServerSession
                If sess.LoadFromFile(sessionPath & f) Then
                    If sess.IsExpired Then
                        On Error Resume Next
                        Kill sessionPath & f
                        On Error GoTo EH
                    End If
                End If
                f = Dir()
            Loop
        End If
    ElseIf SessionStorageType = SessionStorageDatabase Then
        ' 数据库模式，删除过期的 Session 记录
        Dim tblName As String
        tblName = IIf(SessionStoragePath = "", "sessions", SessionStoragePath)
        On Error Resume Next
        Database.Sql("DELETE FROM " & tblName & " WHERE ExpiresAt<?") _
            .Param("ExpiresAt", Now, adDate) _
            .ExecParam
        On Error GoTo EH
    End If
    Exit Sub
EH:
    ' 忽略清理过程中的错误
End Sub

' 获取当前连接数
Public Property Get ConnectionCount() As Long
    If m_cConnPool Is Nothing Then
        ConnectionCount = 0
    Else
        ConnectionCount = m_cConnPool.Count
    End If
End Property
