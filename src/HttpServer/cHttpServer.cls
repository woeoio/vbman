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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Session 存储类型枚举
Public Enum EnumSessionStorageType
    SessionStorageMemory = 0      ' 内存存储（不持久化）
    SessionStorageFileSystem = 1   ' 文件系统存储
    SessionStorageDatabase = 2     ' 数据库存储
End Enum

Dim IsRunning  As Boolean

Private WithEvents m_oServer    As cTlsReMaster
Attribute m_oServer.VB_VarHelpID = -1
Private m_cConnPool             As New Collection
Dim m_Sessions                  As New Scripting.Dictionary

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

' Session 持久化配置
Public SessionStorageType As EnumSessionStorageType ' Session 存储类型
Public SessionStoragePath As String ' 文件系统路径或数据库表名
Public SessionCookieName As String ' Session Cookie 名称
Public SessionAutoCleanup As Boolean ' 是否自动清理过期 Session

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
    ' 逐一释放所有连接，断开循环引用
    Dim i As Long
    For i = m_cConnPool.Count To 1 Step -1
        Dim oCallback As cClientCallback
        Set oCallback = m_cConnPool.Item(i)
        If Not oCallback Is Nothing Then
            oCallback.Release
        End If
    Next
    Set m_cConnPool = Nothing
    Set HttpApp = Nothing
    Set HttpSvr = Nothing
    Set m_oServer = Nothing
    
    Set SSE = Nothing
    Set ScriptEngine = Nothing
    
    Set Router = Nothing
    Set RouteBefore = Nothing
    Set RouteAfter = Nothing
    Set Database = Nothing
    
    Set CrossDomain = Nothing
    
    ' 清理内存 Session
    If Not m_Sessions Is Nothing Then
        m_Sessions.RemoveAll
        Set m_Sessions = Nothing
    End If
    
    StopMe = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
End Function

Private Sub Class_Initialize()
    ' 默认配置
    SessionStorageType = SessionStorageMemory ' 默认内存存储
    SessionCookieName = "VBMAN_SESSIONID"
    SessionAutoCleanup = True
    m_Sessions.CompareMode = TextCompare
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
    Dim oCallback As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    
    ' 1. 获取新数据并追加到接收缓冲区（解决粘包问题）
    Dim newData() As Byte
    oCallback.Socket.GetData newData, vbArray + vbByte
    
    ' 追加数据到缓冲区
    oCallback.AppendToBuffer newData
    
    ' 2. 循环处理缓冲区中的完整请求（解决粘包问题）
    Do While ContainsCompleteRequest(oCallback)
        ' 3. 提取一个完整的 HTTP 请求
        Dim completeRequest As Variant
        completeRequest = ExtractOneRequest(oCallback)
        
        ' 4. 处理该请求
        ProcessHttpRequest Index, completeRequest
        
        ' 5. 从缓冲区移除已处理的数据
        RemoveProcessedData oCallback, completeRequest
    Loop
    Exit Sub
EH:
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    RaiseEvent OnLogs(ToolsLogs.LEVEL_ERROR, "OnDataArrival: " & ErrInfo)
    ' 发生错误时清空缓冲区，避免一直累积错误数据
    oCallback.GetRecvBufferLen = 0
End Sub

' 判断是否包含完整的 HTTP 请求（用于解决粘包问题）
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
    
    If headerEnd = 0 Then
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
    
    ' 计算已接收的体部长度
    Dim bodyStart As Long
    bodyStart = headerEnd + 4  ' \r\n\r\n 的长度（4个字节）
    
    ' 如果是 GET、HEAD、OPTIONS 等没有 body 的请求，contentLength 为 0
    If contentLength = 0 Then
        ' 检查是否是这些没有 body 的请求方法
        If IsMethodWithoutBody(headerStr) Then
            ContainsCompleteRequest = True
            Exit Function
        End If
    End If
    
    ' 判断缓冲区是否包含完整的请求（头部 + 体部）
    Dim totalRequired As Long
    totalRequired = bodyStart + contentLength
    
    ContainsCompleteRequest = (bufferLen >= totalRequired)
End Function

' 查找 CRLFCRLF 标记
Private Function FindCRLFCRLF(buffer() As Byte, bufLen As Long) As Long
    Dim i As Long
    For i = 0 To bufLen - 4
        If buffer(i) = 13 And buffer(i + 1) = 10 And _
           buffer(i + 2) = 13 And buffer(i + 3) = 10 Then
            FindCRLFCRLF = i
            Exit Function
        End If
    Next
    FindCRLFCRLF = 0
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
    
    ' 提取请求头
    Dim headerBytes() As Byte
    ReDim headerBytes(0 To headerEnd - 1)
    CopyMemory headerBytes(0), buffer(0), headerEnd
    result(0) = headerBytes
    
    ' 将请求头转换为字符串以解析 Content-Length
    Dim headerStr As String
    headerStr = ToolsUtf8.Decode(headerBytes)
    Dim contentLength As Long
    contentLength = ParseContentLengthFromHeader(headerStr)
    
    ' 提取请求体
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
    
    ' 返回总长度
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
        
        ' 查找 Content-Length 头（不区分大小写）
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
        
        ' 检查请求方法
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
Private Sub ProcessHttpRequest(Index As Long, completeRequest As Variant)
    On Error GoTo EH
    
    Dim oCallback As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    
    Dim Request As New cHttpServerRequest
    Dim response As New cHttpServerResponse
    
    ' 传入Client给输出对象
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
    
    ' 开始 sse 流程
    If Request.HeaderAccept = "text/event-stream" Then
        oCallback.IsSSE = True
        Set oCallback.Context.Request = Request
        Call SSE.Entry(oCallback)
        Exit Sub
    End If
    
    ' 是否 options 查询
    If Request.Method = ReqOptions Then
        If CrossDomain.Enable = True Then
            response.Text ""
        Else
            response.State403
        End If
        Exit Sub
    End If
    
    ' 开始 http 流程
    Dim Session As cHttpServerSession
    Dim Cookies As New cHttpServerCookies
    
    ' 从请求头解析 Cookie（请求 Cookie 存入 m_RequestCookies，响应 Cookie 通过 Cookie() 设置）
    If Request.Header.Exists("Cookie") = True Then
        Cookies.Decode Request.Header("Cookie")
    End If
    
    ' 从请求 Cookie 读取 SessionID 并加载 Session
    Dim sessionID As String
    Dim sessionCookieName As String
    sessionCookieName = IIf(Me.SessionCookieName = "", "VBMAN_SESSIONID", Me.SessionCookieName)
    
    ' 使用 GetValue 读取请求 Cookie（不创建响应 Cookie）
    sessionID = Cookies.GetValue(sessionCookieName)
    
    If sessionID <> "" Then
        ' 根据配置加载 Session
        If SessionStorageType = SessionStorageMemory Then
            ' 内存模式：直接引用存储的 Session 对象
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
                ' 加载失败或已过期，清理过期文件
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
                ' 加载失败或已过期，清理过期记录
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
    
    ' 如果 Session 未加载（新用户或会话过期），创建空 Session
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
    End With
    
    ' 解析请求体
    If Request.ParseBody(aReqBody) = False Then
        ERR.Raise 500, "REQUEST", Request.LastError
    End If
    
    ' 开始路由
    Call Router.Entry(oCallback)
    
    ' 保存 Session（仅在有 SessionID 时，Item 写入会自动生成 SessionID）
    If Session.sessionID <> "" Then
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
        
        ' 根据配置保存 Session 到存储
        If SessionStorageType = SessionStorageMemory Then
            ' 内存模式：更新或添加到字典（对象引用，无需额外保存）
            Set m_Sessions(Session.sessionID) = Session
        ElseIf Session.Dirty Then
            ' 文件/数据库模式：仅数据变更时持久化，避免不必要的 I/O
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
    
    ' 概率性自动清理过期 Session（约 1% 概率）
    If SessionAutoCleanup Then
        If Rnd < 0.01 Then CleanupExpiredSessions
    End If
    Exit Sub
EH:
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    RaiseEvent OnLogs(ToolsLogs.LEVEL_ERROR, "ProcessHttpRequest: " & ErrInfo)
    response.State500 ErrInfo
End Sub

Friend Sub OnCloseSck(Index As Long)
    On Error GoTo EH
    Dim oCallback           As cClientCallback
    Set oCallback = m_cConnPool.Item(Index)
    If Not oCallback Is Nothing Then
        If oCallback.IsSSE = True Then
            SSE.CloseClient oCallback.hSocket
        End If
        ' 主动断开所有循环引用
        oCallback.Release
    End If
    
    m_cConnPool.Remove Index
    ' 清理尾部空位
    Do While m_cConnPool.Count > 0
        If Not m_cConnPool.Item(m_cConnPool.Count) Is Nothing Then
            Exit Do
        End If
        m_cConnPool.Remove m_cConnPool.Count
    Loop
    Exit Sub
EH:
    ' 出错时仍尝试清理集合
    On Error Resume Next
    m_cConnPool.Remove Index
    Do While m_cConnPool.Count > 0
        If Not m_cConnPool.Item(m_cConnPool.Count) Is Nothing Then
            Exit Do
        End If
        m_cConnPool.Remove m_cConnPool.Count
    Loop
End Sub

' 清理过期的 Session
Public Sub CleanupExpiredSessions()
    On Error GoTo EH
    Dim k As Variant
    Dim Session As cHttpServerSession
    Dim keysToRemove As New Collection
    
    If SessionStorageType = SessionStorageMemory Then
        ' 内存模式：遍历字典清理过期 Session
        For Each k In m_Sessions.Keys
            Set Session = m_Sessions(k)
            If Session.IsExpired Then
                keysToRemove.Add k
            End If
        Next
        Dim i As Long
        For i = 1 To keysToRemove.Count
            m_Sessions.Remove keysToRemove(i)
        Next
    ElseIf SessionStorageType = SessionStorageFileSystem Then
        ' 文件系统模式：删除过期的 Session 文件
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
                Dim sess As New cHttpServerSession
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
        ' 数据库模式：删除过期的 Session 记录
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
    ' 忽略清理错误
End Sub



