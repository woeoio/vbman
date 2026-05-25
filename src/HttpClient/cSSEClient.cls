VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSSEClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'作者：邓伟，QQ：215879458，可自由传播，但请勿删本签名

'SSE (Server-Sent Events) 客户端实现
'支持实时接收服务器推送事件
'2025-01-19 初始版本
'2025-05-25 新增 ConnectPost 方法（支持 AI 流式聊天等 POST 场景）
'           新增 SetHeader/ClearHeaders/RequestTimeOut

'引用 cHttpClient.cls 类文件
'引用 cJson.cls 类文件

'SSE 事件声明
Public Event OnOpen()
Public Event OnMessage(EventName As String, Data As String, Id As String)
Public Event OnError(Description As String, ErrorNumber As Long)
Public Event OnClose()

Dim WithEvents HttpClient As cHttpClient
Attribute HttpClient.VB_VarHelpID = -1
Private WithEvents ReconnectTimer As cTimer

'连接状态
Private IsConnected As Boolean
Private IsConnecting As Boolean
Private ShouldReconnect As Boolean
Private ReconnectDelay As Long
Private LastEventId As String

'SSE 解析缓冲区
Private SSEBuffer As String
Private LastEventType As String
Private LastEventData As String
Private LastEventRetry As Long

'自定义请求头和超时
Private m_Headers As Scripting.Dictionary
Private m_RequestTimeOut As Long

'配置参数
Public AutoReconnect As Boolean
Public ReconnectInterval As Long
Public MaxReconnectAttempts As Long
Public CurrentReconnectAttempts As Long

Public Url As String

Private Sub Class_Initialize()
    IsConnected = False
    IsConnecting = False
    ShouldReconnect = True
    AutoReconnect = True
    ReconnectInterval = 3000  '默认3秒重连间隔
    MaxReconnectAttempts = 10  '最多重连10次
    CurrentReconnectAttempts = 0
    LastEventRetry = 3000       '服务器指定的重连间隔
    ReconnectDelay = 0
    m_RequestTimeOut = 60       '默认60秒超时
    
    Set m_Headers = New Scripting.Dictionary
    m_Headers.CompareMode = TextCompare
    
    '初始化重连定时器
    Set ReconnectTimer = New cTimer
    ReconnectTimer.Interval = ReconnectInterval
    ReconnectTimer.Enabled = False
End Sub

Private Sub Class_Terminate()
    ShouldReconnect = False
    Call Disconnect()
    
    '释放自定义头
    If Not m_Headers Is Nothing Then
        Set m_Headers = Nothing
    End If
    
    '释放定时器
    If Not ReconnectTimer Is Nothing Then
        ReconnectTimer.Enabled = False
        Set ReconnectTimer = Nothing
    End If
End Sub

'=========================================================================
' 公共方法
'=========================================================================

'连接到 SSE 服务端（GET 方式）
Public Function Connect(ByVal ServerUrl As String) As cSSEClient
    Set Connect = Me
    
    If ServerUrl = "" Then
        RaiseEvent OnError("URL 不能为空", 500)
        Exit Function
    End If
    
    Url = ServerUrl
    IsConnecting = True
    
    Set HttpClient = New cHttpClient
    HttpClient.Async(True)
    HttpClient.RequestTimeOut = m_RequestTimeOut
    
    '设置 SSE 请求头
    HttpClient.RequestHeaders("Accept") = Consts.TOOLS_HTTP_CONTENT_TYPE_EVENT_STREAM
    HttpClient.RequestHeaders("Cache-Control") = "no-cache"
    
    '如果有上次的事件ID，设置 Last-Event-ID
    If LastEventId <> "" Then
        HttpClient.RequestHeaders("Last-Event-ID") = LastEventId
    End If
    
    '合并自定义请求头
    Call MergeHeaders
    
    '清空缓冲区和事件状态
    SSEBuffer = ""
    LastEventData = ""
    LastEventType = ""
    
    On Error Resume Next
    Call HttpClient.Fetch(ReqGet, Url, "")
    
    If Err.Number <> 0 Then
        IsConnecting = False
        RaiseEvent OnError(Err.Description, Err.Number)
        '触发重连逻辑
        If AutoReconnect And ShouldReconnect Then
            Call ScheduleReconnect()
        End If
    End If
End Function

'通过 POST 方式连接到 SSE 服务端（用于 AI 流式聊天等场景）
Public Function ConnectPost(ByVal ServerUrl As String, ByVal Body As String) As cSSEClient
    Set ConnectPost = Me
    
    If ServerUrl = "" Then
        RaiseEvent OnError("URL 不能为空", 500)
        Exit Function
    End If
    
    Url = ServerUrl
    IsConnecting = True
    AutoReconnect = False
    ShouldReconnect = False   'POST 方式不需要自动重连
    
    Set HttpClient = New cHttpClient
    HttpClient.Async(True)
    HttpClient.RequestTimeOut = m_RequestTimeOut
    
    '设置 SSE 请求头
    HttpClient.RequestHeaders("Accept") = Consts.TOOLS_HTTP_CONTENT_TYPE_EVENT_STREAM
    HttpClient.RequestHeaders("Cache-Control") = "no-cache"
    
    '合并自定义请求头（包含 Authorization、Content-Type 等）
    Call MergeHeaders
    
    '清空缓冲区和事件状态
    SSEBuffer = ""
    LastEventData = ""
    LastEventType = ""
    
    On Error Resume Next
    Call HttpClient.Fetch(ReqPost, Url, Body)
    
    If Err.Number <> 0 Then
        IsConnecting = False
        RaiseEvent OnError(Err.Description, Err.Number)
    End If
End Function

'断开连接
Public Sub Disconnect()
    ShouldReconnect = False
    IsConnected = False
    IsConnecting = False
    
    '停止重连定时器
    If Not ReconnectTimer Is Nothing Then
        ReconnectTimer.Enabled = False
    End If
    
    If Not HttpClient Is Nothing Then
        Set HttpClient = Nothing
    End If
    RaiseEvent OnClose
End Sub

'添加自定义请求头（链式调用）
Public Function SetHeader(ByVal Name As String, ByVal Value As String) As cSSEClient
    m_Headers(Name) = Value
    Set SetHeader = Me
End Function

'清空自定义请求头
Public Sub ClearHeaders()
    If Not m_Headers Is Nothing Then
        m_Headers.RemoveAll
    End If
End Sub

'=========================================================================
' 属性
'=========================================================================

'获取连接状态
Public Property Get Connected() As Boolean
    Connected = IsConnected
End Property

'获取是否正在连接
Public Property Get Connecting() As Boolean
    Connecting = IsConnecting
End Property

'设置请求超时（秒）
Public Property Let RequestTimeOut(ByVal Value As Long)
    m_RequestTimeOut = Value
End Property

Public Property Get RequestTimeOut() As Long
    RequestTimeOut = m_RequestTimeOut
End Property

'获取最后收到的事件ID
Public Property Get LastReceivedEventId() As String
    LastReceivedEventId = LastEventId
End Property

'获取当前重连延迟（毫秒）
Public Property Get CurrentReconnectDelay() As Long
    CurrentReconnectDelay = ReconnectDelay
End Property

'获取服务器指定的重连间隔（毫秒）
Public Property Get ServerReconnectInterval() As Long
    ServerReconnectInterval = LastEventRetry
End Property

'=========================================================================
' 重连相关方法
'=========================================================================

'关闭自动重连
Public Sub DisableAutoReconnect()
    AutoReconnect = False
    ShouldReconnect = False
End Sub

'启用自动重连
Public Sub EnableAutoReconnect()
    AutoReconnect = True
    ShouldReconnect = True
End Sub

'重置重连计数器
Public Sub ResetReconnectAttempts()
    CurrentReconnectAttempts = 0
End Sub

'定时器回调，用于延迟重连
Public Sub ReconnectCallback()
    If ShouldReconnect = False Then Exit Sub
    
    '关闭定时器
    ReconnectTimer.Enabled = False
    
    If CurrentReconnectAttempts >= MaxReconnectAttempts Then
        RaiseEvent OnError("达到最大重连次数 (" & MaxReconnectAttempts & ")", 0)
        Exit Sub
    End If
    
    CurrentReconnectAttempts = CurrentReconnectAttempts + 1
    Call Connect(Url)
End Sub

'=========================================================================
' 私有方法
'=========================================================================

'计划重连
Private Sub ScheduleReconnect()
    If ShouldReconnect = False Then Exit Sub
    If CurrentReconnectAttempts >= MaxReconnectAttempts Then Exit Sub
    
    Dim Delay As Long
    Delay = IIf(LastEventRetry > 0, LastEventRetry, ReconnectInterval)
    ReconnectDelay = Delay
    
    '设置定时器间隔并启动
    ReconnectTimer.Interval = Delay
    ReconnectTimer.Enabled = True
End Sub

'合并自定义请求头到 HttpClient
Private Sub MergeHeaders()
    If Not m_Headers Is Nothing Then
        If m_Headers.Count > 0 Then
            Dim k As Variant
            For Each k In m_Headers.Keys()
                HttpClient.RequestHeaders(k) = m_Headers(k)
            Next
        End If
    End If
End Sub

'解析 SSE 数据流
Private Sub ParseSSEData(Data As String)
    If Data = "" Then Exit Sub
    
    '将新数据追加到缓冲区
    SSEBuffer = SSEBuffer & Data
    
    '按行分割
    Dim Lines() As String
    Lines = Split(SSEBuffer, vbLf)
    
    Dim Line As Variant
    Dim ProcessedLines As Long
    ProcessedLines = 0
    
    '逐行处理
    For Each Line In Lines
        ProcessedLines = ProcessedLines + 1
        Dim TrimLine As String
        TrimLine = Trim(Line)
        
        '跳过空行和注释行
        If TrimLine = "" Then
            '空行表示一个事件结束
            If LastEventData <> "" Then
                '触发消息事件
                RaiseEvent OnMessage(LastEventType, LastEventData, LastEventId)
                '清空当前事件数据
                LastEventData = ""
                LastEventType = ""
            End If
        ElseIf Left(TrimLine, 1) = ":" Then
            '注释行，忽略
        ElseIf Left(TrimLine, 6) = "event:" Then
            '事件类型
            LastEventType = Trim(Mid(TrimLine, 7))
        ElseIf Left(TrimLine, 5) = "data:" Then
            '事件数据
            If LastEventData <> "" Then
                LastEventData = LastEventData & vbLf
            End If
            LastEventData = LastEventData & Trim(Mid(TrimLine, 6))
        ElseIf Left(TrimLine, 3) = "id:" Then
            '事件ID
            LastEventId = Trim(Mid(TrimLine, 4))
        ElseIf Left(TrimLine, 6) = "retry:" Then
            '重连间隔（毫秒）
            Dim RetryStr As String
            RetryStr = Trim(Mid(TrimLine, 7))
            If IsNumeric(RetryStr) Then
                LastEventRetry = CLng(RetryStr)
            End If
        End If
    Next Line
    
    '保留最后一行（可能是不完整的）
    If UBound(Lines) >= 0 Then
        Dim LastLine As String
        LastLine = Lines(UBound(Lines))
        
        '检查最后一行是否以换行符结尾
        If Right(Data, Len(vbCrLf)) = vbCrLf Or Right(Data, 1) = vbLf Then
            '最后一行是完整的，清空缓冲区
            SSEBuffer = ""
        ElseIf Right(Data, 1) <> vbLf Then
            '最后一行可能不完整，保留在缓冲区
            SSEBuffer = LastLine
        Else
            SSEBuffer = ""
        End If
    End If
End Sub

'=========================================================================
' 事件处理
'=========================================================================

'重连定时器事件处理
Private Sub ReconnectTimer_Timer()
    Call ReconnectCallback
End Sub

'HTTP 客户端响应完成事件
Private Sub HttpClient_OnResponseFinished()
    On Error Resume Next
    
    IsConnecting = False
    IsConnected = True
    CurrentReconnectAttempts = 0
    
    RaiseEvent OnOpen
    
    '获取响应数据
    Dim ResponseText As String
    ResponseText = HttpClient.ReturnText(True, False)
    
    '解析 SSE 数据
    If ResponseText <> "" Then
        Call ParseSSEData(ResponseText)
    End If
    
    'SSE 连接正常关闭
    IsConnected = False
    
    '如果需要自动重连
    If AutoReconnect And ShouldReconnect Then
        Call ScheduleReconnect()
    Else
        RaiseEvent OnClose
    End If
End Sub

'HTTP 客户端错误事件
Private Sub HttpClient_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    IsConnecting = False
    IsConnected = False
    
    RaiseEvent OnError(ErrorDescription, ErrorNumber)
    
    '如果需要自动重连
    If AutoReconnect And ShouldReconnect Then
        Call ScheduleReconnect()
    Else
        RaiseEvent OnClose
    End If
End Sub
