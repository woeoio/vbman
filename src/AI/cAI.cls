VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'作者：邓伟，QQ：215879458，可自由传播，但请勿删本签名

'AI 对象类 - 通用 OpenAI 兼容 API 客户端
'支持：豆包、DeepSeek、OpenAI、通义千问等
'2025-01-19 初始版本
'2025-05-25 修复：流式响应改用 cSSEClient、修复 JSON 解析逻辑、
'           修复系统提示词重复添加、修复 WithEvents 声明

'引用 cHttpClient.cls 类文件
'引用 cJson.cls 类文件
'引用 cSSEClient.cls 类文件

'=========================================================================
' 事件定义
'=========================================================================
Public Event OnStreamChunk(Chunk As String, IsComplete As Boolean)
Public Event OnComplete(Response As String)
Public Event OnError(ErrorMsg As String, ErrorCode As Long)

'=========================================================================
' 私有变量
'=========================================================================
Private WithEvents SSEClient As cSSEClient
Attribute SSEClient.VB_VarHelpID = -1
Dim Json As New cJson

'基础配置
Private m_ApiKey As String
Private m_BaseUrl As String
Private m_ApiPath As String
Private m_Model As String
Private m_Timeout As Long

'请求参数
Private m_Temperature As Double
Private m_MaxTokens As Long
Private m_TopP As Double
Private m_FrequencyPenalty As Double
Private m_PresencePenalty As Double
Private m_Stream As Boolean
Private m_Stop As Variant

'消息管理
Private m_Messages As Collection
Private m_SystemPrompt As String

'高级配置
Private m_CustomHeaders As Scripting.Dictionary
Private m_Proxy As String
Private m_Organization As String

'错误处理
Private m_LastError As String

'流式响应相关
Private m_StreamBuffer As String
Private m_StreamCallback As Object
Private m_StreamCallbackMethod As String
Private m_IsStreaming As Boolean

'=========================================================================
' 属性
'=========================================================================
Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Get IsStreaming() As Boolean
    IsStreaming = m_IsStreaming
End Property

'=========================================================================
' 基础配置方法（链式调用）
'=========================================================================

'设置 API 密钥
Public Function ApiKey(Key As String) As cAI
    m_ApiKey = Key
    Set ApiKey = Me
End Function

'设置基础 URL（用于适配不同服务商）
Public Function BaseUrl(Url As String) As cAI
    m_BaseUrl = Url
    Set BaseUrl = Me
End Function

'设置 API 路径（用于适配不同 API 版本）
Public Function ApiPath(Path As String) As cAI
    m_ApiPath = Path
    Set ApiPath = Me
End Function

'设置模型名称
Public Function Model(ModelName As String) As cAI
    m_Model = ModelName
    Set Model = Me
End Function

'设置超时时间
Public Function Timeout(Seconds As Long) As cAI
    m_Timeout = Seconds
    Set Timeout = Me
End Function

'=========================================================================
' 参数配置方法（链式调用）
'=========================================================================

'设置温度参数
Public Function Temperature(Value As Double) As cAI
    m_Temperature = Value
    Set Temperature = Me
End Function

'设置最大 token 数
Public Function MaxTokens(Value As Long) As cAI
    m_MaxTokens = Value
    Set MaxTokens = Me
End Function

'设置 Top-P
Public Function TopP(Value As Double) As cAI
    m_TopP = Value
    Set TopP = Me
End Function

'设置频率惩罚
Public Function FrequencyPenalty(Value As Double) As cAI
    m_FrequencyPenalty = Value
    Set FrequencyPenalty = Me
End Function

'设置存在惩罚
Public Function PresencePenalty(Value As Double) As cAI
    m_PresencePenalty = Value
    Set PresencePenalty = Me
End Function

'启用/禁用流式响应
Public Function Stream(Enable As Boolean) As cAI
    m_Stream = Enable
    Set Stream = Me
End Function

'设置停止序列
Public Function Stop(Value As Variant) As cAI
    m_Stop = Value
    Set Stop = Me
End Function

'=========================================================================
' 消息管理方法（链式调用）
'=========================================================================

'添加系统提示词
Public Function System(Prompt As String) As cAI
    m_SystemPrompt = Prompt
    Set System = Me
End Function

'添加用户消息
Public Function User(Content As String) As cAI
    If m_Messages Is Nothing Then
        Set m_Messages = New Collection
    End If
    
    Dim Msg As New Scripting.Dictionary
    Msg("role") = "user"
    Msg("content") = Content
    m_Messages.Add Msg
    Set User = Me
End Function

'添加助手消息
Public Function Assistant(Content As String) As cAI
    If m_Messages Is Nothing Then
        Set m_Messages = New Collection
    End If
    
    Dim Msg As New Scripting.Dictionary
    Msg("role") = "assistant"
    Msg("content") = Content
    m_Messages.Add Msg
    Set Assistant = Me
End Function

'清空消息历史
Public Function ClearMessages() As cAI
    If Not m_Messages Is Nothing Then
        Set m_Messages = Nothing
    End If
    m_SystemPrompt = ""
    Set ClearMessages = Me
End Function

'设置消息列表（高级用法）
Public Function Messages(MsgList As Collection) As cAI
    Set m_Messages = MsgList
    Set Messages = Me
End Function

'=========================================================================
' 高级配置方法（链式调用）
'=========================================================================

'添加自定义请求头
Public Function Header(Name As String, Value As String) As cAI
    If m_CustomHeaders Is Nothing Then
        Set m_CustomHeaders = New Scripting.Dictionary
        m_CustomHeaders.CompareMode = TextCompare
    End If
    m_CustomHeaders(Name) = Value
    Set Header = Me
End Function

'设置代理
Public Function Proxy(ProxyUrl As String) As cAI
    m_Proxy = ProxyUrl
    Set Proxy = Me
End Function

'设置组织 ID
Public Function Organization(OrgId As String) As cAI
    m_Organization = OrgId
    Set Organization = Me
End Function

'=========================================================================
' 预设配置方法（便捷方法）
'=========================================================================

'快速配置为豆包 API
Public Function Doubao(ApiKey As String) As cAI
    m_ApiKey = ApiKey
    m_BaseUrl = "https://ark.cn-beijing.volces.com/api/v3"
    m_ApiPath = "/chat/completions"
    Set Doubao = Me
End Function

'快速配置为 DeepSeek API
Public Function DeepSeek(ApiKey As String) As cAI
    m_ApiKey = ApiKey
    m_BaseUrl = "https://api.deepseek.com/v1"
    m_ApiPath = "/chat/completions"
    Set DeepSeek = Me
End Function

'快速配置为 OpenAI API
Public Function OpenAI(ApiKey As String) As cAI
    m_ApiKey = ApiKey
    m_BaseUrl = "https://api.openai.com/v1"
    m_ApiPath = "/chat/completions"
    Set OpenAI = Me
End Function

'快速配置为通义千问 API
Public Function Qwen(ApiKey As String) As cAI
    m_ApiKey = ApiKey
    m_BaseUrl = "https://dashscope.aliyuncs.com/api/v1"
    m_ApiPath = "/chat/completions"
    Set Qwen = Me
End Function

'=========================================================================
' 执行方法（非链式）
'=========================================================================

'发送聊天请求
'Stream=False 时同步返回完整结果
'Stream=True 时通过 SSE 流式接收，阻塞等待直到完成
Public Function Chat(Optional UserMessage As String, Optional Callback As Object, Optional CallbackMethod As String) As String
    On Error GoTo ERR_HANDLER
    
    '添加用户消息（如果提供）
    If UserMessage <> "" Then
        Call Me.User(UserMessage)
    End If
    
    '验证配置
    If m_ApiKey = "" Then
        ERR.Raise 500, , "API Key 未设置，请使用 ApiKey() 或预设配置方法设置"
    End If
    
    '设置默认值
    If m_BaseUrl = "" Then m_BaseUrl = "https://api.openai.com/v1"
    If m_ApiPath = "" Then m_ApiPath = "/chat/completions"
    If m_Model = "" Then m_Model = "gpt-3.5-turbo"
    If m_Timeout = 0 Then m_Timeout = 60
    
    '检查是否启用流式响应
    If m_Stream = True Then
        '检查是否已有流式请求在执行
        If m_IsStreaming Then
            ERR.Raise 500, , "上一次流式请求尚未完成"
        End If
        
        m_StreamCallback = Callback
        m_StreamCallbackMethod = CallbackMethod
        m_StreamBuffer = ""
        
        Call SendStreamRequest
        
        '等待流式响应完成（阻塞等待）
        Do While m_IsStreaming
            DoEvents
        Loop
        
        Chat = m_StreamBuffer
    Else
        '同步请求
        Chat = SendSyncRequest
    End If
    
    Exit Function
    
ERR_HANDLER:
    m_LastError = ERR.Description
    RaiseEvent OnError(ERR.Description, ERR.Number)
    ERR.Raise ERR.Number, ERR.Source, ERR.Description
End Function

'发送聊天请求（返回 JSON 对象）
Public Function ChatJson(Optional UserMessage As String) As cJson
    Dim ResponseText As String
    ResponseText = Me.Chat(UserMessage)
    
    Set ChatJson = New cJson
    Call ChatJson.Decode(ResponseText)
End Function

'重置所有配置为默认值
Public Sub Reset()
    m_ApiKey = ""
    m_BaseUrl = ""
    m_ApiPath = ""
    m_Model = ""
    m_Timeout = 0
    
    m_Temperature = 1#
    m_MaxTokens = 0
    m_TopP = 1#
    m_FrequencyPenalty = 0#
    m_PresencePenalty = 0#
    m_Stream = False
    m_Stop = Empty
    
    Call ClearMessages
    
    If Not m_CustomHeaders Is Nothing Then
        Set m_CustomHeaders = Nothing
    End If
    m_Proxy = ""
    m_Organization = ""
    
    m_LastError = ""
End Sub

'=========================================================================
' 私有方法
'=========================================================================

'构建请求体（不修改 m_Messages，避免系统提示词重复添加）
Private Function BuildRequestBody() As String
    Set Json = New cJson
    
    '设置模型
    Json("model") = m_Model
    
    '构建消息列表（使用临时集合，不修改 m_Messages）
    Dim MsgList As New Collection
    
    '如果有系统提示词，添加到消息列表开头
    If m_SystemPrompt <> "" Then
        Dim SysMsg As New Scripting.Dictionary
        SysMsg("role") = "system"
        SysMsg("content") = m_SystemPrompt
        MsgList.Add SysMsg
    End If
    
    '添加用户/助手消息
    If Not m_Messages Is Nothing Then
        Dim i As Long
        For i = 1 To m_Messages.Count
            MsgList.Add m_Messages(i)
        Next
    End If
    
    '设置消息
    Json("messages") = MsgList
    
    '设置参数
    If m_Temperature <> 1# Then
        Json("temperature") = m_Temperature
    End If
    
    If m_MaxTokens > 0 Then
        Json("max_tokens") = m_MaxTokens
    End If
    
    If m_TopP <> 1# Then
        Json("top_p") = m_TopP
    End If
    
    If m_FrequencyPenalty <> 0 Then
        Json("frequency_penalty") = m_FrequencyPenalty
    End If
    
    If m_PresencePenalty <> 0 Then
        Json("presence_penalty") = m_PresencePenalty
    End If
    
    If m_Stream = True Then
        Json("stream") = True
    End If
    
    If Not IsEmpty(m_Stop) Then
        Json("stop") = m_Stop
    End If
    
    BuildRequestBody = Json.Encode()
End Function

'发送同步请求
Private Function SendSyncRequest() As String
    Dim HttpClient As New cHttpClient
    
    '设置超时
    HttpClient.RequestTimeOut = m_Timeout
    HttpClient.RequestChartSet = "utf-8"
    
    '设置请求头
    HttpClient.RequestHeaders("Authorization") = "Bearer " & m_ApiKey
    HttpClient.RequestHeaders("Content-Type") = "application/json"
    
    '设置组织 ID（如果有）
    If m_Organization <> "" Then
        HttpClient.RequestHeaders("OpenAI-Organization") = m_Organization
    End If
    
    '添加自定义请求头
    If Not m_CustomHeaders Is Nothing Then
        Dim Key As Variant
        For Each Key In m_CustomHeaders.Keys()
            HttpClient.RequestHeaders(Key) = m_CustomHeaders(Key)
        Next
    End If
    
    '构建请求 URL
    Dim Url As String
    Url = m_BaseUrl & m_ApiPath
    
    '发送请求
    Dim Body As String
    Body = BuildRequestBody()
    
    Call HttpClient.SendPost(Url, Body)
    
    '获取响应
    SendSyncRequest = HttpClient.ReturnText(True, False)
    
    '触发完成事件
    RaiseEvent OnComplete(SendSyncRequest)
End Function

'发送流式请求（使用 cSSEClient）
Private Sub SendStreamRequest()
    '先停止流式状态，避免清理旧 SSE 连接时 OnClose 重复触发
    m_IsStreaming = False
    
    '清理旧的 SSE 连接
    If Not SSEClient Is Nothing Then
        On Error Resume Next
        SSEClient.Disconnect
        On Error GoTo 0
        Set SSEClient = Nothing
    End If
    
    '创建新的 SSE 客户端
    Set SSEClient = New cSSEClient
    SSEClient.DisableAutoReconnect
    SSEClient.RequestTimeOut = m_Timeout
    
    '设置请求头
    SSEClient.SetHeader "Authorization", "Bearer " & m_ApiKey
    SSEClient.SetHeader "Content-Type", "application/json"
    
    '设置组织 ID（如果有）
    If m_Organization <> "" Then
        SSEClient.SetHeader "OpenAI-Organization", m_Organization
    End If
    
    '添加自定义请求头
    If Not m_CustomHeaders Is Nothing Then
        Dim Key As Variant
        For Each Key In m_CustomHeaders.Keys()
            SSEClient.SetHeader CStr(Key), CStr(m_CustomHeaders(Key))
        Next
    End If
    
    '构建请求 URL 和请求体
    Dim Url As String
    Url = m_BaseUrl & m_ApiPath
    Dim Body As String
    Body = BuildRequestBody()
    
    '设置流式状态（在新客户端创建后再设置）
    m_IsStreaming = True
    
    '发送流式请求
    SSEClient.ConnectPost Url, Body
End Sub

'解析流式响应的 JSON 数据
'OpenAI SSE 返回格式：{"id":"...","choices":[{"delta":{"content":"..."}}]}
Private Sub ParseStreamJson(JsonData As String)
    If JsonData = "" Then Exit Sub
    
    Dim ChunkJson As New cJson
    On Error Resume Next
    Call ChunkJson.Decode(JsonData)
    If Err.Number <> 0 Then Exit Sub
    
    'SSE 返回的 JSON 是对象格式（不是数组），RootItem 是 Dictionary
    If ChunkJson.RootItem.Count = 0 Then Exit Sub
    If Not ChunkJson.RootItem.Exists("choices") Then Exit Sub
    
    '提取 choices 数组（JSON 数组解析为 Collection）
    Dim vChoices As Variant
    Set vChoices = ChunkJson.Item("choices")
    If Err.Number <> 0 Then Exit Sub
    If vChoices Is Nothing Then Exit Sub
    
    '取第一个 choice（JSON 对象解析为 Dictionary）
    Dim vChoice As Variant
    Set vChoice = vChoices(1)
    If Err.Number <> 0 Then Exit Sub
    If vChoice Is Nothing Then Exit Sub
    
    '提取 delta
    Dim vDelta As Variant
    Set vDelta = vChoice("delta")
    If Err.Number <> 0 Then Exit Sub
    If vDelta Is Nothing Then Exit Sub
    
    '提取 content（可能不存在：首帧只有 role，末帧只有 finish_reason）
    If vDelta.Exists("content") Then
        Dim Content As String
        Content = CStr(vDelta("content"))
        If Err.Number <> 0 Then Exit Sub
        
        If Content <> "" Then
            '追加到缓冲区
            m_StreamBuffer = m_StreamBuffer & Content
            
            '触发流式响应事件
            Call TriggerStreamChunk(Content, False)
        End If
    End If
End Sub

'触发流式响应（回调或事件）
Private Sub TriggerStreamChunk(Chunk As String, IsComplete As Boolean)
    '优先使用回调函数
    If Not m_StreamCallback Is Nothing And m_StreamCallbackMethod <> "" Then
        On Error Resume Next
        Call CallByName(m_StreamCallback, m_StreamCallbackMethod, VbMethod, Chunk, IsComplete)
        On Error GoTo 0
    Else
        '使用事件
        RaiseEvent OnStreamChunk(Chunk, IsComplete)
    End If
    
    If IsComplete Then
        RaiseEvent OnComplete(m_StreamBuffer)
    End If
End Sub

'=========================================================================
' 初始化和清理
'=========================================================================

Private Sub Class_Initialize()
    '初始化默认值
    m_Temperature = 1#
    m_TopP = 1#
    m_FrequencyPenalty = 0#
    m_PresencePenalty = 0#
    m_Stream = False
    m_Timeout = 60
End Sub

Private Sub Class_Terminate()
    '清理 SSE 客户端
    If Not SSEClient Is Nothing Then
        On Error Resume Next
        SSEClient.Disconnect
        On Error GoTo 0
        Set SSEClient = Nothing
    End If
    
    Set Json = Nothing
    Set m_Messages = Nothing
    Set m_CustomHeaders = Nothing
    Set m_StreamCallback = Nothing
End Sub

'=========================================================================
' SSE 客户端事件处理（流式响应）
'=========================================================================

'SSE 连接打开事件
Private Sub SSEClient_OnOpen()
    'SSE 连接已建立，等待数据
End Sub

'SSE 消息事件 - 逐条处理 SSE 数据
Private Sub SSEClient_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    '检查是否为结束标记 [DONE]
    If Data = "[DONE]" Then
        m_IsStreaming = False
        Call TriggerStreamChunk("", True)
        Exit Sub
    End If
    
    '解析 JSON 数据，提取 delta.content
    Call ParseStreamJson(Data)
End Sub

'SSE 错误事件
Private Sub SSEClient_OnError(ByVal Description As String, ByVal ErrorNumber As Long)
    m_LastError = Description
    m_IsStreaming = False
    
    RaiseEvent OnError(Description, ErrorNumber)
End Sub

'SSE 连接关闭事件
Private Sub SSEClient_OnClose()
    '只有在仍在流式状态时才触发完成（防止 [DONE] 已触发过的情况）
    If m_IsStreaming Then
        m_IsStreaming = False
        Call TriggerStreamChunk("", True)
    End If
End Sub
