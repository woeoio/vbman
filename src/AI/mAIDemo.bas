Attribute VB_Name = "mAIDemo"
Option Explicit
'AI 对象使用示例
'作者：邓伟，QQ：215879458

'=========================================================================
' 基础使用示例
'=========================================================================

Sub Example_BasicUsage()
    Dim AI As New cAI
    Dim Response As String
    
    '最简单的使用方式（使用 OpenAI）
    Response = AI.ApiKey("sk-xxx").Chat("你好")
    Debug.Print "响应: " & Response
    
    '使用预设配置
    Response = AI.DeepSeek("your-deepseek-key").Chat("你好")
    Debug.Print "响应: " & Response
    
    '完整配置示例
    Response = AI.ApiKey("sk-xxx") _
        .Model("gpt-4") _
        .Temperature(0.7) _
        .MaxTokens(1000) _
        .System("你是一个专业的编程助手") _
        .Chat("如何实现单例模式？")
    Debug.Print "响应: " & Response
End Sub

'=========================================================================
' 全局对象使用示例
'=========================================================================

Sub Example_GlobalObject()
    Dim Response As String
    
    '使用全局对象 vbman（推荐）
    Response = VBMAN.AI.ApiKey("sk-xxx").Chat("你好")
    Debug.Print "响应: " & Response
    
    '使用预设配置
    Response = VBMAN.AI.Doubao("your-doubao-key").Chat("你好")
    Debug.Print "响应: " & Response
End Sub

'=========================================================================
' 多轮对话示例
'=========================================================================

Sub Example_MultiTurnConversation()
    Dim AI As New cAI
    Dim Response As String
    
    '配置 API
    AI.ApiKey("sk-xxx")
    
    '第一轮
    Response = AI.User("我喜欢编程").Chat()
    Debug.Print "AI: " & Response
    
    '第二轮（自动保留上下文）
    Response = AI.User("推荐一些学习资源").Chat()
    Debug.Print "AI: " & Response
    
    '第三轮
    Response = AI.User("VBA 有什么优势？").Chat()
    Debug.Print "AI: " & Response
    
    '清空历史重新开始
    Response = AI.ClearMessages().User("新话题：天气").Chat()
    Debug.Print "AI: " & Response
End Sub

'=========================================================================
' 流式响应示例 - 使用回调函数
'=========================================================================

Sub Example_StreamWithCallback()
    Dim AI As New cAI
    
    '使用回调函数处理流式响应
    AI.ApiKey("sk-xxx") _
        .Stream(True) _
        .Chat("写一首关于春天的诗", Me, "OnStreamChunk")
    
    Debug.Print "流式响应完成"
End Sub

'回调方法
Private Sub OnStreamChunk(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "流式响应完成"
    Else
        '实时输出流式内容
        Debug.Print Chunk;
    End If
End Sub

'=========================================================================
' 流式响应示例 - 使用事件
'=========================================================================

Sub Example_StreamWithEvents()
    Dim WithEvents AI As cAI
    Set AI = New cAI
    
    '启用流式响应并发送请求（不提供回调，自动触发事件）
    AI.ApiKey("sk-xxx") _
        .Stream(True) _
        .Chat("写一个快速排序算法")
End Sub

'事件处理方法（自动触发）
Private Sub AI_OnStreamChunk(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "流式响应完成"
    Else
        Debug.Print Chunk;
    End If
End Sub

Private Sub AI_OnComplete(Response As String)
    Debug.Print "完整响应: " & Response
End Sub

Private Sub AI_OnError(ErrorMsg As String, ErrorCode As Long)
    Debug.Print "错误: " & ErrorMsg & " (代码: " & ErrorCode & ")"
End Sub

'=========================================================================
' 全局对象流式响应示例
'=========================================================================

Sub Example_GlobalStream()
    '使用全局对象 vbman 一句话使用
    VBMAN.AI.ApiKey("sk-xxx") _
        .Stream(True) _
        .Chat("写一首诗", Me, "OnAIStreamChunk")
    
    '或者使用预设配置
    VBMAN.AI.Doubao("your-key") _
        .Stream(True) _
        .Chat("你好", Me, "OnAIResponse")
End Sub

Private Sub OnAIStreamChunk(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "完成"
    Else
        Debug.Print Chunk;
    End If
End Sub

Private Sub OnAIResponse(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "完成"
    Else
        Debug.Print Chunk;
    End If
End Sub

'=========================================================================
' 获取完整响应 JSON 示例
'=========================================================================

Sub Example_ChatJson()
    Dim AI As New cAI
    Dim ResponseJson As cJson
    
    Set ResponseJson = AI.ApiKey("sk-xxx").ChatJson("你好")
    
    '访问响应数据
    Debug.Print "内容: " & ResponseJson.Item("choices")(0).Item("message").Item("content")
    Debug.Print "总 Token: " & ResponseJson.Item("usage").Item("total_tokens")
    
    '如果响应是数组格式
    If ResponseJson.RootIsArray Then
        Debug.Print "Choices 数量: " & ResponseJson.Items.Count
    End If
End Sub

'=========================================================================
' 预设配置示例
'=========================================================================

Sub Example_Presets()
    Dim AI As New cAI
    Dim Response As String
    
    '使用豆包 API
    Response = AI.Doubao("your-doubao-key") _
        .Model("ep-xxx") _
        .Chat("你好")
    Debug.Print "豆包: " & Response
    
    '使用 DeepSeek API
    Response = AI.DeepSeek("your-deepseek-key") _
        .Model("deepseek-chat") _
        .Chat("你好")
    Debug.Print "DeepSeek: " & Response
    
    '使用 OpenAI API
    Response = AI.OpenAI("your-openai-key") _
        .Model("gpt-4") _
        .Chat("你好")
    Debug.Print "OpenAI: " & Response
    
    '使用通义千问 API
    Response = AI.Qwen("your-qwen-key") _
        .Model("qwen-turbo") _
        .Chat("你好")
    Debug.Print "通义千问: " & Response
End Sub

'=========================================================================
' 自定义配置示例
'=========================================================================

Sub Example_CustomConfig()
    Dim AI As New cAI
    Dim Response As String
    
    '完全自定义配置
    Response = AI.ApiKey("sk-xxx") _
        .BaseUrl("https://your-custom-api.com/v1") _
        .ApiPath("/custom/path") _
        .Model("your-model") _
        .Temperature(0.8) _
        .MaxTokens(2000) _
        .TopP(0.9) _
        .FrequencyPenalty(0.5) _
        .PresencePenalty(0.5) _
        .System("你是一个专业的技术顾问") _
        .Chat("帮我分析一下项目架构")
    
    Debug.Print "响应: " & Response
End Sub

'=========================================================================
' 高级配置示例 - 自定义请求头
'=========================================================================

Sub Example_CustomHeaders()
    Dim AI As New cAI
    Dim Response As String
    
    '添加自定义请求头
    Response = AI.ApiKey("sk-xxx") _
        .Header("X-Custom-Header", "custom-value") _
        .Header("X-Request-ID", "12345") _
        .Organization("org-xxx") _
        .Chat("你好")
    
    Debug.Print "响应: " & Response
End Sub

'=========================================================================
' 错误处理示例
'=========================================================================

Sub Example_ErrorHandling()
    Dim AI As New cAI
    
    On Error Resume Next
    
    Dim Response As String
    Response = AI.ApiKey("").Chat("你好")
    
    If Err.Number <> 0 Then
        Debug.Print "错误: " & Err.Description
        Debug.Print "AI.LastError: " & AI.LastError
    End If
    
    On Error GoTo 0
End Sub

'=========================================================================
' 重置配置示例
'=========================================================================

Sub Example_Reset()
    Dim AI As New cAI
    Dim Response As String
    
    '第一次请求
    Response = AI.ApiKey("sk-xxx") _
        .Model("gpt-4") _
        .System("你是一个英语老师") _
        .Chat("Hello")
    Debug.Print "响应1: " & Response
    
    '重置所有配置
    Call AI.Reset()
    
    '第二次请求（使用默认配置）
    Response = AI.ApiKey("sk-xxx").Chat("你好")
    Debug.Print "响应2: " & Response
End Sub

'=========================================================================
' 系统提示词使用示例
'=========================================================================

Sub Example_SystemPrompt()
    Dim AI As New cAI
    Dim Response As String
    
    '设置系统提示词
    Response = AI.ApiKey("sk-xxx") _
        .System("你是一个专业的翻译助手，只翻译，不解释") _
        .Chat("Hello, how are you?")
    
    Debug.Print "翻译: " & Response
    
    '多轮对话保持系统提示词
    Response = AI.User("What's the weather today?").Chat()
    Debug.Print "翻译: " & Response
End Sub

'=========================================================================
' 停止序列使用示例
'=========================================================================

Sub Example_StopSequence()
    Dim AI As New cAI
    Dim Response As String
    
    '设置停止序列
    Response = AI.ApiKey("sk-xxx") _
        .Stop(Array("\n", "。")) _
        .Chat("写一个简短的故事")
    
    Debug.Print "响应（会在遇到停止序列时停止）: " & Response
End Sub

'=========================================================================
' 窗体中使用 WithEvents 的完整示例
'=========================================================================

'在窗体代码模块中：
'
'Option Explicit
'Dim WithEvents AI As cAI
'
'Private Sub Form_Load()
'    Set AI = New cAI
'End Sub
'
'Private Sub CommandSend_Click()
'    AI.ApiKey("sk-xxx") _
'        .Stream(True) _
'        .Chat(TextQuestion.Text)
'End Sub
'
'Private Sub AI_OnStreamChunk(Chunk As String, IsComplete As Boolean)
'    TextResponse.Text = TextResponse.Text & Chunk
'    TextResponse.SelStart = Len(TextResponse.Text)
'    TextResponse.SelLength = 0
'    DoEvents
'End Sub
'
'Private Sub AI_OnComplete(Response As String)
'    LabelStatus.Caption = "完成"
'End Sub
'
'Private Sub AI_OnError(ErrorMsg As String, ErrorCode As Long)
'    MsgBox "错误: " & ErrorMsg, vbExclamation
'    LabelStatus.Caption = "错误"
'End Sub

'=========================================================================
' 测试函数
'=========================================================================

Sub Test_AllExamples()
    Debug.Print "=== 测试基础使用 ==="
    'Call Example_BasicUsage()
    
    Debug.Print vbCrLf & "=== 测试多轮对话 ==="
    'Call Example_MultiTurnConversation()
    
    Debug.Print vbCrLf & "=== 测试预设配置 ==="
    'Call Example_Presets()
    
    Debug.Print vbCrLf & "=== 测试流式响应 ==="
    'Call Example_StreamWithCallback()
    
    Debug.Print vbCrLf & "=== 测试完成 ==="
End Sub
