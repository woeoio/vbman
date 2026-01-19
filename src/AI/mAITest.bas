Attribute VB_Name = "mAITest"
Option Explicit
'AI 对象测试工具
'作者：邓伟，QQ：215879458

'=========================================================================
' 配置区域 - 在这里修改你的 API Key
'=========================================================================

Private Const TEST_API_KEY As String = "" '在此填入你的 API Key
Private Const TEST_PROVIDER As String = "deepseek" '可选: openai, deepseek, doubao, qwen

'=========================================================================
' 测试函数
'=========================================================================

Sub RunAllTests()
    If TEST_API_KEY = "" Then
        MsgBox "请先在代码中设置 TEST_API_KEY", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "=========================================="
    Debug.Print "开始 AI 对象测试"
    Debug.Print "=========================================="
    Debug.Print ""
    
    '基础功能测试
    Call Test_BasicChat()
    Debug.Print ""
    
    '多轮对话测试
    Call Test_MultiTurn()
    Debug.Print ""
    
    '流式响应测试
    Call Test_StreamResponse()
    Debug.Print ""
    
    '预设配置测试
    Call Test_Presets()
    Debug.Print ""
    
    '错误处理测试
    Call Test_ErrorHandling()
    Debug.Print ""
    
    Debug.Print "=========================================="
    Debug.Print "所有测试完成"
    Debug.Print "=========================================="
End Sub

'测试 1: 基础聊天
Sub Test_BasicChat()
    Debug.Print "--- 测试 1: 基础聊天 ---"
    
    Dim AI As New cAI
    Dim Response As String
    
    On Error Resume Next
    Response = AI.ApiKey(TEST_API_KEY).Chat("你好，请用一句话介绍自己")
    
    If Err.Number <> 0 Then
        Debug.Print "[失败] 错误: " & Err.Description
    Else
        Debug.Print "[成功] 响应: " & Left(Response, 100) & "..."
    End If
    
    On Error GoTo 0
End Sub

'测试 2: 多轮对话
Sub Test_MultiTurn()
    Debug.Print "--- 测试 2: 多轮对话 ---"
    
    Dim AI As New cAI
    Dim Response As String
    
    AI.ApiKey(TEST_API_KEY)
    
    '第一轮
    On Error Resume Next
    Response = AI.User("我喜欢编程").Chat()
    If Err.Number <> 0 Then
        Debug.Print "[失败] 第一轮错误: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    Debug.Print "[成功] 第一轮: " & Left(Response, 50) & "..."
    
    '第二轮
    Response = AI.User("推荐一些 VBA 学习资源").Chat()
    If Err.Number <> 0 Then
        Debug.Print "[失败] 第二轮错误: " & Err.Description
    Else
        Debug.Print "[成功] 第二轮: " & Left(Response, 50) & "..."
    End If
    
    On Error GoTo 0
End Sub

'测试 3: 流式响应
Sub Test_StreamResponse()
    Debug.Print "--- 测试 3: 流式响应 ---"
    Debug.Print "（流式响应将实时输出）"
    
    Dim AI As New cAI
    
    On Error Resume Next
    AI.ApiKey(TEST_API_KEY) _
        .Stream(True) _
        .Chat("请用5个词描述人工智能", Me, "TestStreamCallback")
    
    If Err.Number <> 0 Then
        Debug.Print "[失败] 错误: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub

Private Sub TestStreamCallback(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "[完成] 流式响应测试完成"
    Else
        Debug.Print Chunk;
    End If
End Sub

'测试 4: 预设配置
Sub Test_Presets()
    Debug.Print "--- 测试 4: 预设配置 ---"
    
    Dim AI As New cAI
    Dim Response As String
    
    On Error Resume Next
    
    Select Case LCase(TEST_PROVIDER)
        Case "openai"
            AI.OpenAI TEST_API_KEY
            Debug.Print "使用预设: OpenAI"
        Case "deepseek"
            AI.DeepSeek TEST_API_KEY
            Debug.Print "使用预设: DeepSeek"
        Case "doubao"
            AI.Doubao TEST_API_KEY
            Debug.Print "使用预设: 豆包"
        Case "qwen"
            AI.Qwen TEST_API_KEY
            Debug.Print "使用预设: 通义千问"
        Case Else
            AI.ApiKey TEST_API_KEY
            Debug.Print "使用默认配置"
    End Select
    
    Response = AI.Chat("你好")
    
    If Err.Number <> 0 Then
        Debug.Print "[失败] 错误: " & Err.Description
    Else
        Debug.Print "[成功] 响应: " & Left(Response, 100) & "..."
    End If
    
    On Error GoTo 0
End Sub

'测试 5: 错误处理
Sub Test_ErrorHandling()
    Debug.Print "--- 测试 5: 错误处理 ---"
    
    Dim AI As New cAI
    
    '测试空 API Key
    On Error Resume Next
    AI.ApiKey("").Chat("测试")
    
    If Err.Number <> 0 Then
        Debug.Print "[成功] 正确捕获空 API Key 错误: " & Err.Description
    Else
        Debug.Print "[失败] 未捕获空 API Key 错误"
    End If
    
    On Error GoTo 0
End Sub

'测试 6: 全局对象
Sub Test_GlobalObject()
    Debug.Print "--- 测试 6: 全局对象 ---"
    
    If TEST_API_KEY = "" Then
        Debug.Print "[跳过] 需要设置 API Key"
        Exit Sub
    End If
    
    Dim Response As String
    
    On Error Resume Next
    Response = VBMAN.AI.ApiKey(TEST_API_KEY).Chat("使用全局对象测试")
    
    If Err.Number <> 0 Then
        Debug.Print "[失败] 错误: " & Err.Description
    Else
        Debug.Print "[成功] 响应: " & Left(Response, 100) & "..."
    End If
    
    On Error GoTo 0
End Sub

'测试 7: 参数配置
Sub Test_ParameterConfig()
    Debug.Print "--- 测试 7: 参数配置 ---"
    
    If TEST_API_KEY = "" Then
        Debug.Print "[跳过] 需要设置 API Key"
        Exit Sub
    End If
    
    Dim AI As New cAI
    Dim Response As String
    
    On Error Resume Next
    Response = AI.ApiKey(TEST_API_KEY) _
        .Temperature(0.5) _
        .MaxTokens(100) _
        .System("你是一个简洁的助手") _
        .Chat("用一句话介绍 VBA")
    
    If Err.Number <> 0 Then
        Debug.Print "[失败] 错误: " & Err.Description
    Else
        Debug.Print "[成功] 响应: " & Left(Response, 100) & "..."
    End If
    
    On Error GoTo 0
End Sub

'测试 8: 重置功能
Sub Test_Reset()
    Debug.Print "--- 测试 8: 重置功能 ---"
    
    If TEST_API_KEY = "" Then
        Debug.Print "[跳过] 需要设置 API Key"
        Exit Sub
    End If
    
    Dim AI As New cAI
    
    '第一次配置
    AI.ApiKey(TEST_API_KEY) _
        .Model("gpt-4") _
        .System("系统提示1") _
    
    Debug.Print "[检查] 重置前 Model: " & AI.LastError
    
    '重置
    Call AI.Reset()
    
    '第二次配置
    AI.ApiKey(TEST_API_KEY)
    
    Debug.Print "[成功] 重置完成，可以使用新的配置"
End Sub

'=========================================================================
' 单独测试函数
'=========================================================================

Sub QuickTest()
    If TEST_API_KEY = "" Then
        MsgBox "请先在代码中设置 TEST_API_KEY", vbExclamation
        Exit Sub
    End If
    
    Dim Response As String
    Response = VBMAN.AI.ApiKey(TEST_API_KEY).Chat("你好")
    MsgBox Response, vbInformation, "AI 响应"
End Sub

Sub QuickStreamTest()
    If TEST_API_KEY = "" Then
        MsgBox "请先在代码中设置 TEST_API_KEY", vbExclamation
        Exit Sub
    End If
    
    VBMAN.AI.ApiKey(TEST_API_KEY) _
        .Stream(True) _
        .Chat("写一首关于编程的诗", Me, "QuickStreamCallback")
End Sub

Private Sub QuickStreamCallback(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print vbCrLf & "完成！"
        MsgBox "流式响应完成", vbInformation
    Else
        Debug.Print Chunk;
    End If
End Sub

'=========================================================================
' 性能测试
'=========================================================================

Sub PerformanceTest()
    If TEST_API_KEY = "" Then
        MsgBox "请先在代码中设置 TEST_API_KEY", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "--- 性能测试 ---"
    
    Dim AI As New cAI
    Dim StartTime As Double
    Dim Response As String
    Dim i As Long
    
    AI.ApiKey(TEST_API_KEY)
    
    '测试 5 次请求
    For i = 1 To 5
        StartTime = Timer
        
        On Error Resume Next
        Response = AI.Chat("第 " & i & " 次测试，请回答" & i)
        
        If Err.Number <> 0 Then
            Debug.Print "[失败] 第 " & i & " 次: " & Err.Description
        Else
            Debug.Print "[成功] 第 " & i & " 次 - 耗时: " & Format(Timer - StartTime, "0.000") & "秒 - 长度: " & Len(Response) & "字符"
        End If
        
        On Error GoTo 0
    Next i
End Sub

'=========================================================================
' 快捷菜单
'=========================================================================

Sub ShowTestMenu()
    Dim Menu As String
    Menu = "AI 对象测试工具" & vbCrLf & vbCrLf
    Menu = Menu & "1. 运行所有测试" & vbCrLf
    Menu = Menu & "2. 快速测试（单次请求）" & vbCrLf
    Menu = Menu & "3. 快速流式测试" & vbCrLf
    Menu = Menu & "4. 性能测试" & vbCrLf
    Menu = Menu & "5. 测试全局对象" & vbCrLf
    Menu = Menu & "6. 测试参数配置" & vbCrLf
    Menu = Menu & vbCrLf
    Menu = Menu & "请输入选项（1-6）："
    
    Dim Choice As String
    Choice = InputBox(Menu, "AI 测试工具")
    
    Select Case Choice
        Case "1"
            Call RunAllTests
        Case "2"
            Call QuickTest
        Case "3"
            Call QuickStreamTest
        Case "4"
            Call PerformanceTest
        Case "5"
            Call Test_GlobalObject
        Case "6"
            Call Test_ParameterConfig
        Case Else
            MsgBox "无效选项", vbExclamation
    End Select
End Sub
