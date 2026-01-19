# cAI 类使用说明

## 概述

`cAI` 是一个通用的 AI 对象类，能够对接任何兼容 OpenAI API 格式的接口（如豆包、DeepSeek、OpenAI、通义千问等）。

## 特性

- ✅ **链式调用**：支持流畅的链式 API 设计
- ✅ **多服务商支持**：内置预设配置，支持主流 AI 服务商
- ✅ **流式响应**：支持实时流式输出（回调函数或事件）
- ✅ **多轮对话**：自动维护对话上下文
- ✅ **全局对象**：通过 `VBMAN.AI` 直接使用
- ✅ **灵活配置**：支持完全自定义 API 参数

## 快速开始

### 方式一：使用全局对象（推荐）

```vb
Dim Response As String
Response = VBMAN.AI.ApiKey("sk-xxx").Chat("你好")
```

### 方式二：创建独立对象

```vb
Dim AI As New cAI
Dim Response As String
Response = AI.ApiKey("sk-xxx").Chat("你好")
```

## 基础用法

### 1. 最简单的使用

```vb
'使用 OpenAI
Response = AI.ApiKey("sk-xxx").Chat("你好")

'使用预设配置
Response = AI.DeepSeek("your-key").Chat("你好")
```

### 2. 完整配置

```vb
Response = AI.ApiKey("sk-xxx") _
    .Model("gpt-4") _
    .Temperature(0.7) _
    .MaxTokens(1000) _
    .System("你是一个专业的编程助手") _
    .Chat("如何实现单例模式？")
```

### 3. 多轮对话

```vb
AI.ApiKey("sk-xxx")

'第一轮
AI.User("我喜欢编程").Chat()

'第二轮（自动保留上下文）
AI.User("推荐一些学习资源").Chat()

'清空历史
AI.ClearMessages().User("新话题").Chat()
```

## 流式响应

### 方式一：使用回调函数（推荐）

```vb
'发送请求
AI.ApiKey("sk-xxx") _
    .Stream(True) _
    .Chat("写一首诗", Me, "OnStreamChunk")

'回调函数
Private Sub OnStreamChunk(Chunk As String, IsComplete As Boolean)
    If IsComplete Then
        Debug.Print "完成"
    Else
        Debug.Print Chunk;
    End If
End Sub
```

### 方式二：使用事件

```vb
'声明时使用 WithEvents
Dim WithEvents AI As cAI
Set AI = New cAI

'发送请求
AI.ApiKey("sk-xxx").Stream(True).Chat("写一首诗")

'事件处理
Private Sub AI_OnStreamChunk(Chunk As String, IsComplete As Boolean)
    Debug.Print Chunk;
End Sub
```

## 预设配置

```vb
'豆包
Response = AI.Doubao("your-doubao-key").Chat("你好")

'DeepSeek
Response = AI.DeepSeek("your-deepseek-key").Chat("你好")

'OpenAI
Response = AI.OpenAI("your-openai-key").Chat("你好")

'通义千问
Response = AI.Qwen("your-qwen-key").Chat("你好")
```

## API 参考

### 基础配置方法

| 方法               | 说明          | 示例                                  |
| ------------------ | ------------- | ------------------------------------- |
| `ApiKey(Key)`      | 设置 API 密钥 | `.ApiKey("sk-xxx")`                   |
| `BaseUrl(Url)`     | 设置基础 URL  | `.BaseUrl("https://api.example.com")` |
| `ApiPath(Path)`    | 设置 API 路径 | `.ApiPath("/chat/completions")`       |
| `Model(Name)`      | 设置模型名称  | `.Model("gpt-4")`                     |
| `Timeout(Seconds)` | 设置超时时间  | `.Timeout(120)`                       |

### 参数配置方法

| 方法                      | 说明           | 默认值 |
| ------------------------- | -------------- | ------ |
| `Temperature(Value)`      | 温度参数 (0-2) | 1.0    |
| `MaxTokens(Value)`        | 最大 token 数  | 自动   |
| `TopP(Value)`             | Top-P 采样     | 1.0    |
| `FrequencyPenalty(Value)` | 频率惩罚       | 0      |
| `PresencePenalty(Value)`  | 存在惩罚       | 0      |
| `Stream(Enable)`          | 启用流式响应   | False  |
| `Stop(Value)`             | 停止序列       | 无     |

### 消息管理方法

| 方法                 | 说明           |
| -------------------- | -------------- |
| `System(Prompt)`     | 设置系统提示词 |
| `User(Content)`      | 添加用户消息   |
| `Assistant(Content)` | 添加助手消息   |
| `ClearMessages()`    | 清空消息历史   |
| `Messages(MsgList)`  | 设置消息列表   |

### 高级配置方法

| 方法                  | 说明             |
| --------------------- | ---------------- |
| `Header(Name, Value)` | 添加自定义请求头 |
| `Proxy(Url)`          | 设置代理         |
| `Organization(Id)`    | 设置组织 ID      |

### 执行方法

| 方法                                | 说明                     |
| ----------------------------------- | ------------------------ |
| `Chat([Msg], [Callback], [Method])` | 发送聊天请求             |
| `ChatJson([Msg])`                   | 发送请求并返回 JSON 对象 |
| `Reset()`                           | 重置所有配置             |

### 事件

| 事件                               | 说明           |
| ---------------------------------- | -------------- |
| `OnStreamChunk(Chunk, IsComplete)` | 流式响应数据块 |
| `OnComplete(Response)`             | 请求完成       |
| `OnError(ErrorMsg, ErrorCode)`     | 错误事件       |

## 属性

| 属性        | 说明             |
| ----------- | ---------------- |
| `LastError` | 最后一次错误信息 |

## 窗体中使用示例

```vb
Option Explicit
Dim WithEvents AI As cAI

Private Sub Form_Load()
    Set AI = New cAI
End Sub

Private Sub CommandSend_Click()
    AI.ApiKey("sk-xxx") _
        .Stream(True) _
        .Chat(TextQuestion.Text)
End Sub

Private Sub AI_OnStreamChunk(Chunk As String, IsComplete As Boolean)
    TextResponse.Text = TextResponse.Text & Chunk
    TextResponse.SelStart = Len(TextResponse.Text)
    DoEvents
End Sub
```

## 注意事项

1. **API Key 安全**：不要在代码中硬编码 API Key，建议从配置文件或环境变量读取
2. **错误处理**：所有 API 调用都可能失败，需要完善的错误处理机制
3. **超时设置**：根据模型和网络情况合理设置超时时间
4. **Token 限制**：注意不同模型的 token 限制，避免超出限制
5. **成本控制**：流式响应和长文本会消耗更多 token，注意成本控制
6. **兼容性**：不同服务商的 API 可能有细微差异，需要测试验证

## 完整示例

查看 `mAIDemo.bas` 文件，包含所有使用场景的完整示例代码。

## 版本历史

- **2025-01-19**：初始版本
  - 支持 OpenAI 兼容 API
  - 支持流式响应
  - 支持预设配置
  - 支持链式调用
