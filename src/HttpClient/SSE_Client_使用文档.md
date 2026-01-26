# cSSEClient 使用文档

## 概述

`cSSEClient` 是一个 Server-Sent Events (SSE) 客户端实现，用于接收服务器端实时推送的事件流。

### SSE 协议简介

Server-Sent Events (SSE) 是一种基于 HTTP 的单向服务器推送技术，适用于：
- 实时通知
- 股票价格推送
- 聊天室消息
- 服务器日志流
- 实时数据监控

## 特性

✅ **自动重连** - 连接断开时自动重连，支持自定义重连间隔和最大重连次数
✅ **断点续传** - 使用 `Last-Event-ID` 实现断线重连后继续接收事件
✅ **事件类型支持** - 支持多种事件类型（event、data、id、retry）
✅ **错误处理** - 完善的错误处理和重连机制
✅ **缓冲区管理** - 自动处理不完整的数据行
✅ **心跳保活** - 支持服务器心跳检测

## 快速开始

### 1. 基础使用

```vb
Dim WithEvents SSE As New cSSEClient

Public Sub ConnectToSSE()
    '启用自动重连
    SSE.AutoReconnect = True
    SSE.ReconnectInterval = 3000      '3秒后重连
    SSE.MaxReconnectAttempts = 10     '最多重连10次

    '连接到 SSE 服务端
    SSE.Connect "http://localhost:3000/events"
End Sub

'接收消息事件
Private Sub SSE_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    Debug.Print "收到消息: " & Data
End Sub

'连接建立事件
Private Sub SSE_OnOpen()
    Debug.Print "SSE 连接已建立"
End Sub

'错误事件
Private Sub SSE_OnError(ByVal Description As String, ByVal ErrorNumber As Long)
    Debug.Print "错误: " & Description
End Sub

'连接关闭事件
Private Sub SSE_OnClose()
    Debug.Print "SSE 连接已关闭"
End Sub
```

## API 参考

### 属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `AutoReconnect` | Boolean | 是否自动重连（默认 True） |
| `ReconnectInterval` | Long | 重连间隔（毫秒，默认 3000） |
| `MaxReconnectAttempts` | Long | 最大重连次数（默认 10） |
| `Url` | String | SSE 服务端 URL |
| `Connected` | Boolean | 只读，当前连接状态 |
| `Connecting` | Boolean | 只读，是否正在连接 |
| `LastReceivedEventId` | String | 只读，最后收到的事件 ID |
| `ServerReconnectInterval` | Long | 只读，服务器指定的重连间隔 |
| `CurrentReconnectDelay` | Long | 只读，当前重连延迟 |

### 方法

| 方法 | 参数 | 说明 |
|------|------|------|
| `Connect` | `ServerUrl` | 连接到 SSE 服务端 |
| `Disconnect` | 无 | 断开连接并停止自动重连 |
| `DisableAutoReconnect` | 无 | 关闭自动重连 |
| `EnableAutoReconnect` | 无 | 启用自动重连 |
| `ResetReconnectAttempts` | 无 | 重置重连计数器 |

### 事件

| 事件 | 参数 | 说明 |
|------|------|------|
| `OnOpen` | 无 | 连接建立时触发 |
| `OnMessage` | `EventName`, `Data`, `Id` | 收到消息时触发 |
| `OnError` | `Description`, `ErrorNumber` | 发生错误时触发 |
| `OnClose` | 无 | 连接关闭时触发 |

## SSE 数据格式

SSE 服务器推送的数据格式如下：

```
event: message
id: 12345
retry: 5000
data: {"type": "notification", "content": "Hello World"}

event: heartbeat
data: keep-alive

data: 简单文本消息

```

### 字段说明

- `event` - 事件类型（可选）
- `id` - 事件 ID（可选），用于断点续传
- `retry` - 重连间隔（可选，毫秒）
- `data` - 事件数据（必填），多行数据用多个 `data:` 行
- 空行表示一个事件结束

### 注释行

以 `:` 开头的行是注释，会被忽略：

```
: 这是注释行，客户端不会处理
event: message
data: 实际数据
```

## 使用场景示例

### 1. 聊天室实时消息

```vb
Public Sub JoinChatRoom()
    Set SSE = New cSSEClient
    SSE.AutoReconnect = True
    SSE.Connect "http://localhost:3000/chat/room1"
End Sub

Private Sub SSE_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    If EventName = "chat" Then
        Dim Json As New cJson
        Json.Decode Data
        
        Dim UserName As String
        Dim Message As String
        
        UserName = Json.Item("user")
        Message = Json.Item("message")
        
        Debug.Print UserName & ": " & Message
    End If
End Sub
```

### 2. 股票价格推送

```vb
Public Sub SubscribeStockPrice()
    Set SSE = New cSSEClient
    SSE.AutoReconnect = True
    SSE.ReconnectInterval = 2000
    SSE.Connect "http://localhost:3000/stock/price"
End Sub

Private Sub SSE_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    If EventName = "price" Then
        Dim Json As New cJson
        Json.Decode Data
        
        Dim Symbol As String
        Dim Price As Double
        Dim Change As Double
        
        Symbol = Json.Item("symbol")
        Price = Json.Item("price")
        Change = Json.Item("change")
        
        Debug.Print Symbol & " 价格: " & Price & " 涨跌: " & Change
    End If
End Sub
```

### 3. 服务器日志推送

```vb
Public Sub SubscribeServerLogs()
    Set SSE = New cSSEClient
    SSE.AutoReconnect = True
    SSE.MaxReconnectAttempts = 20
    SSE.Connect "http://localhost:3000/logs/stream"
End Sub

Private Sub SSE_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    If EventName = "log" Then
        Dim Json As New cJson
        Json.Decode Data
        
        Dim Level As String
        Dim Message As String
        
        Level = Json.Item("level")
        Message = Json.Item("message")
        
        Select Case Level
            Case "ERROR"
                Debug.Print "[ERROR] " & Message
            Case "WARN"
                Debug.Print "[WARN] " & Message
            Case "INFO"
                Debug.Print "[INFO] " & Message
        End Select
    End If
End Sub
```

## 服务器端示例（Node.js）

```javascript
const express = require('express');
const app = express();

app.get('/events', (req, res) => {
  // 设置 SSE 响应头
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  let eventId = 0;

  // 定时发送消息
  const interval = setInterval(() => {
    eventId++;
    
    res.write(`event: message\n`);
    res.write(`id: ${eventId}\n`);
    res.write(`data: ${JSON.stringify({ time: Date.now() })}\n\n`);
    
    // 发送心跳
    res.write(`: keep-alive\n\n`);
    
  }, 1000);

  // 客户端断开连接时清理
  req.on('close', () => {
    clearInterval(interval);
  });
});

app.listen(3000, () => {
  console.log('SSE Server running on port 3000');
});
```

## 注意事项

### 1. 连接管理

- 使用 `Disconnect()` 方法优雅地关闭连接
- 调用 `DisableAutoReconnect()` 可以完全停止自动重连
- 对象销毁时会自动清理连接和定时器

### 2. 重连策略

- `ReconnectInterval` 设置默认重连间隔
- 服务器可以通过 `retry:` 字段指定重连间隔
- 达到 `MaxReconnectAttempts` 后会停止重连

### 3. 错误处理

- 网络错误、超时会触发 `OnError` 事件
- 自动重连失败后会停止并通知
- 建议在 `OnError` 中记录日志或提示用户

### 4. 性能优化

- 高频消息可能导致消息处理延迟
- 可以在 `OnMessage` 中实现消息队列
- 考虑使用多线程处理耗时操作

### 5. 与 WebSocket 的区别

| 特性 | SSE | WebSocket |
|------|-----|-----------|
| 方向 | 服务器 → 客户端 | 双向 |
| 协议 | HTTP | WebSocket |
| 自动重连 | 支持 | 需手动实现 |
| 文本/二进制 | 仅文本 | 文本和二进制 |
| 浏览器支持 | 所有现代浏览器 | 所有现代浏览器 |

选择 SSE 的场景：
- 只需要服务器推送
- 需要简单的实现
- 需要自动重连
- 推送的是文本数据

选择 WebSocket 的场景：
- 需要双向通信
- 需要传输二进制数据
- 需要低延迟
- 需要复杂的状态管理

## 完整示例

参考 `src/HttpClient/mSSEDemo.bas` 文件中的完整示例代码。

## 版本历史

- **2025-01-19** - 初始版本
  - 支持 SSE 协议基础功能
  - 实现自动重连机制
  - 支持断点续传
  - 完善的事件处理

## 依赖

- `cHttpClient.cls` - HTTP 客户端
- `cTimer.cls` - 定时器（位于 `Tools/DateTime/Timer/`）
- `cJson.cls` - JSON 解析器

## 作者

邓伟，QQ：215879458

## 许可

可自由传播，但请勿删本签名
