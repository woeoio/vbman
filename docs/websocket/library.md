# WebSocket 类库总览

## 📋 类库组件

### 公共类（用户直接使用）

| 类名 | 文件 | 用途 |
|------|------|------|
| `cWebSocketClient` | cWebSocketClient.cls | WebSocket 客户端实现 |
| `cWebSocketServer` | cWebSocketServer.cls | WebSocket 服务端实现 |

### 辅助类（内部使用）

| 类名 | 文件 | 用途 |
|------|------|------|
| `cWebSocketServerClient` | cWebSocketServerClient.cls | 服务端客户端连接管理 |
| `cWebSocketFrame` | cWebSocketFrame.cls | WebSocket 帧解析和构建 |
| `cByteBuffer` | cByteBuffer.cls | 高效字节缓冲区管理 |

### 工具模块（公共）

| 模块名 | 文件 | 用途 |
|--------|------|------|
| `mWebSocketUtils` | mWebSocketUtils.bas | 公共工具函数（UTF-8, Base64, SHA1 等） |

---

## 📦 类库结构

```
newWebsocket/
├── mWebSocketUtils.bas        # 公共工具模块
├── cByteBuffer.cls            # 高效字节缓冲区
├── cWebSocketFrame.cls        # WebSocket 帧解析器
├── cWebSocketClient.cls       # WebSocket 客户端
├── cWebSocketServer.cls       # WebSocket 服务端
└── cWebSocketServerClient.cls # 服务端客户端连接
```

---

## 🎯 核心设计理念

### 1. 职责分离

每个类只负责一项功能，便于维护和调试：

- **cWebSocketClient** - 仅处理客户端逻辑（连接、发送、接收）
- **cWebSocketServer** - 仅处理服务端逻辑（监听、客户端管理、广播）
- **cWebSocketFrame** - 仅处理帧的解析和构建
- **cByteBuffer** - 仅处理字节缓冲区管理

### 2. 只读解析

`cWebSocketFrame.ParseHeader()` 不会修改输入数据，避免了缓冲区混乱问题：

```vb
' 流程清晰，步骤分明
1. ParseHeader(buffer)    - 解析头部（只读）
2. IsCompleteFrame(buffer) - 检查完整性
3. ExtractPayload(buffer)  - 提取并unmask（消费缓冲区）
```

### 3. 高效缓冲区

`cByteBuffer` 采用预分配策略：

- 初始容量：4KB
- 增长因子：1.5 倍
- 最小化 ReDim Preserve 调用

### 4. 自动握手

- 客户端：自动生成 WebSocket Key、发送握手请求、验证响应
- 服务端：自动检测 Upgrade 请求、计算 Accept Key、发送响应

### 5. 事件驱动

所有操作通过事件通知，无需轮询：

- 连接成功 → `OnOpen`
- 收到消息 → `OnTextMessage` / `OnBinaryMessage`
- 连接关闭 → `OnClose`
- 发生错误 → `OnError`

---

## 🔄 连接状态机

### 客户端状态 (WsState)

```
WS_STATE_CLOSED (0)
    ↓ Connect()
WS_STATE_CONNECTING (1)
    ↓ 握手成功
WS_STATE_OPEN (2)
    ↓ Close()
WS_STATE_CLOSING (3)
    ↓ 关闭完成
WS_STATE_CLOSED (0)
```

### 服务端客户端状态 (WsClientState)

```
WS_CLIENT_PENDING (0)
    ↓ 握手成功
WS_CLIENT_CONNECTED (1)
    ↓ SendClose()
WS_CLIENT_CLOSING (2)
    ↓ 关闭完成
    (对象被移除)
```

---

## 📨 WebSocket 帧

### 帧结构

```
 0                   1                   2                   3
 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
+-+-+-+-+-------+-+-------------+-------------------------------+
|F|R|R|R| opcode|M| Payload len |    Extended payload length    |
|I|S|S|S|  (4)  |A|     (7)     |             (16/64)           |
|N|V|V|V|       |S|             |   (if payload len==126/127)   |
| |1|2|3|       |K|             |                               |
+-+-+-+-+-------+-+-------------+ - - - - - - - - - - - - - - - +
|     Extended payload length continued, if payload len == 127  |
+ - - - - - - - - - - - - - - - +-------------------------------+
|                               |Masking-key, if MASK set to 1  |
+-------------------------------+-------------------------------+
| Masking-key (continued)       |          Payload Data           |
+-------------------------------- - - - - - - - - - - - - - - - +
:                     Payload Data continued ...                :
+ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - +
|                     Payload Data continued ...                |
+---------------------------------------------------------------+
```

### 帧类型 (WsOpCode)

| OpCode | 名称 | 说明 |
|--------|------|------|
| 0 | CONTINUATION | 分片帧的后续帧 |
| 1 | TEXT | 文本数据帧 |
| 2 | BINARY | 二进制数据帧 |
| 8 | CLOSE | 关闭连接帧 |
| 9 | PING | Ping 帧（保活） |
| 10 | PONG | Pong 帧（响应） |

### 关闭代码 (WsCloseCode)

| 代码 | 名称 | 说明 |
|------|------|------|
| 1000 | NORMAL | 正常关闭 |
| 1001 | GOING_AWAY | 端点离开 |
| 1002 | PROTOCOL_ERROR | 协议错误 |
| 1003 | UNSUPPORTED_DATA | 不支持的数据类型 |
| 1006 | ABNORMAL | 异常关闭（本地） |
| 1007 | INVALID_DATA | 无效数据 |
| 1008 | POLICY_VIOLATION | 策略违规 |
| 1009 | MESSAGE_TOO_BIG | 消息过大 |
| 1011 | INTERNAL_ERROR | 内部错误 |

---

## 🚀 性能特点

### 内存效率

- **预分配缓冲区**：减少频繁的内存分配
- **智能增长**：1.5 倍增长因子，平衡空间和性能
- **及时释放**：连接关闭时自动清理资源

### CPU 效率

- **批量处理**：一次解析多个帧
- **最小化拷贝**：使用 CopyMemory API
- **异步处理**：不阻塞主线程

### 网络效率

- **自动分片**：大消息自动分片传输
- **延迟响应**：Ping/Pong 及时响应
- **批量广播**：广播消息一次构建，多次发送

---

## 🛡️ 安全特性

### 握手验证

- **WebSocket Key 验证**：防止跨站 WebSocket 劫持
- **协议版本检查**：确保使用兼容版本
- **HTTP 头解析**：严格验证 Upgrade 请求

### 数据掩码

- **客户端掩码**：所有客户端发送的帧都进行掩码
- **服务端不掩码**：服务端发送的帧不掩码（符合 RFC 6455）

### 错误处理

- **统一错误上报**：所有错误通过 OnError 事件
- **自动关闭**：协议错误自动关闭连接
- **状态码详细**：关闭时提供详细的状态码和原因

---

## 📊 与旧版类库对比

| 特性 | 旧版 | 新版 |
|------|------|------|
| 客户端/服务端 | 混在一个类 | 分离为独立类 |
| 帧解析 | 修改输入数据 | 只读解析 |
| 缓冲区管理 | 动态分配 | 预分配+增长 |
| 职责分离 | 不清晰 | 清晰分离 |
| 调试难度 | 较高 | 较低 |
| 代码可维护性 | 一般 | 优秀 |

---

## 📝 使用场景

### 实时聊天应用

```vb
' 客户端连接聊天服务器
m_Client.Connect "ws://chat.example.com:8080"

' 发送聊天消息
m_Client.SendText "Hello everyone!"

' 接收聊天消息（事件驱动）
Private Sub m_Client_OnTextMessage(ByVal Message As String)
    DisplayMessage Message
End Sub
```

### 实时数据推送

```vb
' 服务端向所有客户端推送数据
Private Sub Timer1_Timer()
    Dim sData As String
    sData = GetLatestData()
    m_Server.BroadcastText sData
End Sub
```

### 在线游戏

```vb
' 发送游戏操作（二进制）
Dim baData() As Byte
baData = SerializeGameAction()
m_Client.SendBinary baData

' 接收其他玩家操作
Private Sub m_Client_OnBinaryMessage(Data() As Byte)
    ProcessGameAction Data
End Sub
```

### 设备控制

```vb
' 控制命令发送
m_Server.SendText ClientID, "TURN_ON"

' 接收设备状态
Private Sub m_Server_OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
    UpdateDeviceStatus ClientID, Message
End Sub
```

---

## 🔧 扩展性

### 添加自定义子协议

```vb
' 在握手时指定子协议
m_Client.Connect "ws://example.com:8080", "chat.v1"
```

### 添加自定义 HTTP 头

```vb
' 扩展 SendHandshake 方法添加自定义头
' 注意：需要在类内部修改
```

### 自定义帧处理

```vb
' 继承 cWebSocketFrame 类添加自定义帧类型
' 或修改 ProcessReceivedData 方法
```

---

## ⚠️ 已知限制

- **不支持 wss://**：SSL/TLS 需要额外实现
- **不支持压缩**：WebSocket 压缩扩展未实现
- **单线程模型**：所有操作在主线程执行
- **最大消息大小**：受 VB6 数组限制（约 2GB）

---

**最后更新**: 2026-01-10
