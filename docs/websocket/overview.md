# WebSocket 类库开发文档

> 🚀 **WebSocket 类库** - 基于 cWinsock 封装的 VB6 WebSocket 实现库，由 woeoio@qq.com 使用 claude ai 辅助 基于 VbAsyncSocket 开发

## 📖 目录

- [概述](#概述)
- [核心亮点](#核心亮点)
- [架构设计](#架构设计)
- [文档索引](#文档索引)

---

## 概述

WebSocket 类库是一个为 VB6 设计的轻量级 WebSocket 通信库，完全符合 RFC 6455 标准。它基于 cWinsock 类实现，提供了简洁易用的 API 和完整的功能支持。

### ✨ 主要特性

- 🔌 **纯类实现** - 无需控件，直接使用对象编程
- 📦 **分离式设计** - 客户端和服务端独立类库，职责清晰
- 🎯 **完整协议支持** - 支持 WebSocket 文本、二进制、Ping/Pong、Close 帧类型
- 🌐 **高效缓冲区** - 预分配字节缓冲区，减少内存分配操作
- 🛡️ **自动握手** - 客户端和服务端自动处理 WebSocket 握手
- 🔄 **消息分片** - 支持大消息的分片传输和自动重组
- 📡 **广播功能** - 服务端支持向所有客户端广播消息
- 🚀 **自动 Ping/Pong** - 支持自动心跳保活

---

## 核心亮点

### 1️⃣ 清晰的职责分离 🎯

类库采用模块化设计，每个类只负责一项功能：

```vb
' cWebSocketServer - 服务器管理
Set m_Server = New cWebSocketServer
m_Server.Listen 8080

' cWebSocketClient - 客户端连接
Set m_Client = New cWebSocketClient
m_Client.Connect "ws://127.0.0.1:8080"

' cWebSocketFrame - 帧解析和构建（内部使用）
' cByteBuffer - 高效字节缓冲区（内部使用）
' mWebSocketUtils - 公共工具函数
```

---

### 2️⃣ 高效的缓冲区管理 📊

`cByteBuffer` 类采用预分配策略，按需增长，最大限度减少 ReDim Preserve 操作：

```vb
' 预分配 4KB，按 1.5 倍增长
Private Const INITIAL_CAPACITY As Long = 4096
Private Const GROWTH_FACTOR As Double = 1.5
```

---

### 3️⃣ 只读帧解析 🔍

`cWebSocketFrame.ParseHeader()` 不会修改输入数据，避免了缓冲区混乱问题：

```vb
' 1. 解析头部（只读）
If oFrame.ParseHeader(RecvBuffer) Then
    ' 2. 检查完整性
    If oFrame.IsCompleteFrame(RecvBuffer) Then
        ' 3. 提取并 unmask（消耗帧）
        baPayload = oFrame.ExtractPayload(RecvBuffer)
    End If
End If
```

---

### 4️⃣ 自动握手处理 🤝

#### 客户端握手

```vb
m_Client.Connect "ws://example.com:8080/chat"
' 自动生成 WebSocket Key
' 自动发送握手请求
' 自动验证服务端响应
```

#### 服务端握手

```vb
m_Server.Listen 8080
' 自动检测 HTTP Upgrade 请求
' 自动计算 Accept Key
' 自动发送 101 Switching Protocols 响应
```

---

### 5️⃣ 消息分片支持 📦

大消息自动分片传输，接收端自动重组：

```vb
' 发送大消息（自动分片）
m_Client.SendText LargeString

' 接收端自动重组完整消息
Private Sub m_Client_OnTextMessage(ByVal Message As String)
    ' Message 已是完整的重组后消息
    ProcessMessage Message
End Sub
```

---

### 6️⃣ 服务端广播功能 📢

```vb
' 向所有客户端广播
m_Server.BroadcastText "Hello everyone!"

' 广播但不包括发送者
m_Server.BroadcastText Message, ExcludeClientID

' 向特定客户端发送
m_Server.SendText ClientID, "Private message"

' 向所有客户端广播二进制数据
m_Server.BroadcastBinary baData
```

---

### 7️⃣ 事件驱动模型 ⚡

```vb
' 客户端事件
Event OnOpen()
Event OnClose(ByVal Code As WsCloseCode, ByVal Reason As String)
Event OnTextMessage(ByVal Message As String)
Event OnBinaryMessage(Data() As Byte)
Event OnError(ByVal Description As String)
Event OnPong(Data() As Byte)

' 服务端事件
Event OnStart(ByVal Port As Long)
Event OnStop()
Event OnClientConnect(ByVal ClientID As String, ByVal RemoteAddress As String, ByVal RemotePort As Long)
Event OnClientDisconnect(ByVal ClientID As String, ByVal Reason As String)
Event OnClientTextMessage(ByVal ClientID As String, ByVal Message As String)
Event OnClientBinaryMessage(ByVal ClientID As String, Data() As Byte)
Event OnError(ByVal Description As String)
```

---

## 架构设计

### 类层次结构

```
WebSocket 类库
├── cWebSocketServer (服务端)
│   ├── m_ListenSocket: cWinsock (监听 Socket)
│   ├── m_Clients: Collection (客户端集合)
│   └── m_FrameBuilder: cWebSocketFrame (帧构建器)
│
├── cWebSocketClient (客户端)
│   ├── m_Socket: cWinsock (连接 Socket)
│   ├── m_RecvBuffer: cByteBuffer (接收缓冲区)
│   └── m_FrameParser: cWebSocketFrame (帧解析器)
│
├── cWebSocketServerClient (服务端客户端)
│   ├── m_Socket: cWinsock
│   ├── m_RecvBuffer: cByteBuffer
│   └── m_FrameParser: cWebSocketFrame
│
├── cWebSocketFrame (帧解析/构建)
│   └── 帧头解析、unmask、帧构建
│
├── cByteBuffer (字节缓冲区)
│   └── 预分配、自动增长、Peek/Consume
│
└── mWebSocketUtils (工具模块)
    ├── UTF-8 编码/解码
    ├── Base64 编码
    ├── SHA1 哈希
    └── WebSocket Key 生成/验证
```

### 对象关系图

```
服务端对象 (cWebSocketServer)
├── Socket (监听套接字: cWinsock)
├── Clients 集合
│   ├── 客户端 1 (cWebSocketServerClient)
│   │   ├── Socket (独立连接: cWinsock)
│   │   ├── RecvBuffer (cByteBuffer)
│   │   └── FrameParser (cWebSocketFrame)
│   ├── 客户端 2 (cWebSocketServerClient)
│   │   ├── Socket (独立连接: cWinsock)
│   │   ├── RecvBuffer (cByteBuffer)
│   │   └── FrameParser (cWebSocketFrame)
│   └── ...
└── FrameBuilder (共享的帧构建器: cWebSocketFrame)
    └── 用于广播时构建帧（一次构建，多客户端发送）
```

### 连接流程

#### 客户端连接流程

```
1. Connect(ws://host:port)
2. 解析 URL (host, port, path)
3. 生成 WebSocket Key
4. cWinsock.Connect()
5. TCP 连接成功
6. 发送握手请求 (HTTP Upgrade)
7. 等待 101 响应
8. 验证 Accept Key
9. 触发 OnOpen 事件
10. 开始发送/接收消息
```

#### 服务端监听流程

```
1. Listen(port)
2. cWinsock.Listen()
3. 等待连接请求
4. 接受连接，创建 cWebSocketServerClient
5. 等待握手请求
6. 验证 WebSocket Key
7. 计算 Accept Key
8. 发送 101 响应
9. 触发 OnClientConnect 事件
10. 开始发送/接收消息
```

---

## 文档索引

|| 文档 | 描述 |
||------|------|
|| [类库总览](./library.md) | WebSocket 类库的整体介绍和设计理念 |
|| [客户端类](./client.md) | cWebSocketClient 类的详细说明 |
|| [服务端类](./server.md) | cWebSocketServer 类的详细说明 |
|| [帧解析类](./frame.md) | cWebSocketFrame 类的详细说明 |
|| [字节缓冲区类](./buffer.md) | cByteBuffer 类的详细说明 |
|| [工具模块](./utils.md) | mWebSocketUtils 模块的详细说明 |
|| [快速开始](./quickstart.md) | 快速入门示例 |
|| [进阶应用](./advanced.md) | 高级功能和最佳实践 |

---

## 依赖关系

| 组件                           | 描述                                                    |
| ------------------------------ | ------------------------------------------------------- |
| **cWinsock.cls**               | 位于 `add/` 目录下的底层 Socket 封装，提供 TCP 连接功能 |
| **cAsyncSocket.cls**           | 位于 `src/` 目录下的异步 Socket 实现，cWinsock 的基础   |
| **mWebSocketUtils.bas**        | WebSocket 工具模块                                      |
| **cByteBuffer.cls**            | 字节缓冲区类                                            |
| **cWebSocketFrame.cls**        | WebSocket 帧解析和构建类                                |
| **cWebSocketClient.cls**       | WebSocket 客户端类                                      |
| **cWebSocketServer.cls**       | WebSocket 服务端类                                      |
| **cWebSocketServerClient.cls** | 服务端客户端连接类                                      |

---

## 兼容性

- **VB6/VBA** - 完全兼容
- **Windows** - Windows XP 及以上版本
- **WebSocket 协议** - RFC 6455 (完全兼容)
- **SSL/TLS** - 暂不支持 (wss:// 需要额外实现)

---

## 许可证

基于 VbAsyncSocket (wqweto@gmail.com) 开发

---

## 作者

**WebSocket 类库**: woeoio@qq.com  
**基础 Socket 库**: woeoio@qq.com  
**原始 Socket 库**: wqweto@gmail.com

---

**最后更新**: 2026-01-10
