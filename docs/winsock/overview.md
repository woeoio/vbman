# cWinsock 类开发文档

> 🚀 **cWinsock** - 简化版 VB6 Winsock 封装库，由 woeoio@qq.com 基于 VbAsyncSocket（作者：wqweto@gmail.com）开发

## 📖 目录

- [概述](#概述)
- [核心亮点](#核心亮点)
- [与原生 Winsock 控件的对比](#与原生-winsock-控件的对比)
- [快速开始](#快速开始)
- [架构设计](#架构设计)
- [文档索引](#文档索引)

---

## 概述

`cWinsock` 是一个为 VB6 设计的轻量级网络通信类，提供了与经典 Winsock 控件相似的事件驱动编程模型，但具有更简洁的 API 和更强大的功能。

### ✨ 主要特性

- 🔌 **纯类实现** - 无需控件，直接使用对象编程
- 🎯 **直接对象引用** - 事件参数直接传递客户端对象，无需通过索引查找
- 🌐 **双协议支持** - 同时支持 TCP 和 UDP 通信
- 🏢 **自动客户端管理** - 服务器模式自动管理所有连接的客户端
- 📦 **智能数据编码** - 支持多种文本编码（GBK/ACP、UTF-8、Unicode）
- 🛡️ **连接拦截能力** - 通过 `ConnectionRequest` 事件的黑名单/白名单机制
- 🔄 **事件代理机制** - 服务器客户端数据统一通过服务器事件触发
- 💾 **灵活数据类型** - 支持字符串和字节数组两种数据格式
- 🚀 **计划功能** - 数据封包协议、TCP 智能心跳、GetData 增强方法

---

## 核心亮点

### 1️⃣ 直接对象引用事件模型 🔗

**传统 Winsock 控件的问题：**
```vb
' 需要通过索引管理客户端
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Integer
    ' 找到空闲的索引或动态加载控件...
End Sub

' 处理数据时需要知道是哪个客户端
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Winsock1(Index).GetData strData
End Sub
```

**cWinsock 的优雅解决方案：**
```vb
' 事件直接传递客户端对象！
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 直接操作 Client 对象，无需索引
    Debug.Print "新客户端: " & Client.RemoteHostIP
    
    ' 拒绝黑名单 IP
    If IsBlacklisted(Client.RemoteHostIP) Then
        DisConnect = True
    End If
End Sub

' 数据事件也是直接传递客户端对象
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    ' 直接从 Client 对象读取数据，无需索引查找
End Sub
```

---

### 2️⃣ 智能的 TCP 客户端事件代理 📡

**问题场景：** 服务器接受新连接后创建的客户端对象，其数据接收事件无法被宿主订阅。

**cWinsock 的解决方案：** 自动通过父服务器对象触发事件

```vb
' 在服务器对象的 DataArrival 事件中
' 可以接收到所有客户端的数据！
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' Client 参数就是具体的客户端对象
    ' 可以直接向该客户端回复
    Client.SendData "Echo: " & sData
End Sub
```

**工作原理：**
1. 服务器接受新连接，创建独立的客户端 socket 对象
2. 客户端接收数据后，通过父服务器的 `RaiseDataArrivalEvent` 方法触发事件
3. 宿主只需订阅服务器对象的事件，即可处理所有客户端的数据

---

### 3️⃣ UDP 服务器虚拟客户端管理 🎭

UDP 是无连接协议，但 `cWinsock` 为每个不同的远程地址:端口创建虚拟客户端对象，模拟连接行为：

```vb
' UDP 服务器模式
Private Sub m_oUdp_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 每个首次通信的远程地址:端口组合
    ' 都会自动创建一个虚拟 Client 对象
    Debug.Print "UDP 客户端: " & Client.RemoteHostIP & ":" & Client.RemotePort
End Sub

Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    
    ' 可以向特定的虚拟客户端回复
    ' cWinsock 会自动使用正确的目标地址:端口
    Client.SendData "Reply: " & sData
End Sub
```

---

### 4️⃣ 连接请求拦截机制 🚦

在 `ConnectionRequest` 事件中通过 `DisConnect` 参数实现连接拦截：

```vb
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 黑名单检查
    If IsInBlacklist(Client.RemoteHostIP) Then
        Debug.Print "拒绝黑名单 IP: " & Client.RemoteHostIP
        DisConnect = True  ' 设置为 True，自动断开并清理资源
        Exit Sub
    End If
    
    ' 端口范围限制
    If Client.RemotePort < 1024 Then
        Debug.Print "拒绝特权端口连接: " & Client.RemotePort
        DisConnect = True
        Exit Sub
    End If
    
    ' 白名单模式
    If m_bWhitelistMode And Not IsInWhitelist(Client.RemoteHostIP) Then
        Debug.Print "不在白名单中，拒绝连接"
        DisConnect = True
        Exit Sub
    End If
    
    ' DisConnect 保持 False，接受连接
    Debug.Print "接受连接: " & Client.RemoteHostIP & ":" & Client.RemotePort
End Sub
```

---

### 5️⃣ 灵活的文本编码支持 🔤

支持多种编码方式，适应不同场景：

```vb
' 默认使用 ACP/GBK 编码（与 VB6 兼容）
Client.SendData "中文测试"
Client.GetData sData  ' 默认 ACP

' 使用 UTF-8 编码（推荐用于网络传输）
Client.SendData "中文测试", ucsScpUtf8
Client.GetData sData, , , ucsScpUtf8

' 使用 Unicode（不转换，保持宽字符）
Client.SendData "中文测试", ScpUnicode
Client.GetData sData, , , ScpUnicode

' 发送字节数组（不涉及编码）
Dim baData() As Byte
baData = GetByteArray()
Client.SendData baData
```

**编码枚举：**
- `ScpAcp` (0) - 系统默认代码页（中文 Windows 上为 GBK）
- `ScpUtf8` (65001) - UTF-8 编码
- `ScpUnicode` (-1) - Unicode，不进行编码转换

---

### 6️⃣ 自动客户端集合管理 📚

服务器模式下，自动维护所有连接的客户端：

```vb
' 启动服务器时自动初始化客户端集合
m_oServer.Listen 8080

' 遍历所有客户端
Dim oClient As cWinsock
For Each oClient In m_oServer.Clients
    Debug.Print "客户端: " & oClient.Tag & " - " & oClient.RemoteHostIP
Next

' 获取客户端数量
Debug.Print "当前连接数: " & m_oServer.ClientCount

' 手动移除客户端（通常由 CloseEvent 自动处理）
m_oServer.RemoveClient oClient
```

---

### 7️⃣ 智能的远程地址解析 🌐

UDP 服务器模式支持域名解析：

```vb
' 设置远程地址（可以是 IP 或域名）
m_oUdp.RemoteHost = "example.com"
m_oUdp.RemotePort = 8888

' 发送时自动解析域名
m_oUdp.SendData "Hello"
```

**内部逻辑：**
```vb
' SendData 方法中的智能选择
If LenB(m_sRemoteHostIP) <> 0 Then
    ' 如果已解析的 IP 存在，优先使用
    m_oSocket.SendText Data, m_sRemoteHostIP, m_lRemotePort, CodePage
ElseIf LenB(m_sRemoteHost) <> 0 Then
    ' 否则使用主机名，底层自动解析域名
    m_oSocket.SendText Data, m_sRemoteHost, m_lRemotePort, CodePage
End If
```

---

### 8️⃣ 数据缓冲区管理 📊

内置数据缓冲区，支持部分读取：

```vb
' 接收数据时只读取前 100 字节
Dim sPartial As String
Client.GetData sPartial, vbString, 100

' 剩余数据自动保存在内部缓冲区
' 下次读取时会继续返回剩余数据
```

**内部缓冲区机制：**
- TCP 和客户端模式：使用 `m_baRecvBuffer` 私有成员
- UDP 服务器虚拟客户端：使用 `UserData` 属性临时存储

---

### 9️⃣ 待开发功能亮点（计划中） 🚀

#### 数据封包协议 📦

**问题场景**：TCP 是流式协议，存在数据分包和粘包问题

```vb
' 发送方连续发送
Client.SendData "Hello"
Client.SendData "World"

' 接收方可能收到
"HelloWorld"  ' 粘包
"Hel"          ' 分包
"loWorld"
```

**计划实现**：
- **字符分隔符协议**
  - 默认分隔符：`\0`（空字符）
  - 支持自定义任意分隔符（如 `\r\n`、`|` 等）
  - 适用于文本协议
- **定长协议** - 适用于固定长度消息
- **长度头协议** - 适用于二进制协议
- **自定义协议** - 支持用户回调函数
- **统一协议接口** - 所有协议类提供统一的 `Encode`（封包）和 `Decode`（解包）函数，内部自动缓存分片数据
- **每个客户端独立协议实例** - 多客户端场景下，每个客户端持有独立的协议实例，缓冲区相互隔离
- **自动处理** - 设置后自动封包/解包

**预期 API**：
```vb
' 设置字符分隔符协议（使用默认 \r\n）
Server.PacketProtocol = ppDelimiter

' 自定义分隔符为空字符
Server.Delimiter = vbNullChar

' 发送自动封包
Client.SendData "Hello World"  ' 自动追加分隔符

' 接收自动解包（协议类内部 Decode 函数处理）
Private Sub Server_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData  ' 自动组装完整消息
End Sub
```

**详细设计**：参见 [开发计划](./roadmap.md#1-数据封包协议高优先级)

---

#### TCP 智能心跳 💓

**问题场景**：TCP 连接可能因网络问题静默断开，需要保活机制

**服务端功能**：
- 定时轮询客户端最后通信时间
- 2 分钟无通信自动断开
- 触发 `ClientTimeout` 事件
- 防止僵尸连接占用资源
- **重要**：每次收发数据后立即重置客户端的 `LastActivityTime`，确保心跳跳过周期后不会被误判为超时

**客户端功能**：
- 每 50 秒自动发送心跳包（1 字节）
- 智能跳过：如有数据收发则跳过当前周期
- 保持连接活跃，防止超时断开

**预期 API**：
```vb
' 服务端配置
Server.HeartbeatTimeout = 120  ' 2 分钟超时
Server.AutoHeartbeat = True    ' 自动开启

' 客户端配置
Client.HeartbeatInterval = 50   ' 50 秒间隔
Client.HeartbeatData = &H0     ' 心跳包内容
Client.AutoHeartbeat = True

' 事件
Private Sub Server_ClientTimeout(Client As cWinsock)
    Debug.Print "客户端超时: " & Client.RemoteHostIP
End Sub

Private Sub Client_HeartbeatSent()
    Debug.Print "心跳已发送"
End Sub
```

**详细设计**：参见 [开发计划](./roadmap.md#2-tcp-智能心跳机制高优先级)

---

#### GetData 增强方法 🎯

**问题场景**：获取数据需要手动转换格式，代码繁琐

**设计原则**：
- 与原始 `GetData` 不同，新方法通过**返回值**输出数据，而非传址参数
- 支持一行代码风格：`Dim Data As String: Data = GetDataText()`

**计划新增方法**：
```vb
' 直接返回文本
Debug.Print Client.GetDataText()                    ' 一行代码
Debug.Print Client.GetDataTextUTF8()                 ' UTF-8 文本
Debug.Print Client.GetDataTextUnicode()              ' Unicode 文本

' 直接返回十六进制
Debug.Print Client.GetDataHex()                      ' "48 65 6C 6C 6F"

' 直接返回字节数组
Dim baData() As Byte
baData = Client.GetDataByteArray()

' 条件判断
If Client.GetDataText() = "Hello" Then
    Debug.Print "收到 Hello"
End If

' 函数调用处理
Dim sReply As String
sReply = ProcessData(Client.GetDataText())
```

**详细设计**：参见 [开发计划](./roadmap.md#3-getdata-增强方法中优先级)

---

## 与原生 Winsock 控件的对比

| 特性 | 原生 Winsock 控件 | cWinsock 类 |
|------|-------------------|-------------|
| **对象模型** | 控件数组，通过索引管理 | 纯类对象，直接引用 |
| **事件参数** | 传递索引，需反查对象 | 直接传递客户端对象 |
| **客户端管理** | 需要手动维护索引和控件 | 自动管理 Clients 集合 |
| **UDP 服务器** | 无连接，无客户端概念 | 虚拟客户端对象 |
| **连接拦截** | 需要在 Accept 后手动关闭 | 事件参数控制，自动清理 |
| **编码支持** | 固定编码 | 多种编码可选 |
| **数据类型** | 字符串/字节数组 | 字符串/字节数组 + 灵活转换 |
| **事件统一** | 每个客户端独立事件 | 服务器统一触发所有客户端事件 |
| **资源管理** | 需要手动 Unload 控件 | 自动清理和垃圾回收 |

---

## 快速开始

### TCP 客户端示例

```vb
Private WithEvents m_oClient As cWinsock

Private Sub Form_Load()
    Set m_oClient = New cWinsock
    m_oClient.Protocol = sckTCPProtocol
    m_oClient.Connect "127.0.0.1", 8080
End Sub

Private Sub m_oClient_Connect(Client As cWinsock)
    Debug.Print "已连接到服务器"
    Client.SendData "Hello, Server!"
End Sub

Private Sub m_oClient_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    Debug.Print "收到数据: " & sData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_oClient.Close_
End Sub
```

### TCP 服务器示例

```vb
Private WithEvents m_oServer As cWinsock

Private Sub Form_Load()
    Set m_oServer = New cWinsock
    m_oServer.Protocol = sckTCPProtocol
    m_oServer.Listen 8080
End Sub

Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    Debug.Print "新客户端连接: " & Client.RemoteHostIP
    ' DisConnect = False 表示接受连接
End Sub

Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    Debug.Print "来自客户端 " & Client.Tag & " 的数据: " & sData
    
    ' 回显
    Client.SendData "Echo: " & sData
End Sub

Private Sub m_oServer_CloseEvent(Client As cWinsock)
    Debug.Print "客户端断开: " & Client.Tag
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_oServer.Close_
End Sub
```

### UDP 通信示例

```vb
Private WithEvents m_oUdp As cWinsock

Private Sub Form_Load()
    Set m_oUdp = New cWinsock
    m_oUdp.Protocol = sckUDPProtocol
    m_oUdp.Bind 8888
End Sub

Private Sub cmdSend_Click()
    m_oUdp.RemoteHost = "127.0.0.1"
    m_oUdp.RemotePort = 9999
    m_oUdp.SendData "Hello, UDP!"
End Sub

Private Sub m_oUdp_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData
    Debug.Print "收到 UDP 数据 (" & Client.RemoteHostIP & ":" & Client.RemotePort & "): " & sData
End Sub
```

---

## 架构设计

### 类层次结构

```
cWinsock (公开类)
    ├── m_oSocket: cAsyncSocket (内部封装)
    ├── m_cClients: Collection (客户端集合)
    ├── m_oParentServer: cWinsock (父服务器引用，仅客户端)
    └── 事件：Connect, CloseEvent, ConnectionRequest, DataArrival, SendProgress, SendComplete, Error
```

### 对象关系图

```
服务器对象
├── Socket (监听套接字)
├── Clients 集合
│   ├── 客户端对象 1 (cWinsock)
│   │   ├── Socket (独立连接)
│   │   └── ParentServer → 服务器对象
│   ├── 客户端对象 2 (cWinsock)
│   │   ├── Socket (独立连接)
│   │   └── ParentServer → 服务器对象
│   └── ...
└── 事件处理器
    └── 所有客户端数据通过此触发
```

### 状态机

```
sckClosed (0)
    ├─ Connect() → sckResolvingHost → sckHostResolved → sckConnecting → sckConnected (7)
    ├─ Listen() → sckListening (2)
    └─ Bind() → sckOpen (1)

sckListening (2)
    └─ OnAccept → 创建客户端 → sckConnected

sckConnected (7)
    └─ OnClose → sckClosed

Error → sckError (9)
```

---

## 文档索引

| 文档 | 描述 |
|------|------|
| [事件详解](./events.md) | 所有事件的详细说明和使用示例 |
| [属性参考](./properties.md) | 所有属性的说明、类型和用途 |
| [方法参考](./methods.md) | 所有方法的参数、返回值和使用示例 |
| [编码指南](./encoding.md) | 文本编码的使用说明和最佳实践 |
| [TCP编程](./tcp.md) | TCP 客户端和服务器编程指南 |
| [UDP编程](./udp.md) | UDP 通信编程指南 |
| [最佳实践](./best-practices.md) | 常见场景的解决方案和性能优化建议 |
| [开发计划](./roadmap.md) | 项目开发进度跟踪和未来功能规划 |

---

## 许可证

基于 VbAsyncSocket (wqweto@gmail.com) 开发

---

## 作者

**cWinsock**: woeoio@qq.com  
**VbAsyncSocket**: wqweto@gmail.com

---

**最后更新**: 2026-01-09
