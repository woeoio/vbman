# cWinsock 开发计划

> 📋 项目开发进度跟踪文档

## 📖 目录

- [已完成功能](#已完成功能)
- [待开发功能](#待开发功能)
- [开发优先级](#开发优先级)
- [技术路线图](#技术路线图)

---

## ✅ 已完成功能

### 核心功能

#### 1. 基础网络通信
- [x] TCP 客户端/服务器通信
- [x] UDP 客户端/服务器通信
- [x] 异步 socket 封装（基于 VbAsyncSocket）
- [x] 连接状态管理
- [x] 异常处理机制

#### 2. 事件驱动模型
- [x] Connect 事件（连接成功）
- [x] CloseEvent 事件（连接关闭）
- [x] ConnectionRequest 事件（新连接请求）
- [x] DataArrival 事件（数据到达）
- [x] SendProgress 事件（发送进度）
- [x] SendComplete 事件（发送完成）
- [x] Error 事件（错误处理）

#### 3. 对象模型设计
- [x] 纯类实现，无需控件
- [x] 直接对象引用传递（无需索引查找）
- [x] 自动客户端集合管理（Clients 集合）
- [x] 服务器-客户端父子关系维护
- [x] 事件代理机制（服务器统一处理客户端事件）

#### 4. TCP 客户端事件代理
- [x] 自动通过父服务器触发客户端事件
- [x] 统一数据到达处理
- [x] 统一连接断开处理
- [x] 统一错误处理

#### 5. UDP 虚拟客户端管理
- [x] 为每个远程地址:端口创建虚拟客户端对象
- [x] 模拟 TCP 连接行为
- [x] 支持 ConnectionRequest 事件
- [x] 自动维护虚拟客户端集合

#### 6. 连接请求拦截
- [x] ConnectionRequest 事件中的 DisConnect 参数
- [x] 黑名单机制
- [x] 白名单机制
- [x] 端口范围限制
- [x] 自动断开和资源清理

#### 7. 文本编码支持
- [x] ScpAcp（系统默认代码页，GBK）
- [x] ScpUtf8（UTF-8 编码）
- [x] ScpUnicode（Unicode 宽字符）
- [x] 字符串/字节数组灵活转换

#### 8. 数据缓冲区管理
- [x] TCP 接收缓冲区（m_baRecvBuffer）
- [x] UDP 虚拟客户端缓冲区（UserData）
- [x] 部分读取支持
- [x] 自动缓冲剩余数据

#### 9. 远程地址解析
- [x] 域名自动解析
- [x] IP 和域名智能选择
- [x] RemoteHost/RemoteHostIP/RemotePort 属性

#### 10. 应用场景
- [x] TCP 服务器模式
- [x] TCP 客户端模式
- [x] UDP 服务器模式
- [x] UDP 客户端模式
- [x] 双向通信
- [x] 广播/多播基础支持

---

## 🚧 待开发功能

### 1. 数据封包协议（高优先级）

**功能描述**：
解决 TCP 分包和粘包问题，提供内置和自定义封包协议支持。

**实现细节**：

#### 1.1 内置协议

##### 字符分隔符协议
- **默认分隔符**：`\0`（空字符）
- **自定义分隔符**：用户可指定任意字符或字符串作为分隔符
- **常见分隔符**：
  - `\r\n` - 回车换行（HTTP、SMTP 等文本协议）
  - `\n` - 换行符
  - `\0` - 空字符（C 字符串风格）
  - `|` - 竖线（自定义协议）
  - 其他任意字符/字符串
- **适用场景**：文本协议、自定义消息格式

##### 定长协议
- 固定长度的消息块
- 适用于已知长度的协议
- 按固定长度切割数据

##### 长度头协议
- 消息头部包含数据长度信息
- 支持不同长度头格式（2字节/4字节整数）
- 自动解析完整消息

#### 1.2 自定义协议
- 用户提供回调函数
- 支持复杂的业务逻辑
- 灵活的封包/解包规则

#### 1.3 协议类设计

**统一接口设计**：
所有协议类都提供 `Encode`（封包）和 `Decode`（解包）两个接口函数，协议类内部自己缓存分片数据。

```vb
' 协议类型枚举
Public Enum PacketProtocol
    ppNone = 0                ' 不处理（默认）
    ppDelimiter = 1           ' 字符分隔符协议
    ppFixedLength = 2         ' 定长协议
    ppLengthHeader = 3        ' 长度头协议
    ppCustom = 4              ' 自定义协议
End Enum

' 协议类统一接口
' Encode：发送数据前调用，用于添加协议标记（封包）
Public Function Encode(ByRef baData() As Byte) As Byte()
    ' 参数：原始数据
    ' 返回值：添加协议标记后的完整数据
End Function

' Decode：接收数据后调用，用于解析完整消息（解包）
Public Function Decode(ByRef baData() As Byte) As Variant
    ' 参数：当前接收到的数据
    ' 返回值：
    '   - 如果数据不完整，返回 Empty（协议类内部缓存数据）
    '   - 如果数据完整，返回完整的消息数据（字节数组）
    '   - 如果有多条完整消息，返回字节数组集合
End Function
```

**协议类内部职责**：
- **Encode（封包）**：在数据末尾或头部添加协议标记
  - 字符分隔符协议：在数据末尾追加分隔符
  - 定长协议：填充或截断到固定长度
  - 长度头协议：在数据前添加长度头
- **Decode（解包）**：
  - 缓存分片数据（类内部维护接收缓冲区）
  - 判断数据是否完整
  - 数据完整后返回完整消息，并清空内部缓冲区
  - 自动处理分包和粘包

**每个客户端独立协议实例**：
- 多客户端场景下，每个 `cWinsock` 客户端对象持有独立的协议类实例
- 协议类的内部缓冲区相互隔离，互不干扰
- 支持不同客户端使用不同的协议类型和参数

#### 1.4 API 设计

```vb
' cWinsock 类中的协议配置
' =============================================

' 设置封包协议类型
Public Property Let PacketProtocol(ByVal eProtocol As PacketProtocol)
Public Property Get PacketProtocol() As PacketProtocol

' 字符分隔符协议参数
Public Property Let Delimiter(ByVal sDelimiter As String)
Public Property Get Delimiter() As String

' 定长协议参数
Public Property Let FixedLength(ByVal lLength As Long)
Public Property Get FixedLength() As Long

' 长度头协议参数
Public Property Let HeaderBytes(ByVal nBytes As Integer)  ' 2 或 4
Public Property Get HeaderBytes() As Integer
Public Property Let HeaderEndian(ByVal eEndian As EndianEnum)
Public Property Get HeaderEndian() As EndianEnum

' 自定义协议处理器对象
Public Property Let CustomPacketHandler(ByVal oHandler As Object)
Public Property Get CustomPacketHandler() As Object

' SendData 自动封包（发送时自动调用协议 Encode 方法）
Public Sub SendData(Data As Variant, Optional ByVal CodePage As ScpEnum = ScpAcp)

' GetData 自动解包（接收时通过协议类 Decode 方法处理）
Public Sub GetData(Data As Variant, Optional ByVal Type_ As VbVarType, Optional ByVal MaxLen As Long, Optional ByVal CodePage As ScpEnum = ScpAcp)
```

#### 1.5 架构设计

**cWinsock 类内部结构**：

```vb
' =============================================
' cWinsock 类定义
' =============================================
Public Class cWinsock
    ' 协议配置属性
    Private m_ePacketProtocol As PacketProtocol
    Private m_oPacketHandler As Object
    Private m_sDelimiter As String
    Private m_lFixedLength As Long
    Private m_nHeaderBytes As Integer
    
    ' 协议类实例（每个客户端对象独立持有）
    Private m_oPacketProtocol As IPacketProtocol
    
    ' 内部发送缓冲区
    Private m_baSendBuffer() As Byte
    
    ' 内部接收缓冲区
    Private m_baRecvBuffer() As Byte
    
    ' ... 其他成员 ...
End Class
```

**每个客户端独立协议实例的机制**：

```vb
' =============================================
' 协议实例创建和管理
' =============================================

' 在设置协议类型时，为当前客户端对象创建独立的协议实例
Public Property Let PacketProtocol(ByVal eProtocol As PacketProtocol)
    m_ePacketProtocol = eProtocol
    
    ' 为当前 cWinsock 对象创建独立的协议实例
    Set m_oPacketProtocol = CreateProtocolInstance(eProtocol)
End Property

' 为当前客户端创建协议实例的私有方法
Private Function CreateProtocolInstance(ByVal eProtocol As PacketProtocol) As IPacketProtocol
    Dim oProtocol As IPacketProtocol
    
    Select Case eProtocol
        Case ppDelimiter
            ' 创建字符分隔符协议实例（每个客户端独立）
            Set oProtocol = New DelimiterProtocol(IIf(LenB(m_sDelimiter) = 0, vbCrLf, m_sDelimiter))
            
        Case ppFixedLength
            ' 创建定长协议实例（每个客户端独立）
            Set oProtocol = New FixedLengthProtocol(m_lFixedLength)
            
        Case ppLengthHeader
            ' 创建长度头协议实例（每个客户端独立）
            Set oProtocol = New LengthHeaderProtocol(m_nHeaderBytes)
            
        Case ppCustom
            ' 使用自定义协议处理器
            Set oProtocol = m_oPacketHandler
            
        Case Else
            Set oProtocol = Nothing
    End Select
    
    Set CreateProtocolInstance = oProtocol
End Function

' 在服务器接受新连接时，新客户端对象自动继承服务器的协议配置
' 并创建自己独立的协议实例
Private Sub OnAcceptClient(ByRef oNewClient As cWinsock)
    ' 新客户端继承服务器协议配置
    oNewClient.PacketProtocol = Me.PacketProtocol
    oNewClient.Delimiter = Me.Delimiter
    oNewClient.FixedLength = Me.FixedLength
    oNewClient.HeaderBytes = Me.HeaderBytes
    
    ' 新客户端创建自己独立的协议实例（缓冲区隔离）
    ' oNewClient 内部会调用 CreateProtocolInstance 创建新实例
End Sub

' 客户端关闭时，清理自己的协议实例
Private Sub OnClose()
    If Not m_oPacketProtocol Is Nothing Then
        m_oPacketProtocol.Clear  ' 清空协议内部缓冲区
        Set m_oPacketProtocol = Nothing
    End If
End Sub
```

**多客户端并发场景下的实例关系**：

```
服务端 cWinsock 对象
├── 协议配置属性
│   ├── PacketProtocol = ppDelimiter
│   └── Delimiter = vbCrLf
│
└── Clients 集合
    ├── 客户端1 (cWinsock)
    │   ├── 协议实例1 (DelimiterProtocol)
    │   │   └── 接收缓冲区1（独立的）
    │   └── Socket 连接1
    │
    ├── 客户端2 (cWinsock)
    │   ├── 协议实例2 (DelimiterProtocol)
    │   │   └── 接收缓冲区2（独立的）
    │   └── Socket 连接2
    │
    └── 客户端3 (cWinsock)
        ├── 协议实例3 (DelimiterProtocol)
        │   └── 接收缓冲区3（独立的）
        └── Socket 连接3
```

**核心优势**：
- ✅ **状态隔离**：每个客户端的协议缓冲区完全独立，互不干扰
- ✅ **协议灵活**：不同客户端可以使用不同的协议类型和参数
- ✅ **线程安全**：VB6 单线程环境下，每个对象独立状态天然安全
- ✅ **内存管理**：客户端断开时自动清理协议实例和缓冲区
- ✅ **配置继承**：新客户端自动继承服务器协议配置，同时创建独立实例

#### 1.6 实现细节

**数据流程**：

1. **发送流程（SendData）**：
   - 将数据转换为字节数组
   - 调用**当前客户端对象**的协议实例 `Encode` 方法进行封包
   - 字符分隔符协议：在数据末尾追加分隔符
   - 定长协议：填充或截断到固定长度
   - 长度头协议：在数据前添加长度头
   - 自定义协议：调用自定义处理器的 `Encode` 方法
   - 发送封包后的完整数据

2. **接收流程（DataArrival）**：
   - 获取原始字节数据
   - 调用**当前客户端对象**的协议实例 `Decode` 方法进行解包
   - 协议实例内部将新数据合并到**自己的**接收缓冲区
   - 判断数据是否完整
   - 数据完整：返回完整消息（去除协议标记），触发事件
   - 数据不完整：返回 Empty，协议实例内部缓存数据
   - 下次数据到达时继续处理

3. **缓冲区管理**：
   - **每个客户端的协议实例独立维护接收缓冲区**
   - 数据完整后自动清空**自己的**缓冲区
   - 支持多条消息连续接收
   - 客户端关闭时清理**自己的**协议状态

4. **多客户端并发处理**：
   ```
   时间轴示例：
   
   T1: 客户端1 收到 "Hel" 
       → 调用客户端1的协议实例1.Decode("Hel")
       → 协议实例1缓冲区：["Hel"]
       → 返回 Empty（不完整）
   
   T2: 客户端2 收到 "Hi\0"
       → 调用客户端2的协议实例2.Decode("Hi\0")
       → 协议实例2缓冲区：["Hi\0"]
       → 解析出完整消息 "Hi"
       → 清空协议实例2缓冲区
       → 触发客户端2的 DataArrival 事件
   
   T3: 客户端1 收到 "lo\0"
       → 调用客户端1的协议实例1.Decode("lo\0")
       → 协议实例1缓冲区：["Hel"] + ["lo\0"] = ["Hello\0"]
       → 解析出完整消息 "Hello"
       → 清空协议实例1缓冲区
       → 触发客户端1的 DataArrival 事件
   
   T4: 客户端3 收到 "Test\0"
       → 调用客户端3的协议实例3.Decode("Test\0")
       → 协议实例3缓冲区：["Test\0"]
       → 解析出完整消息 "Test"
       → 清空协议实例3缓冲区
       → 触发客户端3的 DataArrival 事件
   ```
   - 每个客户端对象持有独立的协议实例
   - 每个协议实例维护独立的内部缓冲区
   - 协议实例之间互不干扰，独立处理各自的数据

#### 1.6 使用示例

**示例1：基本使用（字符分隔符协议）**
```vb
' 服务端配置
Private Sub Form_Load()
    Set m_oServer = New cWinsock
    m_oServer.PacketProtocol = ppDelimiter
    m_oServer.Delimiter = vbNullChar  ' 使用空字符分隔
    m_oServer.Listen 8080
End Sub

' 客户端发送
Client.SendData "Hello"  ' Encode 追加 vbNullChar

' 服务端接收（每个客户端独立的协议实例）
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData  ' Decode 自动解析，去除 vbNullChar
    Debug.Print Client.Tag & ": " & sData
End Sub
```

**示例2：多客户端并发演示**
```vb
' 服务端配置
Private Sub Form_Load()
    Set m_oServer = New cWinsock
    m_oServer.PacketProtocol = ppDelimiter
    m_oServer.Delimiter = vbCrLf  ' 使用回车换行分隔
    m_oServer.Listen 8080
End Sub

' 客户端1 发送（分片）
Client1.SendData "Hel"        ' 协议实例1 缓存："Hel"
Client1.SendData "lo" & vbCrLf  ' 协议实例1 解析："Hello"，清空缓存

' 客户端2 发送（完整）
Client2.SendData "World" & vbCrLf  ' 协议实例2 解析："World"，清空缓存

' 客户端3 发送（超长分片）
Client3.SendData "Go"            ' 协议实例3 缓存："Go"
Client3.SendData "od"            ' 协议实例3 缓存："Good"
Client3.SendData " Morn"         ' 协议实例3 缓存："Good Morn"
Client3.SendData "ing" & vbCrLf  ' 协议实例3 解析："Good Morning"，清空缓存

' 服务端 DataArrival
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData

    ' 每个客户端获取的是自己独立的协议实例解析的完整消息
    ' 协议实例之间互不干扰
    Debug.Print Client.Tag & ": " & sData

    ' 输出顺序取决于网络到达时间，可能是：
    ' 客户端2: World
    ' 客户端1: Hello
    ' 客户端3: Good Morning
End Sub
```

**示例3：不同客户端使用不同协议**
```vb
' 服务端接受连接时可以动态配置每个客户端的协议
Private Sub m_oServer_ConnectionRequest(Client As cWinsock, ByRef DisConnect As Boolean)
    ' 根据客户端 IP 或端口使用不同协议
    Select Case Client.RemotePort
        Case 9001
            ' 客户端1 使用字符分隔符协议
            Client.PacketProtocol = ppDelimiter
            Client.Delimiter = vbNullChar

        Case 9002
            ' 客户端2 使用定长协议
            Client.PacketProtocol = ppFixedLength
            Client.FixedLength = 10

        Case 9003
            ' 客户端3 使用长度头协议
            Client.PacketProtocol = ppLengthHeader
            Client.HeaderBytes = 4
    End Select
End Sub

' 服务端 DataArrival（每个客户端使用自己的协议）
Private Sub m_oServer_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim sData As String
    Client.GetData sData

    ' 根据 Tag 判断是哪个客户端，自动使用对应的协议解析
    Debug.Print Client.Tag & " (" & GetProtocolName(Client) & "): " & sData
End Sub
```

**示例4：粘包场景演示**
```vb
' 客户端快速连续发送（可能导致粘包）
Client.SendData "Msg1" & vbCrLf
Client.SendData "Msg2" & vbCrLf
Client.SendData "Msg3" & vbCrLf

' 服务端可能一次性收到："Msg1" & vbCrLf & "Msg2" & vbCrLf
' 协议实例的 Decode 方法会自动拆分：

' 第一次调用 Decode("Msg1" & vbCrLf & "Msg2" & vbCrLf)
' → 解析出 "Msg1"
' → 缓存："Msg2" & vbCrLf
' → 返回 "Msg1"

' 第二次调用 Decode("Msg3" & vbCrLf)
' → 合并缓存："Msg2" & vbCrLf & "Msg3" & vbCrLf
' → 解析出 "Msg2"
' → 缓存："Msg3" & vbCrLf
' → 返回 "Msg2"

' 第三次调用 Decode(空数据)
' → 解析出 "Msg3"
' → 清空缓存
' → 返回 "Msg3"
```

**示例5：分包场景演示**
```vb
' 客户端发送大数据（可能导致分包）
Dim sBigData As String
sBigData = String(10000, "A") & vbCrLf
Client.SendData sBigData

' 服务端分3次收到：
' DataArrival1: 4000 字节 "AAAA...A"（无分隔符）
' DataArrival2: 4000 字节 "AAAA...A"（无分隔符）
' DataArrival3: 2002 字节 "AAAA...A\r\n"（包含分隔符）

' 协议实例的 Decode 方法会自动拼接：

' 第一次调用 Decode(4000 字节)
' → 缓存：4000 字节
' → 无分隔符，返回 Empty

' 第二次调用 Decode(4000 字节)
' → 合并缓存：8000 字节
' → 无分隔符，返回 Empty

' 第三次调用 Decode(2002 字节)
' → 合并缓存：10002 字节
' → 找到 vbCrLf，解析出完整消息
' → 清空缓存
' → 返回 10000 字节 "AAAA...A"
```

#### 1.7 关键要点总结

**每个客户端独立协议实例的核心优势**：

| 特性 | 说明 |
|------|------|
| **状态隔离** | 每个客户端的协议类实例维护独立的接收缓冲区，互不干扰 |
| **协议灵活** | 不同客户端可以使用不同的协议类型和参数配置 |
| **并发安全** | VB6 单线程环境下，每个对象独立状态天然安全 |
| **内存管理** | 客户端断开时自动清理自己的协议实例和缓冲区 |
| **配置继承** | 新客户端自动继承服务器协议配置，同时创建独立实例 |
| **自动管理** | 无需手动管理协议实例生命周期，cWinsock 内部自动处理 |

**数据流转图**：

```
客户端1 Socket                      客户端2 Socket                      客户端3 Socket
      │                                   │                                   │
      ├─ 收到 "Hel"                       ├─ 收到 "Hi\0"                      ├─ 收到 "Tes"
      │                                   │                                   │
      ▼                                   ▼                                   ▼
协议实例1.Decode                   协议实例2.Decode                   协议实例3.Decode
      │                                   │                                   │
缓冲区1: ["Hel"]                    缓冲区2: ["Hi\0"]                   缓冲区3: ["Tes"]
      │                                   │                                   │
      │                                   ▼                                   ▼
      │                              解析出 "Hi"                         缓冲区3: ["Tes"]
      │                              清空缓冲区2                               │
      │                              触发DataArrival                           │
      │                                   │                                   │
      ├─ 收到 "lo\0"                       │                                   ├─ 收到 "t\0"
      │                                   │                                   │
      ▼                                   │                                   ▼
协议实例1.Decode                          │                           协议实例3.Decode
      │                                   │                                   │
缓冲区1: ["Hello\0"]                     │                           缓冲区3: ["Test\0"]
      │                                   │                                   │
      ▼                                   │                                   ▼
解析出 "Hello"                            │                          解析出 "Test"
清空缓冲区1                              │                          清空缓冲区3
触发DataArrival                           │                          触发DataArrival
```

**最佳实践建议**：

1. **服务器配置协议**：在 `Listen` 之前设置服务器的协议配置，所有新客户端自动继承
2. **动态配置协议**：在 `ConnectionRequest` 事件中根据客户端特征（IP、端口等）动态设置协议
3. **避免混用协议**：同一个客户端在生命周期内建议使用统一的协议类型
4. **客户端关闭清理**：cWinsock 内部自动清理协议实例，无需手动处理
5. **监控协议状态**：可通过协议类实例的 `HasPendingData()` 方法检查是否有未处理的分片数据

**典型应用场景**：

- **HTTP 服务器**：使用字符分隔符协议（`\r\n\r\n`）
- **二进制协议**：使用长度头协议（4字节整数）
- **固定格式协议**：使用定长协议
- **自定义协议**：实现自定义协议类，支持复杂业务逻辑

---

### 2. TCP 智能心跳机制（高优先级）

**功能描述**：
提供自动化的 TCP 连接保活和超时断开机制。

#### 2.1 服务端心跳检测

- **超时机制**
  - 默认超时时间：2 分钟
  - 检测内容：最后通信时间（包含数据收发）
  - 自动断开超时客户端
  - 触发 CloseEvent 事件

- **LastActivityTime 更新规则**
  - **重要**：每次对客户端收发数据后，必须立即重置 `LastActivityTime` 为当前时间
  - 这样客户端跳过心跳周期后，服务端轮询时才不会误判为超时
  - 更新时机：
    - 收到客户端数据时（DataArrival 事件）
    - 向客户端发送数据时（SendData/SendComplete）
    - 收到客户端心跳时
    - 任何有效的通信交互

- **实现方式**
  - 定时器轮询所有客户端的 LastActivityTime 属性
  - 超时则自动 Close 客户端
  - 记录日志（可选）

```vb
' 服务端配置
Public Property Let HeartbeatTimeout(ByVal lSeconds As Long)
Public Property Get HeartbeatTimeout() As Long

' 自动开启/关闭
Public Property Let AutoHeartbeat(ByVal bEnable As Boolean)

' 事件：客户端超时断开
Public Event ClientTimeout(Client As cWinsock)
```

- **LastActivityTime 更新示例**
```vb
' 在 DataArrival 事件中更新
Private Sub Server_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    ' cWinsock 内部自动执行
    Client.LastActivityTime = Now  ' 或 GetTickCount 等时间戳
End Sub

' 在 SendData 方法中更新
Public Sub SendData(Data As Variant, Optional ByVal CodePage As ScpEnum = ScpAcp)
    ' ... 发送数据 ...
    ' cWinsock 内部自动执行
    If IsClient() Then
        m_oParentServer.LastActivityTime = Now
    Else
        m_lLastActivityTime = Now
    End If
End Sub
```

#### 2.2 客户端心跳保活
- **自动 Ping 机制**
  - 默认发送间隔：50 秒
  - 数据内容：1 字节（自定义）
  - 智能跳过：如果最近有数据收发，则跳过当前周期

- **实现方式**
  - 定时器控制发送间隔
  - 检查 LastActivityTime 判断是否需要发送
  - SendData 发送心跳包

```vb
' 客户端配置
Public Property Let HeartbeatInterval(ByVal lSeconds As Long)
Public Property Get HeartbeatInterval() As Long

' 心跳包内容
Public Property Let HeartbeatData(ByVal vData As Variant)
Public Property Get HeartbeatData() As Variant

' 自动开启/关闭
Public Property Let AutoHeartbeat(ByVal bEnable As Boolean)

' 事件：心跳发送
Public Event HeartbeatSent()
```

#### 2.3 心跳状态监控
- 每个客户端维护心跳相关信息
- 属性：
  - `LastHeartbeatTime` - 最后心跳时间
  - `LastActivityTime` - 最后通信时间
  - `HeartbeatCount` - 心跳计数
  - `IsAlive` - 是否存活

---

### 3. GetData 增强方法（中优先级）

**功能描述**：
提供更便捷的数据获取方法，支持不同格式输出。

#### 3.1 设计原则

**重要区别**：
- 原始 `GetData` 使用传址参数输出数据
- 增强方法通过**返回值**方式输出数据
- 方便写成一行代码，如：`Dim Data As String: Data = GetDataText()`

#### 3.2 新增方法

```vb
' 获取文本字符串（使用默认编码 ACP）
Public Function GetDataText(Optional ByVal MaxLen As Long = -1) As String

' 获取文本字符串（指定编码）
Public Function GetDataTextEx(Optional ByVal CodePage As ScpEnum = ScpAcp, Optional ByVal MaxLen As Long = -1) As String

' 获取 UTF-8 字符串
Public Function GetDataTextUTF8(Optional ByVal MaxLen As Long = -1) As String

' 获取 Unicode 字符串
Public Function GetDataTextUnicode(Optional ByVal MaxLen As Long = -1) As String

' 获取十六进制字符串（空格分隔）
Public Function GetDataHex(Optional ByVal MaxLen As Long = -1) As String

' 获取字节数组
Public Function GetDataByteArray(Optional ByVal MaxLen As Long = -1) As Byte()
```

#### 3.3 使用示例

**一行代码风格**：
```vb
' 直接赋值，无需声明后传址
Debug.Print GetDataText()                              ' 输出文本
Debug.Print GetDataHex()                               ' 输出：48 65 6C 6C 6F
Debug.Print GetDataTextUTF8()                          ' 输出 UTF-8 文本

' 判断数据
If GetDataText() = "Hello" Then
    Debug.Print "收到 Hello"
End If

' 处理数据
Dim sReply As String
sReply = ProcessData(GetDataText())

' 赋值给变量
Dim sData As String: sData = GetDataText()
Dim baData() As Byte: baData = GetDataByteArray()
```

**与原始 GetData 对比**：
```vb
' 原始方式（需要声明后传址）
Dim sData As String
Client.GetData sData
Debug.Print sData

' 新方式（直接返回）
Debug.Print Client.GetDataText()

' 原始方式（字节数组）
Dim baData() As Byte
Client.GetData baData

' 新方式（直接返回）
Dim baData() As Byte
baData = Client.GetDataByteArray()
```

**带 MaxLen 参数**：
```vb
' 只读取前 100 字节
Debug.Print Client.GetDataText(100)

' 获取前 50 字节的十六进制表示
Debug.Print Client.GetDataHex(50)
```

#### 3.4 实现细节

**内部实现**：
```vb
Public Function GetDataText(Optional ByVal MaxLen As Long = -1) As String
    Dim sData As String
    GetData sData, vbString, MaxLen, ScpAcp  ' 调用原始 GetData
    GetDataText = sData                      ' 通过返回值返回
End Function

Public Function GetDataHex(Optional ByVal MaxLen As Long = -1) As String
    Dim baData() As Byte
    Dim i As Long
    Dim sHex As String
    
    ' 获取字节数组
    GetData baData, vbByte + vbArray, MaxLen
    
    ' 转换为十六进制字符串
    For i = LBound(baData) To UBound(baData)
        sHex = sHex & Right$("0" & Hex$(baData(i)), 2) & " "
    Next
    
    ' 去除末尾空格
    If Len(sHex) > 0 Then
        GetDataHex = Left$(sHex, Len(sHex) - 1)
    End If
End Function
```

**缓冲区一致性**：
- 所有增强方法内部调用原始 `GetData`
- 保持与原始方法相同的缓冲区行为
- 部分读取后，剩余数据仍保留在内部缓冲区

---

### 4. 性能优化（中优先级）

#### 4.1 批量发送
```vb
' 批量发送多条消息，减少系统调用
Public Sub SendBatch(vData As Variant)
```

#### 4.2 数据压缩
- 可选的压缩算法支持
- 大数据自动压缩
- 透明压缩/解压

#### 4.3 连接池
- 复用 TCP 连接
- 减少 Connect 开销
- 自动负载均衡

---

### 5. 高级功能（低优先级）

#### 5.1 SSL/TLS 加密
- 支持 HTTPS/WSS
- 证书验证
- 安全握手

#### 5.2 WebSocket 协议
- 完整的 WebSocket 支持
- 握手和帧处理
- 自动 Ping/Pong

#### 5.3 断线重连
- 自动重连机制
- 指数退避算法
- 最大重试次数

#### 5.4 限流控制
- 发送速率限制
- 接收速率限制
- 流量统计

---

## 📊 开发优先级

### P0 - 核心功能（必须实现）
1. **数据封包协议**
   - 解决最常见的数据分片问题
   - 提升开发效率
   - 减少错误率

2. **TCP 智能心跳**
   - 保证连接稳定性
   - 及时清理僵尸连接
   - 适用于生产环境

### P1 - 增强功能（重要）
1. **GetData 增强方法**
   - 提升开发体验
   - 减少代码量
   - 降低出错概率

### P2 - 优化功能（可选）
1. **性能优化**
   - 批量发送
   - 数据压缩
   - 连接池

### P3 - 高级功能（长期规划）
1. **SSL/TLS 加密**
2. **WebSocket 协议**
3. **断线重连**
4. **限流控制**

---

## 🗺️ 技术路线图

### 阶段一：封包协议实现（预计 3-5 天）
- Day 1-2: 设计协议接口和数据结构
- Day 3-4: 实现内置协议（CRLF、定长、长度头）
- Day 5: 实现自定义协议接口和测试

### 阶段二：心跳机制实现（预计 2-3 天）
- Day 1-2: 实现服务端超时检测
- Day 3: 实现客户端心跳保活

### 阶段三：GetData 增强（预计 1 天）
- Day 1: 实现 4 个新方法和测试

### 阶段四：性能优化（预计 2-3 天）
- Day 1: 批量发送实现
- Day 2: 数据压缩实现
- Day 3: 连接池设计

### 阶段五：高级功能（长期）
- 根据用户需求和反馈逐步实现

---

## 📝 使用建议

### 开发期间
- 保持向后兼容性
- 不影响现有功能
- 提供充分的单元测试
- 更新文档和示例

### 发布策略
- 分阶段发布
- 收集用户反馈
- 持续优化改进

---

## 🔗 相关文档

- [总览](./overview.md) - 项目概览和核心特性
- [属性参考](./properties.md) - 属性详细说明
- [方法参考](./methods.md) - 方法详细说明
- [TCP编程](./tcp.md) - TCP 开发指南

---

**最后更新**: 2026-01-10
