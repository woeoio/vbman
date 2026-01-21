# NTP 模块

基于 cWinsock 对象实现的 NTP (Network Time Protocol) 客户端和服务端。

## 模块结构

### 类文件

- **cNTPPacket.cls** - NTP 协议数据包封装类
  - 封装 NTP v4 协议数据包的编码和解码
  - 遵循 RFC 5905 标准
  - 支持客户端请求包和服务器响应包的创建

- **cNTPClient.cls** - NTP 客户端
  - 同步/异步时间查询
  - 自动超时处理
  - 支持多服务器查询
  - 提供时间偏移和往返延迟计算
  - 可选的系统时间同步功能

- **cNTPServer.cls** - NTP 服务端
  - 标准 NTP 协议实现 (RFC 5905)
  - 支持多客户端并发
  - 可配置 Stratum 层级
  - 事件驱动架构
  - 完全基于 UDP 协议

- **cNTPTimeProvider.cls** - 时间提供者 (Strategy Pattern)
  - 抽象时间源接口
  - 支持多种时间源：系统时间、外部 NTP 服务器、自定义时间
  - 便于单元测试和自定义时间源

### 演示文件

- **mNTPDemo.bas** - 演示代码和实用函数

## 设计原则

### 职责分离

1. **cNTPPacket** - 只负责数据包的编码和解码
2. **cNTPClient** - 只负责客户端的通信逻辑
3. **cNTPServer** - 只负责服务器的请求处理
4. **cNTPTimeProvider** - 只负责时间源的获取

### 抽象层次

```
应用层 (Demo)
    ↓
业务层 (cNTPClient, cNTPServer)
    ↓
协议层 (cNTPPacket)
    ↓
传输层 (cWinsock)
```

### 设计模式

- **Strategy Pattern**: cNTPTimeProvider 抽象不同的时间源策略
- **Factory Method**: CreateClientRequest/CreateServerResponse 工厂方法
- **Event-Driven**: 客户端和服务端都采用事件驱动架构

## 使用示例

### 客户端使用

#### 同步获取服务器时间

```vb
Dim oClient As New cNTPClient
Dim dtServerTime As Date

dtServerTime = oClient.GetServerTime("pool.ntp.org", 123)
Debug.Print "Server Time: " & Format(dtServerTime, "yyyy-mm-dd hh:mm:ss")
```

#### 获取时间偏移量

```vb
Dim oClient As New cNTPClient
Dim dOffset As Double

dOffset = oClient.GetTimeOffset("pool.ntp.org", 123)

If dOffset > 0 Then
    Debug.Print "Local time is " & dOffset & " seconds ahead of server"
Else
    Debug.Print "Local time is " & Abs(dOffset) & " seconds behind server"
End If
```

#### 异步查询 (需要 WithEvents)

```vb
' 在类模块中实现:
Private WithEvents m_oClient As cNTPClient

Private Sub TestAsync()
    Set m_oClient = New cNTPClient
    m_oClient.RequestTimeAsync "pool.ntp.org", 123
End Sub

Private Sub m_oClient_TimeSyncComplete(ByVal ServerHost As String, _
                                      ByVal ServerTime As Date, _
                                      ByVal Offset As Double, _
                                      ByVal RoundTripDelay As Double, _
                                      ByVal Success As Boolean, _
                                      ByVal ErrorMsg As String)
    If Success Then
        Debug.Print "Async sync completed: " & Format(ServerTime, "hh:mm:ss")
    Else
        Debug.Print "Error: " & ErrorMsg
    End If
End Sub
```

### 服务端使用

#### 基本 NTP 服务器

```vb
Dim WithEvents oServer As cNTPServer

Sub StartServer()
    Set oServer = New cNTPServer
    oServer.Stratum = 2
    oServer.ReferenceID = "LOCL"

    oServer.Start 123
End Sub

Private Sub oServer_ServerStarted(ByVal Port As Long)
    Debug.Print "Server started on port " & Port
End Sub

Private Sub oServer_RequestReceived(ByVal RemoteIP As String, ByVal RemotePort As Long)
    Debug.Print "Request from: " & RemoteIP & ":" & RemotePort
End Sub

Private Sub oServer_ResponseSent(ByVal RemoteIP As String, ByVal RemotePort As Long)
    Debug.Print "Response sent to: " & RemoteIP & ":" & RemotePort
End Sub
```

#### 自定义时间源

```vb
Dim oServer As New cNTPServer
Dim oProvider As New cNTPTimeProvider

' 配置时间提供者
oProvider.TimeSource = TimeSource_Custom
oProvider.UseCustomTime = True
oProvider.CustomTime = DateSerial(2024, 1, 1) + TimeSerial(12, 0, 0)

Set oServer.TimeProvider = oProvider
oServer.Start 1230
```

#### 外部时间源 (作为二级服务器)

```vb
Dim oServer As New cNTPServer
Dim oProvider As New cNTPTimeProvider

oProvider.TimeSource = TimeSource_External
oProvider.ExternalServer = "pool.ntp.org"
oProvider.ExternalPort = 123
oProvider.SyncExternalSource()

Set oServer.TimeProvider = oProvider
oServer.Stratum = 3
oServer.Start 1230
```

## API 参考

### cNTPPacket

#### 属性

| 属性               | 类型    | 说明                 |
| ------------------ | ------- | -------------------- |
| LI                 | Byte    | Leap Indicator (0-3) |
| VN                 | Byte    | Version Number (0-7) |
| Mode               | NTPMode | Mode (0-7)           |
| Stratum            | Byte    | 层级                 |
| Poll               | Byte    | 轮询间隔             |
| Precision          | Byte    | 精度                 |
| RootDelay          | Double  | 根延迟 (秒)          |
| RootDispersion     | Double  | 根分散 (秒)          |
| ReferenceID        | String  | 参考标识符           |
| ReferenceTimestamp | Date    | 参考时间戳           |
| OriginTimestamp    | Date    | 源时间戳             |
| ReceiveTimestamp   | Date    | 接收时间戳           |
| TransmitTimestamp  | Date    | 发送时间戳           |

#### 方法

- `FromBytes(Data() As Byte) As Boolean` - 从字节数组解析
- `ToBytes() As Byte()` - 编码为字节数组
- `CreateClientRequest() As cNTPPacket` - 创建客户端请求包
- `CreateServerResponse(RequestPacket As cNTPPacket) As cNTPPacket` - 创建服务器响应包
- `Clear()` - 清空数据包

### cNTPClient

#### 属性

| 属性       | 类型     | 说明                   |
| ---------- | -------- | ---------------------- |
| ServerHost | String   | 服务器地址             |
| ServerPort | Long     | 服务器端口 (默认: 123) |
| Timeout    | Double   | 超时时间(秒, 默认: 5)  |
| AsyncMode  | Boolean  | 是否异步模式           |
| Socket     | cWinsock | Socket 对象            |

#### 方法

- `GetServerTime(Optional Server, Optional Port) As Date` - 获取服务器时间
- `GetTimeOffset(Optional Server, Optional Port) As Double` - 获取时间偏移
- `RequestTimeAsync(Optional Server, Optional Port, Optional Tag)` - 异步请求
- `SyncSystemTime(Optional Server, Optional Port) As Boolean` - 同步系统时间

#### 事件

- `TimeSyncComplete(ServerHost, ServerTime, Offset, RoundTripDelay, Success, ErrorMsg)`

### cNTPServer

#### 属性

| 属性           | 类型             | 说明        |
| -------------- | ---------------- | ----------- |
| Port           | Long             | 服务器端口  |
| IsRunning      | Boolean          | 是否运行中  |
| Stratum        | Byte             | 层级 (1-15) |
| ReferenceID    | String           | 参考标识符  |
| RootDelay      | Double           | 根延迟 (秒) |
| RootDispersion | Double           | 根分散 (秒) |
| RequestCount   | Long             | 请求计数    |
| TimeProvider   | cNTPTimeProvider | 时间提供者  |
| Socket         | cWinsock         | Socket 对象 |

#### 方法

- `Start(Optional Port As Long = 123)` - 启动服务器
- `Stop()` - 停止服务器
- `ResetRequestCount()` - 重置请求计数

#### 事件

- `ServerStarted(Port)`
- `ServerStopped()`
- `RequestReceived(RemoteIP, RemotePort)`
- `ResponseSent(RemoteIP, RemotePort)`
- `ServerError(Number, Description)`

### cNTPTimeProvider

#### 枚举

- `TimeSourceType` - 时间源类型
  - `TimeSource_System` - 系统时间
  - `TimeSource_External` - 外部 NTP 服务器
  - `TimeSource_Custom` - 自定义时间 (测试用)

#### 属性

| 属性           | 类型           | 说明            |
| -------------- | -------------- | --------------- |
| TimeSource     | TimeSourceType | 时间源类型      |
| ExternalServer | String         | 外部服务器地址  |
| ExternalPort   | Long           | 外部服务器端口  |
| CustomOffset   | Double         | 自定义偏移 (秒) |
| UseCustomTime  | Boolean        | 使用自定义时间  |
| CustomTime     | Date           | 自定义时间      |

#### 方法

- `GetCurrentTime() As Date` - 获取当前时间
- `SyncExternalSource() As Boolean` - 同步外部时间源

## 公共 NTP 服务器

推荐使用以下公共 NTP 服务器：

- pool.ntp.org
- time.windows.com
- time.google.com
- time.nist.gov
- time.cloudflare.com

## 注意事项

1. **端口权限**: NTP 默认端口 123 需要管理员权限，测试时可以使用其他端口 (如 1230)
2. **防火墙**: 确保防火墙允许 UDP 123 端口的入站和出站流量
3. **网络延迟**: NTP 会自动计算网络延迟，但高延迟环境会影响精度
4. **系统时间同步**: 同步系统时间需要管理员权限

## 依赖

- cWinsock - Winsock 封装类

## 参考资料

- RFC 5905 - Network Time Protocol Version 4
- RFC 4330 - Simple Network Time Protocol (SNTP) Version 4
- NTP 协议官方文档: https://www.ntp.org/documentation/
