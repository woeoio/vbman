# 2026-06-09 TCP 粘包/分包问题与 cPacketProtocol 解决方案

## 问题背景

来自用户反馈（群聊记录）：

> "winsock 控件在现场发送大量数据的时候经常发送的不完整，要发好几次可能有一次完整的"
> 
> "之字符格式发送，丢包更严重，后来发送前把字符转为字节再发送，好像效果好点"

这是 TCP 流式协议的典型问题：**TCP 没有消息边界**，数据是作为连续的字节流传输的，不保证一次 Send 对应一次 Receive。

## 粘包与分包是一体两面

| 现象 | 术语 | 说明 |
|------|------|------|
| 一次 Send 的数据，分多次 Receive 到达 | **分包（拆包）** | 发了 1000 字节，先收到 300，再收到 700 |
| 多次 Send 的数据，一次 Receive 全到 | **粘包** | 连续发了 3 条消息，一次收到拼接在一起的数据 |
| 上面两种混合出现 | **最常见** | 收到的数据既不完整又混着下一条的开头 |

用户反馈的"不完整"——每次 `DataArrival` 拿到的数据都不是一条完整消息，**这正是分包问题**。而"要发好几次可能有一次完整的"——说明偶尔凑巧一次收全了，这意味着粘包和分包同时存在。

### 现象图解

```
应用层发送：[Msg1][Msg2][Msg3]
                    ↓ TCP 流式传输（无边界）
接收端可能收到：
  情况1（分包）：  [Msg1前半]  [Msg1后半+Msg2前半]  [Msg2后半+Msg3]
  情况2（粘包）：  [Msg1+Msg2]  [Msg3]
  情况3（混合）：  [Msg1前半]  [Msg1后半+Msg2]  [Msg3前半]  [Msg3后半]
  理想情况（少见）：[Msg1]  [Msg2]  [Msg3]
```

## cPacketProtocol 解决方案

`cPacketProtocol` 通过在数据中定义明确的边界，将无界的字节流还原为有界的消息：

```
原始 TCP 字节流（无边界）：
  [Msg1前半][Msg1后半+Msg2前半][Msg2后半]
                        ↓ cPacketProtocol.Decode()
完整消息：
  [Msg1 完整] → 触发 MessageArrival
  [Msg2 完整] → 触发 MessageArrival
```

### 三种协议对比

| 协议类型 | 原理 | 分包处理 | 粘包处理 | 优缺点 |
|----------|------|----------|----------|--------|
| `ppLengthHeader` | 头部指明消息体长度 | 长度不够则缓存，等数据到齐再提取 | 长度够了就切一条，剩余继续解析 | **推荐**。不依赖数据内容，不限消息长度 |
| `ppDelimiter` | 分隔符标记消息结尾 | 未找到分隔符则缓存 | 找到分隔符就切一条，剩余继续找 | 简单，但分隔符不能出现在消息体中 |
| `ppFixedLength` | 每条消息固定长度 | 不足定长则缓存 | 凑够定长就切一条 | 仅适用于定长消息场景 |

### 推荐方案：ppLengthHeader（4 字节小端头）

不依赖数据内容中出现特殊字符（分隔符协议的硬伤），不限制单条消息长度（定长协议的硬伤），每条消息自带长度，接收端精确知道要读多少字节。

### 使用示例

```vb
'--- 服务端设置 ---
server.PacketProtocol = ppLengthHeader
server.HeaderBytes = 4          ' 4字节长度头（支持最大4GB）
server.HeaderEndian = eeLittleEndian  ' 小端序

'--- 客户端设置 ---
client.PacketProtocol = ppLengthHeader
client.HeaderBytes = 4
client.HeaderEndian = eeLittleEndian

'--- 发送：自动加长度头，无需手动处理 ---
client.SendData myData

'--- 接收：每次一定是完整包 ---
Private Sub server_MessageArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim data() As Byte
    data = Client.GetDataByteArray()  ' 100% 是一条完整消息
End Sub
```

### 安全机制

2026-06-09 同步修复的安全加固：

- `MaxPacketSize`（默认 1MB）：单包最大大小，防止恶意超大包声明耗尽内存
- `MaxBufferSize`（默认 4MB）：缓冲区累积上限，防止大量不完整包慢慢吃内存
- 心跳数据也走协议编码，不会污染协议状态机
- UDP 客户端同样支持分包协议

## 事件模型

| 模式 | 触发事件 | 说明 |
|------|----------|------|
| 无协议（`ppNone`） | `DataArrival` | 原始字节流，可能不完整或粘连 |
| 有协议 | `MessageArrival` | 每次一定是完整的一条消息 |

协议模式下**只触发 `MessageArrival`**，不触发 `DataArrival`，避免同一数据被两个事件重复读取。

## 总结

TCP 粘包/分包不是 Bug，而是 TCP 流式协议的特性。解决方式就是在应用层定义消息边界——这正是 `cPacketProtocol` 的设计目标。设置 `PacketProtocol` 后，开发者无需关心底层字节流的分合，`MessageArrival` 每次触发的都是一条完整消息。
