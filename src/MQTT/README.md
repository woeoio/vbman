# MQTT Protocol Implementation for VB6

基于 `cWinsock` 的完整 MQTT 3.1.1 协议栈实现。

## 文件结构

```
src/MQTT/
├── mMQTTConst.bas          ' MQTT 协议常量定义
├── cMQTTPacket.cls         ' MQTT 报文编解码器
├── cMQTTTopicFilter.cls    ' 主题过滤器
├── cMQTTMessage.cls        ' MQTT 消息对象
├── cMQTTPendingMessage.cls ' 待确认消息
├── cMQTTSession.cls        ' 客户端会话管理
├── cMQTTServer.cls         ' MQTT 服务端 (Broker)
├── cMQTTClient.cls         ' MQTT 客户端
└── frmMQTTDemo.frm         ' 演示窗体
```

## 功能特性

### 已实现

- [x] **报文类型**: CONNECT, CONNACK, PUBLISH, PUBACK, PUBREC, PUBREL, PUBCOMP
- [x] **订阅管理**: SUBSCRIBE, SUBACK, UNSUBSCRIBE, UNSUBACK
- [x] **保活机制**: PINGREQ, PINGRESP
- [x] **QoS 级别**: QoS 0 (最多一次), QoS 1 (至少一次), QoS 2 (恰好一次)
- [x] **主题通配符**: `+` (单级), `#` (多级)
- [x] **保留消息**: Retain 标志支持
- [x] **清理会话**: Clean Session 支持
- [x] **遗嘱消息**: Will Message 框架 (待完善)

### 待实现

- [ ] 遗嘱消息完整实现
- [ ] 消息重传机制
- [ ] 保活超时检测
- [ ] 用户认证
- [ ] TLS/SSL 支持
- [ ] 遗嘱消息延迟发送

## 快速开始

### 服务端使用

```vb
Dim oServer As New cMQTTServer

' 启动服务端
If oServer.StartServer(1883) Then
    Debug.Print "MQTT Server started"
End If

' 发布消息
oServer.PublishMessage "sensor/temperature", "25.5", mqttQoS0, True

' 停止服务端
oServer.StopServer
```

### 客户端使用

```vb
Dim oClient As New cMQTTClient

' 设置连接参数
oClient.ClientId = "MyClient_001"
oClient.UserName = "user"
oClient.Password = "pass"
oClient.CleanSession = True
oClient.KeepAlive = 60

' 连接到服务器
If oClient.Connect("127.0.0.1", 1883) Then
    Debug.Print "Connected!"
End If

' 订阅主题
oClient.Subscribe "sensor/+", mqttQoS1

' 发布消息
oClient.Publish "sensor/temperature", "25.5", mqttQoS1, False

' 断开连接
oClient.Disconnect
```

## 事件处理

### 服务端事件

```vb
Private WithEvents m_oServer As cMQTTServer

Private Sub m_oServer_ClientConnected(ClientId As String, Session As cMQTTSession)
    Debug.Print "Client connected: " & ClientId
End Sub

Private Sub m_oServer_MessageReceived(ClientId As String, Topic As String, Payload As Variant, QoS As MqttQoS)
    Debug.Print "Message from " & ClientId & " on " & Topic
End Sub
```

### 客户端事件

```vb
Private WithEvents m_oClient As cMQTTClient

Private Sub m_oClient_Connected()
    Debug.Print "Connected to server"
End Sub

Private Sub m_oClient_MessageArrived(Topic As String, Payload As Variant, QoS As MqttQoS, Retain As Boolean)
    Debug.Print "Message on " & Topic & ": " & Payload
End Sub
```

## 主题通配符规则

- `+` : 匹配单个层级，如 `sensor/+/temperature` 匹配 `sensor/living/temperature`
- `#` : 匹配多个层级（必须是最后一个），如 `sensor/#` 匹配 `sensor/living/temperature`

## 架构说明

```
┌─────────────────────────────────────────────────────────────┐
│                    应用层 (Application)                      │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐ │
│  │ MQTT Server │  │ MQTT Client │  │    消息路由引擎      │ │
│  └─────────────┘  └─────────────┘  └─────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                    协议层 (Protocol)                         │
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────┐ │
│  │报文编解码器 │  │  会话管理   │  │    QoS 处理器        │ │
│  │ (Encoder)   │  │  (Session)  │  │                     │ │
│  └─────────────┘  └─────────────┘  └─────────────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                    传输层 (Transport)                        │
│              cWinsock (TCP 连接管理)                         │
└─────────────────────────────────────────────────────────────┘
```

## 依赖

- `cWinsock.cls` - 网络传输层
- `cAsyncSocket` - 异步套接字 (cWinsock 依赖)

## 参考

- [MQTT Version 3.1.1 Specification](http://docs.oasis-open.org/mqtt/mqtt/v3.1.1/os/mqtt-v3.1.1-os.html)
