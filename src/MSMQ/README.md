# VB6/VB MSMQ 封装模块

本项目提供了一套完整的 MSMQ (Microsoft Message Queuing) 封装，用于在 VB6/VBA 中方便地使用微软消息队列服务。

## 目录结构

### 文档

| 文件 | 说明 |
|------|------|
| `01_MSMQ概述.md` | MSMQ 基础概念和特点介绍 |
| `02_安装配置.md` | Windows 11/10/Server 安装步骤和配置方法 |
| `03_MSMQ_COM接口.md` | 原始 COM 接口详细说明 |
| `04_VB6封装设计.md` | 封装类设计思路和架构说明 |
| `05_完整代码示例.md` | 各种使用场景的完整代码示例 |
| `06_高级主题.md` | 直接格式名、毒消息处理、HTTP支持等高级话题 |
| `07_最佳实践.md` | 设计原则、性能优化、监控运维等最佳实践 |
| `08_与其他消息队列对比.md` | MSMQ与RabbitMQ/Kafka等对比分析 |

### 封装类

| 文件 | 功能 |
|------|------|
| `cMSMQClient.cls` | 核心客户端，封装连接、发送、接收等操作 |
| `cMSMQMessage.cls` | 消息对象，封装消息属性和数据 |
| `cMSMQManager.cls` | 队列管理，创建/删除/查询队列 |
| `cMSMQAsyncReceiver.cls` | 异步接收器，事件驱动接收消息 |
| `cMSMQTransaction.cls` | 事务处理封装，支持内部和DTC事务 |
| `cMSMQReceiverPool.cls` | 接收池，多实例并行处理消息 |
| `mMSMQConstants.bas` | 常量定义模块，包含所有MSMQ常量 |

## 快速开始

### 1. 系统要求

- Windows 10/11 专业版/企业版 或 Windows Server
- 已安装 MSMQ 服务
- VB6/VBA 开发环境

### 2. 添加引用

在 VB6 中：
```
工程 → 引用 → 勾选 "Microsoft Message Queue Object Library"
```

**导入封装模块：**
```
1. 将所有 .cls 和 .bas 文件导入到 VB6 工程
2. 或者在代码中使用 File → Import File
```

### 3. 基本使用示例

#### 同步发送消息
```vb
Dim mq As New cMSMQClient

If mq.Connect("MyQueue", qtPrivate, maSendOnly, True) Then
    mq.SendMessage "Hello MSMQ!", "测试消息", 3
    mq.CloseConnection
End If
```

#### 异步接收消息（推荐）
```vb
Private WithEvents m_Receiver As cMSMQAsyncReceiver

Private Sub Form_Load()
    Set m_Receiver = New cMSMQAsyncReceiver
    m_Receiver.StartListening "MyQueue", qtPrivate, True
End Sub

Private Sub m_Receiver_OnMessageArrived(ByRef Message As cMSMQMessage)
    Debug.Print "收到: " & Message.Body
    ' 处理消息...
End Sub
```

#### 发送消息
```vb
Dim mq As New cMSMQClient

If mq.Connect("MyQueue", qtPrivate, maSendOnly, True) Then
    mq.SendMessage "Hello MSMQ!", "测试消息", 3
    mq.CloseConnection
End If
```

#### 接收消息
```vb
Dim mq As New cMSMQClient
Dim msg As cMSMQMessage

If mq.Connect("MyQueue", qtPrivate, maReceiveOnly) Then
    Set msg = mq.ReceiveMessage(5000)  ' 等待5秒
    If Not msg Is Nothing Then
        MsgBox "收到: " & msg.Body
    End If
    mq.CloseConnection
End If
```

#### 事务性消息处理
```vb
Dim txn As New cMSMQTransaction
Dim mq As New cMSMQClient

If Not txn.BeginTransaction Then
    MsgBox "开始事务失败: " & txn.LastError
    Exit Sub
End If

mq.Connect "OrderQueue", qtPrivate, maSendOnly

' 发送消息（参与事务）
Dim msg As Object
Set msg = CreateObject("MSMQ.MSMQMessage")
msg.Body = "订单数据"
mq.SendMessageEx msg, txn.TransactionObject

' 提交或回滚
If 处理成功 Then
    txn.Commit
Else
    txn.Abort
End If
```

#### 使用常量模块
```vb
' 引用 mMSMQConstants 模块后
Dim path As String
path = BuildQueuePath("MyQueue", True, ".")
' 结果: .\private$\MyQueue

Dim formatName As String
formatName = BuildDirectFormatName("MyQueue", "Server01", True)
' 结果: DIRECT=OS:Server01\private$\MyQueue

' 获取错误描述
Dim errDesc As String
errDesc = GetMSMQErrorDescription(MQ_ERROR_ACCESS_DENIED)
' 结果: 访问被拒绝
```

#### 管理队列
```vb
Dim mgr As New cMSMQManager

' 创建队列
mgr.CreateQueue "NewQueue", qtPrivate, "我的队列", False

' 检查队列是否存在
If mgr.QueueExists("NewQueue") Then
    ' 获取队列信息
    Dim info As Object
    Set info = mgr.GetQueueInfo("NewQueue")
    MsgBox "消息数: " & info("MessageCount")
End If

' 清空队列
mgr.PurgeQueue "NewQueue"

' 删除队列
mgr.DeleteQueue "NewQueue"
```

## 类说明

### cMSMQClient - 消息队列客户端

主要功能：
- `Connect()` - 连接到队列
- `ConnectByFormatName()` - 通过格式名连接（支持远程队列）
- `SendMessage()` - 发送消息
- `ReceiveMessage()` - 接收消息（同步，阻塞或超时）
- `PeekMessage()` - 查看消息（不删除）
- `Purge()` - 清空队列
- `CloseConnection()` - 关闭连接

事件：
- `OnConnected` - 连接成功时触发
- `OnDisconnected` - 断开连接时触发
- `OnError` - 发生错误时触发

### cMSMQMessage - 消息对象

属性：
- `Body` - 消息正文（可以是任意 Variant 类型）
- `Label` - 消息标签/标题
- `Priority` - 优先级（0-7，默认3）
- `SentTime` - 发送时间
- `ArrivedTime` - 到达时间
- `MessageId` - 消息唯一ID
- `CorrelationId` - 关联ID（用于请求-响应模式）
- `AppSpecific` - 应用程序特定值（可用于消息类型标识）

### cMSMQManager - 队列管理器

主要功能：
- `CreateQueue()` - 创建队列
- `DeleteQueue()` - 删除队列
- `QueueExists()` - 检查队列是否存在
- `GetQueueInfo()` - 获取队列信息（消息数等）
- `PurgeQueue()` - 清空队列
- `GetQueueList()` - 获取队列列表

事件：
- `OnError` - 操作错误时触发
- `OnOperationComplete` - 操作完成时触发

### cMSMQAsyncReceiver - 异步接收器

主要功能：
- `StartListening()` - 开始异步监听队列
- `StartListeningByFormatName()` - 通过格式名监听远程队列
- `StopListening()` - 停止监听

事件：
- `OnMessageArrived` - 消息到达时触发
- `OnReceiveError` - 接收错误时触发
- `OnListenerStarted` - 监听器启动时触发
- `OnListenerStopped` - 监听器停止时触发

属性：
- `AutoReconnect` - 是否自动重连
- `ReceiveTimeout` - 接收超时时间（毫秒）

### cMSMQTransaction - 事务处理器

主要功能：
- `BeginTransaction()` - 开始事务
- `Commit()` - 提交事务
- `Abort()` - 回滚事务
- `Execute()` - 执行事务操作（自动管理）

事件：
- `OnTransactionBegin` - 事务开始时触发
- `OnTransactionCommit` - 事务提交时触发
- `OnTransactionAbort` - 事务回滚时触发

### cMSMQReceiverPool - 接收池

主要功能：
- `Start()` - 启动多个接收器并行处理
- `StopPool()` - 停止所有接收器

事件：
- `OnMessageReceived` - 消息被接收时触发（包含WorkerId）
- `OnWorkerError` - Worker 出错时触发

### mMSMQConstants - 常量模块

提供：
- 所有 MSMQ 常量定义（访问模式、事务类型、优先级等）
- 工具函数：
  - `GetMSMQErrorDescription()` - 获取错误描述
  - `BuildQueuePath()` - 构建队列路径
  - `BuildDirectFormatName()` - 构建直接格式名
  - `GenerateMessageID()` - 生成唯一消息ID

## 枚举常量

```vb
' 访问模式
maSendOnly      ' 仅发送
maReceiveOnly   ' 仅接收
maSendReceive   ' 发送和接收
maPeekOnly      ' 仅查看

' 队列类型
qtPrivate       ' 专用队列（本地）
qtPublic        ' 公共队列（需要域）

' 传递模式
dmExpress       ' 快速（内存中）
dmRecoverable   ' 可恢复（持久化到磁盘）
```

## 文档阅读指南

### 新手入门路线
1. `01_MSMQ概述.md` - 了解 MSMQ 是什么
2. `02_安装配置.md` - 安装和配置 MSMQ
3. `05_完整代码示例.md` - 运行示例代码
4. `07_最佳实践.md` - 学习设计原则

### 进阶学习路线
1. `03_MSMQ_COM接口.md` - 深入了解 COM 接口
2. `04_VB6封装设计.md` - 理解封装架构
3. `06_高级主题.md` - 掌握高级特性
4. `08_与其他消息队列对比.md` - 了解技术选型

## 注意事项

### 使用限制
1. **权限问题**：确保当前用户有操作 MSMQ 的权限
2. **事务性队列**：创建时设置 IsTransactional=True，发送时需指定事务
3. **远程队列**：使用 ConnectByFormatName() 方法，路径格式为 `DIRECT=OS:计算机名\private$\队列名`
4. **消息大小**：默认限制 4MB，大数据量建议分批发送
5. **超时处理**：ReceiveMessage/PeekMessage 的超时单位是毫秒，0表示无限等待
6. **错误处理**：建议订阅 OnError 事件进行统一错误处理

### 已知问题
1. **Windows 版本**：部分功能在 Windows 家庭版不可用
2. **COM 引用**：确保引用正确的 mqoa.dll 版本
3. **防火墙**：远程队列访问需要开放相应端口

### 性能提示
1. 批量处理消息而不是逐条处理
2. 合理使用消息优先级
3. 及时清理日志队列
4. 使用异步接收提高吞吐量

## 进阶阅读

- 查看 `05_完整代码示例.md` 了解更多使用场景：
  - 批量发送/接收
  - 事务性消息
  - 异步接收（事件驱动）
  - 发送对象/结构体
  - 远程队列访问

- 查看 `06_高级主题.md` 了解：
  - 直接格式名（Direct Format Names）
  - 毒消息处理（Poison Messages）
  - HTTP/HTTPS 支持
  - 事务详解
  - 消息确认机制

## 参考

- [Microsoft MSMQ 文档](https://learn.microsoft.com/en-us/windows/win32/msmq/microsoft-message-queue-technology)
- [MSMQ COM API 参考](https://learn.microsoft.com/en-us/windows/win32/msmq/msmq-com-components)
- [Message Queuing Guide](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms700996(v=vs.85))
