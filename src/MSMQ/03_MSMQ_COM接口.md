# MSMQ COM 接口详解

MSMQ 提供 COM 组件接口供应用程序调用。在 VB6/VBA 中使用前需要先引用 MSMQ 库。

## 引用库

### 在 VB6 中添加引用
```
1. 工程 → 引用
2. 勾选 "Microsoft Message Queue Object Library"
3. 通常版本为 3.0 或更高
```

### 库信息
| 属性 | 值 |
|------|-----|
| 类型库 | mqoa.dll |
| ProgID | MSMQ.MSMQQueueInfo |
| 版本 | 3.0/4.0/5.0/6.0（取决于系统） |

## 核心对象

### 1. MSMQQueueInfo - 队列信息对象

用于创建、删除和管理队列。

```vb
Dim qinfo As MSMQQueueInfo
Set qinfo = New MSMQQueueInfo
```

#### 主要属性
| 属性 | 类型 | 说明 |
|------|------|------|
| PathName | String | 队列路径（如 ".\private$\MyQueue"） |
| FormatName | String | 队列格式名 |
| Label | String | 队列标签（描述性名称） |
| Type | String | 队列类型 GUID |

#### 主要方法
```vb
' 创建队列
qinfo.Create Optional IsTransactional

' 删除队列
qinfo.Delete

' 打开队列（返回 MSMQQueue 对象）
Set queue = qinfo.Open(Access, ShareMode)

' 刷新队列信息
qinfo.Refresh
```

#### 打开队列的访问模式
```vb
' Access 参数常量
MQ_RECEIVE_ACCESS = 1    ' 接收访问
MQ_SEND_ACCESS = 2       ' 发送访问
MQ_PEEK_ACCESS = 32      ' 查看访问（只读不删除）

' ShareMode 参数常量
MQ_DENY_NONE = 0         ' 共享模式
MQ_DENY_RECEIVE_SHARE = 1 ' 独占接收
```

### 2. MSMQQueue - 队列对象

表示一个打开的队列，用于发送和接收消息。

```vb
Dim queue As MSMQQueue
Set queue = qinfo.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
```

#### 主要属性
| 属性 | 类型 | 说明 |
|------|------|------|
| Access | Long | 当前访问模式 |
| IsOpen | Boolean | 队列是否打开 |

#### 主要方法
```vb
' 发送消息
queue.Send Message, Optional Transaction

' 接收消息（阻塞）
Set msg = queue.Receive(Optional Transaction, Optional WantDestinationQueue, Optional WantBody)

' 接收消息（超时）
Set msg = queue.Receive(Timeout)

' 查看消息（不删除）
Set msg = queue.Peek(Optional WantDestinationQueue, Optional WantBody, Optional Timeout)

' 关闭队列
queue.Close
```

### 3. MSMQMessage - 消息对象

表示一条消息。

```vb
Dim msg As MSMQMessage
Set msg = New MSMQMessage
```

#### 主要属性
| 属性 | 类型 | 说明 |
|------|------|------|
| Body | Variant | 消息主体内容（任意类型） |
| Label | String | 消息标签 |
| Priority | Long | 优先级（0-7，默认3） |
| Delivery | Long | 传递方式 |
| Acknowledge | Long | 确认类型 |
| CorrelationId | Variant | 关联ID（用于请求-响应模式） |

#### Delivery 常量
```vb
MQMSG_DELIVERY_EXPRESS = 0      ' 快速（内存中，可能丢失）
MQMSG_DELIVERY_RECOVERABLE = 1  ' 可恢复（持久化到磁盘）
```

#### 示例：创建消息
```vb
Dim msg As MSMQMessage
Set msg = New MSMQMessage

msg.Body = "Hello MSMQ"
msg.Label = "测试消息"
msg.Priority = 4
msg.Delivery = MQMSG_DELIVERY_RECOVERABLE
```

### 4. MSMQQuery - 查询对象

用于查询队列。

```vb
Dim query As MSMQQuery
Dim queues As MSMQQueues

Set query = New MSMQQuery
Set queues = query.LookupQueue( _
    Optional QueueGuid, _
    Optional ServiceTypeGuid, _
    Optional Label, _
    Optional CreateTime, _
    Optional ModifyTime, _
    Optional RelServiceType, _
    Optional RelLabel, _
    Optional RelCreateTime, _
    Optional RelModifyTime _
)
```

## 事务处理
n### MSMQTransaction 对象
```vb
Dim txn As MSMQTransaction
Dim txnDisp As MSMQTransactionDispenser

Set txnDisp = New MSMQTransactionDispenser
Set txn = txnDisp.BeginTransaction

' 在事务中发送消息
queue.Send msg, txn

' 提交或中止事务
txn.Commit
txn.Abort
```

### 事务常量
```vb
MQ_MTS_TRANSACTION = 1       ' 使用 MTS 事务
MQ_XA_TRANSACTION = 2        ' 使用 XA 事务
MQ_SINGLE_MESSAGE = 3        ' 单消息事务
```

## 事件处理（异步接收）

### MSMQEvent 对象
```vb
Dim qEvents As MSMQEvent
Set qEvents = New MSMQEvent

' 启用事件通知
queue.EnableNotification qEvents, Optional Cursor, Optional ReceiveTimeout

' 处理 Arrived 事件
Private Sub qEvents_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    ' 消息到达处理
    Set msg = Queue.Receive
End Sub
```

## 常用常量汇总

```vb
' 访问模式
Public Const MQ_RECEIVE_ACCESS = 1
Public Const MQ_SEND_ACCESS = 2
Public Const MQ_PEEK_ACCESS = 32

' 共享模式
Public Const MQ_DENY_NONE = 0
Public Const MQ_DENY_RECEIVE_SHARE = 1

' 传递方式
Public Const MQMSG_DELIVERY_EXPRESS = 0
Public Const MQMSG_DELIVERY_RECOVERABLE = 1

' 确认类型
Public Const MQMSG_ACKNOWLEDGMENT_NONE = 0
Public Const MQMSG_ACKNOWLEDGMENT_FULL_REACH_QUEUE = 1
Public Const MQMSG_ACKNOWLEDGMENT_FULL_RECEIVE = 5

' 日志类型
Public Const MQMSG_JOURNAL_NONE = 0
Public Const MQMSG_DEADLETTER = 1
Public Const MQMSG_JOURNAL = 2

' 优先级（0-7）
Public Const MQMSG_PRIORITY_LOWEST = 0
Public Const MQMSG_PRIORITY_HIGHEST = 7
```

## 错误处理

MSMQ 操作可能产生的错误：

| 错误代码 | 说明 |
|---------|------|
| MQ_ERROR_QUEUE_NOT_FOUND | 队列不存在 |
| MQ_ERROR_ACCESS_DENIED | 访问被拒绝 |
| MQ_ERROR_INVALID_HANDLE | 无效句柄 |
| MQ_ERROR_QUEUE_EXISTS | 队列已存在 |
| MQ_ERROR_INSUFFICIENT_RESOURCES | 资源不足 |
| MQ_ERROR_IO_TIMEOUT | I/O 超时 |
