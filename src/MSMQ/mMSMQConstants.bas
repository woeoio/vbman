Attribute VB_Name = "mMSMQConstants"
'===========================================================================
' 模块名:  mMSMQConstants
' 描述:    MSMQ 常量定义模块
' 作者:    [作者]
' 日期:    [日期]
' 说明:    定义 MSMQ 相关的所有常量，便于统一引用
'===========================================================================

'----------------------------------
' 队列访问模式
'----------------------------------
Public Const MQ_RECEIVE_ACCESS = 1
Public Const MQ_SEND_ACCESS = 2
Public Const MQ_PEEK_ACCESS = 32

'----------------------------------
' 队列共享模式
'----------------------------------
Public Const MQ_DENY_NONE = 0
Public Const MQ_DENY_RECEIVE_SHARE = 1

'----------------------------------
' 事务类型
'----------------------------------
Public Const MQ_NO_TRANSACTION = 0
Public Const MQ_MTS_TRANSACTION = 1
Public Const MQ_XA_TRANSACTION = 2
Public Const MQ_SINGLE_MESSAGE = 3

'----------------------------------
' 消息传递类型
'----------------------------------
Public Const MQMSG_DELIVERY_EXPRESS = 0
Public Const MQMSG_DELIVERY_RECOVERABLE = 1

'----------------------------------
' 消息确认类型
'----------------------------------
Public Const MQMSG_ACKNOWLEDGMENT_NONE = 0
Public Const MQMSG_ACKNOWLEDGMENT_FULL_REACH_QUEUE = 1
Public Const MQMSG_ACKNOWLEDGMENT_FULL_RECEIVE = 5
Public Const MQMSG_ACKNOWLEDGMENT_NACK_REACH_QUEUE = 4

'----------------------------------
' 日志类型
'----------------------------------
Public Const MQMSG_JOURNAL_NONE = 0
Public Const MQMSG_DEADLETTER = 1
Public Const MQMSG_JOURNAL = 2

'----------------------------------
' 消息优先级 (0-7)
'----------------------------------
Public Const MQMSG_PRIORITY_LOWEST = 0
Public Const MQMSG_PRIORITY_VERY_LOW = 1
Public Const MQMSG_PRIORITY_LOW = 2
Public Const MQMSG_PRIORITY_NORMAL = 3
Public Const MQMSG_PRIORITY_ABOVE_NORMAL = 4
Public Const MQMSG_PRIORITY_HIGH = 5
Public Const MQMSG_PRIORITY_VERY_HIGH = 6
Public Const MQMSG_PRIORITY_HIGHEST = 7

'----------------------------------
' 发送者ID类型
'----------------------------------
Public Const MQMSG_SENDERID_TYPE_NONE = 0
Public Const MQMSG_SENDERID_TYPE_SID = 1

'----------------------------------
' 消息类别（确认消息）
'----------------------------------
Public Const MQMSG_CLASS_NORMAL = 0
Public Const MQMSG_CLASS_REPORT = 1
Public Const MQMSG_CLASS_ACK_REACH_QUEUE = 2
Public Const MQMSG_CLASS_ACK_RECEIVE = 16384
Public Const MQMSG_CLASS_NACK_BAD_DST_Q = 32768
Public Const MQMSG_CLASS_NACK_PURGED = 32769
Public Const MQMSG_CLASS_NACK_REACH_QUEUE_TIMEOUT = 32770
Public Const MQMSG_CLASS_NACK_Q_EXCEED_QUOTA = 32771
Public Const MQMSG_CLASS_NACK_ACCESS_DENIED = 32772
Public Const MQMSG_CLASS_NACK_HOP_COUNT_EXCEEDED = 32773
Public Const MQMSG_CLASS_NACK_BAD_SIGNATURE = 32774
Public Const MQMSG_CLASS_NACK_BAD_ENCRYPTION = 32775
Public Const MQMSG_CLASS_NACK_COULD_NOT_ENCRYPT = 32776
Public Const MQMSG_CLASS_NACK_NOT_TRANSACTIONAL_Q = 32777
Public Const MQMSG_CLASS_NACK_NOT_TRANSACTIONAL_MSG = 32778
Public Const MQMSG_CLASS_NACK_UNSUPPORTED_CRYPTO_PROVIDER = 32779
Public Const MQMSG_CLASS_NACK_SOURCE_COMPUTER_GUID_CHANGED = 32780
Public Const MQMSG_CLASS_NACK_Q_DELETED = 32781
Public Const MQMSG_CLASS_NACK_Q_PURGED = 32782
Public Const MQMSG_CLASS_NACK_RECEIVE_TIMEOUT = 32783
Public Const MQMSG_CLASS_NACK_RECEIVE_REJECTED = 32784

'----------------------------------
' 错误代码
'----------------------------------
Public Const MQ_ERROR_QUEUE_NOT_FOUND = -1072824317
Public Const MQ_ERROR_QUEUE_EXISTS = -1072824315
Public Const MQ_ERROR_INVALID_PARAMETER = -1072824313
Public Const MQ_ERROR_INVALID_HANDLE = -1072824312
Public Const MQ_ERROR_SERVICE_NOT_AVAILABLE = -1072824309
Public Const MQ_ERROR_IO_TIMEOUT = -1072824293
Public Const MQ_ERROR_ACCESS_DENIED = -1072824283
Public Const MQ_ERROR_INSUFFICIENT_RESOURCES = -1072824281
Public Const MQ_ERROR_MESSAGE_ALREADY_RECEIVED = -1072824277
Public Const MQ_ERROR_INVALID_TRANSACTION = -1072824231
Public Const MQ_ERROR_TRANSACTION_NOT_FOUND = -1072824229

'----------------------------------
' 格式名类型
'----------------------------------
Public Const MQF_PRIVATE_QUEUE = "PRIVATE$"
Public Const MQF_JOURNAL_SUFFIX = ";JOURNAL"
Public Const MQF_DEADLETTER = "SYSTEM$;DEADLETTER"
Public Const MQF_DEADXACT = "SYSTEM$;DEADXACT"

'----------------------------------
' 属性 ID (用于高级操作)
'----------------------------------
Public Const PROPID_M_ABORT_COUNT = 73
Public Const PROPID_M_MOVE_COUNT = 74

'----------------------------------
' 队列类型枚举 (供代码使用)
'----------------------------------
Public Enum MSMQQueueTypeConst
    MQ_PRIVATE_QUEUE = 0
    MQ_PUBLIC_QUEUE = 1
End Enum

'----------------------------------
' 访问模式枚举
'----------------------------------
Public Enum MSMQAccessModeConst
    MQ_MODE_SEND = 0
    MQ_MODE_RECEIVE = 1
    MQ_MODE_SEND_RECEIVE = 2
    MQ_MODE_PEEK = 3
End Enum

'----------------------------------
' 工具函数：获取错误描述
'----------------------------------
Public Function GetMSMQErrorDescription(ByVal ErrorCode As Long) As String
    Select Case ErrorCode
        Case MQ_ERROR_QUEUE_NOT_FOUND
            GetMSMQErrorDescription = "队列不存在"
        Case MQ_ERROR_QUEUE_EXISTS
            GetMSMQErrorDescription = "队列已存在"
        Case MQ_ERROR_INVALID_PARAMETER
            GetMSMQErrorDescription = "无效参数"
        Case MQ_ERROR_INVALID_HANDLE
            GetMSMQErrorDescription = "无效句柄"
        Case MQ_ERROR_SERVICE_NOT_AVAILABLE
            GetMSMQErrorDescription = "MSMQ 服务不可用"
        Case MQ_ERROR_IO_TIMEOUT
            GetMSMQErrorDescription = "I/O 超时"
        Case MQ_ERROR_ACCESS_DENIED
            GetMSMQErrorDescription = "访问被拒绝"
        Case MQ_ERROR_INSUFFICIENT_RESOURCES
            GetMSMQErrorDescription = "资源不足"
        Case MQ_ERROR_MESSAGE_ALREADY_RECEIVED
            GetMSMQErrorDescription = "消息已被接收"
        Case MQ_ERROR_INVALID_TRANSACTION
            GetMSMQErrorDescription = "无效的事务"
        Case MQ_ERROR_TRANSACTION_NOT_FOUND
            GetMSMQErrorDescription = "事务未找到"
        Case Else
            GetMSMQErrorDescription = "未知错误 (" & ErrorCode & ")"
    End Select
End Function

'----------------------------------
' 工具函数：构建队列路径
'----------------------------------
Public Function BuildQueuePath( _
    ByVal QueueName As String, _
    Optional ByVal IsPrivate As Boolean = True, _
    Optional ByVal MachineName As String = "." _
) As String
    
    If IsPrivate Then
        BuildQueuePath = MachineName & "\PRIVATE$\" & QueueName
    Else
        BuildQueuePath = MachineName & "\" & QueueName
    End If
End Function

'----------------------------------
' 工具函数：构建直接格式名
'----------------------------------
Public Function BuildDirectFormatName( _
    ByVal QueueName As String, _
    ByVal MachineName As String, _
    Optional ByVal IsPrivate As Boolean = True _
) As String
    
    Dim queuePath As String
    If IsPrivate Then
        queuePath = "PRIVATE$\" & QueueName
    Else
        queuePath = QueueName
    End If
    
    BuildDirectFormatName = "DIRECT=OS:" & MachineName & "\" & queuePath
End Function

'----------------------------------
' 工具函数：生成唯一消息ID
'----------------------------------
Public Function GenerateMessageID() As String
    ' 格式：YYYYMMDDHHNNSS-Random-Random
    GenerateMessageID = Format(Now, "yyyymmddhhnnss") & _
                       "-" & CStr(Int(Rnd * 1000000)) & _
                       "-" & CStr(GetTickCount())
End Function

Private Declare Function GetTickCount Lib "kernel32" () As Long
