VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQTransaction
' 描述:    MSMQ 事务处理封装类
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    封装 MSMQ 事务操作，支持内部事务和 DTC 事务
'===========================================================================

'----------------------------------
' 成员变量
'----------------------------------
Private m_Transaction As Object       ' MSMQTransaction
Private m_Dispenser As Object         ' MSMQTransactionDispenser
Private m_IsActive As Boolean
Private m_TransactionType As MSMQTransactionType
Private m_LastError As String

'----------------------------------
' 枚举
'----------------------------------
Public Enum MSMQTransactionType
    ttInternal = 0      ' MSMQ 内部事务
    ttDTC = 1           ' DTC (Distributed Transaction Coordinator) 事务
End Enum

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnTransactionBegin()
Public Event OnTransactionCommit()
Public Event OnTransactionAbort()
Public Event OnTransactionError(ByVal Operation As String, ByVal ErrorDescription As String)

'----------------------------------
' 属性
'----------------------------------
Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Get TransactionType() As MSMQTransactionType
    TransactionType = m_TransactionType
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

'----------------------------------
' 属性: 获取事务对象（供 Send/Receive 使用）
'----------------------------------
Public Property Get TransactionObject() As Object
    If m_IsActive Then
        Set TransactionObject = m_Transaction
    Else
        Set TransactionObject = Nothing
    End If
End Property

'----------------------------------
' 方法: 开始事务
'----------------------------------
Public Function BeginTransaction( _
    Optional ByVal TransactionType As MSMQTransactionType = ttInternal _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 如果已有活动事务，先回滚
    If m_IsActive Then Abort
    
    m_TransactionType = TransactionType
    
    If TransactionType = ttInternal Then
        ' 使用 MSMQ 内部事务
        Set m_Dispenser = CreateObject("MSMQ.MSMQTransactionDispenser")
        Set m_Transaction = m_Dispenser.BeginTransaction
    Else
        ' DTC 事务 - 需要 COM+ 或 MSDTC 支持
        ' 注：DTC 事务较为复杂，需要额外配置
        Set m_Dispenser = CreateObject("MSMQ.MSMQCoordinatedTransactionDispenser")
        Set m_Transaction = m_Dispenser.BeginTransaction
    End If
    
    m_IsActive = True
    RaiseEvent OnTransactionBegin
    BeginTransaction = True
    Exit Function
    
ErrorHandler:
    m_LastError = "开始事务失败: " & Err.Description
    RaiseEvent OnTransactionError("BeginTransaction", m_LastError)
    BeginTransaction = False
End Function

'----------------------------------
' 方法: 提交事务
'----------------------------------
Public Function Commit() As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsActive Then
        m_LastError = "没有活动事务可提交"
        Commit = False
        Exit Function
    End If
    
    m_Transaction.Commit
    m_IsActive = False
    
    Set m_Transaction = Nothing
    Set m_Dispenser = Nothing
    
    RaiseEvent OnTransactionCommit
    Commit = True
    Exit Function
    
ErrorHandler:
    m_LastError = "提交事务失败: " & Err.Description
    RaiseEvent OnTransactionError("Commit", m_LastError)
    Commit = False
End Function

'----------------------------------
' 方法: 回滚事务
'----------------------------------
Public Function Abort() As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsActive Then
        Abort = True
        Exit Function
    End If
    
    m_Transaction.Abort
    m_IsActive = False
    
    Set m_Transaction = Nothing
    Set m_Dispenser = Nothing
    
    RaiseEvent OnTransactionAbort
    Abort = True
    Exit Function
    
ErrorHandler:
    m_LastError = "回滚事务失败: " & Err.Description
    RaiseEvent OnTransactionError("Abort", m_LastError)
    Abort = False
End Function

'----------------------------------
' 方法: 执行事务操作（自动回滚/提交）
'----------------------------------
Public Function Execute( _
    ByRef Operation As Object, _
    Optional ByVal AutoCommit As Boolean = True _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 开始事务
    If Not BeginTransaction() Then
        Execute = False
        Exit Function
    End If
    
    ' 执行操作
    ' 注：Operation 应该是一个实现了特定接口的对象
    ' 这里简化处理，实际使用时需要配合具体业务逻辑
    
    If AutoCommit Then
        Execute = Commit()
    Else
        Execute = True
    End If
    Exit Function
    
ErrorHandler:
    Abort
    m_LastError = "执行事务操作失败: " & Err.Description
    RaiseEvent OnTransactionError("Execute", m_LastError)
    Execute = False
End Function

'----------------------------------
' 清理
'----------------------------------
Private Sub Class_Terminate()
    If m_IsActive Then
        On Error Resume Next
        Abort
    End If
End Sub
