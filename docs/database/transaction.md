# 事务处理

本文档详细介绍 cDataBase 类的事务处理功能，包括事务的开始、提交、回滚等。

---

## ? 目录

- [事务概述](#事务概述)
- [事务方法](#事务方法)
- [基本使用](#基本使用)
- [错误处理](#错误处理)
- [嵌套事务](#嵌套事务)
- [最佳实践](#最佳实践)

---

## 事务概述

### 什么是事务

事务是一组数据库操作，要么全部成功，要么全部失败。事务具有以下特性（ACID）：

- **原子性 (Atomicity)** - 事务中的所有操作要么全部执行，要么全部不执行
- **一致性 (Consistency)** - 事务执行前后数据库保持一致状态
- **隔离性 (Isolation)** - 并发事务之间相互隔离
- **持久性 (Durability)** - 事务提交后，数据永久保存

### 事务状态

```
开始事务
    ↓
执行操作 1
    ↓
执行操作 2
    ↓
执行操作 3
    ↓
提交事务 ──→ 成功：所有操作生效
    │
    └──→ 失败：自动回滚，所有操作撤销
```

---

## 事务方法

### TransBegin 方法

`TransBegin` 方法开始一个新事务。

#### 语法

```vb
Function TransBegin() As Boolean
```

#### 返回值

- `True` - 事务开始成功
- `False` - 事务开始失败（可通过 `LastErr` 查看错误信息）

#### 示例

```vb
' 开始事务
If db.TransBegin Then
    Debug.Print "事务已开始"
Else
    Debug.Print "事务开始失败: " & db.LastErr
End If
```

### TransCommit 方法

`TransCommit` 方法提交事务。

#### 语法

```vb
Function TransCommit() As Boolean
```

#### 功能

- 提交所有事务中的操作
- 如果提交失败，自动回滚
- 清除事务标记

#### 返回值

- `True` - 提交成功
- `False` - 提交失败（已自动回滚）

#### 示例

```vb
' 提交事务
If db.TransCommit Then
    Debug.Print "事务提交成功"
Else
    Debug.Print "事务提交失败，已自动回滚: " & db.LastErr
End If
```

### TransRollback 方法

`TransRollback` 方法回滚事务。

#### 语法

```vb
Function TransRollback() As Boolean
```

#### 功能

- 撤销所有事务中的操作
- 清除事务标记

#### 返回值

- `True` - 回滚成功
- `False` - 回滚失败（可通过 `LastErr` 查看错误信息）

#### 示例

```vb
' 回滚事务
If db.TransRollback Then
    Debug.Print "事务已回滚"
Else
    Debug.Print "回滚失败: " & db.LastErr
End If
```

---

## 基本使用

### 示例 1：简单事务

```vb
' 开始事务
If db.TransBegin Then
    ' 执行操作 1
    If db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec Then
        ' 执行操作 2
        If db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec Then
            ' 提交事务
            If db.TransCommit Then
                Debug.Print "事务提交成功"
            Else
                Debug.Print "提交失败，已自动回滚"
            End If
        Else
            ' 操作 2 失败，回滚
            db.TransRollback
        End If
    Else
        ' 操作 1 失败，回滚
        db.TransRollback
    End If
End If
```

### 示例 2：使用错误处理

```vb
On Error GoTo ErrHandler

' 开始事务
If Not db.TransBegin Then
    Debug.Print "事务开始失败: " & db.LastErr
    Exit Sub
End If

' 执行多个操作
db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
db.Sql("INSERT INTO users (name) VALUES ('李四')").Exec
db.Sql("UPDATE users SET status = 'active' WHERE name = '张三'").Exec

' 提交事务
If db.TransCommit Then
    Debug.Print "所有操作成功"
Else
    Debug.Print "提交失败: " & db.LastErr
End If

Exit Sub

ErrHandler:
' 发生错误，回滚事务
db.TransRollback
Debug.Print "发生错误，事务已回滚: " & Err.Description
```

### 示例 3：转账操作

```vb
' 转账操作：从账户 A 转 100 元到账户 B
Function TransferMoney(lFromAccount As Long, lToAccount As Long, dAmount As Double) As Boolean
    On Error GoTo ErrHandler
    
    ' 开始事务
    If Not db.TransBegin Then
        TransferMoney = False
        Exit Function
    End If
    
    ' 从账户 A 扣除金额
    Dim sSql As String
    sSql = "UPDATE accounts SET balance = balance - " & dAmount & _
           " WHERE id = " & lFromAccount & " AND balance >= " & dAmount
    If Not db.Sql(sSql).Exec Then
        db.TransRollback
        TransferMoney = False
        Exit Function
    End If
    
    ' 检查是否有记录被更新
    Dim lAffected As Long
    If db.Sql("SELECT @@ROWCOUNT").Query Then
        If db.Rs.EOF Or db.Rs(0) = 0 Then
            db.TransRollback
            TransferMoney = False
            Exit Function
        End If
    End If
    
    ' 向账户 B 增加金额
    sSql = "UPDATE accounts SET balance = balance + " & dAmount & _
           " WHERE id = " & lToAccount
    If Not db.Sql(sSql).Exec Then
        db.TransRollback
        TransferMoney = False
        Exit Function
    End If
    
    ' 记录转账日志
    sSql = "INSERT INTO transfers (from_account, to_account, amount) VALUES (" & _
           lFromAccount & ", " & lToAccount & ", " & dAmount & ")"
    If Not db.Sql(sSql).Exec Then
        db.TransRollback
        TransferMoney = False
        Exit Function
    End If
    
    ' 提交事务
    If db.TransCommit Then
        TransferMoney = True
    Else
        TransferMoney = False
    End If
    
    Exit Function
    
ErrHandler:
    db.TransRollback
    TransferMoney = False
End Function
```

---

## 错误处理

### 自动回滚机制

类库提供了自动回滚机制：

1. **提交失败自动回滚** - `TransCommit` 失败时自动调用 `TransRollback`
2. **断开连接自动回滚** - `Disconnect` 时自动回滚未完成的事务

```vb
' 提交失败时自动回滚
If db.TransBegin Then
    db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
    ' 如果提交失败，会自动回滚
    If Not db.TransCommit Then
        Debug.Print "提交失败，已自动回滚: " & db.LastErr
    End If
End If
```

### 错误处理示例

```vb
Function ExecuteTransaction() As Boolean
    On Error GoTo ErrHandler
    
    ' 开始事务
    If Not db.TransBegin Then
        Debug.Print "事务开始失败: " & db.LastErr
        ExecuteTransaction = False
        Exit Function
    End If
    
    ' 执行操作
    If Not db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec Then
        db.TransRollback
        Debug.Print "操作失败: " & db.LastErr
        ExecuteTransaction = False
        Exit Function
    End If
    
    ' 提交事务
    If db.TransCommit Then
        ExecuteTransaction = True
    Else
        Debug.Print "提交失败: " & db.LastErr
        ExecuteTransaction = False
    End If
    
    Exit Function
    
ErrHandler:
    ' 发生异常，回滚事务
    db.TransRollback
    Debug.Print "异常发生，已回滚: " & Err.Description
    ExecuteTransaction = False
End Function
```

---

## 嵌套事务

### 注意事项

ADO 支持嵌套事务，但需要注意：

1. **嵌套级别** - ADO 支持多级嵌套事务
2. **提交顺序** - 必须按相反顺序提交（内层先提交）
3. **回滚影响** - 回滚会影响所有嵌套级别

### 嵌套事务示例

```vb
' 外层事务
If db.TransBegin Then
    db.Sql("INSERT INTO users (name) VALUES ('用户1')").Exec
    
    ' 内层事务（嵌套）
    If db.TransBegin Then
        db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题1')").Exec
        
        ' 提交内层事务
        If db.TransCommit Then
            Debug.Print "内层事务提交成功"
        End If
    End If
    
    ' 提交外层事务
    If db.TransCommit Then
        Debug.Print "外层事务提交成功"
    End If
End If
```

---

## 最佳实践

### 1. 始终使用事务处理多个相关操作

```vb
' ? 推荐：使用事务保证一致性
db.TransBegin
db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec
db.TransCommit

' ? 不推荐：不使用事务
db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec
' 如果第二个操作失败，第一个操作已经提交，数据不一致
```

### 2. 检查每个操作的返回值

```vb
' ? 推荐：检查每个操作
db.TransBegin

If Not db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec Then
    db.TransRollback
    Exit Sub
End If

If Not db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec Then
    db.TransRollback
    Exit Sub
End If

db.TransCommit
```

### 3. 使用错误处理

```vb
' ? 推荐：使用错误处理
On Error GoTo ErrHandler

db.TransBegin
db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec
db.TransCommit

Exit Sub

ErrHandler:
    db.TransRollback
    Debug.Print "错误: " & Err.Description
```

### 4. 及时提交或回滚

```vb
' ? 推荐：操作完成后立即提交或回滚
db.TransBegin
' ... 执行操作 ...
db.TransCommit  ' 或 db.TransRollback

' ? 不推荐：长时间保持事务打开
db.TransBegin
' ... 执行操作 ...
' 等待用户输入（事务保持打开状态）
' ...
db.TransCommit
```

### 5. 使用批量操作的事务

```vb
' ? 推荐：批量操作使用事务
db.TransBegin

Dim i As Long
For i = 1 To 1000
    If Not db.Sql("INSERT INTO users (name) VALUES ('用户" & i & "')").Exec Then
        db.TransRollback
        Exit For
    End If
Next

If i > 1000 Then
    db.TransCommit
End If
```

---

## 常见问题

### Q1: 事务提交后数据没有保存？

**原因**: 可能是数据库不支持事务，或者连接字符串配置问题。

**解决**: 
- 检查数据库类型是否支持事务
- 验证连接字符串配置

### Q2: 如何检查事务状态？

```vb
' 注意：类库内部使用 IsTrans 标记，但不对外暴露
' 可以通过尝试提交来检查
If db.TransCommit Then
    Debug.Print "有事务且提交成功"
Else
    Debug.Print "没有事务或提交失败"
End If
```

### Q3: 事务会影响性能吗？

**回答**: 是的，事务会锁定资源，影响并发性能。

**建议**: 
- 尽量缩短事务时间
- 只包含必要的操作
- 避免在事务中进行长时间操作

---

**最后更新**: 2026-01-21
