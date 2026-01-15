# 高级功能

本文档介绍 cDataBase 类的高级功能和最佳实践，帮助您更好地使用数据库类库。

---

## ? 目录

- [异步执行](#异步执行)
- [连接池管理](#连接池管理)
- [批量操作优化](#批量操作优化)
- [性能优化技巧](#性能优化技巧)
- [错误处理策略](#错误处理策略)
- [设计模式应用](#设计模式应用)
- [常见问题解决](#常见问题解决)

---

## 异步执行

### Async 属性

使用 `Async` 属性可以异步执行 SQL 语句，不阻塞主线程。

#### 语法

```vb
Property Get Async() As cDataBase
```

#### 示例

```vb
' 异步执行 INSERT
db.Sql("INSERT INTO logs (message) VALUES ('test')").Async.Exec

' 监听完成事件
Private Sub db_AsyncExecuteComplete(ByVal RecordsAffected As Long, _
                                    ByVal pError As ADODB.Error, _
                                    adStatus As ADODB.EventStatusEnum, _
                                    ByVal pCommand As ADODB.Command, _
                                    ByVal pRecordset As ADODB.Recordset, _
                                    ByVal pConnection As ADODB.Connection)
    If pError Is Nothing Then
        Debug.Print "异步执行成功，影响行数: " & RecordsAffected
    Else
        Debug.Print "异步执行失败: " & pError.Description
    End If
End Sub
```

### 异步执行场景

```vb
' 场景 1：日志记录（不阻塞主流程）
Sub LogMessage(sMessage As String)
    db.Sql("INSERT INTO logs (message, created_at) VALUES (?, ?)") _
        .Param("message", sMessage, VBMAN.adVarWChar) _
        .Param("created_at", Now, VBMAN.adDate) _
        .Async.ExecParam
    ' 不等待执行完成，继续执行后续代码
End Sub

' 场景 2：批量数据处理
Sub ProcessLargeDataset()
    ' 主线程继续处理
    ProcessData
    
    ' 异步保存结果
    db.Sql("INSERT INTO results SELECT * FROM temp_table").Async.Exec
End Sub
```

---

## 连接池管理

### 多数据库连接

使用连接池管理多个数据库连接。

```vb
Dim db As New VBMAN.cDataBase

' 主数据库
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "maindb"

' 日志数据库
Dim dbLog As VBMAN.cDataBase
Set dbLog = db.ConnInst("log", False)
dbLog.Connect VBMAN.enumDbType_Mysql, "192.168.1.100:3306", "loguser", "pwd", "logdb"

' 缓存数据库
Dim dbCache As VBMAN.cDataBase
Set dbCache = db.ConnInst("cache", False)
dbCache.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "cachedb"

' 使用不同的数据库
db.Sql("SELECT * FROM users").Query
dbLog.Sql("INSERT INTO logs (msg) VALUES ('test')").Exec
dbCache.Sql("SELECT * FROM cache_data").Query
```

### 动态连接管理

```vb
' 根据配置动态创建连接
Function GetDatabase(sConfigName As String) As VBMAN.cDataBase
    Dim dbInst As VBMAN.cDataBase
    Set dbInst = db.ConnInst(sConfigName, False)
    
    ' 根据配置连接不同的数据库
    Select Case sConfigName
    Case "main"
        dbInst.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "maindb"
    Case "backup"
        dbInst.Connect VBMAN.enumDbType_Mysql, "192.168.1.100:3306", "user", "pwd", "backupdb"
    Case "readonly"
        dbInst.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "readonly", "pwd", "maindb"
    End Select
    
    Set GetDatabase = dbInst
End Function

' 使用
Dim dbMain As VBMAN.cDataBase
Set dbMain = GetDatabase("main")
dbMain.Sql("SELECT * FROM users").Query
```

### 连接清理

```vb
' 移除指定连接
db.ConnInstRemove "log"

' 移除所有连接
db.ConnInstRemove ""
```

---

## 批量操作优化

### 批量插入优化

```vb
' 方法 1：使用 BatchInsert（推荐）
Sub BatchInsertUsers(colUsers As Collection)
    If db.BatchInsert("users", colUsers) Then
        Debug.Print "批量插入成功"
    End If
End Sub

' 方法 2：使用事务 + 循环
Sub BatchInsertUsers2(colUsers As Collection)
    db.TransBegin
    
    Dim i As Long
    For i = 1 To colUsers.Count
        Dim dictUser As Scripting.Dictionary
        Set dictUser = colUsers(i)
        
        Dim sSql As String
        sSql = "INSERT INTO users (name, age) VALUES ('" & _
               db.Escape(dictUser("name")) & "', " & dictUser("age") & ")"
        
        If Not db.Sql(sSql).Exec Then
            db.TransRollback
            Exit Sub
        End If
    Next
    
    db.TransCommit
End Sub

' 方法 3：使用 VALUES 子句（SQL Server/MySQL）
Sub BatchInsertUsers3(colUsers As Collection)
    Dim sSql As String
    sSql = "INSERT INTO users (name, age) VALUES "
    
    Dim i As Long
    For i = 1 To colUsers.Count
        Dim dictUser As Scripting.Dictionary
        Set dictUser = colUsers(i)
        
        If i > 1 Then sSql = sSql & ", "
        sSql = sSql & "('" & db.Escape(dictUser("name")) & "', " & dictUser("age") & ")"
    Next
    
    db.Sql(sSql).Exec
End Sub
```

### 批量更新优化

```vb
' 使用事务批量更新
Sub BatchUpdateUsers(colUpdates As Collection)
    db.TransBegin
    
    Dim i As Long
    For i = 1 To colUpdates.Count
        Dim dictUpdate As Scripting.Dictionary
        Set dictUpdate = colUpdates(i)
        
        If Not db.Sql("UPDATE users SET name = ?, age = ? WHERE id = ?") _
            .Param("name", dictUpdate("name"), VBMAN.adVarWChar) _
            .Param("age", dictUpdate("age"), VBMAN.adInteger) _
            .Param("id", dictUpdate("id"), VBMAN.adInteger) _
            .ExecParam Then
            
            db.TransRollback
            Exit Sub
        End If
    Next
    
    db.TransCommit
End Sub
```

---

## 性能优化技巧

### 1. 使用索引字段

```vb
' ? 推荐：使用索引字段作为条件
db.Sql("SELECT * FROM users WHERE id = 1").Query

' ? 不推荐：使用非索引字段
db.Sql("SELECT * FROM users WHERE name = '张三'").Query
```

### 2. 只查询需要的字段

```vb
' ? 推荐：只查询需要的字段
db.Sql("SELECT id, name FROM users").Query

' ? 不推荐：查询所有字段
db.Sql("SELECT * FROM users").Query
```

### 3. 使用合适的游标类型

```vb
' ? 推荐：只读查询使用 ForwardOnly
db.Sql("SELECT * FROM users").Query adOpenForwardOnly, adLockReadOnly

' ? 不推荐：使用默认游标（可能较慢）
db.Sql("SELECT * FROM users").Query
```

### 4. 限制结果集大小

```vb
' ? 推荐：使用 TOP/LIMIT 限制结果
db.Sql("SELECT TOP 100 * FROM users").Query

' ? 不推荐：查询所有数据
db.Sql("SELECT * FROM users").Query
```

### 5. 使用分页

```vb
' ? 推荐：使用分页
db.Sql("SELECT * FROM users").Page(1, 20).Query

' ? 不推荐：一次性查询所有数据
db.Sql("SELECT * FROM users").Query
```

### 6. 缓存查询结果

```vb
' 缓存查询结果
Private m_colCachedUsers As Collection
Private m_dtCacheTime As Date

Function GetUsers() As Collection
    ' 缓存 5 分钟
    If DateDiff("s", m_dtCacheTime, Now) > 300 Or m_colCachedUsers Is Nothing Then
        If db.Sql("SELECT * FROM users").Fetch Then
            Set m_colCachedUsers = db.Rows
            m_dtCacheTime = Now
        End If
    End If
    
    Set GetUsers = m_colCachedUsers
End Function
```

---

## 错误处理策略

### 统一错误处理

```vb
' 统一错误处理函数
Function ExecuteSQL(sSql As String) As Boolean
    On Error GoTo ErrHandler
    
    If db.Sql(sSql).Exec Then
        ExecuteSQL = True
    Else
        LogError "SQL执行失败", db.LastErr
        ExecuteSQL = False
    End If
    
    Exit Function
    
ErrHandler:
    LogError "发生异常", Err.Description
    ExecuteSQL = False
End Function

' 错误日志记录
Sub LogError(sOperation As String, sError As String)
    ' 记录到文件或数据库
    Debug.Print Now & " - " & sOperation & ": " & sError
End Sub
```

### 重试机制

```vb
' 带重试的查询
Function QueryWithRetry(sSql As String, Optional lMaxRetries As Long = 3) As Boolean
    Dim lRetry As Long
    For lRetry = 1 To lMaxRetries
        If db.Sql(sSql).Query Then
            QueryWithRetry = True
            Exit Function
        End If
        
        ' 检查连接
        If Not db.CheckConnection Then
            db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
        End If
        
        ' 等待后重试
        Sleep 1000
    Next
    
    QueryWithRetry = False
End Function
```

---

## 设计模式应用

### 单例模式

```vb
' 数据库单例
Private m_DB As VBMAN.cDataBase

Function GetDatabase() As VBMAN.cDataBase
    If m_DB Is Nothing Then
        Set m_DB = New VBMAN.cDataBase
        m_DB.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
    End If
    
    Set GetDatabase = m_DB
End Function
```

### 工厂模式

```vb
' 数据库工厂
Function CreateDatabase(sType As String) As VBMAN.cDataBase
    Dim db As New VBMAN.cDataBase
    
    Select Case sType
    Case "main"
        db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "maindb"
    Case "log"
        db.Connect VBMAN.enumDbType_Mysql, "192.168.1.100:3306", "loguser", "pwd", "logdb"
    Case "cache"
        db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "cachedb"
    End Select
    
    Set CreateDatabase = db
End Function
```

### 仓储模式

```vb
' 用户仓储
Class cUserRepository
    Private m_DB As VBMAN.cDataBase
    
    Private Sub Class_Initialize()
        Set m_DB = New VBMAN.cDataBase
        m_DB.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
    End Sub
    
    Function GetById(lId As Long) As Scripting.Dictionary
        If m_DB.Sql("SELECT * FROM users WHERE id = ?") _
            .Param("id", lId, VBMAN.adInteger) _
            .QueryParam Then
            
            If m_DB.Rows.Count > 0 Then
                Set GetById = m_DB.Row
            End If
        End If
    End Function
    
    Function GetAll() As Collection
        If m_DB.Sql("SELECT * FROM users").Fetch Then
            Set GetAll = m_DB.Rows
        End If
    End Function
    
    Function Save(dictUser As Scripting.Dictionary) As Boolean
        If m_DB.Sql("INSERT INTO users (name, age) VALUES (?, ?)") _
            .Param("name", dictUser("name"), VBMAN.adVarWChar) _
            .Param("age", dictUser("age"), VBMAN.adInteger) _
            .ExecParam Then
            
            Save = True
        Else
            Save = False
        End If
    End Function
End Class
```

---

## 常见问题解决

### Q1: 连接超时

**问题**: 连接数据库时超时。

**解决**:
```vb
' 设置连接超时
db.Conn.ConnectionTimeout = 30  ' 30 秒
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
```

### Q2: 查询超时

**问题**: 查询执行时间过长。

**解决**:
```vb
' 设置命令超时
db.Conn.CommandTimeout = 60  ' 60 秒
db.Sql("SELECT * FROM large_table").Query
```

### Q3: 内存不足

**问题**: 查询大量数据导致内存不足。

**解决**:
```vb
' 使用分页查询
db.Sql("SELECT * FROM large_table").Page(1, 1000).Query

' 或使用流式处理
If db.Sql("SELECT * FROM large_table").Query Then
    Do Until db.Rs.EOF
        ' 处理单条记录
        ProcessRecord db.Rs
        db.Rs.MoveNext
    Loop
End If
```

### Q4: 并发冲突

**问题**: 多个操作同时执行导致冲突。

**解决**:
```vb
' 使用事务和锁定
db.TransBegin
db.Sql("SELECT * FROM users WHERE id = 1").Query adOpenKeyset, adLockPessimistic
' 处理数据
db.Sql("UPDATE users SET ...").Exec
db.TransCommit
```

---

## 最佳实践总结

### 1. 连接管理

- ? 在需要时连接，使用完毕后断开
- ? 对于频繁操作，保持连接
- ? 使用连接池管理多个连接

### 2. 查询优化

- ? 使用索引字段作为条件
- ? 只查询需要的字段
- ? 使用分页限制结果集
- ? 使用合适的游标类型

### 3. 安全防护

- ? 始终使用参数化查询
- ? 验证用户输入
- ? 使用事务保证一致性

### 4. 错误处理

- ? 始终检查返回值
- ? 记录错误日志
- ? 实现重试机制

### 5. 代码组织

- ? 使用设计模式
- ? 封装常用操作
- ? 保持代码清晰

---

**最后更新**: 2026-01-21
