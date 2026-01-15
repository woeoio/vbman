# 执行操作

本文档详细介绍 cDataBase 类的数据修改操作，包括 INSERT、UPDATE、DELETE 等。

---

## ? 目录

- [Exec 方法](#exec-方法)
- [INSERT 操作](#insert-操作)
- [UPDATE 操作](#update-操作)
- [DELETE 操作](#delete-操作)
- [获取影响行数](#获取影响行数)
- [获取最后插入ID](#获取最后插入id)
- [批量操作](#批量操作)

---

## Exec 方法

### 基本语法

`Exec` 方法用于执行 INSERT、UPDATE、DELETE 等非查询 SQL 语句。

#### 语法

```vb
Function Exec(Optional RecordsAffected, Optional Options As Long = -1) As Boolean
```

#### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `RecordsAffected` | `Variant` | 返回受影响的行数（可选） |
| `Options` | `Long` | 执行选项（可选，默认 -1） |

#### 返回值

- `True` - 执行成功
- `False` - 执行失败（可通过 `LastErr` 查看错误信息）

#### 示例

```vb
' 执行 INSERT 语句
If db.Sql("INSERT INTO users (name, age) VALUES ('张三', 25)").Exec Then
    Debug.Print "插入成功"
Else
    Debug.Print "插入失败: " & db.LastErr
End If
```

---

## INSERT 操作

### 基本插入

```vb
' 插入单条记录
If db.Sql("INSERT INTO users (name, age, email) VALUES ('张三', 25, 'zhangsan@example.com')").Exec Then
    Debug.Print "插入成功"
End If
```

### 插入多条记录

```vb
' 使用 VALUES 子句插入多条
Dim sSql As String
sSql = "INSERT INTO users (name, age) VALUES " & _
       "('张三', 25), " & _
       "('李四', 30), " & _
       "('王五', 28)"
If db.Sql(sSql).Exec Then
    Debug.Print "批量插入成功"
End If
```

### 使用参数化插入

```vb
' 使用参数化查询（推荐，防止 SQL 注入）
If db.Sql("INSERT INTO users (name, age, email) VALUES (?, ?, ?)") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .Param("age", 25, VBMAN.adInteger) _
    .Param("email", "zhangsan@example.com", VBMAN.adVarWChar) _
    .ExecParam Then
    
    Debug.Print "插入成功，ID: " & db.LastInsertId
End If
```

### INSERT SELECT

```vb
' 从其他表插入数据
If db.Sql("INSERT INTO users_backup SELECT * FROM users WHERE age > 30").Exec Then
    Debug.Print "数据复制成功"
End If
```

---

## UPDATE 操作

### 基本更新

```vb
' 更新单条记录
If db.Sql("UPDATE users SET age = 26 WHERE id = 1").Exec Then
    Debug.Print "更新成功"
End If
```

### 更新多条记录

```vb
' 批量更新
If db.Sql("UPDATE users SET status = 'active' WHERE age > 18").Exec Then
    Debug.Print "批量更新成功"
End If
```

### 使用参数化更新

```vb
' 使用参数化查询
If db.Sql("UPDATE users SET age = ?, email = ? WHERE id = ?") _
    .Param("age", 26, VBMAN.adInteger) _
    .Param("email", "newemail@example.com", VBMAN.adVarWChar) _
    .Param("id", 1, VBMAN.adInteger) _
    .ExecParam Then
    
    Debug.Print "更新成功"
End If
```

### 条件更新

```vb
' 复杂条件更新
Dim sSql As String
sSql = "UPDATE users SET status = 'inactive' " & _
       "WHERE last_login < DATEADD(day, -30, GETDATE())"
If db.Sql(sSql).Exec Then
    Debug.Print "过期用户已标记"
End If
```

---

## DELETE 操作

### 基本删除

```vb
' 删除单条记录
If db.Sql("DELETE FROM users WHERE id = 1").Exec Then
    Debug.Print "删除成功"
End If
```

### 批量删除

```vb
' 删除多条记录
If db.Sql("DELETE FROM users WHERE age < 18").Exec Then
    Debug.Print "批量删除成功"
End If
```

### 使用参数化删除

```vb
' 使用参数化查询
If db.Sql("DELETE FROM users WHERE id = ?") _
    .Param("id", 1, VBMAN.adInteger) _
    .ExecParam Then
    
    Debug.Print "删除成功"
End If
```

### 清空表

```vb
' 清空表（注意：会删除所有数据）
If db.Sql("DELETE FROM users").Exec Then
    Debug.Print "表已清空"
End If

' 或者使用 TRUNCATE（更快，但不可回滚）
If db.Sql("TRUNCATE TABLE users").Exec Then
    Debug.Print "表已清空"
End If
```

---

## 获取影响行数

### RecordsAffected 参数

```vb
' 获取受影响的行数
Dim lAffected As Long
If db.Sql("UPDATE users SET status = 'active' WHERE age > 18").Exec(lAffected) Then
    Debug.Print "更新了 " & lAffected & " 条记录"
End If
```

### 示例

```vb
' INSERT 操作
Dim lAffected As Long
If db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec(lAffected) Then
    Debug.Print "插入了 " & lAffected & " 条记录"
End If

' UPDATE 操作
If db.Sql("UPDATE users SET status = 'active'").Exec(lAffected) Then
    Debug.Print "更新了 " & lAffected & " 条记录"
End If

' DELETE 操作
If db.Sql("DELETE FROM users WHERE age < 18").Exec(lAffected) Then
    Debug.Print "删除了 " & lAffected & " 条记录"
End If
```

---

## 获取最后插入ID

### LastInsertId 方法

`LastInsertId` 方法获取最后插入的自增主键 ID。

#### 语法

```vb
Function LastInsertId() As Variant
```

#### 支持的数据库

- SQL Server - 使用 `SCOPE_IDENTITY()`
- MySQL - 使用 `LAST_INSERT_ID()`
- Access - 使用 `@@IDENTITY`

#### 示例

```vb
' 插入数据
If db.Sql("INSERT INTO users (name, age) VALUES ('张三', 25)").Exec Then
    ' 获取最后插入的 ID
    Dim lId As Variant
    lId = db.LastInsertId
    Debug.Print "新用户 ID: " & lId
End If
```

### 注意事项

1. **必须在 INSERT 后立即调用**
   ```vb
   ' ? 正确
   db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
   Dim lId As Variant
   lId = db.LastInsertId
   
   ' ? 错误：中间有其他操作
   db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
   db.Sql("SELECT * FROM users").Query  ' 这会重置 ID
   Dim lId As Variant
   lId = db.LastInsertId  ' 可能获取不到正确的 ID
   ```

2. **仅支持自增主键**
   ```vb
   ' 表必须有自增主键
   CREATE TABLE users (
       id INT IDENTITY(1,1) PRIMARY KEY,  -- SQL Server
       name NVARCHAR(50)
   )
   ```

---

## 批量操作

### BatchInsert 方法

`BatchInsert` 方法批量插入数据，使用事务保证数据一致性。

#### 语法

```vb
Function BatchInsert(ByVal TableName As String, ByVal Data As Collection) As Boolean
```

#### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `TableName` | `String` | 表名（必需） |
| `Data` | `Collection` | 数据集合，每个元素是 Dictionary（必需） |

#### 示例

```vb
' 准备数据
Dim colData As New Collection
Dim dictRow As Scripting.Dictionary

' 第一行
Set dictRow = New Scripting.Dictionary
dictRow.Add "name", "张三"
dictRow.Add "age", 25
dictRow.Add "email", "zhangsan@example.com"
colData.Add dictRow

' 第二行
Set dictRow = New Scripting.Dictionary
dictRow.Add "name", "李四"
dictRow.Add "age", 30
dictRow.Add "email", "lisi@example.com"
colData.Add dictRow

' 批量插入
If db.BatchInsert("users", colData) Then
    Debug.Print "批量插入成功"
Else
    Debug.Print "批量插入失败: " & db.LastErr
End If
```

### 批量更新

```vb
' 使用事务批量更新
db.TransBegin

Dim i As Long
For i = 1 To 100
    Dim sSql As String
    sSql = "UPDATE users SET status = 'active' WHERE id = " & i
    If Not db.Sql(sSql).Exec Then
        db.TransRollback
        Exit For
    End If
Next

If db.TransCommit Then
    Debug.Print "批量更新成功"
End If
```

---

## 异步执行

### Async 属性

使用 `Async` 属性可以异步执行 SQL 语句。

```vb
' 异步执行
db.Sql("INSERT INTO users (name) VALUES ('张三')").Async.Exec

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

---

## 错误处理

```vb
' 执行错误处理
If Not db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec Then
    Debug.Print "执行失败"
    Debug.Print "错误代码: " & db.LastErrNumber
    Debug.Print "错误描述: " & db.LastErrDescription
    Debug.Print "完整信息: " & db.LastErr
    Exit Sub
End If
```

---

## 最佳实践

### 1. 使用参数化查询

```vb
' ? 推荐：使用参数化查询
db.Sql("INSERT INTO users (name, age) VALUES (?, ?)") _
    .Param("name", txtName.Text, VBMAN.adVarWChar) _
    .Param("age", CLng(txtAge.Text), VBMAN.adInteger) _
    .ExecParam

' ? 不推荐：直接拼接 SQL（容易 SQL 注入）
db.Sql("INSERT INTO users (name, age) VALUES ('" & txtName.Text & "', " & txtAge.Text & ")").Exec
```

### 2. 使用事务保证一致性

```vb
' ? 推荐：使用事务
db.TransBegin
db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec
db.Sql("INSERT INTO posts (user_id, title) VALUES (1, '标题')").Exec
If Not db.TransCommit Then
    Debug.Print "事务失败，已回滚"
End If
```

### 3. 检查返回值

```vb
' ? 推荐：始终检查返回值
If db.Sql("INSERT INTO users (name) VALUES ('张三')").Exec Then
    ' 成功处理
Else
    ' 错误处理
End If
```

### 4. 获取影响行数

```vb
' ? 推荐：检查影响行数
Dim lAffected As Long
If db.Sql("UPDATE users SET status = 'active'").Exec(lAffected) Then
    If lAffected > 0 Then
        Debug.Print "更新了 " & lAffected & " 条记录"
    Else
        Debug.Print "没有记录被更新"
    End If
End If
```

---

**最后更新**: 2026-01-21
