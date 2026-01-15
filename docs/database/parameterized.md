# 参数化查询

本文档详细介绍 cDataBase 类的参数化查询功能，这是防止 SQL 注入攻击的重要安全特性。

---

## ? 目录

- [参数化查询概述](#参数化查询概述)
- [Param 方法](#param-方法)
- [ExecParam 方法](#execparam-方法)
- [QueryParam 方法](#queryparam-方法)
- [数据类型](#数据类型)
- [安全优势](#安全优势)
- [使用示例](#使用示例)
- [最佳实践](#最佳实践)

---

## 参数化查询概述

### 什么是参数化查询

参数化查询是将 SQL 语句和参数值分开处理的技术，参数值通过占位符（通常是 `?`）传递。

### 为什么使用参数化查询

1. **防止 SQL 注入** - 参数值会被转义，无法执行恶意 SQL 代码
2. **性能优化** - 数据库可以缓存执行计划
3. **类型安全** - 自动处理数据类型转换
4. **代码清晰** - SQL 语句和参数值分离，易于维护

### SQL 注入示例

```vb
' ? 危险：直接拼接 SQL（容易 SQL 注入）
Dim sName As String
sName = "'; DROP TABLE users; --"
db.Sql("SELECT * FROM users WHERE name = '" & sName & "'").Query
' 实际执行的 SQL: SELECT * FROM users WHERE name = ''; DROP TABLE users; --'
' 结果：users 表被删除！

' ? 安全：使用参数化查询
db.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", sName, VBMAN.adVarWChar) _
    .QueryParam
' 参数值会被安全处理，无法执行恶意代码
```

---

## Param 方法

### 语法

```vb
Function Param(ByVal ParamName As String, _
               ByVal ParamValue As Variant, _
               Optional ByVal ParamType As DataTypeEnum = adVarChar) As cDataBase
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `ParamName` | `String` | 参数名称（用于标识，实际使用 `?` 占位符） |
| `ParamValue` | `Variant` | 参数值（必需） |
| `ParamType` | `DataTypeEnum` | 参数数据类型（可选，默认 `adVarChar`） |

### 返回值

返回 `cDataBase` 对象，支持链式调用。

### 示例

```vb
' 添加单个参数
db.Sql("SELECT * FROM users WHERE id = ?") _
    .Param("id", 1, VBMAN.adInteger)

' 添加多个参数（链式调用）
db.Sql("SELECT * FROM users WHERE name = ? AND age > ?") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .Param("age", 18, VBMAN.adInteger)
```

---

## ExecParam 方法

### 语法

`ExecParam` 方法执行带参数的 INSERT、UPDATE、DELETE 操作。

```vb
Function ExecParam(Optional RecordsAffected) As Boolean
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `RecordsAffected` | `Variant` | 返回受影响的行数（可选） |

### 返回值

- `True` - 执行成功
- `False` - 执行失败（可通过 `LastErr` 查看错误信息）

### 示例

```vb
' 插入数据
If db.Sql("INSERT INTO users (name, age, email) VALUES (?, ?, ?)") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .Param("age", 25, VBMAN.adInteger) _
    .Param("email", "zhangsan@example.com", VBMAN.adVarWChar) _
    .ExecParam Then
    
    Debug.Print "插入成功，ID: " & db.LastInsertId
End If

' 更新数据
If db.Sql("UPDATE users SET age = ?, email = ? WHERE id = ?") _
    .Param("age", 26, VBMAN.adInteger) _
    .Param("email", "newemail@example.com", VBMAN.adVarWChar) _
    .Param("id", 1, VBMAN.adInteger) _
    .ExecParam Then
    
    Debug.Print "更新成功"
End If

' 删除数据
If db.Sql("DELETE FROM users WHERE id = ?") _
    .Param("id", 1, VBMAN.adInteger) _
    .ExecParam Then
    
    Debug.Print "删除成功"
End If
```

---

## QueryParam 方法

### 语法

`QueryParam` 方法执行带参数的 SELECT 查询。

```vb
Function QueryParam(Optional CurType As CursorTypeEnum = adOpenKeyset, _
                    Optional LockType As LockTypeEnum = adLockOptimistic) As Boolean
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `CurType` | `CursorTypeEnum` | 游标类型（可选，默认 `adOpenKeyset`） |
| `LockType` | `LockTypeEnum` | 锁定类型（可选，默认 `adLockOptimistic`） |

### 返回值

- `True` - 查询成功，结果存储在 `db.Rs` 中
- `False` - 查询失败（可通过 `LastErr` 查看错误信息）

**重要**：`QueryParam` 方法返回的是布尔值，不是 Recordset。查询结果需要通过 `db.Rs` 属性访问。

### 示例

```vb
' 查询单条记录（使用 db.Rs 访问结果）
If db.Sql("SELECT * FROM users WHERE id = ?") _
    .Param("id", 1, VBMAN.adInteger) _
    .QueryParam Then
    
    ' 使用 db.Rs 访问 Recordset
    If Not db.Rs.EOF Then
        Debug.Print db.Rs("name")
    End If
    db.Rs.Close
End If

' 查询多条记录（使用 db.Rs 访问结果）
If db.Sql("SELECT * FROM users WHERE age > ? AND status = ?") _
    .Param("age", 18, VBMAN.adInteger) _
    .Param("status", "active", VBMAN.adVarWChar) _
    .QueryParam Then
    
    ' 使用 db.Rs 遍历结果
    Do Until db.Rs.EOF
        Debug.Print db.Rs("name")
        db.Rs.MoveNext
    Loop
    db.Rs.Close
End If

' 如果使用 Fetch 方法，会自动转换为 db.Rows 和 db.Row
If db.Sql("SELECT * FROM users WHERE id = ?") _
    .Param("id", 1, VBMAN.adInteger) _
    .QueryParam Then
    
    ' 使用 Fetch 后可以访问 db.Rows 和 db.Row
    If db.Fetch Then
        If db.Rows.Count > 0 Then
            Debug.Print db.Row("name")
        End If
    End If
End If
```

---

## 数据类型

### 常用数据类型

| 类型 | 值 | 说明 | 示例 |
|------|-----|------|------|
| `adVarChar` | 200 | 可变长度字符串 | "张三" |
| `adVarWChar` | 202 | Unicode 字符串 | "张三" |
| `adInteger` | 3 | 32 位整数 | 25 |
| `adBigInt` | 20 | 64 位整数 | 1234567890 |
| `adDouble` | 5 | 双精度浮点数 | 3.14 |
| `adDate` | 7 | 日期时间 | Now |
| `adBoolean` | 11 | 布尔值 | True |
| `adDecimal` | 14 | 精确数值 | 99.99 |

### 数据类型选择

```vb
' 字符串
db.Param("name", "张三", VBMAN.adVarWChar)  ' Unicode 字符串（推荐）
db.Param("name", "张三", VBMAN.adVarChar)   ' ANSI 字符串

' 整数
db.Param("age", 25, VBMAN.adInteger)        ' 32 位整数
db.Param("id", 1234567890, VBMAN.adBigInt)   ' 64 位整数

' 浮点数
db.Param("price", 99.99, VBMAN.adDouble)    ' 双精度
db.Param("amount", 99.99, VBMAN.adDecimal)  ' 精确数值

' 日期时间
db.Param("created", Now, VBMAN.adDate)      ' 日期时间

' 布尔值
db.Param("active", True, VBMAN.adBoolean)   ' 布尔值
```

---

## 安全优势

### SQL 注入防护

```vb
' ? 危险：直接拼接
Dim sInput As String
sInput = "'; DROP TABLE users; --"
db.Sql("SELECT * FROM users WHERE name = '" & sInput & "'").Query
' 结果：users 表被删除

' ? 安全：参数化查询
db.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", sInput, VBMAN.adVarWChar) _
    .QueryParam
' 参数值被安全处理，无法执行恶意代码
```

### 特殊字符处理

```vb
' ? 危险：特殊字符可能导致错误
Dim sName As String
sName = "O'Brien"
db.Sql("SELECT * FROM users WHERE name = '" & sName & "'").Query
' SQL: SELECT * FROM users WHERE name = 'O'Brien'
' 错误：单引号未转义

' ? 安全：参数化查询自动处理
db.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", sName, VBMAN.adVarWChar) _
    .QueryParam
' 特殊字符被自动转义
```

---

## 使用示例

### 示例 1：用户登录

```vb
Function UserLogin(sUsername As String, sPassword As String) As Boolean
    ' 使用参数化查询防止 SQL 注入
    If db.Sql("SELECT * FROM users WHERE username = ? AND password = ?") _
        .Param("username", sUsername, VBMAN.adVarWChar) _
        .Param("password", sPassword, VBMAN.adVarWChar) _
        .QueryParam Then
        
        If db.Rows.Count > 0 Then
            UserLogin = True
        Else
            UserLogin = False
        End If
    Else
        UserLogin = False
    End If
End Function
```

### 示例 2：搜索功能

```vb
Function SearchUsers(sKeyword As String) As Collection
    Set SearchUsers = New Collection
    
    ' 使用 LIKE 和参数化查询
    Dim sSql As String
    sSql = "SELECT * FROM users WHERE name LIKE ? OR email LIKE ?"
    
    If db.Sql(sSql) _
        .Param("name", "%" & sKeyword & "%", VBMAN.adVarWChar) _
        .Param("email", "%" & sKeyword & "%", VBMAN.adVarWChar) _
        .QueryParam Then
        
        Set SearchUsers = db.Rows
    End If
End Function
```

### 示例 3：批量插入

```vb
Sub BatchInsertUsers(colUsers As Collection)
    db.TransBegin
    
    Dim i As Long
    For i = 1 To colUsers.Count
        Dim dictUser As Scripting.Dictionary
        Set dictUser = colUsers(i)
        
        If Not db.Sql("INSERT INTO users (name, age, email) VALUES (?, ?, ?)") _
            .Param("name", dictUser("name"), VBMAN.adVarWChar) _
            .Param("age", dictUser("age"), VBMAN.adInteger) _
            .Param("email", dictUser("email"), VBMAN.adVarWChar) _
            .ExecParam Then
            
            db.TransRollback
            Exit Sub
        End If
    Next
    
    db.TransCommit
End Sub
```

### 示例 4：动态查询

```vb
Function GetUsers(Optional sName As String = "", _
                  Optional lMinAge As Long = 0, _
                  Optional sStatus As String = "") As Collection
    Set GetUsers = New Collection
    
    Dim sSql As String
    sSql = "SELECT * FROM users WHERE 1=1"
    
    ' 动态构建 SQL 和参数
    If sName <> "" Then
        sSql = sSql & " AND name LIKE ?"
    End If
    If lMinAge > 0 Then
        sSql = sSql & " AND age >= ?"
    End If
    If sStatus <> "" Then
        sSql = sSql & " AND status = ?"
    End If
    
    ' 设置 SQL
    db.Sql(sSql)
    
    ' 添加参数
    If sName <> "" Then
        db.Param("name", "%" & sName & "%", VBMAN.adVarWChar)
    End If
    If lMinAge > 0 Then
        db.Param("age", lMinAge, VBMAN.adInteger)
    End If
    If sStatus <> "" Then
        db.Param("status", sStatus, VBMAN.adVarWChar)
    End If
    
    ' 执行查询
    If db.QueryParam Then
        Set GetUsers = db.Rows
    End If
End Function
```

---

## 最佳实践

### 1. 始终使用参数化查询处理用户输入

```vb
' ? 推荐：使用参数化查询
db.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", txtName.Text, VBMAN.adVarWChar) _
    .QueryParam

' ? 不推荐：直接拼接用户输入
db.Sql("SELECT * FROM users WHERE name = '" & txtName.Text & "'").Query
```

### 2. 选择正确的数据类型

```vb
' ? 推荐：明确指定数据类型
db.Param("age", 25, VBMAN.adInteger)
db.Param("name", "张三", VBMAN.adVarWChar)
db.Param("price", 99.99, VBMAN.adDecimal)

' ? 不推荐：使用默认类型（可能类型不匹配）
db.Param("age", 25)  ' 默认 adVarChar，可能出错
```

### 3. 使用 Unicode 字符串类型

```vb
' ? 推荐：使用 adVarWChar 支持中文
db.Param("name", "张三", VBMAN.adVarWChar)

' ? 不推荐：使用 adVarChar（可能中文乱码）
db.Param("name", "张三", VBMAN.adVarChar)
```

### 4. 参数顺序要与 SQL 中的占位符顺序一致

```vb
' ? 正确：参数顺序与 ? 顺序一致
db.Sql("SELECT * FROM users WHERE name = ? AND age > ?") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .Param("age", 18, VBMAN.adInteger)

' ? 错误：参数顺序错误
db.Sql("SELECT * FROM users WHERE name = ? AND age > ?") _
    .Param("age", 18, VBMAN.adInteger) _
    .Param("name", "张三", VBMAN.adVarWChar)
```

### 5. 清理参数

```vb
' 注意：ExecParam 和 QueryParam 会自动清理参数
' 如果只调用 Param 而不执行，需要手动清理（通过执行或重新创建对象）
```

---

## 常见问题

### Q1: 参数化查询比直接拼接慢吗？

**回答**: 不会，参数化查询通常更快，因为：
- 数据库可以缓存执行计划
- 减少了 SQL 解析时间
- 避免了字符串拼接开销

### Q2: 如何在 LIKE 查询中使用参数？

```vb
' ? 正确：在参数值中包含通配符
db.Sql("SELECT * FROM users WHERE name LIKE ?") _
    .Param("name", "%" & sKeyword & "%", VBMAN.adVarWChar) _
    .QueryParam

' ? 错误：在 SQL 中使用通配符
db.Sql("SELECT * FROM users WHERE name LIKE '%?%'") _
    .Param("name", sKeyword, VBMAN.adVarWChar) _
    .QueryParam
```

### Q3: 如何处理 NULL 值？

```vb
' 使用 Null 值
db.Sql("SELECT * FROM users WHERE email = ?") _
    .Param("email", Null, VBMAN.adVarWChar) _
    .QueryParam

' 或者使用 IS NULL
db.Sql("SELECT * FROM users WHERE email IS NULL").Query
```

### Q4: 参数化查询支持 IN 子句吗？

```vb
' 注意：ADO 参数化查询对 IN 子句支持有限
' 建议：使用多个 OR 条件或动态构建 SQL

' 方法 1：使用多个 OR
db.Sql("SELECT * FROM users WHERE id = ? OR id = ? OR id = ?") _
    .Param("id1", 1, VBMAN.adInteger) _
    .Param("id2", 2, VBMAN.adInteger) _
    .Param("id3", 3, VBMAN.adInteger) _
    .QueryParam

' 方法 2：动态构建（需要验证输入）
Dim sIds As String
sIds = "1,2,3"  ' 需要验证格式
db.Sql("SELECT * FROM users WHERE id IN (" & sIds & ")").Query
```

---

**最后更新**: 2026-01-21
