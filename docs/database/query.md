# 查询操作

本文档详细介绍 cDataBase 类的查询功能，包括基本查询、结果集处理等。

---

## ? 目录

- [基本查询](#基本查询)
- [查询方法](#查询方法)
- [结果集处理](#结果集处理)
- [查询选项](#查询选项)
- [常见查询场景](#常见查询场景)

---

## 基本查询

### Sql 方法

`Sql` 方法用于设置 SQL 查询语句。

#### 语法

```vb
Function Sql(ByVal RawSqlString As String) As cDataBase
```

#### 示例

```vb
' 设置 SQL 语句
db.Sql("SELECT * FROM users WHERE age > 18")
```

### Query 方法

`Query` 方法执行查询并返回布尔值。**查询结果存储在 `db.Rs` 属性中，而不是作为返回值**。

#### 语法

```vb
Function Query(Optional CurType As CursorTypeEnum = adOpenKeyset, _
              Optional LockType As LockTypeEnum = adLockOptimistic, _
              Optional Options As Long = -1) As Boolean
```

#### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `CurType` | `CursorTypeEnum` | 游标类型（可选，默认 `adOpenKeyset`） |
| `LockType` | `LockTypeEnum` | 锁定类型（可选，默认 `adLockOptimistic`） |
| `Options` | `Long` | 查询选项（可选，默认 -1） |

#### 返回值

- `True` - 查询成功，结果存储在 `db.Rs` 中
- `False` - 查询失败（可通过 `LastErr` 查看错误信息）

**重要**：`Query` 方法返回的是布尔值，不是 Recordset。查询结果需要通过 `db.Rs` 属性访问。

#### 示例

```vb
' 基本查询
If db.Sql("SELECT * FROM users").Query Then
    ' 查询成功，使用 db.Rs 访问 Recordset（不是返回值）
    Do Until db.Rs.EOF
        Debug.Print db.Rs("name")
        db.Rs.MoveNext
    Loop
    ' 使用完毕后关闭 Recordset
    db.Rs.Close
End If

' ? 错误示例：不要将 Query 的返回值赋给 Recordset 变量
' Dim Rs As ADODB.Recordset
' Set Rs = db.Query("SELECT * FROM users")  ' 错误！Query 返回布尔值

' ? 正确示例：使用 db.Rs 访问结果
If db.Sql("SELECT * FROM users").Query Then
    ' 使用 db.Rs 访问结果
    Do Until db.Rs.EOF
        Debug.Print db.Rs("name")
        db.Rs.MoveNext
    Loop
    db.Rs.Close
End If
```

### Fetch 方法

`Fetch` 方法执行查询并自动转换为 Dictionary 集合。**主要场景是用于 JSON 输出和数据交换**，特别适合与 `cHttpServer` 等组件配合使用。

#### 语法

```vb
Function Fetch(Optional CurType As CursorTypeEnum = adOpenKeyset, _
               Optional LockType As LockTypeEnum = adLockOptimistic, _
               Optional Options As Long = -1) As Boolean
```

#### 功能

- 执行查询
- 自动将 Recordset 转换为 Collection（存储在 `Rows` 属性中）
- 自动设置第一行到 `Row` 属性
- **可直接用于 JSON 序列化**，配合 `cJson` 和 `cHttpServerResponse` 使用

#### 示例

```vb
' 查询并自动转换
If db.Sql("SELECT * FROM users WHERE age > 18").Fetch Then
    ' 访问第一行
    Debug.Print db.Row("name")
    Debug.Print db.Row("age")
    
    ' 遍历所有行
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

#### JSON 输出场景（主要用途）

`Fetch` 方法的主要优势是方便将查询结果转换为 JSON 字符串，用于外部数据交换：

```vb
' 在 HttpServer 路由中使用
Private Sub Server_OnRoute(ctx As VBMAN.cHttpServerContext)
    ' 查询数据
    If ctx.Db.Sql("SELECT * FROM users").Fetch Then
        ' 直接输出 JSON（一句话完成）
        ctx.Response.Json ctx.Db.Rows
    End If
End Sub

' 或者使用 cJson 对象
Dim json As New VBMAN.cJson
If db.Sql("SELECT * FROM users").Fetch Then
    Dim sJson As String
    sJson = json.Encode(db.Rows)  ' 转换为 JSON 字符串
    Debug.Print sJson
End If
```

**注意**：`cHttpServerResponse.Json` 方法会自动识别 Recordset 并转换，但使用 `Fetch` 后直接传入 `db.Rows`（Collection）更高效，因为已经完成转换。

---

## 查询方法对比

| 方法 | 返回类型 | 自动转换 | 使用场景 |
|------|----------|----------|----------|
| `Query` | `Recordset` | 否 | 需要直接操作 Recordset，需要 Recordset 的高级功能 |
| `Fetch` | `Collection` | 是 | **主要用于 JSON 输出和数据交换**，配合 cHttpServer、cJson 使用 |

---

## 结果集处理

### Rs 属性

`Rs` 属性提供对 ADO Recordset 的直接访问。

```vb
' 使用 Recordset
If db.Sql("SELECT * FROM users").Query Then
    Do Until db.Rs.EOF
        Debug.Print db.Rs("name")
        Debug.Print db.Rs("age")
        db.Rs.MoveNext
    Loop
    
    ' 获取记录数
    Debug.Print "总记录数: " & db.Rs.RecordCount
End If
```

### Rows 属性

`Rows` 属性是转换后的 Collection，包含所有行的 Dictionary。**可以直接用于 JSON 序列化**。

```vb
' 使用 Collection
If db.Sql("SELECT * FROM users").Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
        Debug.Print db.Rows(i)("age")
    Next
End If
```

#### JSON 输出示例

```vb
' 场景 1：在 HttpServer 中直接输出 JSON
Private Sub Server_OnRoute(ctx As VBMAN.cHttpServerContext)
    If ctx.Db.Sql("SELECT * FROM users").Fetch Then
        ' 直接输出为 JSON（推荐方式）
        ctx.Response.Json ctx.Db.Rows
    End If
End Sub

' 场景 2：使用 cJson 对象转换为 JSON 字符串
Dim json As New VBMAN.cJson
If db.Sql("SELECT * FROM users").Fetch Then
    Dim sJson As String
    sJson = json.Encode(db.Rows)
    ' 输出: [{"id":1,"name":"张三","age":25},{"id":2,"name":"李四","age":30}]
End If

' 场景 3：带分页的 JSON 输出
If db.Sql("SELECT * FROM users").Page(1, 10).Fetch Then
    Dim lTotal As Long
    lTotal = db.Count("users")
    ' 输出带总数和分页信息的 JSON
    ctx.Response.Json db.Rows, 200, "成功", lTotal
End If
```

### Row 属性

`Row` 属性是当前第一行的 Dictionary。

```vb
' 访问第一行
If db.Sql("SELECT * FROM users WHERE id = 1").Fetch Then
    Debug.Print db.Row("name")
    Debug.Print db.Row("age")
End If
```

### 结果集结构

```
Rows (Collection)
├── Rows(1) (Dictionary)
│   ├── "id" => 1
│   ├── "name" => "张三"
│   └── "age" => 25
├── Rows(2) (Dictionary)
│   ├── "id" => 2
│   ├── "name" => "李四"
│   └── "age" => 30
└── ...
```

---

## 查询选项

### 游标类型 (CursorTypeEnum)

| 类型 | 值 | 说明 |
|------|-----|------|
| `adOpenForwardOnly` | 0 | 仅向前游标（最快） |
| `adOpenKeyset` | 1 | 键集游标（默认） |
| `adOpenDynamic` | 2 | 动态游标 |
| `adOpenStatic` | 3 | 静态游标 |

```vb
' 使用仅向前游标（性能最佳）
db.Sql("SELECT * FROM users").Query adOpenForwardOnly

' 使用静态游标（支持 RecordCount）
db.Sql("SELECT * FROM users").Query adOpenStatic
```

### 锁定类型 (LockTypeEnum)

| 类型 | 值 | 说明 |
|------|-----|------|
| `adLockReadOnly` | 1 | 只读（默认查询） |
| `adLockPessimistic` | 2 | 悲观锁定 |
| `adLockOptimistic` | 3 | 乐观锁定 |
| `adLockBatchOptimistic` | 4 | 批量乐观锁定 |

```vb
' 只读查询（性能最佳）
db.Sql("SELECT * FROM users").Query adOpenKeyset, adLockReadOnly
```

---

## 常见查询场景

### 场景 1：单条记录查询

```vb
' 查询单条记录
If db.Sql("SELECT * FROM users WHERE id = 1").Fetch Then
    If db.Rows.Count > 0 Then
        Debug.Print "用户名: " & db.Row("name")
        Debug.Print "年龄: " & db.Row("age")
    Else
        Debug.Print "未找到记录"
    End If
End If
```

### 场景 2：条件查询

```vb
' 多条件查询
Dim sSql As String
sSql = "SELECT * FROM users WHERE age > 18 AND status = 'active'"
If db.Sql(sSql).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 场景 3：排序查询

```vb
' 按年龄降序排列
If db.Sql("SELECT * FROM users ORDER BY age DESC").Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name") & " - " & db.Rows(i)("age")
    Next
End If
```

### 场景 4：聚合查询

```vb
' 统计查询
If db.Sql("SELECT COUNT(*) AS cnt, AVG(age) AS avg_age FROM users").Fetch Then
    If db.Rows.Count > 0 Then
        Debug.Print "总用户数: " & db.Row("cnt")
        Debug.Print "平均年龄: " & db.Row("avg_age")
    End If
End If
```

### 场景 5：分组查询

```vb
' 分组统计
If db.Sql("SELECT status, COUNT(*) AS cnt FROM users GROUP BY status").Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("status") & ": " & db.Rows(i)("cnt")
    Next
End If
```

### 场景 6：联表查询

```vb
' 内连接查询
Dim sSql As String
sSql = "SELECT u.name, p.title " & _
       "FROM users u " & _
       "INNER JOIN posts p ON u.id = p.user_id"
If db.Sql(sSql).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name") & " - " & db.Rows(i)("title")
    Next
End If
```

### 场景 7：子查询

```vb
' 子查询
Dim sSql As String
sSql = "SELECT * FROM users " & _
       "WHERE id IN (SELECT user_id FROM orders WHERE amount > 1000)"
If db.Sql(sSql).Fetch Then
    ' 处理结果
End If
```

### 场景 8：模糊查询

```vb
' LIKE 查询
If db.Sql("SELECT * FROM users WHERE name LIKE '%张%'").Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 场景 9：分页查询

```vb
' 使用 Page 方法（详见 pagination.md）
If db.Sql("SELECT * FROM users").Page(1, 10).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 场景 10：参数化查询

```vb
' 使用参数化查询（详见 parameterized.md）
If db.Sql("SELECT * FROM users WHERE name = ? AND age > ?") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .Param("age", 18, VBMAN.adInteger) _
    .QueryParam Then
    
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 场景 11：JSON 输出（Fetch 的主要场景）

```vb
' 在 HttpServer 路由中输出 JSON
Private Sub Server_OnRoute(ctx As VBMAN.cHttpServerContext)
    ' 查询数据并转换为 Collection
    If ctx.Db.Sql("SELECT * FROM users WHERE status = 'active'").Fetch Then
        ' 直接输出 JSON（一句话完成）
        ctx.Response.Json ctx.Db.Rows
    End If
End Sub

' 带分页的 JSON API
Private Sub Server_OnRoute(ctx As VBMAN.cHttpServerContext)
    Dim lPage As Long
    Dim lPageSize As Long
    lPage = CLng(ctx.Request.Query("page"))
    lPageSize = CLng(ctx.Request.Query("pageSize"))
    
    ' 查询分页数据
    If ctx.Db.Sql("SELECT * FROM users ORDER BY id").Page(lPage, lPageSize).Fetch Then
        Dim lTotal As Long
        lTotal = ctx.Db.Count("users")
        ' 输出 JSON，包含数据、总数等信息
        ctx.Response.Json ctx.Db.Rows, 200, "成功", lTotal
    End If
End Sub

' 使用 cJson 对象转换为 JSON 字符串
Dim json As New VBMAN.cJson
If db.Sql("SELECT * FROM users").Fetch Then
    Dim sJson As String
    sJson = json.Encode(db.Rows)
    ' 可以用于文件保存、网络传输等
    Debug.Print sJson
End If
```

---

## 性能优化

### 1. 只查询需要的字段

```vb
' ? 不推荐：查询所有字段
db.Sql("SELECT * FROM users").Query

' ? 推荐：只查询需要的字段
db.Sql("SELECT id, name FROM users").Query
```

### 2. 使用索引字段

```vb
' ? 推荐：使用索引字段作为条件
db.Sql("SELECT * FROM users WHERE id = 1").Query
```

### 3. 限制结果集大小

```vb
' ? 推荐：使用 TOP 限制结果
db.Sql("SELECT TOP 100 * FROM users").Query
```

### 4. 使用合适的游标类型

```vb
' ? 推荐：只读查询使用 ForwardOnly
db.Sql("SELECT * FROM users").Query adOpenForwardOnly, adLockReadOnly
```

---

## 错误处理

```vb
' 查询错误处理
If Not db.Sql("SELECT * FROM users").Query Then
    Debug.Print "查询失败"
    Debug.Print "错误代码: " & db.LastErrNumber
    Debug.Print "错误描述: " & db.LastErrDescription
    Debug.Print "完整信息: " & db.LastErr
    Exit Sub
End If
```

---

## 最佳实践

### 1. 始终检查返回值

```vb
' ? 推荐
If db.Sql("SELECT * FROM users").Query Then
    ' 处理结果
Else
    ' 处理错误
End If
```

### 2. 使用 Fetch 进行 JSON 输出

```vb
' ? 推荐：使用 Fetch 进行 JSON 输出（主要场景）
If db.Sql("SELECT * FROM users").Fetch Then
    ' 在 HttpServer 中直接输出
    ctx.Response.Json db.Rows

    ' 只输出一行记录的
    ctx.Response.Json db.Row
    
    ' 或使用 cJson 对象
    Dim json As New VBMAN.cJson
    Dim sJson As String
    sJson = json.Encode(db.Rows)
End If
```

### 3. 及时释放资源

```vb
' ? 推荐：查询完成后关闭 Recordset
If db.Sql("SELECT * FROM users").Query Then
    ' 处理结果
    If db.Rs.State <> adStateClosed Then
        db.Rs.Close
    End If
End If
```

---

**最后更新**: 2026-01-21
