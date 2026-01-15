# 工具方法

本文档详细介绍 cDataBase 类提供的各种工具方法，用于简化常用数据库操作。

---

## ? 目录

- [Count 方法](#count-方法)
- [LastInsertId 方法](#lastinsertid-方法)
- [TableExists 方法](#tableexists-方法)
- [GetTableFields 方法](#gettablefields-方法)
- [GetTables 方法](#gettables-方法)
- [GetDatabases 方法](#getdatabases-方法)
- [GetVersion 方法](#getversion-方法)
- [Escape 方法](#escape-方法)
- [CheckConnection 方法](#checkconnection-方法)

---

## Count 方法

### 语法

统计表中的记录数。

```vb
Function Count(Optional ByVal TableName As String = "") As Long
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `TableName` | `String` | 表名（可选，为空时使用当前 SQL） |

### 返回值

返回记录数（`Long` 类型）。

### 示例

```vb
' 统计指定表的记录数
Dim lCount As Long
lCount = db.Count("users")
Debug.Print "用户总数: " & lCount

' 统计当前 SQL 查询结果的记录数
db.Sql("SELECT * FROM users WHERE age > 18")
lCount = db.Count()  ' 使用当前 SQL
Debug.Print "成年用户数: " & lCount
```

### 实现原理

```vb
' 如果指定了表名
SELECT COUNT(*) AS cnt FROM users

' 如果使用当前 SQL
SELECT COUNT(*) AS cnt FROM (SELECT * FROM users WHERE age > 18) AS T
```

---

## LastInsertId 方法

### 语法

获取最后插入的自增主键 ID。

```vb
Function LastInsertId() As Variant
```

### 返回值

返回最后插入的 ID（`Variant` 类型）。

### 支持的数据库

| 数据库 | 实现方式 |
|--------|----------|
| SQL Server | `SCOPE_IDENTITY()` |
| MySQL | `LAST_INSERT_ID()` |
| Access | `@@IDENTITY` |

### 示例

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

## TableExists 方法

### 语法

判断表是否存在。

```vb
Function TableExists(ByVal TableName As String) As Boolean
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `TableName` | `String` | 表名（必需） |

### 返回值

- `True` - 表存在
- `False` - 表不存在

### 支持的数据库

| 数据库 | 实现方式 |
|--------|----------|
| SQL Server | `INFORMATION_SCHEMA.TABLES` |
| MySQL | `INFORMATION_SCHEMA.TABLES` |
| Access | `MSysObjects` |

### 示例

```vb
' 检查表是否存在
If db.TableExists("users") Then
    Debug.Print "users 表存在"
Else
    Debug.Print "users 表不存在"
    ' 创建表
    db.Sql("CREATE TABLE users (id INT PRIMARY KEY, name NVARCHAR(50))").Exec
End If
```

### 使用场景

```vb
' 场景 1：创建表前检查
If Not db.TableExists("users") Then
    db.Sql("CREATE TABLE users (...)").Exec
End If

' 场景 2：删除表前检查
If db.TableExists("temp_table") Then
    db.Sql("DROP TABLE temp_table").Exec
End If
```

---

## GetTableFields 方法

### 语法

获取表的字段列表。

```vb
Function GetTableFields(ByVal TableName As String) As Collection
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `TableName` | `String` | 表名（必需） |

### 返回值

返回字段名集合（`Collection` 类型）。

### 支持的数据库

| 数据库 | 实现方式 |
|--------|----------|
| SQL Server | `INFORMATION_SCHEMA.COLUMNS` |
| MySQL | `INFORMATION_SCHEMA.COLUMNS` |
| Access | `MSysObjects` |

### 示例

```vb
' 获取字段列表
Dim colFields As Collection
Set colFields = db.GetTableFields("users")

' 遍历字段
Dim i As Long
For i = 1 To colFields.Count
    Debug.Print "字段 " & i & ": " & colFields(i)
Next
```

### 使用场景

```vb
' 场景 1：动态生成 SQL
Function BuildSelectSQL(sTableName As String) As String
    Dim colFields As Collection
    Set colFields = db.GetTableFields(sTableName)
    
    Dim sFields As String
    Dim i As Long
    For i = 1 To colFields.Count
        If sFields <> "" Then sFields = sFields & ", "
        sFields = sFields & colFields(i)
    Next
    
    BuildSelectSQL = "SELECT " & sFields & " FROM " & sTableName
End Function

' 场景 2：验证字段是否存在
Function FieldExists(sTableName As String, sFieldName As String) As Boolean
    Dim colFields As Collection
    Set colFields = db.GetTableFields(sTableName)
    
    Dim i As Long
    For i = 1 To colFields.Count
        If colFields(i) = sFieldName Then
            FieldExists = True
            Exit Function
        End If
    Next
    
    FieldExists = False
End Function
```

---

## GetTables 方法

### 语法

获取数据库中的表名列表。

```vb
Function GetTables(Optional ByVal DatabaseName As String = "") As Collection
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `DatabaseName` | `String` | 数据库名（可选，为空时使用当前数据库） |

### 返回值

返回表名集合（`Collection` 类型）。

### 支持的数据库

| 数据库 | 实现方式 |
|--------|----------|
| SQL Server | `INFORMATION_SCHEMA.TABLES` |
| MySQL | `INFORMATION_SCHEMA.TABLES` |
| Access | `MSysObjects` |

### 示例

```vb
' 获取当前数据库的表列表
Dim colTables As Collection
Set colTables = db.GetTables()

' 遍历表名
Dim i As Long
For i = 1 To colTables.Count
    Debug.Print "表 " & i & ": " & colTables(i)
Next

' 获取指定数据库的表列表（SQL Server/MySQL）
Set colTables = db.GetTables("mydb")
```

### 使用场景

```vb
' 场景 1：列出所有表
Sub ListAllTables()
    Dim colTables As Collection
    Set colTables = db.GetTables()
    
    Dim i As Long
    For i = 1 To colTables.Count
        Debug.Print colTables(i)
    Next
End Sub

' 场景 2：备份所有表
Sub BackupAllTables()
    Dim colTables As Collection
    Set colTables = db.GetTables()
    
    Dim i As Long
    For i = 1 To colTables.Count
        Dim sTableName As String
        sTableName = colTables(i)
        db.Sql("SELECT * INTO " & sTableName & "_backup FROM " & sTableName).Exec
    Next
End Sub
```

---

## GetDatabases 方法

### 语法

获取数据库服务器中的数据库列表。

```vb
Function GetDatabases() As Collection
```

### 返回值

返回数据库名集合（`Collection` 类型）。

### 支持的数据库

| 数据库 | 实现方式 |
|--------|----------|
| SQL Server | `sys.databases` |
| MySQL | `SHOW DATABASES` |

### 示例

```vb
' 获取数据库列表
Dim colDatabases As Collection
Set colDatabases = db.GetDatabases()

' 遍历数据库名
Dim i As Long
For i = 1 To colDatabases.Count
    Debug.Print "数据库 " & i & ": " & colDatabases(i)
Next
```

### 使用场景

```vb
' 场景 1：列出所有数据库
Sub ListAllDatabases()
    Dim colDatabases As Collection
    Set colDatabases = db.GetDatabases()
    
    Dim i As Long
    For i = 1 To colDatabases.Count
        Debug.Print colDatabases(i)
    Next
End Sub

' 场景 2：切换数据库
Sub SwitchDatabase(sDatabaseName As String)
    ' 检查数据库是否存在
    Dim colDatabases As Collection
    Set colDatabases = db.GetDatabases()
    
    Dim i As Long
    Dim bExists As Boolean
    bExists = False
    For i = 1 To colDatabases.Count
        If colDatabases(i) = sDatabaseName Then
            bExists = True
            Exit For
        End If
    Next
    
    If bExists Then
        db.Sql("USE " & sDatabaseName).Exec
    Else
        Debug.Print "数据库不存在: " & sDatabaseName
    End If
End Sub
```

---

## GetVersion 方法

### 语法

获取数据库版本信息。

```vb
Function GetVersion() As String
```

### 返回值

返回版本字符串（`String` 类型）。

### 示例

```vb
' 获取数据库版本
Dim sVersion As String
sVersion = db.GetVersion
Debug.Print "数据库版本: " & sVersion
```

### 使用场景

```vb
' 场景 1：检查数据库版本
Sub CheckDatabaseVersion()
    Dim sVersion As String
    sVersion = db.GetVersion
    Debug.Print "当前数据库版本: " & sVersion
End Sub

' 场景 2：版本兼容性检查
Function IsVersionCompatible(sMinVersion As String) As Boolean
    Dim sVersion As String
    sVersion = db.GetVersion
    ' 比较版本逻辑...
    IsVersionCompatible = True
End Function
```

---

## Escape 方法

### 语法

转义 SQL 字符串中的特殊字符（防止 SQL 注入）。

```vb
Function Escape(ByVal Str As String) As String
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `Str` | `String` | 要转义的字符串（必需） |

### 返回值

返回转义后的字符串（`String` 类型）。

### 实现原理

将单引号 `'` 转义为两个单引号 `''`。

### 示例

```vb
' 转义字符串
Dim sName As String
sName = "O'Brien"
Dim sEscaped As String
sEscaped = db.Escape(sName)
' 结果: "O''Brien"

' 使用转义后的字符串
db.Sql("SELECT * FROM users WHERE name = '" & sEscaped & "'").Query
```

### 注意事项

**推荐使用参数化查询而不是 Escape 方法**：

```vb
' ? 推荐：使用参数化查询
db.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", "O'Brien", VBMAN.adVarWChar) _
    .QueryParam

' ?? 可以使用：使用 Escape（但不如参数化查询安全）
db.Sql("SELECT * FROM users WHERE name = '" & db.Escape("O'Brien") & "'").Query
```

---

## CheckConnection 方法

### 语法

检查连接状态并尝试重连。

```vb
Function CheckConnection() As Boolean
```

### 返回值

- `True` - 连接正常或重连成功
- `False` - 连接断开且重连失败

### 功能

1. 检查连接状态
2. 如果断开，尝试重新连接
3. 更新连接标记

### 示例

```vb
' 检查连接
If Not db.CheckConnection Then
    Debug.Print "连接已断开，尝试重连..."
    ' 重新连接
    db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
End If
```

### 使用场景

```vb
' 场景 1：定期检查连接
Private Sub Timer1_Timer()
    If Not db.CheckConnection Then
        Debug.Print "连接断开，需要重新连接"
    End If
End Sub

' 场景 2：操作前检查连接
Sub ExecuteQuery()
    ' 检查连接
    If Not db.CheckConnection Then
        Debug.Print "连接不可用"
        Exit Sub
    End If
    
    ' 执行查询
    db.Sql("SELECT * FROM users").Query
End Sub
```

---

## 综合示例

### 示例 1：数据库信息查看器

```vb
Sub ShowDatabaseInfo()
    ' 显示数据库版本
    Debug.Print "数据库版本: " & db.GetVersion
    
    ' 显示数据库列表
    Dim colDatabases As Collection
    Set colDatabases = db.GetDatabases()
    Debug.Print "数据库列表:"
    Dim i As Long
    For i = 1 To colDatabases.Count
        Debug.Print "  - " & colDatabases(i)
    Next
    
    ' 显示表列表
    Dim colTables As Collection
    Set colTables = db.GetTables()
    Debug.Print "表列表:"
    For i = 1 To colTables.Count
        Debug.Print "  - " & colTables(i)
        
        ' 显示每个表的字段
        Dim colFields As Collection
        Set colFields = db.GetTableFields(colTables(i))
        Dim j As Long
        For j = 1 To colFields.Count
            Debug.Print "    * " & colFields(j)
        Next
    Next
End Sub
```

### 示例 2：表结构比较

```vb
Function CompareTableStructure(sTable1 As String, sTable2 As String) As Boolean
    Dim colFields1 As Collection
    Dim colFields2 As Collection
    Set colFields1 = db.GetTableFields(sTable1)
    Set colFields2 = db.GetTableFields(sTable2)
    
    ' 比较字段数量
    If colFields1.Count <> colFields2.Count Then
        CompareTableStructure = False
        Exit Function
    End If
    
    ' 比较字段名
    Dim i As Long
    For i = 1 To colFields1.Count
        If colFields1(i) <> colFields2(i) Then
            CompareTableStructure = False
            Exit Function
        End If
    Next
    
    CompareTableStructure = True
End Function
```

---

**最后更新**: 2026-01-21
