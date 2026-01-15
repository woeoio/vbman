# 连接管理

本文档详细介绍 cDataBase 类的数据库连接管理功能，包括连接、断开、连接池等。

---

## ? 目录

- [数据库连接](#数据库连接)
- [支持的数据库类型](#支持的数据库类型)
- [连接字符串配置](#连接字符串配置)
- [连接状态管理](#连接状态管理)
- [连接池管理](#连接池管理)
- [错误处理](#错误处理)

---

## 数据库连接

### Connect 方法

`Connect` 方法用于建立数据库连接。

#### 语法

```vb
Function Connect( _
    ByVal DbType As enumDbType, _
    Optional ByVal DbAddress As String = "127.0.0.1,1433", _
    Optional ByVal username As String = "sa", _
    Optional ByVal password As String = "Sa123456", _
    Optional ByVal DefaultDataBase As String = "master") As Boolean
```

#### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `DbType` | `enumDbType` | 数据库类型（必需） |
| `DbAddress` | `String` | 数据库地址（可选，默认 "127.0.0.1,1433"） |
| `username` | `String` | 用户名（可选，默认 "sa"） |
| `password` | `String` | 密码（可选，默认 "Sa123456"） |
| `DefaultDataBase` | `String` | 默认数据库（可选，默认 "master"） |

#### 返回值

- `True` - 连接成功
- `False` - 连接失败（可通过 `LastErr` 查看错误信息）

#### 示例

```vb
Dim db As New VBMAN.cDataBase

' SQL Server 连接
If db.Connect(VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "password", "mydb") Then
    Debug.Print "连接成功"
Else
    Debug.Print "连接失败: " & db.LastErr
End If
```

---

## 支持的数据库类型

### enumDbType 枚举

```vb
Public Enum enumDbType
    Access = 1    ' Microsoft Access
    Mysql = 2     ' MySQL
    MsSql = 3     ' Microsoft SQL Server
    Csv = 4       ' CSV 文件
End Enum
```

### SQL Server 连接

```vb
' 基本连接
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "password", "master"

' 使用命名实例
db.Connect VBMAN.enumDbType_MsSql, "SERVER\INSTANCE,1433", "sa", "password", "mydb"

' 使用 Windows 身份验证（需要修改连接字符串）
' 注意：需要在类内部修改连接字符串以支持 Windows 身份验证
```

### MySQL 连接

```vb
' 基本连接（默认端口 3306）
db.Connect VBMAN.enumDbType_Mysql, "localhost:3306", "root", "password", "testdb"

' 指定端口
db.Connect VBMAN.enumDbType_Mysql, "192.168.1.100:3306", "user", "pwd", "mydb"

' 使用中文逗号（自动转换）
db.Connect VBMAN.enumDbType_Mysql, "localhost，3306", "root", "pwd", "testdb"
```

### Access 连接

```vb
' 连接 Access 数据库文件
db.Connect VBMAN.enumDbType_Access, "C:\data\mydb.mdb"

' 使用相对路径（自动转换为绝对路径）
db.Connect VBMAN.enumDbType_Access, "data\mydb.mdb"
```

### CSV 连接

```vb
' 连接 CSV 文件目录
db.Connect VBMAN.enumDbType_Csv, "C:\data\csvfiles"

' CSV 文件会被当作表来查询
db.Sql("SELECT * FROM data.csv").Query
```

---

## 连接字符串配置

### 自动生成的连接字符串

类库会根据数据库类型自动生成连接字符串：

#### SQL Server

```
Driver={SQL Server};Server=127.0.0.1,1433;Uid=sa;pwd=password;Database=mydb;
```

#### MySQL

```
Driver={MySQL ODBC 5.1 Driver};Server=localhost:3306;Uid=root;pwd=password;Database=testdb;
```

#### Access

```
Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\data\mydb.mdb;
```

#### CSV

```
Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=C:\data\csvfiles
```

### 自定义连接字符串

如果需要使用自定义连接字符串，可以在连接前设置：

```vb
' 注意：需要在类内部修改以支持自定义连接字符串
' 或者直接使用 ADO Connection 对象
Set db.Conn = New ADODB.Connection
db.Conn.ConnectionString = "Provider=SQLOLEDB;Data Source=...;..."
db.Conn.Open
db.IsConnect = True
```

---

## 连接状态管理

### IsConnect 属性

`IsConnect` 属性表示当前连接状态。

```vb
' 检查连接状态
If db.IsConnect Then
    Debug.Print "已连接"
Else
    Debug.Print "未连接"
End If
```

### CheckConnection 方法

`CheckConnection` 方法检查连接状态并尝试重连。

```vb
' 检查并重连
If Not db.CheckConnection Then
    Debug.Print "连接已断开，尝试重连..."
    db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
End If
```

### 连接状态检查

在执行操作前检查连接状态：

```vb
If Not db.IsConnect Then
    Debug.Print "数据库未连接"
    Exit Sub
End If

' 执行查询
db.Sql("SELECT * FROM users").Query
```

---

## 连接池管理

### ConnInst 方法

`ConnInst` 方法创建或获取连接池中的数据库实例。

#### 语法

```vb
Function ConnInst(ByVal InstName As String, Optional ByVal IsCloneMasterConnection As Boolean = True) As cDataBase
```

#### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `InstName` | `String` | 实例名称（必需） |
| `IsCloneMasterConnection` | `Boolean` | 是否克隆主连接（可选，默认 True） |

#### 示例

```vb
Dim db As New VBMAN.cDataBase

' 主连接
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"

' 创建连接实例 1（克隆主连接）
Dim db1 As VBMAN.cDataBase
Set db1 = db.ConnInst("db1", True)
db1.Sql("SELECT * FROM table1").Query

' 创建连接实例 2（独立连接）
Dim db2 As VBMAN.cDataBase
Set db2 = db.ConnInst("db2", False)
db2.Connect VBMAN.enumDbType_Mysql, "localhost:3306", "root", "pwd", "testdb"
db2.Sql("SELECT * FROM table2").Query

' 获取已存在的实例
Set db1 = db.ConnInst("db1")
```

### ConnInstRemove 方法

`ConnInstRemove` 方法移除连接池中的实例。

#### 语法

```vb
Sub ConnInstRemove(Optional ByVal InstName As String)
```

#### 示例

```vb
' 移除指定实例
db.ConnInstRemove "db1"

' 移除所有实例
db.ConnInstRemove ""
```

### 连接池使用场景

#### 场景 1：多数据库操作

```vb
Dim db As New VBMAN.cDataBase

' 主数据库
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"

' 日志数据库
Dim dbLog As VBMAN.cDataBase
Set dbLog = db.ConnInst("log", False)
dbLog.Connect VBMAN.enumDbType_Mysql, "192.168.1.100:3306", "loguser", "pwd", "logdb"

' 使用不同的数据库
db.Sql("SELECT * FROM users").Query
dbLog.Sql("INSERT INTO logs (msg) VALUES ('test')").Exec
```

#### 场景 2：动态数据库切换

```vb
Dim db As New VBMAN.cDataBase

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
    End Select
    
    Set GetDatabase = dbInst
End Function

' 使用
Dim dbMain As VBMAN.cDataBase
Set dbMain = GetDatabase("main")
dbMain.Sql("SELECT * FROM users").Query
```

---

## 断开连接

### Disconnect 方法

`Disconnect` 方法断开数据库连接。

#### 语法

```vb
Function Disconnect() As Boolean
```

#### 功能

- 关闭数据库连接
- 关闭记录集
- 释放资源
- 自动回滚未完成的事务

#### 示例

```vb
' 断开连接
If db.Disconnect Then
    Debug.Print "已断开连接"
End If
```

### 自动断开

类在销毁时会自动断开连接：

```vb
Private Sub Form_Unload(Cancel As Integer)
    ' 不需要手动调用 Disconnect，类会自动处理
    Set db = Nothing
End Sub
```

---

## 错误处理

### 错误属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `LastErr` | `String` | 最后错误信息（完整描述） |
| `LastErrNumber` | `Long` | 最后错误代码 |
| `LastErrDescription` | `String` | 最后错误描述 |

### 错误处理示例

```vb
' 连接错误处理
If Not db.Connect(VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb") Then
    Debug.Print "连接失败"
    Debug.Print "错误代码: " & db.LastErrNumber
    Debug.Print "错误描述: " & db.LastErrDescription
    Debug.Print "完整信息: " & db.LastErr
    Exit Sub
End If

' 查询错误处理
If Not db.Sql("SELECT * FROM users").Query Then
    Debug.Print "查询失败: " & db.LastErr
    Exit Sub
End If
```

### 常见连接错误

#### 错误 1：无法连接到服务器

```
错误代码: -2147467259
错误描述: [Microsoft][ODBC SQL Server Driver][DBNETLIB]SQL Server 不存在或访问被拒绝
```

**解决方案**：
- 检查数据库服务是否启动
- 验证服务器地址和端口
- 检查防火墙设置

#### 错误 2：登录失败

```
错误代码: -2147467259
错误描述: [Microsoft][ODBC SQL Server Driver][SQL Server]用户 'sa' 登录失败
```

**解决方案**：
- 验证用户名和密码
- 检查 SQL Server 身份验证模式
- 确认用户权限

#### 错误 3：数据库不存在

```
错误代码: -2147467259
错误描述: [Microsoft][ODBC SQL Server Driver][SQL Server]无法打开登录所请求的数据库
```

**解决方案**：
- 验证数据库名称
- 检查数据库是否存在
- 确认用户有访问权限

---

## 最佳实践

### 1. 连接管理

```vb
' ? 推荐：在需要时连接，使用完毕后断开
Private Sub ProcessData()
    Dim db As New VBMAN.cDataBase
    
    If db.Connect(VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb") Then
        ' 执行操作
        db.Sql("SELECT * FROM users").Query
    End If
    
    db.Disconnect
    Set db = Nothing
End Sub
```

### 2. 连接复用

```vb
' ? 推荐：对于频繁操作，保持连接
Private m_DB As VBMAN.cDataBase

Private Sub Form_Load()
    Set m_DB = New VBMAN.cDataBase
    m_DB.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_DB Is Nothing Then
        m_DB.Disconnect
        Set m_DB = Nothing
    End If
End Sub
```

### 3. 错误处理

```vb
' ? 推荐：始终检查返回值并处理错误
If Not db.Connect(...) Then
    MsgBox "连接失败: " & db.LastErr, vbCritical
    Exit Sub
End If
```

---

**最后更新**: 2026-01-21
