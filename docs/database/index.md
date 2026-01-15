# 数据库类库开发文档

> ? **cDataBase 类库** - 基于 ADO 封装的 VB6 数据库操作类，由 215879458@qq.com 开发

## ? 目录

- [概述](#概述)
- [核心亮点](#核心亮点)
- [架构设计](#架构设计)
- [文档索引](#文档索引)

---

## 概述

cDataBase 类库是一个为 VB6 设计的轻量级数据库操作封装类，完全基于 ADO (ActiveX Data Objects) 实现，提供了简洁易用的 API 和完整的功能支持。

### ? 主要特性

- ? **多数据库支持** - 支持 Access、MySQL、SQL Server、CSV 等多种数据库
- ? **连接池管理** - 支持多数据库连接实例池，动态管理数据库对象
- ? **分页查询** - 内置分页功能，支持 SQL Server、MySQL、Access
- ? **参数化查询** - 支持参数化查询，有效防止 SQL 注入攻击
- ? **事务处理** - 完整的事务支持，自动回滚机制
- ? **异步执行** - 支持异步 SQL 执行，提升性能
- ? **结果集转换** - 自动将 Recordset 转换为 Dictionary 集合
- ?? **工具方法** - 丰富的工具方法，简化常用操作

---

## 核心亮点

### 1?? 简洁的 API 设计 ?

类库采用链式调用设计，代码简洁易读：

```vb
' 在 VB6 项目中引用 VBMAN.dll 后使用
Dim db As New VBMAN.cDataBase

' 连接数据库
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "password", "mydb"

' 查询数据（链式调用）
db.Sql("SELECT * FROM users WHERE id > ?").Param("id", 100).QueryParam

' 获取结果
Dim i As Long
For i = 1 To db.Rows.Count
    Debug.Print db.Rows(i)("name")
Next
```

---

### 2?? 多数据库类型支持 ?

```vb
' SQL Server
db.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "master"

' MySQL
db.Connect VBMAN.enumDbType_Mysql, "localhost:3306", "root", "pwd", "testdb"

' Access
db.Connect VBMAN.enumDbType_Access, "C:\data\mydb.mdb"

' CSV
db.Connect VBMAN.enumDbType_Csv, "C:\data\csvfiles"
```

---

### 3?? 分页查询支持 ?

```vb
' 第 2 页，每页 10 条
db.Sql("SELECT * FROM users").Page(2, 10).Query

' 自动转换为对应数据库的分页 SQL
' SQL Server: OFFSET 10 ROWS FETCH NEXT 10 ROWS ONLY
' MySQL: LIMIT 10 OFFSET 10
```

---

### 4?? 参数化查询（防 SQL 注入） ?

```vb
' 使用参数化查询，安全可靠
db.Sql("SELECT * FROM users WHERE name = ? AND age > ?") _
    .Param("name", "张三", VBMAN.adVarChar) _
    .Param("age", 18, VBMAN.adInteger) _
    .QueryParam
```

---

### 5?? 事务处理 ?

```vb
' 开始事务
db.TransBegin

' 执行多个操作
db.Sql("INSERT INTO users (name) VALUES ('user1')").Exec
db.Sql("INSERT INTO users (name) VALUES ('user2')").Exec

' 提交事务（失败自动回滚）
If db.TransCommit Then
    Debug.Print "事务提交成功"
End If
```

---

### 6?? 结果集自动转换 ?

```vb
' 查询后自动转换为 Dictionary 集合
db.Sql("SELECT * FROM users").Fetch

' 访问第一行数据
Debug.Print db.Row("name")
Debug.Print db.Row("age")

' 遍历所有行
Dim i As Long
For i = 1 To db.Rows.Count
    Debug.Print db.Rows(i)("name")
Next
```

---

### 7?? 连接池管理 ?

```vb
' 创建连接实例
Dim db1 As VBMAN.cDataBase
Set db1 = db.ConnInst("db1")

' 使用独立的连接实例
db1.Sql("SELECT * FROM table1").Query

' 移除连接实例
db.ConnInstRemove "db1"
```

---

## 架构设计

### 类层次结构

```
cDataBase (数据库操作类)
├── Connection (ADODB.Connection) - 数据库连接
├── Recordset (ADODB.Recordset) - 记录集
├── Command (ADODB.Command) - 参数化查询命令
└── Connections (Dictionary) - 连接池
```

### 对象关系图

```
数据库对象 (cDataBase)
├── 主连接 (Conn)
│   ├── Recordset (Rs)
│   └── Command (Cmd) - 参数化查询
└── 连接池 (Connections)
    ├── 连接实例 1 (cDataBase)
    │   └── 独立连接
    ├── 连接实例 2 (cDataBase)
    │   └── 独立连接
    └── ...
```

### 数据流程

#### 查询流程

```
1. Sql("SELECT ...") - 设置 SQL 语句
2. Page(1, 10) - 可选：设置分页
3. Query() - 执行查询
4. Rs - 获取 Recordset
5. Rows - 自动转换为 Dictionary 集合
```

#### 执行流程

```
1. Sql("INSERT/UPDATE/DELETE ...") - 设置 SQL 语句
2. Param("name", value) - 可选：添加参数
3. Exec() / ExecParam() - 执行操作
4. LastInsertId() - 可选：获取最后插入的 ID
```

---

## 文档索引

| 文档 | 描述 |
|------|------|
| [快速入门](./quickstart.md) | 快速上手指南 |
| [连接管理](./connection.md) | 数据库连接和连接池管理 |
| [查询操作](./query.md) | SELECT 查询的详细说明 |
| [执行操作](./execute.md) | INSERT、UPDATE、DELETE 操作 |
| [事务处理](./transaction.md) | 事务的完整使用指南 |
| [分页功能](./pagination.md) | 分页查询的详细说明 |
| [参数化查询](./parameterized.md) | 参数化查询和 SQL 注入防护 |
| [工具方法](./utilities.md) | 常用工具方法说明 |
| [高级功能](./advanced.md) | 高级功能和最佳实践 |

---

## 依赖关系

| 组件 | 描述 |
| ------------------------------ | ------------------------------------------------------- |
| **Microsoft ActiveX Data Objects 2.8 Library** | ADO 核心库，必须引用 |
| **Microsoft Scripting Runtime** | Dictionary 对象支持 |
| **ToolsList.bas** | Recordset 转 Collection 工具 |
| **ToolsFso.bas** | 文件路径处理工具 |

---

## 兼容性

- **VB6/VBA** - 完全兼容
- **Windows** - Windows XP 及以上版本
- **数据库** - Access、MySQL、SQL Server、CSV
- **ADO 版本** - ADO 2.8 及以上

---

## 许可

基于 VBMAN 项目开发

---

## 作者

**数据库类库**: 215879458@qq.com

---

**最后更新**: 2026-01-21
