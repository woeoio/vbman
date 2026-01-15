# 分页功能

本文档详细介绍 cDataBase 类的分页查询功能，支持 SQL Server、MySQL、Access 等数据库。

---

## ? 目录

- [分页概述](#分页概述)
- [Page 方法](#page-方法)
- [支持的数据库](#支持的数据库)
- [使用示例](#使用示例)
- [性能优化](#性能优化)
- [常见问题](#常见问题)

---

## 分页概述

### 什么是分页

分页是将大量数据分成多个页面显示的技术，每页显示固定数量的记录。

### 分页的优势

- **性能优化** - 只查询需要的数据，减少内存占用
- **用户体验** - 快速加载，避免长时间等待
- **资源节约** - 减少网络传输和数据库负载

### 分页原理

```
总记录数: 1000
每页显示: 10
总页数: 100

第 1 页: 记录 1-10   (OFFSET 0, LIMIT 10)
第 2 页: 记录 11-20  (OFFSET 10, LIMIT 10)
第 3 页: 记录 21-30  (OFFSET 20, LIMIT 10)
...
```

---

## Page 方法

### 语法

```vb
Function Page(Optional num As Long = 1, Optional Limit As Long = 10) As cDataBase
```

### 参数说明

| 参数 | 类型 | 说明 |
|------|------|------|
| `num` | `Long` | 页码（可选，默认 1） |
| `Limit` | `Long` | 每页记录数（可选，默认 10） |

### 返回值

返回 `cDataBase` 对象，支持链式调用。

### 示例

```vb
' 查询第 1 页，每页 10 条
db.Sql("SELECT * FROM users").Page(1, 10).Query

' 查询第 2 页，每页 20 条
db.Sql("SELECT * FROM users").Page(2, 20).Query
```

---

## 支持的数据库

### SQL Server (2012+)

使用 `OFFSET ... ROWS FETCH NEXT ... ROWS ONLY` 语法。

```vb
' 原始 SQL
db.Sql("SELECT * FROM users").Page(2, 10).Query

' 自动转换为
SELECT * FROM users 
ORDER BY (SELECT NULL)
OFFSET 10 ROWS FETCH NEXT 10 ROWS ONLY
```

**注意**: SQL Server 2012 及以上版本才支持 `OFFSET FETCH`。

### MySQL

使用 `LIMIT ... OFFSET ...` 语法。

```vb
' 原始 SQL
db.Sql("SELECT * FROM users").Page(2, 10).Query

' 自动转换为
SELECT * FROM users 
LIMIT 10 OFFSET 10
```

### Access

使用 `TOP` 和子查询（简化实现）。

```vb
' 原始 SQL
db.Sql("SELECT * FROM users").Page(2, 10).Query

' 自动转换为
SELECT TOP 20 * FROM (
    SELECT * FROM users
) AS T
```

**注意**: Access 的分页实现是简化版本，可能不是完全准确的分页。

---

## 使用示例

### 示例 1：基本分页

```vb
' 查询第 1 页，每页 10 条
If db.Sql("SELECT * FROM users").Page(1, 10).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 示例 2：带排序的分页

```vb
' 按年龄降序排列，分页显示
If db.Sql("SELECT * FROM users ORDER BY age DESC").Page(2, 10).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name") & " - " & db.Rows(i)("age")
    Next
End If
```

### 示例 3：条件查询分页

```vb
' 查询年龄大于 18 的用户，分页显示
If db.Sql("SELECT * FROM users WHERE age > 18").Page(1, 20).Fetch Then
    Dim i As Long
    For i = 1 To db.Rows.Count
        Debug.Print db.Rows(i)("name")
    Next
End If
```

### 示例 4：获取总记录数

```vb
' 查询总数
Dim lTotal As Long
lTotal = db.Count("users")

' 计算总页数
Dim lPageSize As Long
lPageSize = 10
Dim lTotalPages As Long
lTotalPages = Int((lTotal + lPageSize - 1) / lPageSize)

' 分页查询
If db.Sql("SELECT * FROM users").Page(1, lPageSize).Fetch Then
    ' 显示数据
End If
```

### 示例 5：完整分页函数

```vb
' 分页查询函数
Function GetUsersPage(lPage As Long, lPageSize As Long) As Collection
    Set GetUsersPage = New Collection
    
    ' 查询数据
    If db.Sql("SELECT * FROM users ORDER BY id").Page(lPage, lPageSize).Fetch Then
        Set GetUsersPage = db.Rows
    End If
End Function

' 使用
Dim colUsers As Collection
Set colUsers = GetUsersPage(1, 10)
```

### 示例 6：分页导航

```vb
' 分页导航类
Private m_lCurrentPage As Long
Private m_lPageSize As Long
Private m_lTotalRecords As Long

Private Sub LoadPage(lPage As Long)
    ' 验证页码
    If lPage < 1 Then lPage = 1
    
    Dim lTotalPages As Long
    lTotalPages = Int((m_lTotalRecords + m_lPageSize - 1) / m_lPageSize)
    If lPage > lTotalPages Then lPage = lTotalPages
    
    m_lCurrentPage = lPage
    
    ' 查询数据
    If db.Sql("SELECT * FROM users ORDER BY id").Page(lPage, m_lPageSize).Fetch Then
        ' 显示数据
        DisplayUsers
    End If
End Sub

Private Sub cmdNextPage_Click()
    LoadPage m_lCurrentPage + 1
End Sub

Private Sub cmdPrevPage_Click()
    LoadPage m_lCurrentPage - 1
End Sub
```

---

## 性能优化

### 1. 使用索引字段排序

```vb
' ? 推荐：使用索引字段排序
db.Sql("SELECT * FROM users ORDER BY id").Page(1, 10).Query

' ? 不推荐：使用非索引字段排序
db.Sql("SELECT * FROM users ORDER BY name").Page(1, 10).Query
```

### 2. 只查询需要的字段

```vb
' ? 推荐：只查询需要的字段
db.Sql("SELECT id, name FROM users").Page(1, 10).Query

' ? 不推荐：查询所有字段
db.Sql("SELECT * FROM users").Page(1, 10).Query
```

### 3. 使用 WHERE 条件限制

```vb
' ? 推荐：使用 WHERE 条件
db.Sql("SELECT * FROM users WHERE status = 'active'").Page(1, 10).Query

' ? 不推荐：查询所有数据再分页
db.Sql("SELECT * FROM users").Page(1, 10).Query
```

### 4. 合理设置每页记录数

```vb
' ? 推荐：合理的每页记录数（10-50）
db.Sql("SELECT * FROM users").Page(1, 20).Query

' ? 不推荐：每页记录数过大
db.Sql("SELECT * FROM users").Page(1, 1000).Query
```

---

## 常见问题

### Q1: SQL Server 分页报错

**错误**: `'OFFSET' 附近有语法错误`

**原因**: SQL Server 版本低于 2012，不支持 `OFFSET FETCH`。

**解决**: 
- 升级到 SQL Server 2012 或更高版本
- 或使用 `ROW_NUMBER()` 实现分页（需要修改类库代码）

### Q2: Access 分页不准确

**原因**: Access 的分页实现是简化版本，使用 `TOP` 查询。

**解决**: 
- 对于 Access，建议使用 `ROW_NUMBER()` 或手动实现分页
- 或升级到 SQL Server/MySQL

### Q3: 如何获取总记录数？

```vb
' 方法 1：使用 Count 方法
Dim lTotal As Long
lTotal = db.Count("users")

' 方法 2：使用 COUNT(*) 查询
If db.Sql("SELECT COUNT(*) AS cnt FROM users").Fetch Then
    lTotal = db.Row("cnt")
End If
```

### Q4: 分页后如何保持排序？

```vb
' ? 正确：在 SQL 中包含 ORDER BY
db.Sql("SELECT * FROM users ORDER BY age DESC").Page(1, 10).Query

' ? 错误：分页后再排序（会丢失排序）
db.Sql("SELECT * FROM users").Page(1, 10).Query
' 然后对结果排序（只对当前页排序，不是全局排序）
```

### Q5: 如何实现跳转到指定页？

```vb
Function GoToPage(lPage As Long, lPageSize As Long) As Boolean
    ' 验证页码
    If lPage < 1 Then lPage = 1
    
    ' 查询指定页
    If db.Sql("SELECT * FROM users ORDER BY id").Page(lPage, lPageSize).Fetch Then
        GoToPage = True
    Else
        GoToPage = False
    End If
End Function
```

---

## 最佳实践

### 1. 始终使用 ORDER BY

```vb
' ? 推荐：使用 ORDER BY 保证顺序
db.Sql("SELECT * FROM users ORDER BY id").Page(1, 10).Query

' ? 不推荐：不使用 ORDER BY（顺序不确定）
db.Sql("SELECT * FROM users").Page(1, 10).Query
```

### 2. 验证页码和每页记录数

```vb
' ? 推荐：验证参数
Function GetPage(lPage As Long, lPageSize As Long) As Collection
    If lPage < 1 Then lPage = 1
    If lPageSize < 1 Then lPageSize = 10
    If lPageSize > 100 Then lPageSize = 100  ' 限制最大每页记录数
    
    If db.Sql("SELECT * FROM users ORDER BY id").Page(lPage, lPageSize).Fetch Then
        Set GetPage = db.Rows
    End If
End Function
```

### 3. 缓存总记录数

```vb
' ? 推荐：缓存总记录数，避免频繁查询
Private m_lCachedTotal As Long
Private m_dtCacheTime As Date

Function GetTotalRecords() As Long
    ' 缓存 5 分钟
    If DateDiff("s", m_dtCacheTime, Now) > 300 Or m_lCachedTotal = 0 Then
        m_lCachedTotal = db.Count("users")
        m_dtCacheTime = Now
    End If
    GetTotalRecords = m_lCachedTotal
End Function
```

---

**最后更新**: 2026-01-21
