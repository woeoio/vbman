# 快速入门

本指南将帮助您快速上手 cDataBase 类库，创建基本的数据库应用程序。

---

## ? 前置准备

### 必需文件

确保以下文件已添加到项目中：

| 文件 | 位置 | 说明 |
|------|------|------|
| `VBMAN.dll` | 项目引用 | 编译后的 COM 组件 |
| `Microsoft ActiveX Data Objects 2.8 Library` | 项目引用 | ADO 核心库 |

### 添加到项目

1. 打开 VB6 项目
2. 菜单：**项目** → **引用**
3. 勾选以下引用：
   - ? **VBMAN** (VBMAN.dll)
   - ? **Microsoft ActiveX Data Objects 2.8 Library**

---

## ? 客户端快速入门

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtLog`) - 显示日志（MultiLine = True）
- 1 个 CommandButton (`cmdConnect`) - 连接数据库
- 1 个 CommandButton (`cmdQuery`) - 执行查询

### 步骤 2：编写代码

```vb
Option Explicit

' 声明数据库对象（使用 VBMAN.dll 中的类）
Private WithEvents m_DB As VBMAN.cDataBase

Private Sub Form_Load()
    ' 创建数据库对象
    Set m_DB = New VBMAN.cDataBase
    
    ' 连接 SQL Server 数据库
    If m_DB.Connect(VBMAN.enumDbType_MsSql, _
                    "127.0.0.1,1433", _
                    "sa", _
                    "Sa123456", _
                    "master") Then
        LogMessage "数据库连接成功"
    Else
        LogMessage "数据库连接失败: " & m_DB.LastErr
    End If
End Sub

Private Sub cmdQuery_Click()
    On Error GoTo EH
    
    ' 执行查询
    If m_DB.Sql("SELECT TOP 10 * FROM sys.tables").Query Then
        ' 获取结果集
        Dim i As Long
        For i = 1 To m_DB.Rows.Count
            LogMessage "表名: " & m_DB.Rows(i)("name")
        Next
    Else
        LogMessage "查询失败: " & m_DB.LastErr
    End If
    
    Exit Sub
    
EH:
    LogMessage "错误: " & Err.Description
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 断开连接
    If Not m_DB Is Nothing Then
        m_DB.Disconnect
    End If
End Sub
```

### 步骤 3：运行测试

1. 按 F5 运行程序
2. 点击"连接数据库"
3. 点击"执行查询"
4. 查看日志输出

---

## ? 服务端快速入门

### 步骤 1：创建窗体

创建一个新窗体，添加以下控件：

- 1 个 TextBox (`txtPort`) - 端口号
- 1 个 CommandButton (`cmdStart`) - 启动服务
- 1 个 ListBox (`lstResults`) - 显示结果
- 1 个 TextBox (`txtLog`) - 显示日志

### 步骤 2：编写代码

```vb
Option Explicit

Private WithEvents m_DB As VBMAN.cDataBase

Private Sub Form_Load()
    Set m_DB = New VBMAN.cDataBase
    txtPort.Text = "1433"
End Sub

Private Sub cmdStart_Click()
    On Error GoTo EH
    
    ' 连接数据库
    If m_DB.Connect(VBMAN.enumDbType_MsSql, _
                    "127.0.0.1," & txtPort.Text, _
                    "sa", _
                    "Sa123456", _
                    "master") Then
        LogMessage "数据库服务已启动"
        LoadData
    Else
        LogMessage "启动失败: " & m_DB.LastErr
    End If
    
    Exit Sub
    
EH:
    LogMessage "错误: " & Err.Description
End Sub

Private Sub LoadData()
    ' 查询数据
    If m_DB.Sql("SELECT name FROM sys.tables ORDER BY name").Fetch Then
        Dim i As Long
        lstResults.Clear
        For i = 1 To m_DB.Rows.Count
            lstResults.AddItem m_DB.Rows(i)("name")
        Next
        LogMessage "已加载 " & m_DB.Rows.Count & " 条记录"
    End If
End Sub

' ====== 辅助函数 ======

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_DB Is Nothing Then
        m_DB.Disconnect
    End If
End Sub
```

---

## ? 完整示例：用户管理

### 服务端代码

```vb
Option Explicit

Private WithEvents m_DB As VBMAN.cDataBase

Private Sub Form_Load()
    Set m_DB = New VBMAN.cDataBase
    
    ' 连接数据库
    m_DB.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
    
    ' 创建用户表（如果不存在）
    CreateUserTable
    
    ' 加载用户列表
    LoadUsers
End Sub

' 创建用户表
Private Sub CreateUserTable()
    Dim sSql As String
    sSql = "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'users') " & _
           "CREATE TABLE users (id INT IDENTITY(1,1) PRIMARY KEY, " & _
           "name NVARCHAR(50), age INT, email NVARCHAR(100))"
    
    m_DB.Sql(sSql).Exec
End Sub

' 加载用户列表
Private Sub LoadUsers()
    If m_DB.Sql("SELECT * FROM users ORDER BY id").Fetch Then
        Dim i As Long
        lstUsers.Clear
        For i = 1 To m_DB.Rows.Count
            lstUsers.AddItem m_DB.Rows(i)("name") & " - " & m_DB.Rows(i)("age")
        Next
    End If
End Sub

' 添加用户
Private Sub cmdAddUser_Click()
    ' 使用参数化查询防止 SQL 注入
    If m_DB.Sql("INSERT INTO users (name, age, email) VALUES (?, ?, ?)") _
        .Param("name", txtName.Text, VBMAN.adVarWChar) _
        .Param("age", CLng(txtAge.Text), VBMAN.adInteger) _
        .Param("email", txtEmail.Text, VBMAN.adVarWChar) _
        .ExecParam Then
        
        LogMessage "用户添加成功，ID: " & m_DB.LastInsertId
        LoadUsers
    Else
        LogMessage "添加失败: " & m_DB.LastErr
    End If
End Sub

' 删除用户
Private Sub cmdDeleteUser_Click()
    If m_DB.Sql("DELETE FROM users WHERE id = ?") _
        .Param("id", CLng(txtId.Text), VBMAN.adInteger) _
        .ExecParam Then
        
        LogMessage "用户删除成功"
        LoadUsers
    End If
End Sub

Private Sub LogMessage(sMessage As String)
    txtLog.Text = txtLog.Text & Format$(Now, "hh:mm:ss") & " - " & sMessage & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
```

### 客户端代码

```vb
Option Explicit

Private WithEvents m_DB As VBMAN.cDataBase

Private Sub Form_Load()
    Set m_DB = New VBMAN.cDataBase
    m_DB.Connect VBMAN.enumDbType_MsSql, "127.0.0.1,1433", "sa", "pwd", "mydb"
    
    LoadUsers
End Sub

Private Sub LoadUsers()
    ' 使用分页查询
    If m_DB.Sql("SELECT * FROM users").Page(1, 10).Fetch Then
        Dim i As Long
        lstUsers.Clear
        For i = 1 To m_DB.Rows.Count
            lstUsers.AddItem m_DB.Rows(i)("name") & " - " & m_DB.Rows(i)("age")
        Next
    End If
End Sub

Private Sub cmdSearch_Click()
    ' 使用参数化查询搜索
    If m_DB.Sql("SELECT * FROM users WHERE name LIKE ?") _
        .Param("name", "%" & txtSearch.Text & "%", VBMAN.adVarWChar) _
        .QueryParam Then
        
        Dim i As Long
        lstResults.Clear
        For i = 1 To m_DB.Rows.Count
            lstResults.AddItem m_DB.Rows(i)("name")
        Next
    End If
End Sub
```

---

## ? 常见问题

### Q1: 编译错误"用户定义类型未定义"

**原因**: 未引用 `VBMAN.dll` 或 `Microsoft ActiveX Data Objects 2.8 Library`

**解决**: 
1. 菜单：**项目** → **引用**
2. 勾选 **VBMAN** 和 **Microsoft ActiveX Data Objects 2.8 Library**

---

### Q2: 连接失败"无法连接到数据库"

**原因**: 连接字符串错误或数据库服务未启动

**解决**: 
- 检查数据库服务是否运行
- 验证连接参数（地址、端口、用户名、密码）
- 检查防火墙设置

---

### Q3: 查询返回空结果

**原因**: SQL 语句错误或表不存在

**解决**: 
- 使用 `m_DB.LastErr` 查看错误信息
- 检查 SQL 语句语法
- 验证表名和字段名

---

### Q4: 如何执行事务操作

```vb
' 开始事务
m_DB.TransBegin

' 执行多个操作
m_DB.Sql("INSERT INTO table1 ...").Exec
m_DB.Sql("INSERT INTO table2 ...").Exec

' 提交事务（失败自动回滚）
If m_DB.TransCommit Then
    Debug.Print "成功"
Else
    Debug.Print "失败: " & m_DB.LastErr
End If
```

---

### Q5: 如何防止 SQL 注入

```vb
' ? 错误：直接拼接 SQL（不安全）
m_DB.Sql("SELECT * FROM users WHERE name = '" & txtName.Text & "'").Query

' ? 正确：使用参数化查询（安全）
m_DB.Sql("SELECT * FROM users WHERE name = ?") _
    .Param("name", txtName.Text, VBMAN.adVarWChar) _
    .QueryParam
```

---

### Q6: 如何获取最后插入的 ID

```vb
' 插入数据
m_DB.Sql("INSERT INTO users (name) VALUES (?)") _
    .Param("name", "张三", VBMAN.adVarWChar) _
    .ExecParam

' 获取最后插入的 ID
Dim lId As Variant
lId = m_DB.LastInsertId
Debug.Print "新用户 ID: " & lId
```

---

## ? 下一步

- 查看 [连接管理](./connection.md) 了解数据库连接
- 查看 [查询操作](./query.md) 了解查询功能
- 查看 [参数化查询](./parameterized.md) 了解安全查询
- 查看 [高级功能](./advanced.md) 了解最佳实践

---

**最后更新**: 2026-01-21
