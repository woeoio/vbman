Attribute VB_Name = "Demo_Database"
'===============================================================
' cDatabase 类完整测试用例模块
' 作者：邓伟，215879458@qq.com
' 日期：2026-02-26
' 说明：测试 cDatabase 类的所有功能，包括连接、查询、事务、分页等
'===============================================================

Option Explicit

'数据库配置 - 请根据实际环境修改
Private Const TEST_DB_TYPE As Long = enumDbType.Mysql                           '1:Access, 2:Mysql, 3:MsSql, 4:Csv
Private Const TEST_DB_ADDRESS As String = "10.0.0.252,3306"
Private Const TEST_DB_USERNAME As String = "root"
Private Const TEST_DB_PASSWORD As String = "root"
Private Const TEST_DB_DATABASE As String = "testdb"

'测试数据库对象
Private TestDB As cDataBase
Private TestResult As String

'数据库配置 - 用于创建数据库的连接（不指定数据库）
Private Const ADMIN_DB_DATABASE As String = ""                                  ' 连接时不指定数据库，用于创建新数据库

'===============================================================
' 主测试入口
'===============================================================
Public Sub RunAllTests()
'    On Error GoTo ErrHandler
    
    Dim StartTime As Double
    StartTime = Timer
    
    Debug.Print "======================================================"
    Debug.Print "开始执行 cDatabase 类完整测试"
    Debug.Print "======================================================"
    Debug.Print ""
    
    '初始化测试数据库
    Set TestDB = New cDataBase
    
    '执行所有测试
    Call Test_CreateDatabase
    Call Test_ConnectDisconnect
    Call Test_CreateTestTable
    Call Test_InsertData
    Call Test_QueryData
    Call Test_UpdateData
    Call Test_DeleteData
    Call Test_Transaction
    Call Test_ParameterizedQuery
    Call Test_Pagination
    Call Test_LastInsertId
    Call Test_Count
    Call Test_TableExists
    Call Test_GetTableFields
    Call Test_BatchInsert
    Call Test_GetDatabases
    Call Test_GetTables
    Call Test_Escape
    Call Test_CheckConnection
    Call Test_GetVersion
    Call Test_ConnectionPool
    Call Test_AsyncExecute
    
    '清理
    Call Test_Cleanup
    Call Test_Disconnect
    
    Dim EndTime As Double
    EndTime = Timer
    
    Debug.Print ""
    Debug.Print "======================================================"
    Debug.Print "所有测试完成！耗时: " & Format(EndTime - StartTime, "0.00") & " 秒"
    Debug.Print "======================================================"
    
    Exit Sub
    
ErrHandler:
    Debug.Print "运行测试时发生错误: " & ERR.Description
    Debug.Print "错误号: " & ERR.Number
End Sub

'===============================================================
' 测试0: 创建测试数据库
'===============================================================
Private Sub Test_CreateDatabase()
    Debug.Print ">>> 测试0: 创建测试数据库"
    
    On Error GoTo ErrHandler
    
    '先连接到服务器（不指定数据库）
    Dim AdminDB As New cDataBase
    Dim bConnect As Boolean
    bConnect = AdminDB.Connect(TEST_DB_TYPE, TEST_DB_ADDRESS, TEST_DB_USERNAME, TEST_DB_PASSWORD, ADMIN_DB_DATABASE)
    
    If bConnect Then
        Debug.Print "  [OK] 已连接到数据库服务器"
    Else
        Debug.Print "  [FAIL] 连接数据库服务器失败"
        Debug.Print "       错误信息: " & AdminDB.LastErr
        Exit Sub
    End If
    
    '检查数据库是否已存在
    Dim bExists As Boolean
    bExists = DatabaseExists(AdminDB, TEST_DB_DATABASE)
    
    If bExists Then
        Debug.Print "       数据库 '" & TEST_DB_DATABASE & "' 已存在，跳过创建"
    Else
        '创建数据库
        Dim sqlCreate As String
        sqlCreate = "CREATE DATABASE `" & TEST_DB_DATABASE & "` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        
        Dim bCreate As Boolean
        bCreate = AdminDB.Sql(sqlCreate).Exec()
        
        If bCreate Then
            Debug.Print "  [OK] 数据库 '" & TEST_DB_DATABASE & "' 创建成功"
        Else
            Debug.Print "  [FAIL] 创建数据库失败"
            Debug.Print "       错误信息: " & AdminDB.LastErr
        End If
    End If
    
    '断开管理连接
    Call AdminDB.DisConnect
    Set AdminDB = Nothing
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 创建数据库时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'辅助函数：检查数据库是否存在
Private Function DatabaseExists(ByVal Db As cDataBase, ByVal DBName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sqlCheck As String
    sqlCheck = "SELECT COUNT(*) AS cnt FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & DBName & "'"
    
    Call Db.Sql(sqlCheck).Query
    If Db.Rs.RecordCount > 0 Then
        DatabaseExists = (Db.Rs("cnt") > 0)
    Else
        DatabaseExists = False
    End If
    
    Exit Function
    
ErrHandler:
    DatabaseExists = False
End Function

'===============================================================
' 测试1: 连接和断开数据库
'===============================================================
Private Sub Test_ConnectDisconnect()
    Debug.Print ">>> 测试1: 连接和断开数据库"
    
    On Error GoTo ErrHandler
    
    '测试连接
    Dim bConnect As Boolean
    bConnect = TestDB.Connect(TEST_DB_TYPE, TEST_DB_ADDRESS, TEST_DB_USERNAME, TEST_DB_PASSWORD, TEST_DB_DATABASE)
    
    If bConnect Then
        Debug.Print "  [OK] 数据库连接成功"
        Debug.Print "       连接状态: " & IIf(TestDB.IsConnect, "已连接", "未连接")
        Debug.Print "       连接版本: " & TestDB.GetVersion()
    Else
        Debug.Print "  [FAIL] 数据库连接失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '测试连接状态检查
    Dim bCheck As Boolean
    bCheck = TestDB.CheckConnection()
    Debug.Print "       连接状态检查: " & IIf(bCheck, "正常", "异常")
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 测试连接时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试2: 创建测试表
'===============================================================
Private Sub Test_CreateTestTable()
    Debug.Print ">>> 测试2: 创建测试表"
    
    On Error GoTo ErrHandler
    
    '先删除表（如果存在）
    Dim sqlDrop As String
    sqlDrop = "DROP TABLE IF EXISTS test_users"
    Call TestDB.Sql(sqlDrop).Exec
    
    '创建测试表
    Dim sqlCreate As String
    sqlCreate = "CREATE TABLE test_users (" & _
                "id INT PRIMARY KEY AUTO_INCREMENT," & _
                "username VARCHAR(50) NOT NULL," & _
                "email VARCHAR(100)," & _
                "age INT," & _
                "score DECIMAL(10,2)," & _
                "is_active BOOLEAN DEFAULT TRUE," & _
                "create_time DATETIME DEFAULT CURRENT_TIMESTAMP" & _
                ")"
    
    Dim bCreate As Boolean
    bCreate = TestDB.Sql(sqlCreate).Exec()
    
    If bCreate Then
        Debug.Print "  [OK] 测试表创建成功"
    Else
        Debug.Print "  [FAIL] 测试表创建失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '验证表是否存在
    Dim bExists As Boolean
    bExists = TestDB.TableExists("test_users")
    Debug.Print "       表存在检查: " & IIf(bExists, "存在", "不存在")
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 创建测试表时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试3: 插入数据
'===============================================================
Private Sub Test_InsertData()
    Debug.Print ">>> 测试3: 插入数据"
    
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim sqlInsert As String
    Dim bInsert As Boolean
    Dim lAffected As Long
    
    '插入单条记录
    sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('张三', 'zhangsan@test.com', 25, 95.5)"
    bInsert = TestDB.Sql(sqlInsert).Exec(lAffected)
    
    If bInsert Then
        Debug.Print "  [OK] 插入单条记录成功"
        Debug.Print "       影响行数: " & lAffected
        Debug.Print "       插入ID: " & TestDB.LastInsertId()
    Else
        Debug.Print "  [FAIL] 插入单条记录失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '批量插入记录
    Debug.Print ""
    Debug.Print "  开始批量插入记录..."
    For i = 1 To 10
        sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES " & _
                    "('用户" & i & "', 'user" & i & "@test.com', " & (20 + i) & ", " & (80 + i) & ".5)"
        bInsert = TestDB.Sql(sqlInsert).Exec(lAffected)
        If Not bInsert Then
            Debug.Print "  [FAIL] 批量插入第 " & i & " 条记录失败: " & TestDB.LastErr
            Exit For
        End If
    Next i
    
    Debug.Print "  [OK] 批量插入完成"
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 插入数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试4: 查询数据
'===============================================================
Private Sub Test_QueryData()
    Debug.Print ">>> 测试4: 查询数据"
    
    On Error GoTo ErrHandler
    
    '查询所有记录
    Dim sqlSelect As String
    sqlSelect = "SELECT * FROM test_users"
    
    Dim bQuery As Boolean
    bQuery = TestDB.Sql(sqlSelect).Query()
    
    If bQuery Then
        Debug.Print "  [OK] 查询成功"
        Debug.Print "       记录总数: " & TestDB.Rs.RecordCount
        
        '遍历记录
        Dim i As Long
        If TestDB.Rs.RecordCount > 0 Then
            TestDB.Rs.MoveFirst
            Debug.Print ""
            Debug.Print "  前3条记录:"
            For i = 1 To 3
                If TestDB.Rs.EOF Then Exit For
                Debug.Print "    ID: " & TestDB.Rs("id") & ", 用户名: " & TestDB.Rs("username") & ", 年龄: " & TestDB.Rs("age")
                TestDB.Rs.MoveNext
            Next i
        End If
        
        '关闭记录集
        If TestDB.Rs.State <> adStateClosed Then
            TestDB.Rs.Close
        End If
    Else
        Debug.Print "  [FAIL] 查询失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '使用 Fetch 方法查询
    Debug.Print ""
    bQuery = TestDB.Sql(sqlSelect).Fetch()
    
    If bQuery Then
        Debug.Print "  [OK] Fetch 查询成功"
        Debug.Print "       Rows集合数量: " & TestDB.Rows.Count
        
        If TestDB.Rows.Count > 0 Then
            Debug.Print "       第一行数据:"
            Dim Key As Variant
            Dim Row As Scripting.Dictionary
            Set Row = TestDB.Rows(1)
            For Each Key In Row.Keys
                Debug.Print "         " & Key & ": " & Row(Key)
            Next
        End If
    Else
        Debug.Print "  [FAIL] Fetch 查询失败"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 查询数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试5: 更新数据
'===============================================================
Private Sub Test_UpdateData()
    Debug.Print ">>> 测试5: 更新数据"
    
    On Error GoTo ErrHandler
    
    '更新单条记录
    Dim sqlUpdate As String
    sqlUpdate = "UPDATE test_users SET score = 99.9 WHERE username = '张三'"
    
    Dim bUpdate As Boolean
    Dim lAffected As Long
    bUpdate = TestDB.Sql(sqlUpdate).Exec(lAffected)
    
    If bUpdate Then
        Debug.Print "  [OK] 更新成功"
        Debug.Print "       影响行数: " & lAffected
        
        '验证更新
        Dim sqlCheck As String
        sqlCheck = "SELECT score FROM test_users WHERE username = '张三'"
        Call TestDB.Sql(sqlCheck).Fetch
        If TestDB.Rows.Count > 0 Then
            Debug.Print "       更新后的分数: " & TestDB.Rows(1)("score")
        End If
    Else
        Debug.Print "  [FAIL] 更新失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 更新数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试6: 删除数据
'===============================================================
Private Sub Test_DeleteData()
    Debug.Print ">>> 测试6: 删除数据"
    
    On Error GoTo ErrHandler
    
    '先查询总记录数
    Dim lCountBefore As Long
    lCountBefore = TestDB.Count("test_users")
    Debug.Print "  删除前记录数: " & lCountBefore
    
    '删除指定记录
    Dim sqlDelete As String
    sqlDelete = "DELETE FROM test_users WHERE username = '张三'"
    
    Dim bDelete As Boolean
    Dim lAffected As Long
    bDelete = TestDB.Sql(sqlDelete).Exec(lAffected)
    
    If bDelete Then
        Debug.Print "  [OK] 删除成功"
        Debug.Print "       影响行数: " & lAffected
        
        '验证删除
        Dim lCountAfter As Long
        lCountAfter = TestDB.Count("test_users")
        Debug.Print "  删除后记录数: " & lCountAfter
    Else
        Debug.Print "  [FAIL] 删除失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 删除数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试7: 事务处理
'===============================================================
Private Sub Test_Transaction()
    Debug.Print ">>> 测试7: 事务处理"
    
    On Error GoTo ErrHandler
    
    '测试事务提交
    Debug.Print "  测试事务提交:"
    
    Dim bBegin As Boolean
    bBegin = TestDB.TransBegin()
    Debug.Print "       开始事务: " & IIf(bBegin, "成功", "失败")
    
    If bBegin Then
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('事务用户1', 'trans1@test.com', 30, 88)"
        Call TestDB.Sql(sqlInsert).Exec
        
        sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('事务用户2', 'trans2@test.com', 31, 89)"
        Call TestDB.Sql(sqlInsert).Exec
        
        Dim bCommit As Boolean
        bCommit = TestDB.TransCommit()
        Debug.Print "       提交事务: " & IIf(bCommit, "成功", "失败")
        
        If bCommit Then
            Dim lCount As Long
            lCount = TestDB.Count("test_users")
            Debug.Print "       当前记录数: " & lCount
        End If
    End If
    
    '测试事务回滚
    Debug.Print ""
    Debug.Print "  测试事务回滚:"
    
    bBegin = TestDB.TransBegin()
    Debug.Print "       开始事务: " & IIf(bBegin, "成功", "失败")
    
    If bBegin Then
        lCount = TestDB.Count("test_users")
        
        sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('回滚用户', 'rollback@test.com', 32, 90)"
        Call TestDB.Sql(sqlInsert).Exec
        
        Dim bRollback As Boolean
        bRollback = TestDB.TransRollback()
        Debug.Print "       回滚事务: " & IIf(bRollback, "成功", "失败")
        
        Dim lCount2 As Long
        lCount2 = TestDB.Count("test_users")
        Debug.Print "       记录数变化: " & lCount & " -> " & lCount2 & " (应保持不变)"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 事务处理时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试8: 参数化查询
'===============================================================
Private Sub Test_ParameterizedQuery()
    Debug.Print ">>> 测试8: 参数化查询"
    
    On Error GoTo ErrHandler
    
    '测试参数化插入
    Dim sqlInsert As String
    sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES (?, ?, ?, ?)"
    
    Dim bExec As Boolean
    bExec = TestDB.Sql(sqlInsert) _
        .Param("username", "参数用户1", adVarChar) _
        .Param("email", "param1@test.com", adVarChar) _
        .Param("age", 35, adInteger) _
        .Param("score", 92.5, adDouble) _
        .ExecParam()
    
    If bExec Then
        Debug.Print "  [OK] 参数化插入成功"
        Debug.Print "       插入ID: " & TestDB.LastInsertId()
    Else
        Debug.Print "  [FAIL] 参数化插入失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '测试参数化查询
    Dim sqlSelect As String
    sqlSelect = "SELECT * FROM test_users WHERE username = ? AND age > ?"
    
    Dim bQuery As Boolean
    bQuery = TestDB.Sql(sqlSelect) _
        .Param("username", "参数用户1", adVarChar) _
        .Param("age", 30, adInteger) _
        .QueryParam()
    
    If bQuery Then
        Debug.Print "  [OK] 参数化查询成功"
        Debug.Print "       查询结果数量: " & TestDB.Rs.RecordCount
    Else
        Debug.Print "  [FAIL] 参数化查询失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 参数化查询时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试9: 分页查询
'===============================================================
Private Sub Test_Pagination()
    Debug.Print ">>> 测试9: 分页查询"
    
    '    On Error GoTo ErrHandler
    
    '获取总记录数
    Dim lTotalCount As Long
    lTotalCount = TestDB.Count("test_users")
    Debug.Print "  总记录数: " & lTotalCount
    
    '查询第1页，每页3条
    Dim lPageSize As Long
    lPageSize = 3
    
    Dim sqlSelect As String
    sqlSelect = "SELECT * FROM test_users ORDER BY id"
    
    Dim bQuery As Boolean
    bQuery = TestDB.Sql(sqlSelect).Page(1, lPageSize).Query()
    
    If bQuery Then
        Debug.Print "  [OK] 第1页查询成功"
        Debug.Print "       记录数: " & TestDB.Rs.RecordCount
        
        '显示第1页数据
        If TestDB.Rs.RecordCount > 0 Then
            TestDB.Rs.MoveFirst
            Debug.Print "  第1页数据:"
            Do While Not TestDB.Rs.EOF
                Debug.Print "    ID: " & TestDB.Rs("id") & ", 用户名: " & TestDB.Rs("username")
                TestDB.Rs.MoveNext
            Loop
        End If
        
        '关闭记录集以释放资源
        If TestDB.Rs.State <> adStateClosed Then
            TestDB.Rs.Close
        End If
    Else
        Debug.Print "  [FAIL] 第1页查询失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    '查询第2页
    bQuery = TestDB.Sql(sqlSelect).Page(2, lPageSize).Fetch()
    
    If bQuery Then
        Debug.Print ""
        Debug.Print "  [OK] 第2页查询成功"
        Debug.Print "       记录数: " & TestDB.Rows.Count
        
        '显示第2页数据
        If TestDB.Rows.Count > 0 Then
            Debug.Print "  第2页数据:"
            Dim i As Long
            For i = 1 To TestDB.Rows.Count
                Dim Row As Scripting.Dictionary
                Set Row = TestDB.Rows(i)
                Debug.Print "    ID: " & Row("id") & ", 用户名: " & Row("username")
            Next
        End If
    Else
        Debug.Print ""
        Debug.Print "  [FAIL] 第2页查询失败"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 分页查询时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试10: 获取最后插入ID
'===============================================================
Private Sub Test_LastInsertId()
    Debug.Print ">>> 测试10: 获取最后插入ID"
    
    On Error GoTo ErrHandler
    
    Dim sqlInsert As String
    sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('测试ID', 'testid@test.com', 40, 100)"
    
    Dim bInsert As Boolean
    bInsert = TestDB.Sql(sqlInsert).Exec()
    
    If bInsert Then
        Dim vLastId As Variant
        vLastId = TestDB.LastInsertId()
        
        If Not IsEmpty(vLastId) Then
            Debug.Print "  [OK] 获取最后插入ID成功"
            Debug.Print "       最后插入ID: " & vLastId
        Else
            Debug.Print "  [FAIL] 获取最后插入ID失败（返回空）"
        End If
    Else
        Debug.Print "  [FAIL] 插入失败，无法测试获取ID"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 测试获取最后插入ID时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试11: 统计记录数
'===============================================================
Private Sub Test_Count()
    Debug.Print ">>> 测试11: 统计记录数"
    
    On Error GoTo ErrHandler
    
    '使用表名统计
    Dim lCount1 As Long
    lCount1 = TestDB.Count("test_users")
    Debug.Print "  [OK] 表名统计: " & lCount1 & " 条记录"
    
    '使用SQL统计
    Dim sqlCount As String
    sqlCount = "SELECT * FROM test_users WHERE age > 25"
    
    Dim lCount2 As Long
    lCount2 = TestDB.Sql(sqlCount).Count()
    Debug.Print "  [OK] SQL统计（年龄>25）: " & lCount2 & " 条记录"
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 统计记录数时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试12: 检查表是否存在
'===============================================================
Private Sub Test_TableExists()
    Debug.Print ">>> 测试12: 检查表是否存在"
    
    On Error GoTo ErrHandler
    
    '检查存在的表
    Dim bExists1 As Boolean
    bExists1 = TestDB.TableExists("test_users")
    Debug.Print "  test_users 表: " & IIf(bExists1, "存在", "不存在")
    
    '检查不存在的表
    Dim bExists2 As Boolean
    bExists2 = TestDB.TableExists("not_exists_table")
    Debug.Print "  not_exists_table 表: " & IIf(bExists2, "存在", "不存在")
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 检查表是否存在时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试13: 获取表的字段列表
'===============================================================
Private Sub Test_GetTableFields()
    Debug.Print ">>> 测试13: 获取表的字段列表"
    
    On Error GoTo ErrHandler
    
    Dim fields As Collection
    Set fields = TestDB.GetTableFields("test_users")
    
    '检查是否返回了有效的Collection
    If fields Is Nothing Then
        Debug.Print "  [FAIL] GetTableFields返回Nothing"
        Debug.Print "       错误信息: " & TestDB.LastErr
        Debug.Print ""
        Exit Sub
    End If
    
    If fields.Count > 0 Then
        Debug.Print "  [OK] 获取字段列表成功"
        Debug.Print "       字段数量: " & fields.Count
        Debug.Print "       字段列表:"
        
        Dim i As Long
        Dim lMaxFields As Long
        lMaxFields = fields.Count
        '限制显示的字段数量
        If lMaxFields > 50 Then lMaxFields = 50
        
        For i = 1 To lMaxFields
            Debug.Print "         " & i & ". " & fields(i)
        Next
        
        If fields.Count > 50 Then
            Debug.Print "         ... 还有 " & (fields.Count - 50) & " 个字段未显示"
        End If
    Else
        Debug.Print "  [FAIL] 获取字段列表失败或表没有字段"
        Debug.Print "       Collection.Count: " & fields.Count
        If TestDB.LastErr <> "" Then
            Debug.Print "       错误信息: " & TestDB.LastErr
        End If
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 获取表的字段列表时发生错误: " & ERR.Description
    Debug.Print "       错误号: " & ERR.Number
    Debug.Print ""
End Sub

'===============================================================
' 测试14: 批量插入数据
'===============================================================
Private Sub Test_BatchInsert()
    Debug.Print ">>> 测试14: 批量插入数据"
    
    On Error GoTo ErrHandler
    
    '准备批量数据
    Dim Data As New Collection
    Dim Row As Scripting.Dictionary
    Dim i As Long
    
    For i = 1 To 5
        Set Row = New Scripting.Dictionary
        Row.Add "username", "批量用户" & i
        Row.Add "email", "batch" & i & "@test.com"
        Row.Add "age", 50 + i
        Row.Add "score", 90 + i
        Data.Add Row
    Next
    
    '获取插入前记录数
    Dim lCountBefore As Long
    lCountBefore = TestDB.Count("test_users")
    
    '执行批量插入
    Dim bBatch As Boolean
    bBatch = TestDB.BatchInsert("test_users", Data)
    
    If bBatch Then
        Debug.Print "  [OK] 批量插入成功"
        Debug.Print "       插入数量: " & Data.Count
        
        '验证插入结果
        Dim lCountAfter As Long
        lCountAfter = TestDB.Count("test_users")
        Debug.Print "       记录数变化: " & lCountBefore & " -> " & lCountAfter
    Else
        Debug.Print "  [FAIL] 批量插入失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 批量插入数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试15: 获取数据库列表
'===============================================================
Private Sub Test_GetDatabases()
    Debug.Print ">>> 测试15: 获取数据库列表"
    
    On Error GoTo ErrHandler
    
    Dim Databases As Collection
    Set Databases = TestDB.GetDatabases()
    
    '检查是否返回了有效的Collection
    If Databases Is Nothing Then
        Debug.Print "  [FAIL] GetDatabases返回Nothing"
        Debug.Print "       错误信息: " & TestDB.LastErr
        Debug.Print ""
        Exit Sub
    End If
    
    If Databases.Count > 0 Then
        Debug.Print "  [OK] 获取数据库列表成功"
        Debug.Print "       数据库数量: " & Databases.Count
        Debug.Print "       数据库列表（前10个）:"
        
        Dim i As Long
        Dim lMaxCount As Long
        lMaxCount = Databases.Count
        If lMaxCount > 10 Then lMaxCount = 10
        
        For i = 1 To lMaxCount
            Debug.Print "         " & i & ". " & Databases(i)
        Next
        
        If Databases.Count > 10 Then
            Debug.Print "         ... 还有 " & (Databases.Count - 10) & " 个数据库未显示"
        End If
    Else
        Debug.Print "  [FAIL] 获取数据库列表失败或没有数据库"
        Debug.Print "       Collection.Count: " & Databases.Count
        If TestDB.LastErr <> "" Then
            Debug.Print "       错误信息: " & TestDB.LastErr
        End If
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 获取数据库列表时发生错误: " & ERR.Description
    Debug.Print "       错误号: " & ERR.Number
    Debug.Print ""
End Sub

'===============================================================
' 测试16: 获取表列表
'===============================================================
Private Sub Test_GetTables()
    Debug.Print ">>> 测试16: 获取表列表"
    
    On Error GoTo ErrHandler
    
    Dim Tables As Collection
    Set Tables = TestDB.GetTables()
    
    '检查是否返回了有效的Collection
    If Tables Is Nothing Then
        Debug.Print "  [FAIL] GetTables返回Nothing"
        Debug.Print "       错误信息: " & TestDB.LastErr
        Debug.Print ""
        Exit Sub
    End If
    
    If Tables.Count > 0 Then
        Debug.Print "  [OK] 获取表列表成功"
        Debug.Print "       表数量: " & Tables.Count
        Debug.Print "       表列表（前10个）:"
        
        Dim i As Long
        Dim lMaxCount As Long
        lMaxCount = Tables.Count
        If lMaxCount > 10 Then lMaxCount = 10
        
        For i = 1 To lMaxCount
            Debug.Print "         " & i & ". " & Tables(i)
        Next
        
        If Tables.Count > 10 Then
            Debug.Print "         ... 还有 " & (Tables.Count - 10) & " 个表未显示"
        End If
    Else
        Debug.Print "  [FAIL] 获取表列表失败或没有表"
        Debug.Print "       Collection.Count: " & Tables.Count
        If TestDB.LastErr <> "" Then
            Debug.Print "       错误信息: " & TestDB.LastErr
        End If
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 获取表列表时发生错误: " & ERR.Description
    Debug.Print "       错误号: " & ERR.Number
    Debug.Print ""
End Sub

'===============================================================
' 测试17: SQL转义
'===============================================================
Private Sub Test_Escape()
    Debug.Print ">>> 测试17: SQL转义"
    
    On Error GoTo ErrHandler
    
    Dim inputStr As String
    inputStr = "O'Reilly's Book"
    
    Dim EscapedStr As String
    EscapedStr = TestDB.Escape(inputStr)
    
    Debug.Print "  [OK] SQL转义测试"
    Debug.Print "       原字符串: " & inputStr
    Debug.Print "       转义后: " & EscapedStr
    
    '测试带单引号的插入
    Dim sqlInsert As String
    sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('" & EscapedStr & "', 'escape@test.com', 45, 95)"
    
    Dim bInsert As Boolean
    bInsert = TestDB.Sql(sqlInsert).Exec()
    
    If bInsert Then
        Debug.Print "       插入成功"
    Else
        Debug.Print "       插入失败: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] SQL转义测试时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试18: 检查连接状态
'===============================================================
Private Sub Test_CheckConnection()
    Debug.Print ">>> 测试18: 检查连接状态"
    
    On Error GoTo ErrHandler
    
    '检查当前连接状态
    Dim bConnected As Boolean
    bConnected = TestDB.CheckConnection()
    Debug.Print "  [OK] 连接状态检查: " & IIf(bConnected, "已连接", "未连接")
    
    '执行简单查询验证连接
    Dim sqlTest As String
    sqlTest = "SELECT 1"
    Dim bTest As Boolean
    bTest = TestDB.Sql(sqlTest).Query()
    
    If bTest Then
        Debug.Print "       查询测试: 成功"
    Else
        Debug.Print "       查询测试: 失败"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 检查连接状态时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试19: 获取版本信息
'===============================================================
Private Sub Test_GetVersion()
    Debug.Print ">>> 测试19: 获取版本信息"
    
    On Error GoTo ErrHandler
    
    Dim sVersion As String
    sVersion = TestDB.GetVersion()
    
    If sVersion <> "" Then
        Debug.Print "  [OK] 获取版本信息成功"
        Debug.Print "       ADO版本: " & sVersion
    Else
        Debug.Print "  [FAIL] 获取版本信息失败"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 获取版本信息时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试20: 连接池功能
'===============================================================
Private Sub Test_ConnectionPool()
    Debug.Print ">>> 测试20: 连接池功能"
    
    On Error GoTo ErrHandler
    
    '创建连接池实例
    Dim dbPool1 As cDataBase
    Dim dbPool2 As cDataBase
    
    Set dbPool1 = TestDB.ConnInst("Pool1")
    Set dbPool2 = TestDB.ConnInst("Pool2")
    
    Debug.Print "  [OK] 创建连接池实例"
    Debug.Print "       Pool1 连接状态: " & IIf(dbPool1.IsConnect, "已连接", "未连接")
    Debug.Print "       Pool2 连接状态: " & IIf(dbPool2.IsConnect, "已连接", "未连接")
    
    '使用连接池实例查询
    Dim sqlSelect As String
    sqlSelect = "SELECT COUNT(*) AS cnt FROM test_users"
    
    Dim bQuery As Boolean
    bQuery = dbPool1.Sql(sqlSelect).Query()
    
    If bQuery Then
        Debug.Print "       Pool1 查询成功，记录数: " & dbPool1.Rs("cnt")
    Else
        Debug.Print "       Pool1 查询失败"
    End If
    
    '移除连接池实例
    TestDB.ConnInstRemove "Pool1"
    TestDB.ConnInstRemove "Pool2"
    
    Debug.Print "       连接池实例已移除"
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 测试连接池功能时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 测试21: 异步执行
'===============================================================
Private Sub Test_AsyncExecute()
    Debug.Print ">>> 测试21: 异步执行"
    
    On Error GoTo ErrHandler
    
    Dim sqlInsert As String
    sqlInsert = "INSERT INTO test_users (username, email, age, score) VALUES ('异步用户', 'async@test.com', 55, 98)"
    
    '异步执行SQL
    Debug.Print "  开始异步执行SQL..."
    Dim bExec As Boolean
    bExec = TestDB.Sql(sqlInsert).Async().Exec()
    
    If bExec Then
        Debug.Print "  [OK] 异步执行已启动"
        Debug.Print "       注意：异步执行结果通过事件回调处理"
    Else
        Debug.Print "  [FAIL] 异步执行失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 测试异步执行时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 清理测试数据
'===============================================================
Private Sub Test_Cleanup()
    Debug.Print ">>> 清理测试数据"
    
    On Error GoTo ErrHandler
    
    Dim sqlDrop As String
    sqlDrop = "DROP TABLE IF EXISTS test_users"
    
    Dim bDrop As Boolean
    bDrop = TestDB.Sql(sqlDrop).Exec()
    
    If bDrop Then
        Debug.Print "  [OK] 测试表已删除"
    Else
        Debug.Print "  [WARN] 删除测试表失败（可能表不存在）"
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [WARN] 清理测试数据时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

'===============================================================
' 断开数据库连接
'===============================================================
Private Sub Test_Disconnect()
    Debug.Print ">>> 断开数据库连接"
    
    On Error GoTo ErrHandler
    
    Dim bDisConnect As Boolean
    bDisConnect = TestDB.DisConnect()
    
    If bDisConnect Then
        Debug.Print "  [OK] 数据库连接已断开"
        Debug.Print "       连接状态: " & IIf(TestDB.IsConnect, "已连接", "未连接")
    Else
        Debug.Print "  [FAIL] 断开数据库连接失败"
        Debug.Print "       错误信息: " & TestDB.LastErr
    End If
    
    Debug.Print ""
    
    Exit Sub
    
ErrHandler:
    Debug.Print "  [FAIL] 断开数据库连接时发生错误: " & ERR.Description
    Debug.Print ""
End Sub

''===============================================================
'' 异步执行完成事件处理
''===============================================================
'Private Sub TestDB_AsyncExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
'    Debug.Print "  [AsyncEvent] 异步执行完成"
'
'    If adStatus = adErrorsOccurred Then
'        Debug.Print "       发生错误: " & pError.Description
'    Else
'        Debug.Print "       影响行数: " & RecordsAffected
'    End If
'End Sub

'===============================================================
' 单独测试函数
'===============================================================
Public Sub TestSpecificCase(ByVal TestCase As String)
    Select Case LCase(TestCase)
    Case "database"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_Disconnect
        
    Case "connect"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_Disconnect
        
    Case "insert"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_CreateTestTable
        Call Test_InsertData
        Call Test_Cleanup
        Call Test_Disconnect
        
    Case "query"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_CreateTestTable
        Call Test_InsertData
        Call Test_QueryData
        Call Test_Cleanup
        Call Test_Disconnect
        
    Case "transaction"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_CreateTestTable
        Call Test_Transaction
        Call Test_Cleanup
        Call Test_Disconnect
        
    Case "pagination"
        Set TestDB = New cDataBase
        Call Test_CreateDatabase
        Call Test_ConnectDisconnect
        Call Test_CreateTestTable
        Call Test_InsertData
        Call Test_Pagination
        Call Test_Cleanup
        Call Test_Disconnect
        
    Case Else
        Debug.Print "未知测试用例: " & TestCase
        Debug.Print "可用测试用例: database, connect, insert, query, transaction, pagination"
    End Select
End Sub

'===============================================================
' 测试说明文档
'===============================================================
Public Sub ShowTestHelp()
    Debug.Print "======================================================"
    Debug.Print "cDatabase 类测试用例说明"
    Debug.Print "======================================================"
    Debug.Print ""
    Debug.Print "使用方法:"
    Debug.Print "1. 运行所有测试: RunAllTests"
    Debug.Print "2. 运行单个测试: TestSpecificCase ""测试名称"""
    Debug.Print ""
    Debug.Print "可用测试名称:"
    Debug.Print "  - database   : 测试创建数据库"
    Debug.Print "  - connect    : 测试连接"
    Debug.Print "  - insert     : 测试插入"
    Debug.Print "  - query      : 测试查询"
    Debug.Print "  - transaction: 测试事务"
    Debug.Print "  - pagination : 测试分页"
    Debug.Print ""
    Debug.Print "注意事项:"
    Debug.Print "1. 运行前请修改模块顶部的数据库配置参数"
    Debug.Print "2. 测试会自动创建 '" & TEST_DB_DATABASE & "' 数据库（如不存在）"
    Debug.Print "3. 测试会创建和删除 test_users 表"
    Debug.Print ""
    Debug.Print "======================================================"
End Sub
