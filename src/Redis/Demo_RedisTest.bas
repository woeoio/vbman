Attribute VB_Name = "Module1"
'=========================================================================
' cRedisClient 全量单元测试
'=========================================================================
Public Sub TestRedisClient()
    Dim Redis As New cRedisClient
    Dim bResult As Boolean
    Dim vResult As Variant
    Dim sResult As String
    Dim lResult As Long
    Dim i As Long
    Dim Dict As Scripting.Dictionary
    Dim vArray() As Variant
    
    Debug.Print "=========================================="
    Debug.Print "cRedisClient 全量单元测试"
    Debug.Print "=========================================="
    
    On Error GoTo TestError
    
    '=====================================================================
    ' 1. 连接测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[1] 连接测试"
    Debug.Print "----------------------------------------"
    
    bResult = Redis.Connect("dbserver.cc", 9736)
    Debug.Print "连接: " & IIf(bResult, "成功", "失败")
    Debug.Print "Connected: " & Redis.Connected
    Debug.Print "Host: " & Redis.Host
    Debug.Print "Port: " & Redis.Port
    
    If Redis.Connected = False Then
        Debug.Print "连接失败，跳过后续测试"
        Exit Sub
    End If
    
    ' PING 测试
    sResult = Redis.Ping()
    Debug.Print "PING: " & sResult
    
    '=====================================================================
    ' 2. 基础命令测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[2] 基础命令测试"
    Debug.Print "----------------------------------------"
    
    ' SET/GET
    bResult = Redis.Set_("test_key", "test_value")
    Debug.Print "SET test_key: " & IIf(bResult, "成功", "失败")
    
    sResult = Redis.Get_("test_key")
    Debug.Print "GET test_key: " & sResult
    
    ' SET with EX (过期时间)
    bResult = Redis.Set_("expire_key", "will_expire", 10, 0)
    Debug.Print "SET with EX(10s): " & IIf(bResult, "成功", "失败")
    
    lResult = Redis.TTL("expire_key")
    Debug.Print "TTL expire_key: " & lResult
    
    ' EXISTS - 单个键测试
    lResult = Redis.Exists("test_key")
    Debug.Print "EXISTS test_key: " & lResult
    
    lResult = Redis.Exists("expire_key")
    Debug.Print "EXISTS expire_key: " & lResult
    
    lResult = Redis.Exists("nonexistent")
    Debug.Print "EXISTS nonexistent: " & lResult
    
    ' DEL - 单个键
    lResult = Redis.Del("test_key")
    Debug.Print "DEL test_key: " & lResult & " 个键被删除"
    
    lResult = Redis.Exists("test_key")
    Debug.Print "EXISTS test_key after DEL: " & lResult
    
    ' KEYS
    bResult = Redis.Set_("pattern_test1", "value1")
    bResult = Redis.Set_("pattern_test2", "value2")
    vArray = Redis.Keys("pattern_test*")
    Debug.Print "KEYS pattern_test*: 找到 " & UBound(vArray) + 1 & " 个键"
    For i = 0 To UBound(vArray)
        Debug.Print "  - " & vArray(i)
    Next
    Redis.Del "pattern_test1"
    Redis.Del "pattern_test2"
    
    '=====================================================================
    ' 3. String 操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[3] String 操作测试"
    Debug.Print "----------------------------------------"
    
    ' INCR/DECR
    Call Redis.Set_("counter", "10")
    lResult = Redis.Incr("counter")
    Debug.Print "INCR counter: " & lResult
    
    lResult = Redis.Incr("counter")
    Debug.Print "INCR counter again: " & lResult
    
    lResult = Redis.Decr("counter")
    Debug.Print "DECR counter: " & lResult
    
    ' MGET/MSET - 先用单个SET设置
    Call Redis.Set_("mkey1", "value1")
    Call Redis.Set_("mkey2", "value2")
    Call Redis.Set_("mkey3", "value3")
    
    vArray = Redis.MGet("mkey1", "mkey2", "mkey3", "nonexistent")
    Debug.Print "MGET mkey1, mkey2, mkey3, nonexistent:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  [" & i & "]: " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    Redis.Del "mkey1"
    Redis.Del "mkey2"
    Redis.Del "mkey3"
    
    '=====================================================================
    ' 4. Hash 操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[4] Hash 操作测试"
    Debug.Print "----------------------------------------"
    
    ' HSET/HGET
    bResult = Redis.HSet("user:1", "name", "张三")
    Debug.Print "HSET user:1 name: " & IIf(bResult, "成功", "失败")
    
    bResult = Redis.HSet("user:1", "age", "25")
    bResult = Redis.HSet("user:1", "city", "北京")
    
    sResult = Redis.HGet("user:1", "name")
    Debug.Print "HGET user:1 name: " & sResult
    
    sResult = Redis.HGet("user:1", "age")
    Debug.Print "HGET user:1 age: " & sResult
    
    sResult = Redis.HGet("user:1", "city")
    Debug.Print "HGET user:1 city: " & sResult
    
    ' HEXISTS
    bResult = Redis.HExists("user:1", "name")
    Debug.Print "HEXISTS user:1 name: " & IIf(bResult, "True", "False")
    
    bResult = Redis.HExists("user:1", "email")
    Debug.Print "HEXISTS user:1 email: " & IIf(bResult, "True", "False")
    
    ' HMGET
    vArray = Redis.HMGet("user:1", "name", "age", "city")
    Debug.Print "HMGET user:1 name, age, city:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  [" & i & "]: " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' HGETALL
    Set Dict = Redis.HGetAll("user:1")
    Debug.Print "HGETALL user:1:"
    Dim Key As Variant
    For Each Key In Dict.Keys
        Debug.Print "  " & Key & ": " & Dict(Key)
    Next
    
    ' HDEL
    lResult = Redis.HDel("user:1", "age")
    Debug.Print "HDEL user:1 age: 删除 " & lResult & " 个字段"
    
    Redis.Del "user:1"
    
    '=====================================================================
    ' 5. List 操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[5] List 操作测试"
    Debug.Print "----------------------------------------"
    
    ' LPUSH - 逐个添加
    lResult = Redis.LPush("mylist", "item1")
    lResult = Redis.LPush("mylist", "item2")
    lResult = Redis.LPush("mylist", "item3")
    Debug.Print "LPUSH mylist 3次: 长度=" & lResult
    
    ' RPUSH - 逐个添加
    lResult = Redis.RPush("mylist", "item4")
    lResult = Redis.RPush("mylist", "item5")
    Debug.Print "RPUSH mylist 2次: 长度=" & lResult
    
    ' LLEN
    lResult = Redis.lLen("mylist")
    Debug.Print "LLEN mylist: " & lResult
    
    ' LRANGE
    vArray = Redis.LRange("mylist", 0, -1)
    Debug.Print "LRANGE mylist 0 -1:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  [" & i & "]: " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' LPOP/RPOP
    sResult = Redis.LPop("mylist")
    Debug.Print "LPOP mylist: " & sResult
    
    sResult = Redis.RPop("mylist")
    Debug.Print "RPOP mylist: " & sResult
    
    Redis.Del "mylist"
    
    '=====================================================================
    ' 6. Set 操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[6] Set 操作测试"
    Debug.Print "----------------------------------------"
    
    ' SADD - 逐个添加
    lResult = Redis.SAdd("myset", "apple")
    lResult = lResult + Redis.SAdd("myset", "banana")
    lResult = lResult + Redis.SAdd("myset", "orange")
    Debug.Print "SADD myset apple, banana, orange: 新增 " & lResult & " 个成员"
    
    ' SMEMBERS
    vArray = Redis.SMembers("myset")
    Debug.Print "SMEMBERS myset:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  - " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' SCARD
    lResult = Redis.SCard("myset")
    Debug.Print "SCARD myset: " & lResult
    
    ' SISMEMBER
    bResult = Redis.SIsMember("myset", "apple")
    Debug.Print "SISMEMBER myset apple: " & IIf(bResult, "True", "False")
    
    bResult = Redis.SIsMember("myset", "grape")
    Debug.Print "SISMEMBER myset grape: " & IIf(bResult, "True", "False")
    
    ' SREM - 逐个删除
    lResult = Redis.SRem("myset", "banana")
    Debug.Print "SREM myset banana: 删除 " & lResult & " 个成员"
    
    Redis.Del "myset"
    
    '=====================================================================
    ' 7. Sorted Set 操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[7] Sorted Set 操作测试"
    Debug.Print "----------------------------------------"
    
    ' ZADD
    lResult = Redis.ZAdd("leaderboard", 100, "player1")
    lResult = Redis.ZAdd("leaderboard", 200, "player2")
    lResult = Redis.ZAdd("leaderboard", 150, "player3")
    Debug.Print "ZADD leaderboard 3 members: 最新新增 " & lResult & " 个成员"
    
    ' ZRANGE
    vArray = Redis.ZRange("leaderboard", 0, -1)
    Debug.Print "ZRANGE leaderboard 0 -1:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  - " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' ZRANGE with scores
    vArray = Redis.ZRange("leaderboard", 0, -1, True)
    Debug.Print "ZRANGE leaderboard 0 -1 WITHSCORES:"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray) Step 2
            Debug.Print "  - " & vArray(i) & ": " & vArray(i + 1)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' ZCARD
    lResult = Redis.ZCard("leaderboard")
    Debug.Print "ZCARD leaderboard: " & lResult
    
    ' ZREM
    lResult = Redis.ZRem("leaderboard", "player3")
    Debug.Print "ZREM leaderboard player3: 删除 " & lResult & " 个成员"
    
    Redis.Del "leaderboard"
    
    '=====================================================================
    ' 8. 事务操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[8] 事务操作测试"
    Debug.Print "----------------------------------------"
    
    Debug.Print "InTransaction: " & Redis.InTransaction
    
    ' MULTI
    bResult = Redis.Multi()
    Debug.Print "MULTI: " & IIf(bResult, "成功", "失败")
    Debug.Print "InTransaction: " & Redis.InTransaction
    
    ' 在事务中执行命令
    Call Redis.Set_("trans_key1", "value1")
    Call Redis.Set_("trans_key2", "value2")
    Call Redis.Set_("trans_key3", "value3")
    Debug.Print "在事务中执行了3个 SET 命令"
    
    ' EXEC
    vArray = Redis.Exec()
    Debug.Print "EXEC: 返回 " & UBound(vArray) + 1 & " 个结果"
    If IsArray(vArray) Then
        For i = 0 To UBound(vArray)
            Debug.Print "  [" & i & "]: " & vArray(i)
        Next
    Else
        Debug.Print "  结果: " & Join(vArray, " | ")
    End If
    
    ' 验证结果
    sResult = Redis.Get_("trans_key1")
    Debug.Print "GET trans_key1: " & sResult
    Debug.Print "InTransaction: " & Redis.InTransaction
    
    ' 清理
    Redis.Del "trans_key1"
    Redis.Del "trans_key2"
    Redis.Del "trans_key3"
    
    ' 测试 DISCARD
    Call Redis.Multi
    Call Redis.Set_("discard_key", "will_be_discarded")
    bResult = Redis.Discard()
    Debug.Print "DISCARD: " & IIf(bResult, "成功", "失败")
    
    lResult = Redis.Exists("discard_key")
    Debug.Print "EXISTS discard_key after DISCARD: " & lResult & " (应为0)"
    
    '=====================================================================
    ' 9. 服务器操作测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[9] 服务器操作测试"
    Debug.Print "----------------------------------------"
    
    ' PING
    sResult = Redis.Ping()
    Debug.Print "PING: " & sResult
    
    ' INFO
    sResult = Redis.Info("server")
    Debug.Print "INFO server: (前100字符) " & Left(sResult, 100) & "..."
    
    ' INFO (全部)
    sResult = Redis.Info()
    Debug.Print "INFO: " & Len(sResult) & " 字节"
    
    '=====================================================================
    ' 10. 数据库切换测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[10] 数据库切换测试"
    Debug.Print "----------------------------------------"
    
    Debug.Print "当前 DbIndex: " & Redis.DbIndex
    
    bResult = Redis.SelectDb(1)
    Debug.Print "SELECT 1: " & IIf(bResult, "成功", "失败")
    Debug.Print "DbIndex: " & Redis.DbIndex
    
    Call Redis.Set_("db1_key", "in_database_1")
    sResult = Redis.Get_("db1_key")
    Debug.Print "GET db1_key: " & sResult
    
    ' 切回数据库0
    Redis.DbIndex = 0
    bResult = Redis.SelectDb(0)
    Debug.Print "SELECT 0: " & IIf(bResult, "成功", "失败")
    Debug.Print "DbIndex: " & Redis.DbIndex
    
    lResult = Redis.Exists("db1_key")
    Debug.Print "EXISTS db1_key in db0: " & lResult & " (应为0)"
    
    ' 清理数据库1
    Redis.SelectDb (1)
    Redis.FlushDb
    Redis.SelectDb (0)
    
    '=====================================================================
    ' 11. 属性测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[11] 属性测试"
    Debug.Print "----------------------------------------"
    
    Debug.Print "Host: " & Redis.Host
    Debug.Print "Port: " & Redis.Port
    Debug.Print "Connected: " & Redis.Connected
    Debug.Print "DbIndex: " & Redis.DbIndex
    Debug.Print "Timeout: " & Redis.Timeout
    Debug.Print "InTransaction: " & Redis.InTransaction
    Debug.Print "InPipeline: " & Redis.InPipeline
    
    '=====================================================================
    ' 12. 错误处理测试
    '=====================================================================
    Debug.Print ""
    Debug.Print "[12] 错误处理测试"
    Debug.Print "----------------------------------------"
    
    ' 获取不存在的键
    sResult = Redis.Get_("nonexistent_key_12345")
    Debug.Print "GET nonexistent_key_12345: '" & sResult & "' (应为空)"
    
    ' 删除不存在的键
    lResult = Redis.Del("nonexistent_key_12345")
    Debug.Print "DEL nonexistent_key_12345: " & lResult & " (应为0)"
    
    '=====================================================================
    ' 13. 清理测试数据
    '=====================================================================
    Debug.Print ""
    Debug.Print "[13] 清理测试数据"
    Debug.Print "----------------------------------------"
    
    ' 清理所有测试键
    Redis.Del "expire_key"
    Redis.Del "counter"
    Redis.Del "discard_key"
    Redis.Del "db1_key"
    
    '=====================================================================
    ' 14. 断开连接
    '=====================================================================
    Debug.Print ""
    Debug.Print "[14] 断开连接"
    Debug.Print "----------------------------------------"
    
    Redis.DisConnect
    Debug.Print "Connected after Disconnect: " & Redis.Connected
    
    '=====================================================================
    ' 测试完成
    '=====================================================================
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "测试完成！"
    Debug.Print "=========================================="
    
    Exit Sub
    
TestError:
    Debug.Print ""
    Debug.Print "!!! 测试出错 !!!"
    Debug.Print "错误: " & ERR.Description
    Debug.Print "错误号: " & ERR.Number
    Debug.Print "=========================================="
End Sub
