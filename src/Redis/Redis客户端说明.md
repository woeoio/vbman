# cRedisClient - Redis 客户端说明文档

## 简介

`cRedisClient` 是一个简单易用的 Redis 数据库客户端，采用 VB6/VBA 编写。它支持 Redis RESP 协议，提供了常用的 Redis 命令接口，包括连接管理、基础操作、数据结构操作和事务等功能。

## 特性

- 支持基础 Redis 命令 (GET, SET, DEL, EXISTS, KEYS 等)
- 支持多种数据结构：String、Hash、List、Set、Sorted Set
- 支持管道模式
- 支持事务模式
- 自动重连机制
- 支持 Redis 密码认证
- 支持多数据库切换
- 同步操作，简单易用

## 连接方法

### 基本连接

```vb
Dim oRedis As New cRedisClient

' 连接到本地 Redis 服务器
If oRedis.Connect() Then
    Debug.Print "连接成功!"
Else
    Debug.Print "连接失败: " & oRedis.LastError
End If
```

### 连接到指定服务器

```vb
Dim oRedis As New cRedisClient

' 连接到指定地址和端口
If oRedis.Connect("192.168.1.100", 6379) Then
    Debug.Print "连接成功!"
End If
```

### 带密码认证的连接

```vb
Dim oRedis As New cRedisClient

' 连接并使用密码认证
If oRedis.Connect("127.0.0.1", 6379, "mypassword") Then
    Debug.Print "连接成功!"
End If
```

### 设置超时时间

```vb
Dim oRedis As New cRedisClient

' 设置超时时间为 10 秒
oRedis.Timeout = 10
If oRedis.Connect() Then
    Debug.Print "连接成功!"
End If
```

### 断开连接

```vb
oRedis.DisConnect
```

## 属性说明

| 属性          | 类型     | 说明                            |
| ------------- | -------- | ------------------------------- |
| Host          | String   | 获取 Redis 服务器地址           |
| Port          | Long     | 获取 Redis 服务器端口           |
| Connected     | Boolean  | 获取是否已连接                  |
| DbIndex       | Long     | 获取或设置当前数据库索引 (0-15) |
| Timeout       | Double   | 获取或设置超时时间（秒）        |
| InTransaction | Boolean  | 获取是否在事务中                |
| InPipeline    | Boolean  | 获取是否在管道模式              |
| LastError     | String   | 获取最后一次错误信息            |
| Socket        | cWinsock | 获取底层的 Socket 对象          |

## 基础命令

### Auth - 认证

```vb
' 使用密码认证
If oRedis.Auth("mypassword") Then
    Debug.Print "认证成功"
End If
```

### SelectDb - 选择数据库

```vb
' 切换到数据库 1
If oRedis.SelectDb(1) Then
    Debug.Print "已切换到数据库 1"
End If

' 或直接设置属性
oRedis.DbIndex = 2  ' 切换到数据库 2
```

### Ping - 测试连接

```vb
Dim sResult As String
sResult = oRedis.Ping()
Debug.Print sResult  ' 输出: PONG
```

### Info - 获取服务器信息

```vb
' 获取所有信息
Dim sInfo As String
sInfo = oRedis.Info()
Debug.Print sInfo

' 获取特定部分的信息
sInfo = oRedis.Info("server")
Debug.Print sInfo
```

## String 操作

### Set\_ - 设置键值

```vb
' 基本设置
oRedis.Set_ "name", "张三"

' 设置过期时间（秒）
oRedis.Set_ "session", "abc123", , 3600  ' 1小时后过期

' 设置过期时间（毫秒）
oRedis.Set_ "token", "xyz789", , 60000  ' 1分钟后过期
```

### Get\_ - 获取键值

```vb
Dim sValue As String
sValue = oRedis.Get_("name")
Debug.Print sValue  ' 输出: 张三
```

### Del - 删除键

```vb
' 删除单个键
oRedis.Del "name"

' 删除多个键
oRedis.Del "key1", "key2", "key3"
```

### Exists - 检查键是否存在

```vb
' 检查单个键
If oRedis.Exists("name") > 0 Then
    Debug.Print "键存在"
End If

' 检查多个键
Dim lCount As Long
lCount = oRedis.Exists("key1", "key2", "key3")
Debug.Print "存在 " & lCount & " 个键"
```

### Keys - 查找键

```vb
' 查找所有键
Dim vKeys As Variant
vKeys = oRedis.Keys("*")

' 查找以 "user:" 开头的键
vKeys = oRedis.Keys("user:*")

' 输出所有键
Dim i As Long
If IsArray(vKeys) Then
    For i = 0 To UBound(vKeys)
        Debug.Print vKeys(i)
    Next
End If
```

### Expire - 设置过期时间

```vb
' 设置键的过期时间（秒）
oRedis.Expire "name", 300  ' 5分钟后过期
```

### TTL - 获取剩余生存时间

```vb
Dim lTTL As Long
lTTL = oRedis.TTL("name")

If lTTL = -1 Then
    Debug.Print "键永不过期"
ElseIf lTTL = -2 Then
    Debug.Print "键不存在"
Else
    Debug.Print "剩余 " & lTTL & " 秒"
End If
```

### Incr - 自增

```vb
' 初始化计数器
oRedis.Set_ "counter", "10"

' 自增
Dim lValue As Long
lValue = oRedis.Incr("counter")
Debug.Print lValue  ' 输出: 11
```

### Decr - 自减

```vb
' 自减
lValue = oRedis.Decr("counter")
Debug.Print lValue  ' 输出: 10
```

### MGet - 批量获取

```vb
' 同时获取多个键
Dim vValues As Variant
vValues = oRedis.MGet("name", "age", "city")

If IsArray(vValues) Then
    Dim i As Long
    For i = 0 To UBound(vValues)
        Debug.Print vValues(i)
    Next
End If
```

### MSet - 批量设置

```vb
' 批量设置键值对
oRedis.MSet "name", "李四", "age", "25", "city", "北京"
```

## Hash 操作

### HSet - 设置 Hash 字段

```vb
' 设置单个字段
oRedis.HSet "user:1", "name", "张三"
oRedis.HSet "user:1", "age", "25"
oRedis.HSet "user:1", "email", "zhangsan@example.com"
```

### HGet - 获取 Hash 字段

```vb
Dim sValue As String
sValue = oRedis.HGet("user:1", "name")
Debug.Print sValue  ' 输出: 张三
```

### HMGet - 批量获取 Hash 字段

```vb
' 批量获取字段
Dim vValues As Variant
vValues = oRedis.HMGet("user:1", "name", "age", "email")

If IsArray(vValues) Then
    Dim i As Long
    For i = 0 To UBound(vValues)
        Debug.Print vValues(i)
    Next
End If
```

### HGetAll - 获取所有 Hash 字段

```vb
' 获取所有字段和值
Dim oDict As Scripting.Dictionary
Set oDict = oRedis.HGetAll("user:1")

' 遍历字典
Dim vKey As Variant
For Each vKey In oDict.Keys
    Debug.Print vKey & ": " & oDict(vKey)
Next
```

### HDel - 删除 Hash 字段

```vb
' 删除单个字段
oRedis.HDel "user:1", "email"

' 删除多个字段
oRedis.HDel "user:1", "age", "email"
```

### HExists - 检查 Hash 字段是否存在

```vb
If oRedis.HExists("user:1", "name") Then
    Debug.Print "字段存在"
End If
```

## List 操作

### LPush - 从左侧插入

```vb
' 插入单个值
oRedis.LPush "mylist", "item1"

' 插入多个值
oRedis.LPush "mylist", "item2", "item3"
```

### RPush - 从右侧插入

```vb
' 从右侧插入
oRedis.RPush "mylist", "item4"
```

### LPop - 从左侧弹出

```vb
Dim sValue As String
sValue = oRedis.LPop("mylist")
Debug.Print sValue  ' 输出最左侧的元素
```

### RPop - 从右侧弹出

```vb
sValue = oRedis.RPop("mylist")
Debug.Print sValue  ' 输出最右侧的元素
```

### LLen - 获取列表长度

```vb
Dim lLen As Long
lLen = oRedis.lLen("mylist")
Debug.Print "列表长度: " & lLen
```

### LRange - 获取列表范围内的元素

```vb
' 获取所有元素
Dim vItems As Variant
vItems = oRedis.LRange("mylist", 0, -1)

If IsArray(vItems) Then
    Dim i As Long
    For i = 0 To UBound(vItems)
        Debug.Print vItems(i)
    Next
End If

' 获取前 10 个元素
vItems = oRedis.LRange("mylist", 0, 9)

' 获取从第 6 个到第 10 个元素
vItems = oRedis.LRange("mylist", 5, 9)
```

## Set 操作

### SAdd - 添加集合成员

```vb
' 添加单个成员
oRedis.SAdd "myset", "apple"

' 添加多个成员
oRedis.SAdd "myset", "banana", "orange", "grape"
```

### SMembers - 获取所有集合成员

```vb
' 获取所有成员
Dim vMembers As Variant
vMembers = oRedis.SMembers("myset")

If IsArray(vMembers) Then
    Dim i As Long
    For i = 0 To UBound(vMembers)
        Debug.Print vMembers(i)
    Next
End If
```

### SCard - 获取集合成员数量

```vb
Dim lCount As Long
lCount = oRedis.SCard("myset")
Debug.Print "集合成员数: " & lCount
```

### SIsMember - 检查成员是否存在

```vb
If oRedis.SIsMember("myset", "apple") Then
    Debug.Print "成员存在"
End If
```

### SRem - 删除集合成员

```vb
' 删除单个成员
oRedis.SRem "myset", "apple"

' 删除多个成员
oRedis.SRem "myset", "banana", "orange"
```

## Sorted Set 操作

### ZAdd - 添加有序集合成员

```vb
' 添加带分数的成员
oRedis.ZAdd "mysortedset", 100, "member1"
oRedis.ZAdd "mysortedset", 200, "member2"
oRedis.ZAdd "mysortedset", 150, "member3"
```

### ZRange - 获取范围内的成员

```vb
' 获取所有成员（按分数升序）
Dim vMembers As Variant
vMembers = oRedis.ZRange("mysortedset", 0, -1)

If IsArray(vMembers) Then
    Dim i As Long
    For i = 0 To UBound(vMembers)
        Debug.Print vMembers(i)
    Next
End If

' 获取带分数的成员
vMembers = oRedis.ZRange("mysortedset", 0, -1, True)
```

### ZRem - 删除有序集合成员

```vb
' 删除单个成员
oRedis.ZRem "mysortedset", "member1"

' 删除多个成员
oRedis.ZRem "mysortedset", "member2", "member3"
```

### ZCard - 获取有序集合成员数量

```vb
Dim lCount As Long
lCount = oRedis.ZCard("mysortedset")
Debug.Print "有序集合成员数: " & lCount
```

## 事务操作

### Multi - 开始事务

```vb
' 开始事务
oRedis.Multi
```

### Exec - 提交事务

```vb
' 提交事务并获取执行结果
Dim vResults As Variant
vResults = oRedis.Exec()

If IsArray(vResults) Then
    Dim i As Long
    For i = 0 To UBound(vResults)
        Debug.Print vResults(i)
    Next
End If
```

### Discard - 取消事务

```vb
' 取消事务
oRedis.Discard
```

### 事务使用示例

```vb
' 开始事务
oRedis.Multi

' 执行多个命令（不会被立即执行）
oRedis.Set_ "key1", "value1"
oRedis.Set_ "key2", "value2"
oRedis.Incr "counter"

' 提交事务
Dim vResults As Variant
vResults = oRedis.Exec()

' vResults 包含所有命令的执行结果
```

## 服务器操作

### FlushDb - 清空当前数据库

```vb
' 清空当前数据库中的所有键
If oRedis.FlushDb() Then
    Debug.Print "数据库已清空"
End If
```

## 完整示例

### 示例 1：基本的 String 操作

```vb
Sub Example1_BasicString()
    Dim oRedis As New cRedisClient

    ' 连接
    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 设置值
    oRedis.Set_ "name", "张三"
    oRedis.Set_ "age", "25"

    ' 获取值
    Debug.Print "姓名: " & oRedis.Get_("name")
    Debug.Print "年龄: " & oRedis.Get_("age")

    ' 设置过期时间
    oRedis.Expire "name", 300

    ' 检查剩余时间
    Debug.Print "name 的剩余生存时间: " & oRedis.TTL("name") & " 秒"

    ' 断开连接
    oRedis.DisConnect
End Sub
```

### 示例 2：Hash 操作

```vb
Sub Example2_HashOperations()
    Dim oRedis As New cRedisClient

    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 设置用户信息
    oRedis.HSet "user:1001", "name", "李四"
    oRedis.HSet "user:1001", "age", "30"
    oRedis.HSet "user:1001", "city", "上海"
    oRedis.HSet "user:1001", "email", "lisi@example.com"

    ' 获取单个字段
    Debug.Print "用户姓名: " & oRedis.HGet("user:1001", "name")

    ' 获取所有字段
    Dim oDict As Scripting.Dictionary
    Set oDict = oRedis.HGetAll("user:1001")

    Debug.Print vbCrLf & "用户详细信息:"
    Dim vKey As Variant
    For Each vKey In oDict.Keys
        Debug.Print "  " & vKey & ": " & oDict(vKey)
    Next

    oRedis.DisConnect
End Sub
```

### 示例 3：List 操作

```vb
Sub Example3_ListOperations()
    Dim oRedis As New cRedisClient

    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 清空现有列表
    oRedis.Del "tasks"

    ' 添加任务
    oRedis.RPush "tasks", "任务1"
    oRedis.RPush "tasks", "任务2"
    oRedis.RPush "tasks", "任务3"
    oRedis.RPush "tasks", "任务4"
    oRedis.RPush "tasks", "任务5"

    ' 获取列表长度
    Debug.Print "任务列表长度: " & oRedis.lLen("tasks")

    ' 获取所有任务
    Dim vTasks As Variant
    vTasks = oRedis.LRange("tasks", 0, -1)

    Debug.Print vbCrLf & "所有任务:"
    Dim i As Long
    For i = 0 To UBound(vTasks)
        Debug.Print "  " & (i + 1) & ". " & vTasks(i)
    Next

    ' 处理第一个任务
    Debug.Print vbCrLf & "处理任务: " & oRedis.LPop("tasks")
    Debug.Print "剩余任务数: " & oRedis.lLen("tasks")

    oRedis.DisConnect
End Sub
```

### 示例 4：事务操作

```vb
Sub Example4_Transaction()
    Dim oRedis As New cRedisClient

    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 初始化计数器
    oRedis.Set_ "counter", "0"
    Debug.Print "初始计数: " & oRedis.Get_("counter")

    ' 开始事务
    oRedis.Multi

    ' 执行多个自增操作
    oRedis.Incr "counter"
    oRedis.Incr "counter"
    oRedis.Incr "counter"

    ' 提交事务
    Dim vResults As Variant
    vResults = oRedis.Exec()

    Debug.Print "事务执行结果:"
    If IsArray(vResults) Then
        For i = 0 To UBound(vResults)
            Debug.Print "  操作 " & (i + 1) & " 结果: " & vResults(i)
        Next
    End If

    Debug.Print "最终计数: " & oRedis.Get_("counter")

    oRedis.DisConnect
End Sub
```

### 示例 5：Set 和 Sorted Set 操作

```vb
Sub Example5_Sets()
    Dim oRedis As New cRedisClient

    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 使用 Set 存储标签
    oRedis.SAdd "post:1:tags", "技术", "编程", "Redis", "VB6"

    Debug.Print "文章标签数: " & oRedis.SCard("post:1:tags")

    Dim vTags As Variant
    vTags = oRedis.SMembers("post:1:tags")

    Debug.Print "所有标签:"
    Dim i As Long
    For i = 0 To UBound(vTags)
        Debug.Print "  - " & vTags(i)
    Next

    ' 使用 Sorted Set 存储排行榜
    oRedis.ZAdd "leaderboard", 1000, "玩家A"
    oRedis.ZAdd "leaderboard", 1500, "玩家B"
    oRedis.ZAdd "leaderboard", 800, "玩家C"
    oRedis.ZAdd "leaderboard", 2000, "玩家D"

    Debug.Print vbCrLf & "排行榜（按分数升序）:"
    vTags = oRedis.ZRange("leaderboard", 0, -1)
    For i = 0 To UBound(vTags)
        Debug.Print "  " & (i + 1) & ". " & vTags(i)
    Next

    Debug.Print vbCrLf & "排行榜（带分数）:"
    vTags = oRedis.ZRange("leaderboard", 0, -1, True)
    For i = 0 To UBound(vTags) Step 2
        If i + 1 <= UBound(vTags) Then
            Debug.Print "  " & vTags(i) & ": " & vTags(i + 1) & " 分"
        End If
    Next

    oRedis.DisConnect
End Sub
```

### 示例 6：数据库切换和批量操作

```vb
Sub Example6_MultipleDb()
    Dim oRedis As New cRedisClient

    If Not oRedis.Connect() Then
        Debug.Print "连接失败: " & oRedis.LastError
        Exit Sub
    End If

    ' 在数据库 0 中存储用户数据
    oRedis.SelectDb 0
    oRedis.Set_ "user:1", "张三"
    oRedis.Set_ "user:2", "李四"
    Debug.Print "数据库0中的用户: " & oRedis.Keys("user:*")

    ' 在数据库 1 中存储配置数据
    oRedis.SelectDb 1
    oRedis.Set_ "config:appname", "MyApp"
    oRedis.Set_ "config:version", "1.0.0"
    oRedis.Set_ "config:debug", "false"
    Debug.Print "数据库1中的配置: " & oRedis.Keys("config:*")

    ' 批量获取配置
    Dim vConfigs As Variant
    vConfigs = oRedis.MGet("config:appname", "config:version")
    Debug.Print vbCrLf & "配置信息:"
    If IsArray(vConfigs) Then
        For i = 0 To UBound(vConfigs)
            Debug.Print "  " & vConfigs(i)
        Next
    End If

    ' 切换回数据库 0
    oRedis.SelectDb 0
    Debug.Print vbCrLf & "当前数据库: " & oRedis.DbIndex

    oRedis.DisConnect
End Sub
```

## 事件处理

`cRedisClient` 提供了以下事件用于处理连接状态和错误：

### OnDisconnected - 连接断开事件

```vb
' 在类模块中声明 WithEvents 变量
Private WithEvents m_oRedis As cRedisClient

Private Sub m_oRedis_OnDisconnected()
    Debug.Print "Redis 连接已断开"
    ' 可以在这里实现重连逻辑
End Sub
```

### OnError - 错误事件

```vb
Private Sub m_oRedis_OnError(ByVal ErrorMsg As String)
    Debug.Print "Redis 错误: " & ErrorMsg
    ' 可以在这里实现错误处理逻辑
End Sub
```

## 注意事项

1. **连接管理**：使用完毕后记得调用 `DisConnect` 方法释放连接
2. **错误处理**：建议在操作前后检查连接状态和 `LastError` 属性
3. **超时设置**：根据网络环境合理设置 `Timeout` 属性
4. **数据类型**：Redis 中的所有值都以字符串形式存储，数值操作需要转换
5. **事务安全**：事务中的命令不会立即执行，直到调用 `Exec` 或 `Discard`
6. **批量操作**：使用 `MGet`、`MSet` 等批量操作可以提高性能
7. **资源释放**：在对象不再使用时，将其设为 `Nothing` 释放资源

## 依赖项

- `cWinsock` - Socket 通信组件
- `ToolsUtf8` - UTF-8 编解码工具
- `Scripting.Dictionary` - 用于 `HGetAll` 返回字典对象

## 许可证

请参考项目主文档了解许可证信息。

