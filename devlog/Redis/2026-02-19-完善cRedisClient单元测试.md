# 2026-02-19-完善cRedisClient单元测试

## src/Redis/cRedisClient.cls

### 问题修复

#### 1. ParamArray 嵌套数组处理

**问题**: DEL/EXISTS 等命令使用 ParamArray 时，Array() 包装导致参数被当作单个字符串传递，而非展开为多个参数。

**解决方案**: 新增 `FlattenArray()` 函数递归展平嵌套数组。

```vb
Private Function FlattenArray(ByVal vInput As Variant) As Variant()
    ' 递归展平嵌套数组
End Function
```

#### 2. UTF-8 编码支持

**问题**: BuildRedisCommand() 返回的字符串未正确处理 UTF-8 编码，导致中文字符乱码。

**解决方案**: 修改 `BuildRedisCommand()` 直接返回 UTF-8 编码的 Byte() 数组。

```vb
Private Function BuildRedisCommand(ByVal Command As Variant) As Byte()
    ' 使用 ToolsUtf8.Encode 编码字符串为 UTF-8
    baPart = ToolsUtf8.Encode(sPart)
End Function
```

#### 3. RESP 协议批量字符串解析

**问题**: ParseResp() 使用 RESP 协议的 UTF-8 字节长度通过 Mid() 截取字符串，但数据已被 ReceiveElement() 解码为 VB6 字符串，导致字符长度与字节长度不匹配。

**解决方案**: 修改 ParseResp() 中的批量字符串解析逻辑，通过查找下一个 \r\n 的位置确定字符串实际长度。

```vb
Case "$"
    ' 批量字符串
    lCrLf = InStr(lPos, Buffer, vbCrLf)
    lLength = Val(Mid(Buffer, lPos, lCrLf - lPos))
    lPos = lCrLf + 2
    If lLength = -1 Then
        ParseResp = Null
    Else
        ' 查找下一个 \r\n 确定实际长度
        Dim lActualCrLf As Long
        lActualCrLf = InStr(lPos, Buffer, vbCrLf)
        ParseResp = Mid(Buffer, lPos, lActualCrLf - lPos)
        lPos = lActualCrLf + 2
    End If
```

## src/Tools/Demo_RedisTest.bas

### 新增功能

创建完整的 cRedisClient 单元测试，包含 14 个测试模块：

1. **连接测试** - Connect/Disconnect/PING
2. **基础命令测试** - SET/GET/DEL/EXISTS/KEYS/TTL
3. **String 操作测试** - INCR/DECR/MGET
4. **Hash 操作测试** - HSET/HGET/HMGET/HGETALL/HEXISTS/HDEL
5. **List 操作测试** - LPUSH/RPUSH/LLEN/LRANGE/LPOP/RPOP
6. **Set 操作测试** - SADD/SMEMBERS/SCARD/SISMEMBER/SREM
7. **Sorted Set 操作测试** - ZADD/ZRANGE/ZCARD/ZREM
8. **事务操作测试** - MULTI/EXEC/DISCARD
9. **服务器操作测试** - PING/INFO
10. **数据库切换测试** - SELECT/DbIndex
11. **属性测试** - Host/Port/Connected/Timeout
12. **错误处理测试** - 不存在的键处理
13. **清理测试数据** - 测试后清理
14. **断开连接** - 验证连接状态

### 测试结果

✅ 所有 14 个测试模块通过
✅ 中文内容（张三、北京）正确存储和读取
✅ Hash 操作（HMGET/HGETALL）正确返回数组
✅ 所有数据结构操作正常
✅ 事务、数据库切换等高级功能正常

## 总结

- 修复了 ParamArray、UTF-8 编码、RESP 协议解析三个核心问题
- 完善了覆盖所有 Redis 数据结构和功能的单元测试
- 验证了 cRedisClient 对中文内容的完整支持
