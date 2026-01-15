# cByteBuffer 类参考

## 📋 类概述

`cByteBuffer` 是高效的字节缓冲区类，用于管理二进制数据流。

### 设计特点

- **预分配** - 初始分配 4KB，避免频繁内存分配
- **智能增长** - 按 1.5 倍增长，平衡空间和性能
- **最小化拷贝** - 使用 CopyMemory API 提高效率
- **Peek/Consume** - 支持查看数据而不消费

---

## 🔧 属性参考

### Size - 当前数据大小

**类型**: `Long`  
**读写**: 只读

**说明**: 缓冲区中当前存储的数据字节数。

```vb
Debug.Print "缓冲区大小: " & oBuffer.Size & " 字节"

If oBuffer.Size > 0 Then
    ProcessData oBuffer
End If
```

---

### Capacity - 缓冲区容量

**类型**: `Long`  
**读写**: 只读

**说明**: 缓冲区当前分配的总容量（字节）。

```vb
Debug.Print "容量: " & oBuffer.Capacity & " 字节"
Debug.Print "使用率: " & (oBuffer.Size / oBuffer.Capacity * 100) & "%"
```

---

### IsEmpty - 是否为空

**类型**: `Boolean`  
**读写**: 只读

**说明**: 缓冲区是否为空（Size = 0）。

```vb
If oBuffer.IsEmpty Then
    Debug.Print "缓冲区为空"
Else
    Debug.Print "缓冲区有 " & oBuffer.Size & " 字节数据"
End If
```

---

## 🚀 方法参考

### Append - 追加数据

**语法**:

```vb
Public Sub Append(ByRef Data() As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Data()` | Byte() | 要追加的字节数组 |

**说明**: 
- 如果需要，自动扩容
- 数据追加到缓冲区末尾

**示例**:

```vb
' 追加字节数组
Dim baData() As Byte
baData = StringToUTF8("Hello")
oBuffer.Append baData

' 追加接收到的网络数据
Private Sub Socket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Client.GetData baData, vbByte + vbArray
    oBuffer.Append baData
    
    ' 处理缓冲区
    ProcessBuffer
End Sub
```

---

### AppendByte - 追加单个字节

**语法**:

```vb
Public Sub AppendByte(ByVal Value As Byte)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Value` | Byte | 要追加的字节值 |

**示例**:

```vb
' 构建协议头
oBuffer.AppendByte &H01  ' 版本
oBuffer.AppendByte &H02  ' 类型
oBuffer.AppendByte &H03  ' 标志
```

---

### Peek - 查看数据（不消费）

**语法**:

```vb
Public Function Peek(ByVal Offset As Long, ByVal Length As Long, ByRef OutData() As Byte) As Boolean
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Offset` | Long | 偏移位置（从 0 开始） |
| `Length` | Long | 要查看的字节数 |
| `OutData()` | Byte() | 输出字节数组 |

**返回值**: `Boolean` - 成功返回 `True`，否则返回 `False`

**说明**: 
- 只读操作，不修改缓冲区
- 适合用于预检查数据

**示例**:

```vb
' 检查前 4 字节是否为特定值
Dim baHeader() As Byte
If oBuffer.Peek(0, 4, baHeader) Then
    If baHeader(0) = &HDE And baHeader(1) = &HAD Then
        Debug.Print "检测到魔数"
    End If
End If
```

---

### PeekByte - 查看单个字节

**语法**:

```vb
Public Function PeekByte(ByVal Offset As Long) As Byte
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Offset` | Long | 偏移位置 |

**返回值**: 读取到的字节值

**示例**:

```vb
' 检查第一个字节
Dim bFirst As Byte
bFirst = oBuffer.PeekByte(0)
Debug.Print "第一个字节: " & Hex$(bFirst)
```

---

### Consume - 消费数据

**语法**:

```vb
Public Sub Consume(ByVal Length As Long)
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Length` | Long | 要消费的字节数 |

**说明**: 
- 从缓冲区前面移除指定字节数
- 剩余数据会向前移动
- 如果消费全部，缓冲区变空

**示例**:

```vb
' 消费前 4 字节
oBuffer.Consume 4

' 消费已处理的数据
oBuffer.Consume lProcessedBytes
```

---

### Extract - 提取并消费数据

**语法**:

```vb
Public Function Extract(ByVal Length As Long) As Byte()
```

**参数**:

| 参数 | 类型 | 说明 |
|------|------|------|
| `Length` | Long | 要提取的字节数 |

**返回值**: 提取的字节数组

**说明**: 
- 返回指定长度的数据
- 自动从缓冲区中移除

**示例**:

```vb
' 提取并消费前 10 字节
Dim baData() As Byte
baData = oBuffer.Extract(10)
Debug.Print "提取了 " & (UBound(baData) + 1) & " 字节"

' WebSocket 帧提取示例
Private Sub ExtractFrame(oBuffer As cByteBuffer)
    Dim oFrame As New cWebSocketFrame
    
    If oFrame.ParseHeader(oBuffer) Then
        If oFrame.IsCompleteFrame(oBuffer) Then
            Dim baFrame() As Byte
            baFrame = oBuffer.Extract(oFrame.TotalFrameLength)
            ' 处理帧...
        End If
    End If
End Sub
```

---

### ToArray - 获取所有数据

**语法**:

```vb
Public Function ToArray() As Byte()
```

**返回值**: 缓冲区所有数据的副本

**说明**: 
- 返回数据的副本，不影响缓冲区
- 空缓冲区返回空数组

**示例**:

```vb
' 获取所有数据
Dim baAll() As Byte
baAll = oBuffer.ToArray
Debug.Print "总数据: " & (UBound(baAll) + 1) & " 字节"

' 保存到文件
SaveToFile "data.bin", baAll
```

---

### Clear - 清空缓冲区

**语法**:

```vb
Public Sub Clear()
```

**说明**: 清除所有数据，但保留容量。

**示例**:

```vb
' 清空缓冲区
oBuffer.Clear
Debug.Print "缓冲区已清空，大小: " & oBuffer.Size
```

---

### Reset - 重置缓冲区

**语法**:

```vb
Public Sub Reset()
```

**说明**: 
- 清除所有数据
- 重置容量为初始值（4KB）

**示例**:

```vb
' 完全重置
oBuffer.Reset
Debug.Print "缓冲区已重置，容量: " & oBuffer.Capacity
```

---

### GetBufferPtr - 获取缓冲区指针

**语法**:

```vb
Public Function GetBufferPtr() As Long
```

**返回值**: 内部缓冲区的内存地址

**说明**: 
- 用于高性能场景
- ⚠️ 危险操作，不要超出 Size 范围写入

**示例**:

```vb
' 高性能填充（仅高级用法）
Dim pBuffer As Long
pBuffer = oBuffer.GetBufferPtr

If pBuffer <> 0 Then
    ' 使用 CopyMemory API 直接写入
    CopyMemory ByVal pBuffer, baData(0), UBound(baData) + 1
    oBuffer.Size = oBuffer.Size + UBound(baData) + 1
End If
```

---

## 📝 使用示例

### WebSocket 帧解析

```vb
Private Sub ProcessWebSocketData(oBuffer As cByteBuffer)
    Dim oFrame As New cWebSocketFrame
    
    Do While oBuffer.Size >= 2
        ' 解析头部
        If Not oFrame.ParseHeader(oBuffer) Then
            Exit Do ' 需要更多数据
        End If
        
        ' 检查完整性
        If Not oFrame.IsCompleteFrame(oBuffer) Then
            Exit Do ' 需要更多数据
        End If
        
        ' 提取帧
        Dim baFrame() As Byte
        baFrame = oBuffer.Extract(oFrame.TotalFrameLength)
        
        ' 处理帧
        ProcessFrame baFrame, oFrame
    Loop
End Sub
```

### 网络数据接收

```vb
Private WithEvents m_Socket As cWinsock
Private m_RecvBuffer As cByteBuffer

Private Sub Form_Load()
    Set m_RecvBuffer = New cByteBuffer
End Sub

Private Sub m_Socket_DataArrival(Client As cWinsock, ByVal bytesTotal As Long)
    Dim baData() As Byte
    Client.GetData baData, vbByte + vbArray
    
    ' 追加到缓冲区
    m_RecvBuffer.Append baData
    
    ' 处理缓冲区中的完整消息
    ProcessBuffer
End Sub

Private Sub ProcessBuffer()
    Do While m_RecvBuffer.Size >= 4
        ' 读取消息长度（假设前 4 字节是长度）
        Dim lLen As Long
        lLen = CLng(m_RecvBuffer.PeekByte(0)) * 256& ^ 3 + _
                 CLng(m_RecvBuffer.PeekByte(1)) * 256& ^ 2 + _
                 CLng(m_RecvBuffer.PeekByte(2)) * 256& + _
                 CLng(m_RecvBuffer.PeekByte(3))
        
        ' 检查是否有完整消息
        If m_RecvBuffer.Size < 4 + lLen Then
            Exit Do ' 需要更多数据
        End If
        
        ' 提取消息头
        m_RecvBuffer.Consume 4
        
        ' 提取消息体
        Dim baMessage() As Byte
        baMessage = m_RecvBuffer.Extract(lLen)
        
        ' 处理消息
        ProcessMessage baMessage
    Loop
End Sub
```

### 协议头构建

```vb
Private Function BuildProtocolHeader() As cByteBuffer
    Dim oBuffer As New cByteBuffer
    
    ' 构建协议头
    oBuffer.AppendByte &H01              ' 版本
    oBuffer.AppendByte &H00              ' 类型
    oBuffer.AppendByte &H00              ' 标志
    oBuffer.AppendByte &H00              ' 保留
    
    ' 添加长度（4 字节）
    Dim lLen As Long
    lLen = 1234
    oBuffer.AppendByte (lLen And &HFF000000) \ &H1000000
    oBuffer.AppendByte (lLen And &HFF0000) \ &H10000
    oBuffer.AppendByte (lLen And &HFF00&) \ &H100&
    oBuffer.AppendByte (lLen And &HFF&)
    
    Set BuildProtocolHeader = oBuffer
End Function
```

### 数据分片处理

```vb
Private m_FragmentBuffer As cByteBuffer

Private Sub HandleFragmentedFrame(oFrame As cWebSocketFrame, oBuffer As cByteBuffer)
    ' 提取帧
    Dim baPayload() As Byte
    baPayload = oBuffer.Extract(oFrame.TotalFrameLength)
    
    If m_FragmentBuffer.IsEmpty Then
        ' 第一个分片
        If Not oFrame.FIN Then
            ' 开始分片消息
            m_FragmentBuffer.Clear
            m_FragmentBuffer.Append baPayload
        Else
            ' 单一帧（无分片）
            ProcessCompleteMessage baPayload, oFrame.OpCode
        End If
    Else
        ' 后续分片
        m_FragmentBuffer.Append baPayload
        
        If oFrame.FIN Then
            ' 最后一个分片
            Dim baComplete() As Byte
            baComplete = m_FragmentBuffer.ToArray
            ProcessCompleteMessage baComplete, oFrame.OpCode
            m_FragmentBuffer.Clear
        End If
    End If
End Sub
```

---

## ⚠️ 注意事项

1. **容量自动增长** - Append 时如果超出容量，会自动扩容
2. **Peek 是只读** - Peek 操作不会修改缓冲区
3. **Consume 会移除** - Consume 后数据会被删除
4. **Extract 会消费** - Extract 相当于 Peek + Consume
5. **GetBufferPtr 危险** - 仅限高级使用，确保不越界

---

## 🔍 性能优化

### 批量 Append

```vb
' ✅ 好的做法：一次 Append
Dim baData() As Byte
baData = BuildLargeData()
oBuffer.Append baData

' ❌ 不好的做法：多次 Append
For i = 0 To 10000
    oBuffer.AppendByte baData(i)
Next i
```

### 避免频繁 ToArray

```vb
' ✅ 好的做法：直接使用 Peek/Consume
If oBuffer.Peek(0, 4, baHeader) Then
    ProcessHeader baHeader
End If

' ❌ 不好的做法：频繁 ToArray
Dim baAll() As Byte
baAll = oBuffer.ToArray
For i = 0 To 100
    ProcessSegment baAll, i * 100, 100
Next i
```

---

**最后更新**: 2026-01-10
