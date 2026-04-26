# cDelay 延时对象开发文档

```vb
'===========================================================================
' 名称: cDelay
' 描述: 延时对象类 - 支持事件触发、回调函数和同步等待三种模式
' 作者: 邓伟，QQ: 215879458
' 网站: https://vb6.pro
' 日期: 2026-03-31
' 参考: cTimer.cls 架构
'===========================================================================
```

## 概述

`cDelay` 是一个 VB6 延时对象类，提供三种使用模式：

- **事件模式**：通过 `OnTime` 事件触发
- **回调模式**：通过 `CallByName` 调用指定函数
- **同步模式**：阻塞等待，不卡 UI

## 文件结构

```
Delay/
├── cDelay.cls      ' 延时对象类
├── mDelay.bas      ' 管理模块（全局集合、定时器回调）
└── cDelayDemo.frm  ' 演示窗体
```

## 特性

- ✅ 自动生命周期管理（引用计数）
- ✅ 支持最多 9 个回调参数（使用 ParamArray）
- ✅ 同步模式支持取消操作
- ✅ 消息泵实现，不依赖 DoEvents
- ✅ 线程安全的消息循环

---

## 快速开始

### 1. 事件模式

```vb
Private WithEvents m_Delay As cDelay

Private Sub Form_Load()
    Set m_Delay = New cDelay
End Sub

Private Sub btnStart_Click()
    m_Delay.CountDown 3000  ' 3秒后触发 OnTime
End Sub

Private Sub m_Delay_OnTime()
    MsgBox "时间到了！"
End Sub
```

### 2. 回调模式

```vb
Private m_Delay As cDelay

Private Sub Form_Load()
    Set m_Delay = New cDelay
End Sub

Private Sub btnStart_Click()
    ' 链式调用，传递多个参数
    m_Delay.Callback(Me, "MyFunc", "参数1", 123, Now).CountDown 3000
End Sub

Public Sub MyFunc(ByVal s As String, ByVal n As Long, ByVal d As Date)
    MsgBox s & ", " & n
End Sub
```

### 3. 同步模式

```vb
Private m_DelaySync As cDelay  ' 必须是模块级变量才能取消

Private Sub btnStart_Click()
    lblStatus.Caption = "等待中..."
    lblStatus.Refresh

    m_DelaySync.Sync().CountDown 3000  ' 等待3秒，UI保持响应

    ' 检查是否被取消
    If m_DelaySync.IsCancelled Then
        lblStatus.Caption = "已取消"
    Else
        lblStatus.Caption = "完成！"
    End If
End Sub

Private Sub btnCancel_Click()
    m_DelaySync.Cancel  ' 取消同步等待
End Sub
```

---

## API 参考

### 属性

| 属性          | 类型        | 说明                         |
| ------------- | ----------- | ---------------------------- |
| `Mode`        | `DelayMode` | 当前模式（只读）             |
| `IsActive`    | `Boolean`   | 是否处于活动状态（只读）     |
| `IsCancelled` | `Boolean`   | 同步模式下是否被取消（只读） |
| `DelayMs`     | `Long`      | 延迟毫秒数（只读）           |

> 📊 **对比分析**：[cDelay vs 原生 DoEvents 对比](./whyDelay.md)

### 方法

#### `Callback(CallbackObject, ProcName, [ParamArray P()])`

设置回调模式。

**参数：**

- `CallbackObject` - 回调方法所在的对象（通常是 `Me`）
- `ProcName` - 回调方法名称（字符串）
- `P()` - 可选参数，最多 9 个

**返回：** `cDelay` 对象自身（支持链式调用）

**示例：**

```vb
delay.Callback(Me, "ProcessData", userId, userName).CountDown(2000)
```

#### `Sync()`

设置同步模式。

**返回：** `cDelay` 对象自身（支持链式调用）

**示例：**

```vb
delay.Sync().CountDown(5000)
```

#### `CountDown(Milliseconds)`

开始倒计时。

**参数：**

- `Milliseconds` - 延迟时间（毫秒）

**注意：**

- 事件/回调模式：立即返回，时间到后触发事件/回调
- 同步模式：阻塞等待，UI保持响应

#### `Cancel()`

取消延迟。

**行为：**

- 停止定时器（事件/回调模式）
- 设置取消标志（同步模式）
- 清理回调信息
- 触发 `OnCancel` 事件

### 事件

#### `OnTime()`

时间到达时触发（仅事件模式）。

#### `OnCancel()`

延迟被取消时触发。

---

## 同步模式详解

### 为什么需要模块级变量？

同步模式下 `CountDown` 是**阻塞调用**：

```vb
' 错误：局部变量无法在阻塞期间操作
Private Sub btnStart_Click()
    Dim delay As New cDelay
    delay.Sync().CountDown(3000)  ' 阻塞在这里！
    ' 下面的代码在3秒后才执行
End Sub

' 正确：模块级变量可以在其他事件中访问
Private m_Delay As cDelay

Private Sub btnStart_Click()
    Set m_Delay = New cDelay
    m_Delay.Sync().CountDown(3000)  ' 阻塞
End Sub

Private Sub btnCancel_Click()
    m_Delay.Cancel  ' 可以取消上面的等待
End Sub
```

### 取消机制

```vb
Private Sub btnStart_Click()
    m_Delay.Sync().CountDown(10000)  ' 10秒等待

    ' 检查取消状态
    If m_Delay.IsCancelled Then
        ' 用户提前取消了
    Else
        ' 正常完成
    End If
End Sub
```

---

## 技术实现

### 消息泵原理

同步模式使用 Windows API 实现消息循环：

```vb
Private Declare Function PeekMessage Lib "user32" ...
Private Declare Function DispatchMessage Lib "user32" ...

Do While Not bDone
    ' 检查取消标志
    If m_Cancelled Then Exit Do

    ' 检查时间
    If GetTickCount() >= EndTick Then Exit Do

    ' 处理消息队列
    Do While PeekMessage(uMsg, 0, 0, 0, PM_REMOVE) <> 0
        TranslateMessage uMsg
        DispatchMessage uMsg
    Loop

    Sleep 1  ' 避免CPU占用过高
Loop
```

相比 `DoEvents` 的优势：

- 更精细的控制
- 可以处理特定消息
- 避免重入问题

### 自动生命周期管理

```
第一个实例创建
    ↓
AutoInit() 调用 InitDelaySystem()
    ↓
引用计数 = 1

...

最后一个实例销毁
    ↓
AutoCleanup() 检测到计数 = 0
    ↓
调用 TermDelaySystem() 清理全局资源
```

---

## 注意事项

1. **同步模式必须配合模块级变量使用**，否则无法取消

2. **取消后回调被清理**，需要重新设置 `Callback` 才能再次使用

3. **参数传递使用 Variant**，回调函数参数类型要匹配

4. **避免在回调中阻塞**，否则会影响其他定时器

5. **窗体卸载时清理**：
   ```vb
   Private Sub Form_Unload(Cancel As Integer)
       m_Delay.Cancel
       Set m_Delay = Nothing
   End Sub
   ```

---

## 常见问题

### Q: 同步模式下 UI 还会响应吗？

A: 会。消息泵会处理 Windows 消息，包括鼠标、键盘、重绘等。

### Q: 可以同时运行多个延迟吗？

A: 可以。每个 `cDelay` 实例独立运行，通过全局集合管理。（将来增加 cDelays.cls 使用 Tag 管理多实例）

### Q: 回调模式可以调用私有方法吗？

A: 不可以。回调方法必须是 `Public` 的，因为 `CallByName` 需要外部可见。

### Q: 同步模式的精度如何？

A: 受 `Sleep 1` 影响，精度约为 1-15 毫秒（取决于系统调度）。

---

## 更新日志

### v1.0 (2026-03-31)

- 初始版本
- 支持三种模式：事件、回调、同步
- 自动生命周期管理
- 同步模式支持取消

---

## 未来扩展功能预览

以下是计划中将与 `CountDown` 同级的新增延迟方式：

### 新增延迟方式概览

| 方法            | 用途         | 示例场景              |
| --------------- | ------------ | --------------------- |
| **ClockAt**     | 指定时间执行 | 每天14:30提醒         |
| **ClockAtDate** | 指定日期时间 | 2026-04-01会议提醒    |
| **Interval**    | 重复间隔执行 | 每5秒检查状态         |
| **Random**      | 随机时间范围 | 模拟人工操作防检测    |
| **WaitUntil**   | 条件满足执行 | 等待数据准备好        |
| **Debounce**    | 防抖         | 搜索框输入停止后搜索  |
| **Throttle**    | 节流         | 滚动事件限制频率      |
| **AfterLast**   | 空闲后执行   | 无操作3秒后自动保存   |
| **Retry**       | 重试机制     | 网络请求失败自动重试  |
| **Schedule**    | 多时间点     | 每天9/12/18点检查邮件 |
| **DelayAfter**  | 事件后延迟   | 文件修改后5秒备份     |

### 代码示例

```vb
' 每天14:30执行
delay.ClockAt("14:30:00").RepeatDaily().Callback(Me, "DailyReport")

' 每5秒检查一次，最多10次
delay.Interval(5000).RepeatCount(10).Callback(Me, "CheckStatus")

' 防抖：输入停止500ms后才搜索
delay.Debounce(500).Callback(Me, "DoSearch")

' 节流：最多每2秒记录一次日志
delay.Throttle(2000).Callback(Me, "SaveLog")

' 空闲3秒后自动保存
delay.AfterLastInput(3000).Callback(Me, "AutoSave")

' 失败后重试3次，间隔递增(1s, 2s, 4s)
delay.Retry(3, 1000, 2).Callback(Me, "ConnectServer")
```

### 这些功能让 cDelay 从一个简单的延时工具变成强大的调度系统！

路线图：[./TODO.md](./TODO.md)

_最后更新: 2026-03-31_

