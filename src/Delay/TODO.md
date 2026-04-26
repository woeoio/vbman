# cDelay 延时对象 - TODO 与路线图

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

## 当前版本: v1.0

### 已实现功能

- [x] 三种工作模式：事件、回调、同步
- [x] ParamArray 变长参数支持（最多9个）
- [x] 同步模式消息泵实现（无 DoEvents）
- [x] 同步模式取消功能
- [x] 自动生命周期管理（引用计数）
- [x] 链式调用 API 设计

---

## 路线图

### v1.1 增强版多实例管理

参考 Toast 实现，实现更强大的多实例管理：

#### 1.1.1 实例命名与追踪

```vb
' 为延迟实例设置名称，便于后续管理
Dim delay As New cDelay
delay.Tag("AutoSaveTimer").Callback(Me, "AutoSave").CountDown(5000)

' 通过名称获取实例状态
If Delays.Exists("AutoSaveTimer") Then
    Debug.Print "剩余时间: " & Delays("AutoSaveTimer").RemainingTime
End If
```

**实现要点：**

- 添加 `Tag` 属性/方法，支持链式调用
- 全局集合使用 Key（如 "id:" & TimerID 或自定义 Tag）
- 提供 `Delays.Exists(Tag)` 方法检查实例是否存在
- 提供 `Delays(Tag)` 索引器获取实例引用

#### 1.1.2 按类型分组管理

参考 Toast 的 9 个位置集合，按用途分组：

```vb
' 枚举定义
Public Enum DelayGroup
    dgSystem = 1      ' 系统级延时
    dgUser = 2        ' 用户操作相关
    dgBackground = 3  ' 后台任务
    dgAnimation = 4   ' 动画相关
End Enum

' 使用
delay.Group(dgBackground).Callback(Me, "SyncData").CountDown(60000)

' 批量操作
delay.Group(dgBackground).CancelAll   ' 取消所有后台任务
```

**实现要点：**

- 每个 Group 对应一个独立的 Collection
- 支持分组级别的批量操作（CancelAll、PauseAll、ResumeAll）

#### 1.1.3 实例状态查询

```vb
' 获取当前活动的延迟数量
Dim count As Long
count = Delays.Count                    ' 总数
count = Delays.CountByGroup(dgSystem)   ' 按分组

' 枚举所有活动实例
Dim d As cDelay
For Each d In Delays
    Debug.Print d.Tag & " - 剩余: " & d.RemainingTime
Next
```

**新增属性/方法：**

- `RemainingTime` - 剩余毫秒数（计算属性）
- `ElapsedTime` - 已过去毫秒数
- `Progress` - 完成百分比（0-100）

---

### v1.2 新增延迟方式（与 CountDown 同级）

#### 1.2.1 ClockAt - 指定绝对时间

```vb
' 今天14:30执行
delay.ClockAt "14:30:00".Callback(Me, "DoSomething")

' 指定日期时间
delay.ClockAtDate "2026-04-01 09:00:00".Callback(Me, "MeetingReminder")

' 明天早上8点
delay.ClockAtDate DateAdd("d", 1, Date) & " 08:00:00"
```

**实现要点：**

- 计算目标时间与当前时间的差值
- 如果目标时间已过，可选择立即执行或次日执行
- 支持重复：每天/每周/每月的同时间执行

#### 1.2.2 Interval - 重复间隔

```vb
' 每5秒执行一次
delay.Interval(5000).Callback(Me, "CheckStatus")

' 限制次数
delay.Interval(1000).RepeatCount(10).Callback(Me, "UpdateProgress")

' 停止
delay.StopInterval
```

**实现要点：**

- 内部使用 Windows 定时器或循环
- 支持无限重复或指定次数
- 可以动态修改间隔时间

#### 1.2.3 Random - 随机延迟

```vb
' 在1-5秒之间随机
delay.Random(1000, 5000).Callback(Me, "RandomTask")

' 用于防检测、模拟人工操作
```

#### 1.2.4 WaitUntil - 条件等待

```vb
' 等待某个条件为真（最多等10秒）
delay.WaitUntil(AddressOf IsDataReady, 10000).Callback(Me, "ProcessData")

' 配合函数
Public Function IsDataReady() As Boolean
    IsDataReady = (FileLen("data.txt") > 0)
End Function
```

**实现要点：**

- 在消息循环中轮询检查条件
- 支持超时
- 可设置检查间隔（默认100ms）

#### 1.2.5 Debounce - 防抖

```vb
' 输入停止500ms后才执行搜索
delay.Debounce(500).Callback(Me, "DoSearch")

' 使用场景：搜索框、窗口调整大小
' 频繁触发时，只有最后一次会真正执行
```

#### 1.2.6 Throttle - 节流

```vb
' 最多每2秒执行一次
delay.Throttle(2000).Callback(Me, "SaveLog")

' 使用场景：日志记录、滚动事件
' 频繁触发时，按固定频率执行
```

#### 1.2.7 AfterLast - 空闲后执行

```vb
' 鼠标键盘无操作3秒后执行
delay.AfterLastInput(3000).Callback(Me, "AutoSave")

' 配合特定事件
delay.AfterLast("MouseMove", 1000).Callback(Me, "ShowTooltip")
```

**实现要点：**

- 需要监听系统消息或轮询 GetLastInputInfo
- 重置计时器当有新输入时

#### 1.2.8 Retry - 重试机制

```vb
' 失败后自动重试，最多3次，间隔递增
delay.Retry(3, 1000, 2).Callback(Me, "ConnectServer")
' 第1次失败后等1秒，第2次等2秒，第3次等4秒

' 配合回调返回值判断是否成功
Public Function ConnectServer() As Boolean
    ' 返回 False 会自动触发重试
    ConnectServer = AttemptConnection()
End Function
```

#### 1.2.9 Schedule - 多时间点调度

```vb
' 一天内的多个时间点
delay.Schedule(Array("09:00", "12:00", "18:00")).Callback(Me, "CheckMail")

' 工作日调度
delay.ScheduleWeekday("09:00,12:00,18:00").Callback(Me, "CheckMail")
```

#### 1.2.10 DelayAfter - 事件后延迟

```vb
' 文件修改后5秒自动备份
delay.AfterEvent("FileChanged", "C:\data.txt", 5000).Callback(Me, "BackupFile")

' 窗口失焦后自动保存
delay.AfterEvent("LostFocus", Me.hWnd, 1000).Callback(Me, "AutoSave")
```

---

### v1.3 高级控制功能

#### 1.3.1 暂停与恢复

```vb
delay.CountDown(10000)      ' 10秒延时

' 暂停
delay.Pause                 ' 暂停计时
Debug.Print delay.RemainingTime  ' 还剩多少

' 恢复
delay.Resume                ' 从暂停处继续
```

**技术方案：**

- 暂停时记录 `PausedTick` 和 `RemainingAtPause`
- 恢复时重新计算 `EndTick = GetTickCount() + RemainingAtPause`
- 仅适用于 Windows 定时器模式（事件/回调）

#### 1.3.2 重置与修改

```vb
' 重置当前延时（从头开始）
delay.Reset

' 修改剩余时间
delay.ChangeDelay 5000      ' 修改为还剩5秒

' 增加时间
delay.AddTime 2000          ' 增加2秒
```

---

### v1.4 性能与调试增强

#### 1.4.1 调试模式

```vb
' 全局开启调试
Delays.DebugMode = True

' 输出日志
' [cDelay] AutoSaveTimer: CountDown started, 5000ms
' [cDelay] AutoSaveTimer: Tick, remaining 3000ms
' [cDelay] AutoSaveTimer: OnTime fired
' [cDelay] AutoSaveTimer: Cancelled by user
```

#### 1.4.2 性能统计

```vb
' 获取统计信息
Dim stats As DelayStats
stats = Delays.GetStats

Debug.Print "总创建数: " & stats.TotalCreated
Debug.Print "总完成数: " & stats.TotalCompleted
Debug.Print "总取消数: " & stats.TotalCancelled
Debug.Print "平均执行时间: " & stats.AvgExecutionTime
```

#### 1.4.3 内存监控

```vb
' 自动清理僵尸实例（可选）
Delays.AutoCleanup = True
Delays.AutoCleanupInterval = 60000  ' 每60秒检查一次
```

---

### v1.5 同步模式增强

#### 1.5.1 超时处理

```vb
' 同步模式支持超时回调
delay.Sync().Timeout(5000).OnTimeout(Me, "OnSyncTimeout").CountDown(10000)
' 如果10秒内未完成，5秒后触发 OnSyncTimeout
```

#### 1.5.2 条件等待

```vb
' 等待直到条件满足或超时
delay.Sync().WaitUntil(AddressOf CheckCondition, 5000)
' 每帧检查 CheckCondition 函数，返回 True 则提前结束
```

#### 1.5.3 多任务等待

```vb
' 等待多个延时全部完成
Dim tasks As New Collection
tasks.Add delay1
tasks.Add delay2
tasks.Add delay3

Delay.WaitAll(tasks, 10000)  ' 最多等10秒

' 或等待任一完成
Delay.WaitAny(tasks, 10000)
```

---

### v1.6 回调增强

#### 1.6.1 Lambda/匿名函数支持（VB6 模拟）

```vb
' 使用 AddressOf 传递函数指针
delay.Callback(AddressOf MyStaticFunc).CountDown(1000)

' 或使用类实例方法
delay.Callback2(Me, "MethodName", param1, param2)
```

#### 1.6.2 异步回调

```vb
' 回调在独立线程执行（需谨慎）
delay.AsyncCallback(Me, "HeavyTask").CountDown(100)
```

#### 1.6.3 回调链

```vb
' 多个回调按顺序执行
delay.Callback(Me, "Step1") _
      .Then(Me, "Step2") _
      .Then(Me, "Step3") _
      .CountDown(1000)
' 1000ms 后依次执行 Step1 -> Step2 -> Step3
```

---

### v1.7 集成与扩展

#### 1.7.1 与 UI 控件集成

```vb
' 自动更新进度条
delay.BindProgressBar(ProgressBar1).CountDown(5000)

' 自动更新标签
delay.BindLabel(Label1, "剩余时间: {remaining}s").CountDown(5000)
```

#### 1.7.2 音效支持

```vb
delay.OnTimeSound("C:\Windows\Media\chimes.wav").CountDown(5000)
delay.OnCancelSound("C:\Windows\Media\ding.wav")
```

#### 1.7.3 配置文件支持

```vb
' 从配置加载
delay.LoadConfig("AutoSaveDelay")
' 读取注册表/INI：Delay.AutoSaveDelay = 5000
```

---

### v2.0 架构重构（远期）

#### 2.0.1 COM 组件化

- 编译为 ActiveX DLL
- 支持跨项目复用
- 提供强类型接口

#### 2.0.2 多线程支持

- 真正的后台线程定时器
- 不依赖 Windows 消息队列
- 适用于无窗口环境（如 Windows Service）

#### 2.0.3 .NET 互操作

- 提供 .NET 封装类
- 支持从 VB.NET/C# 调用

---

## 优先级建议

### 高优先级（近期实现）

1. **实例命名与追踪** - 基础功能，很多高级特性依赖于此
2. **ClockAt** - 指定时间执行，非常实用的功能
3. **Interval** - 重复间隔执行，类似定时器
4. **Debounce/Throttle** - 防抖节流，UI开发必备
5. **调试模式** - 便于开发和排查问题

### 中优先级（中期实现）

6. 按类型分组管理
7. 暂停/恢复功能
8. Random 随机延迟
9. WaitUntil 条件等待
10. 性能统计

### 低优先级（远期考虑）

11. Retry 重试机制
12. Schedule 多时间点调度
13. 多任务等待
14. 回调链
15. COM 组件化

---

## 设计原则

1. **向后兼容**：所有新增功能都是可选的，不影响现有 API
2. **链式调用**：保持 `object.Method().Another().Action()` 风格
3. **默认值合理**：不设置时提供智能默认行为
4. **错误处理**：所有操作都有明确的错误状态和返回值
5. **资源安全**：确保没有内存泄漏，清理彻底

---

## 贡献建议

如果你想参与开发，以下是一些切入点：

- **文档完善**：补充更多使用示例和最佳实践
- **测试用例**：编写边界情况测试（如极限时间值、并发场景）
- **性能优化**：优化消息循环，减少 CPU 占用
- **代码审查**：检查线程安全和资源释放

---

## 参考资源

- Toast 多实例管理：`D:\code\vi\vbmanlib\vbman\src\Toast\cToast.cls`
- Windows Timer API：MSDN `SetTimer` / `KillTimer`
- VB6 消息循环：`PeekMessage` / `DispatchMessage`
- 设计模式：观察者模式、工厂模式、对象池

---

_最后更新: 2026-03-31_

