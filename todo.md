### 2026-03-20

- httpclient对象解析cookies会丢失同名键值对。

### 2026-01-30

- 增加 m_oServer.ClearCache 函数
- 协议常量改为和 Winsock.ocx 一致性

### 2026-01-26

- 给VBMAN增加标准dll函数，开发者可以免注册调用一些功能，比如完成注册，下载webview2运行时等
- 给VBMAN2增加免注册清单文件，配置“兼容性”为“xp”。可以解决家庭版和教育版无法注册ocx问题

### 2026-01-21

- 给 httpClient 增加批量请求模式，支持同步获取（竞赛或者全部）和异步获取
- 增加 vbman.codegen 模块，用于生成代码（从json转对象，从curl转httpclient）
- csv对象写入换行符单元格的时候没处理引号。需要修复
- csv对象的showto函数需要支持列宽记忆，
- csv对象的行列号支持开关配置
- csv对象内置一个开源表格
- csv对象支持rs对象，可任意筛选
- 增加纤程对象，配合单元线程
- 增加线程池对象
- 增加异步模块
- 增加子类化对象
- 基于子类化对象增加一个滚动条控制类，用于给控件添加滚动条滚轮控制（比如dataRgid）
- cHttpClient 的错误处理需要改进，对于非 2xx 状态的返回不能处理body，现在返回429会抛出类型不匹配。
- toolsArray 增加数组操作类，包括判空，等
- cWinsock 增加大文件分块传输

### 2026-01-20

今天看到一个牛逼的函数，来自维托的：

```vb
Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must--you'll probably still crash,
   ' but it will be because of the subclass, not the uncounted reference
End Property
```

### 2026-01-19

- 给所有字典对象改为大小写不敏感

### 2026-01-18

- 给cWinsock添加一个函数，用于绑定业务 userID，方便针对用户发送，模仿 webman
- 给cHttpServer增加 view 函数，用于渲染html

### 2025-06-22

- CGI模块

### 2025-01-03

- 系统服务模块，支持一键安装为服务，可获取运行状态， 可控制服务
- sqlite模块
- 动态路由参数支持
- webdav客户端
- ctx.request 默认成员
- ctx.request 未处理 utf8
