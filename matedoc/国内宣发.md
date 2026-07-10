# 用 VB6 写 Web 后端？没错，这事真有人干成了

你有多久没听到有人聊 VB6 了？

我猜大部分人的反应是："那玩意不是早进博物馆了吗？" 确实，在 Python 和 Node.js 横行的年代，VB6 听起来像是一个上古词汇。但现实是——国内还有大量的工控系统、企业内部工具、老旧业务系统跑在 VB6 上，这些代码还在干活，也没有人敢动它们。

问题来了：当你用 VB6 维护的旧系统需要对接互联网，需要开个接口给前端调，需要接个硬件设备的实时数据推送……你怎么搞？以前的做法是写个中间层，用 C# 或者 Python 转一道，又多了一套东西要维护。

今天给大家介绍的东西，就是为了解决这个事儿。

---

## VBMAN：让 VB6 干它以前干不了的事

VBMAN 是一个专门给 VB6/VBA 用的网络开发框架，它把一堆现代网络能力打包成了 VB6 能直接用的对象。不用装 Python，不用装 Node，不需要中间层，VB6 自己就能搞定。

这项目从 2017 年的一个念头和一张架构图开始，2020年动手编写，2023年重构，今年 6 月刚开源，GPL v3 协议，二进制永久免费。

先别急着划走，看看它能干啥：

---

## 重头戏：HTTP 服务器，VB6 直接当后端

这是我觉得最炸裂的功能。以前想用 VB6 跑个 HTTP 服务？做梦。现在几行代码就起来了：

```vb
Dim Server As New cHttpServer

Private Sub Form_Load()

    ' 把当前窗体注册为控制器
    Call Server.Router.Reg("Home", Me)

    ' 配路由——跟那些现代框架一个路子
    Call Server.Router.Add("/", "Home@Index")
    Call Server.Router.Add("/api/user/{id}", "Home@GetUser", OnlyGet)
    
    ' 如果需要提供静态服务器，只需要指定静态目录
    call Server.WebRoot("C:\WebRoot")
    
    ' 启动！
    If Server.Start(8080) Then
        Debug.Print "跑起来了: http://localhost:8080"
    Else
        Msgbox Server.LastError
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Server.StopMe
End Sub
```

控制器写法也很自然，直接在窗体里写方法就行：
（实际业务接口多的时候可以独立类文件，路由注册：`Server.Router.Reg "User", New bUser`）

```vb
' 首页
Public Sub Index(ctx As cHttpServerContext)
    ctx.Response.Html "<h1>Hello World!</h1>"
End Sub

' 带参数的 RESTful 路由
Public Sub GetUser(ctx As cHttpServerContext)
    Dim id As String
    id = ctx.Request.RouteParams("id")

    With New cJson
        .item("userId") = id
        .item("name") = "张三"
        ctx.Response.Json .Root
    End With
End Sub
```

看到没？路由、参数解析、JSON 响应，该有的都有。`ctx` 上下文对象里塞了 Request、Response、Session、Cookie、数据库引用，一套拿到手，写法跟 Express.js 的 `req/res` 思路差不多，只不过换成了 VB6 的语法。

当然，正式项目建议用独立的类模块当控制器，代码更干净：

```vb
' cUserController.cls
Public Sub List(ctx As cHttpServerContext)
    With New cJson
        .item("team") = Array("张三", "李四", "王五")
        .item("total") = 3
        ctx.Response.Json .Root, 0, "Success"
    End With
End Sub
```

而且这个 HTTP 服务器不是玩具，该有的东西一样不少：

- **路由系统**——手动注册、参数路由 `{id}`、自动路由、HTTP 方法限定、路由组和中间件
- **Session 管理**——内存、文件、数据库三种存储，随你选
- **静态文件服务**——设个 WebRoot 目录就能托管网页，自带 ETag 缓存和 304 协商，该省的流量都省了
- **SSE 实时推送**——服务端主动往浏览器推数据，做实时大屏、消息通知轻轻松松
- **CORS 跨域**——前后端分离开发时不用头疼跨域了
- **HTTPS/TLS**——链式调用配证书，PFX、PEM、Windows 证书存储三种方式都支持
- **连接管理**——最大连接数、空闲超时、请求体大小限制，内置定时器自动清理
- **MVC 模式**——路由-控制器-视图，项目大了也好管

甚至还有个性能统计对象 `Statistics`，请求量、状态码分布、QPS、峰值连接数全都能查，做个运维监控面板都有数据源了。

---

## 不止 HTTP 服务器，全家桶了解一下

VBMAN 不是一个只做 HTTP 的库，它是一整套工具集。下面把所有对象给大家列一遍，按用途分个类，一看就明白：

### 网络通信

| 对象 | 说明 |
|------|------|
| **cHttpServer** | HTTP 服务器，上面重点讲了这个 |
| **cHttpClient** | HTTP 客户端，发请求、调 API、带 Cookie 管理、SSL 都行 |
| **cWebSocketServer** | WebSocket 服务端，双向实时通信，符合 RFC 6455 |
| **cWebSocketClient** | WebSocket 客户端 |
| **cWinsock** | 底层网络通信，TCP/UDP 双协议，自带 TLS 和心跳 |
| **cSSEClient** | SSE 客户端，自动重连、续传 |
| **cRedisClient** | Redis 客户端，String/Hash/List/Set/SortedSet 全覆盖，支持管道和事务 |

### 工业通信

| 对象 | 说明 |
|------|------|
| **cModbusMaster** | Modbus 主站，完整功能码支持，TCP/RTU 都行 |
| **cModbusSlave** | Modbus 从站 |
| **cSerialPort** | 串口通信，纯 Win32 API 实现，零 OCX 依赖，COM10+ 也能用 |

### 数据处理

| 对象 | 说明 |
|------|------|
| **cJson** | JSON 解析和构建，链式调用，中文不会变 \uXXXX |
| **cDataBase** | 数据库操作，ADO 封装，连接池、链式 CRUD、分页、事务全有 |
| **cCollection** | 增强版集合，比 VB6 自带的好用太多 |

### 加密和安全

| 对象 | 说明 |
|------|------|
| **cAes / cAesCBC** | AES 对称加密 |
| **cCryptoHash** | 哈希计算，MD5 到 SHA-512 |
| **cCryptoHMAC** | HMAC 消息认证码 |
| **cPassword** | 密码处理工具 |
| **cTlsReMaster** | TLS 证书统一配置，HttpServer/WebSocket/Winsock 共享 |

### 工具箱

| 对象 | 说明 |
|------|------|
| **cToolsStr** | 字符串处理 |
| **cToolsArray** | 数组操作 |
| **cToolsDic** | 字典操作 |
| **cToolsList** | 列表操作 |
| **cToolsMath** | 数学运算 |
| **cToolsDateTime** / **cTimeUse** / **cTimer** | 时间日期、计时器 |
| **cToolsBase64** | Base64 编解码 |
| **cToolsUtf8** | UTF-8 编解码 |
| **cToolsCrc** | CRC 校验 |
| **cToolsFso** / **cFileEx** / **cToolsStream** | 文件系统操作 |
| **cToolsHttp** | HTTP 工具函数 |
| **cToolsSystem** | 系统信息 |
| **cToolsWindow** | 窗口操作 |
| **cImage** | 图片格式互转 |
| **cFormater** | 格式化工具 |

### 界面和交互

| 对象 | 说明 |
|------|------|
| **cToast** | 消息提示弹窗，9 种位置、4 种主题，堆叠显示 |
| **cDialog** | 系统对话框，打开/保存文件、选文件夹 |
| **cQRcode** | 二维码生成，纯 VB 实现 |

### 运维和扩展

| 对象 | 说明 |
|------|------|
| **cLogs** | 四通道日志——文件、窗体、DbgView、网页远程调试 |
| **cStartUp** | 开机自启管理 |
| **cStdIO** / **cCmd** | 命令行执行，捕获输出 |
| **cPLI** | 调用外部程序（Python/Node），拿返回值 |
| **cDelay** | 延时执行，同步模式用消息泵不卡 UI |
| **cIni** / **cCsv** | INI 和 CSV 文件读写 |
| **cRegedit** | 注册表操作 |
| **cAI** | 通用 AI 客户端，兼容 OpenAI 协议，豆包/DeepSeek/通义都能接 |
| **cBaidu** | 百度服务接口 |

### 还有

| 对象 | 说明 |
|------|------|
| **cZipArchive** | ZIP 压缩解压，纯 VB 实现 |
| **cSSEServer** | SSE 服务端推送（集成在 HttpServer 中，也可独立用） |

有没有发现？加起来快 50 个对象了。从网络到数据库到加密到串口到 AI，日常开发用到的基本上都给你备齐了。而且这些对象的调用方式都是 VB6 原生风格，不会让你觉得突兀。

---

## VBMAN2：升级版来了，还能内嵌浏览器

如果 VBMAN 已经让你觉得有点意思了，那还有一个好消息——它的升级版 **VBMAN2** 已经在路上。

VBMAN2 基于 TwinBASIC 开发，编译产物是 DLL，同时支持 32 位和 64 位。它是 VBMAN 的升级版本，不是另起炉灶，而是在 VBMAN 的基础上继续进化。

### 目前已完成：WebView2 对象

VBMAN2 第一个拿出来的是 `cWebView2Host`，一句话就能在你的窗体里嵌入一个完整的 Chromium 浏览器控件：

```vb
Dim wv As New cWebView2Host

Private Sub Form_Load()
    wv.Initialize Me.hWnd, "https://vb6.pro"
End Sub
```

对，就这两行代码，窗体中就跑起来了一个 Edge 内核的浏览器。不用注册 OCX，不用拖控件（当然你也可以初始化到任何具有 hWnd 句柄的控件上，比如 PictureBox，Frame...），不用配置。

而且 VB 和 JS 之间的交互能力非常完整：

```vb
' VB 同步调用 JS，直接拿返回值
Dim title As String
title = wv.JsRun("document.title") ' 实际提供了直接的属性 wv.DocumentTitle

' 双向数据绑定——VB6 的变量和网页 UI 自动联动
wv.BindData "name", "#name-input", "value"    ' VB → UI
wv.BindUI Me, "OnInput", "#name-input", EventName:="input"  ' UI → VB

Public Sub OnInput(ByVal EventName As String, ByVal Detail As String)
    wv.SetData "name", GetValue(Detail, "value")
End Sub
```

还有 40+ 事件（导航、脚本、鼠标、键盘、下载、PDF 打印）、CDP 全能力调用、Cookie 管理、本地资源映射、多宿主适配（VB6/Excel/Access 全支持）等，功能很扎实。

### 后续计划

VBMAN2 会逐步把 VBMAN 的所有对象迁移过来，而且多数对象会以更强劲的性能和更丰富的功能重新实现。远期规划包括：

- 基于 IOCP 的高性能网络库和 HTTP 服务器
- 真正可调试的多线程池
- 更多数据库驱动的集成
- 更强的 AI 对象

VBMAN2 没有开源计划，但和 VBMAN 一样，**二进制 DLL 永久免费使用**。

欢迎到文档站 https://doc.vb6.pro 下载 VBMAN2 的示例源码，跑起来就知道了。

---

## 实际能拿来干嘛？

说点接地气的场景：

**1. 给老系统加 API 接口**

你那个跑了十年的 VB6 业务系统，老板突然说要做个手机端查看数据。不用迁移，不用中间层，VB6 里启动个 HttpServer，写几个 API 路由，前端直接调。

**2. 工控设备数据采集和展示**

PLC、传感器通过 Modbus 采集数据，VBMAN 的 cModbusMaster 读回来，通过 HttpServer + SSE 实时推到网页上，一个车间大屏就出来了。

**3. 做个简单的后台管理**

写几个接口，前端用 Vue/React 也行，用纯 HTML 也行，VB6 当后端完全够用。Session、Cookie、数据库全有，一个人就能搞定前后端。

**4. IoT 网关**

串口读设备数据，MQTT 转发，HTTP 暴露接口，一个 VB6 程序全搞定。

**5. 内嵌浏览器做混合应用**

用 VBMAN2 的 WebView2 对象，VB6 窗体里嵌入网页界面，VB 处理底层逻辑，HTML/CSS/JS 做展示层，两边的交互还能双向绑定。

---

## 是真开源，不是那种"开源"

有些项目说开源，结果核心功能要收费，或者社区版各种阉割。VBMAN 不是这种套路：

- **二进制 DLL 永久免费**，拿来用就行，没功能限制
- **源代码开放**，GPL v3 协议，个人随便用，改了要开源
- 商业闭源需要授权，但那是真把源码拿去改完闭源卖的情况，正常用根本不涉及

---

## 上手试试？

如果你手头还有 VB6 开发环境，真的建议试一下。最简单的体验方式：

```vb
Dim Server As New cHttpServer

Private Sub Form_Load()
    With Server
    ' 把窗体本身注册为控制器
        .Router.Reg("Home", Me)
        .Router.Add("/", "Home@Hello")
        .Start 8080
    End With
End Sub

Public Sub Hello(ctx As cHttpServerContext)
    ctx.Response.Html "<h1>你好，VB6 也能写 Web！</h1>"
End Sub
```

运行后浏览器打开 `http://localhost:8080`，看到页面的那一刻，你大概率会说："卧槽，还真能跑。"

---

## 致谢

**首先致谢本公众号站长的发表，才使得VBMAN和大家在此见面，真心感谢**

VBMAN 的底层依赖了几个优秀的开源项目，向这些作者致敬：

| 项目 | 用途 |
|------|------|
| [VbAsyncSocket](https://github.com/wqweto/VbAsyncSocket) | 所有 Socket 通信对象的基础 |
| [ZipArchive](https://github.com/wqweto/ZipArchive) | cZipArchive 压缩解压基于它 |
| [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) | cJson 对象的后端解析引擎 |
| [HttpMimeType](mailto://jason@bitspaces.com) | HTTP MIME 类型识别 |
| [cTimer](http://sandsprite.com) | cTimer 高精度定时器 |

VBMAN2 同样站在巨人的肩膀上：

| 项目 | 用途 |
|------|------|
| [twinBASIC WebView2Package](https://docs.twinbasic.com/WebView2) | cWebView2Host 的核心 WebView2 基础能力 |

特别感谢群友 **周杰**，提供了 Web 页面任意元素绑定到 VB6/VBA 函数的建议，这是 BindUI 功能的灵感来源。

---

## 相关链接

- 国内仓库：https://gitcode.com/woeoio/vbman
- 开发文档：https://doc.vb6.pro
- QQ 交流群：915520648
- 作者网站：https://a-vi.com

---

*VB6 可能老了，但用 VB6 的人还在。VBMAN 就是给这些还在用 VB6 干活的人准备的。*