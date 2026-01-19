### 2025-01-19

- ? 实现了 SSE (Server-Sent Events) 客户端 cSSEClient.cls
- ? 支持自动重连、断点续传、事件类型解析
- ? 添加 SSE 客户端使用文档

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
