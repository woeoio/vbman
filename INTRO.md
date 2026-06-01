# VBMAN 简介

VBMAN 是一个使用 BASIC 语言构建的网络应用开发框架，为开发者提供简洁高效的服务器端和客户端开发工具集，支持 JSON 处理、HTTP 请求/服务器、数据库操作等功能，适用于 VB6 和 VBA 开发环境。

---

## 发布话术

### 简短版（适合微博、Twitter 等）

> 开源项目 VBMAN 正式发布！一个专为 VB6/VBA 打造的网络应用开发框架，提供 HTTP 服务器/客户端、JSON 处理、数据库操作、WebSocket 等现代开发能力。老牌 BASIC 语言也能玩现代网络编程！
> 
> 项目仓库：https://gitcode.com/woeoio/vbman
> 开发文档：https://doc.vb6.pro

---

### 完整版（适合知乎、掘金、CSDN 等）

> ## VBMAN 开源发布 - VB6/VBA 的现代网络开发框架
> 
> 还在为 VB6/VBA 缺乏现代网络库而苦恼？VBMAN 来了！
> 
> **VBMAN** 是一个专为 BASIC 语言设计的网络应用开发框架，让 VB6/VBA 开发者也能轻松构建现代网络应用：
> 
> - HTTP 服务器/客户端 - 快速搭建 Web 服务或调用 API
> - WebSocket / SSE - 实时通信能力
> - JSON 处理 - 现代数据交换格式支持
> - 数据库操作 - 简化 SQL 操作
> - Modbus/MQTT/Redis - 工业物联网协议支持
> - 影子窗口、序列化工具等实用组件
> 
> 项目基于 GPL-3.0 协议开源，二进制文件永久免费使用。
> 
> 项目仓库：https://gitcode.com/woeoio/vbman
> 开发文档：https://doc.vb6.pro
> 
> 欢迎各位 VB6/VBA 开发者试用、提 Issue、提 PR！
> 
> #VB6 #VBA #开源 #网络编程 #BASIC

---

### 技术社区版（适合 V2EX、掘金等）

> 给还在维护 VB6/VBA 老项目的同学们推荐一个实用的开源框架。
> 
> VBMAN 是我开发的一个网络应用框架，目的是让 VB6/VBA 能方便地做现代网络开发。主要功能包括 HTTP 服务器（可以当 Web 后端）、HTTP 客户端、WebSocket、SSE、JSON 处理、数据库操作、Modbus TCP/RTU、MQTT、Redis 等。
> 
> 框架设计遵循 VB6 的开发习惯，API 风格尽量贴近原生，降低学习成本。
> 
> 目前项目已开源在 GitCode，文档站点也已上线，有需要的可以看看：
> 
> - 源码：https://gitcode.com/woeoio/vbman
> - 文档：https://doc.vb6.pro
> 
> 欢迎 Star 和提 Issue。

---

## Reddit 英文发布版

### 标题建议（任选其一）

1. `VBMAN - A Modern Network Development Framework for VB6/VBA (Open Source)`
2. `Open Sourced VBMAN: Bringing HTTP/WebSocket/JSON to VB6/VBA`
3. `VBMAN - Finally, a proper network framework for VB6/VBA developers`
4. `After 8 years of development, I'm open sourcing my VB6/VBA network framework`

### 正文（Markdown 格式）

```markdown
Hey r/vba and r/visualbasic!

I'm excited to share that I've just open-sourced **VBMAN**, a network application development framework I've been working on since 2017. It's designed specifically for VB6 and VBA developers who need modern networking capabilities.

## What is VBMAN?

VBMAN is a comprehensive network development framework that brings modern web technologies to the BASIC ecosystem. If you've ever struggled with making HTTP requests, hosting a web server, or handling WebSockets in VB6/VBA, this might save you some headaches.

## Key Features

### Web Development
- **HTTP Server** - Build lightweight web services directly in VB6/VBA
- **HTTP Client** - Modern HTTP/HTTPS requests with JSON support
- **WebSocket** - Real-time bidirectional communication
- **Server-Sent Events (SSE)** - Server-to-client streaming
- **JSON Processing** - Native JSON serialization/deserialization

### Database & Storage
- **Database Access** - Simplified SQL operations
- **Redis Client** - Cache and message broker support
- **INI File Handler** - Structured configuration management

### Industrial IoT Protocols
- **Modbus TCP/RTU** - Master and Slave implementations
- **MQTT Client/Server** - Lightweight messaging for IoT
- **Serial Port** - Direct hardware communication

### Utilities
- **Shadow Window** - Modern UI effects for VB6 forms
- **Collection Tools** - Enhanced data structures
- **Cryptography** - AES, Hash, HMAC support
- **Logging System** - Structured application logging

## Why VBMAN?

I started this project in 2017 because I needed to build network applications for industrial automation, but VB6's built-in networking capabilities were... limited. Rather than migrate everything to .NET, I decided to extend VB6's capabilities.

The framework is designed to feel native to VB6 developers - the API style follows familiar patterns, so you don't need to learn a completely new way of thinking.

## VBMAN2 - The Next Generation

I'm also working on **VBMAN2**, which adds WebView2 support with two-way data binding. This means you can build modern web-based UIs that communicate seamlessly with your VB6/VBA code:

```vb
' Bind UI elements to VB6 code
wv.BindData "username", "#user-name", "textContent"
wv.SetData "username", "John"  ' UI updates automatically

' Two-way binding
wv.BindUI Me, "OnSearch", "#search-input", EventName:="input"
```

## Project Info

- **License**: GPL-3.0 (Binary files are free forever, source available for personal use)
- **Documentation**: https://doc.vb6.pro
- **Repository**: https://gitcode.com/woeoio/vbman

## Who is this for?

- Maintaining legacy VB6/VBA applications
- Industrial automation systems
- Quick prototyping for internal tools
- Anyone who still enjoys BASIC syntax (no judgment here!)

## Questions?

Happy to answer any questions! The documentation site has detailed API references and examples. I've been using this in production environments for years, so it's battle-tested for industrial scenarios.

---

**Fun fact**: This project started as "BSMAN" (Basic Server Man) back in 2017, with grand plans for ASPMAN, VBSMAN, and VBAMAN sub-projects. Only VBMAN survived... and thrived!
```

---

## 链接汇总

| 用途 | 地址 |
|------|------|
| 项目仓库 | https://gitcode.com/woeoio/vbman |
| 开发文档 | https://doc.vb6.pro |
| 作者网站 | https://a-vi.com |

