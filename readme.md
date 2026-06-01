# VBMAN - BASIC 网络应用开发框架

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-orange.svg)](https://www.microsoft.com/windows)

## 项目简介

VBMAN 是一个使用 BASIC 语言构建的网络应用开发框架，旨在为开发者提供简洁、高效的服务器端和客户端开发工具集。

## 项目历史

### 起源 (2017)

本项目起源于 2017 年的一个念头——希望使用 BASIC 语言构建一系列网络应用开发框架。当时项目起名为 **BSMAN**，是 **Basic Server Man** 的缩写。

BSMAN 旗下规划了一系列子项目：

- **ASPMAN** - Active Server Pages 框架
- **VBMAN** - Visual Basic 组件库 (本项目)
- **VBSMAN** - VBScript 工具集
- **VBAMAN** - VBA (Visual Basic for Applications) 扩展库
- **TBMAN** - TwinBasic 版本迁移

### 发展历程

| 时间       | 里程碑                                                                       |
| ---------- | ---------------------------------------------------------------------------- |
| 2023-10-26 | 编写了第一个 **ASPMAN** 子项目                                               |
| 2024-09-30 | 以此为起点，正式开发 **VBMAN** 子项目                                        |
| 2025       | 开始将 VBMAN 迁移到 **TwinBasic** 平台，命名为 **TBMAN**，源码发布但进度较慢 |
| 2026-06-01 | **VBMAN 正式开源**                                                           |

## 功能特性

- **丰富的工具类库** - 提供字符串处理、日期时间、文件操作、HTTP 请求等常用工具
- **数据库支持** - 简化的数据库连接和操作接口
- **JSON 处理** - 内置 JSON 序列化/反序列化功能
- **HTTP 客户端** - 便捷的 HTTP 请求封装
- **HTTP 服务器** - 轻量级 HTTP 服务器功能
- **日志系统** - 分级日志记录功能
- **注册表操作** - 简化的 Windows 注册表读写接口
- **更多功能** - 请参考开发文档...

## 快速开始

### 环境要求

- Windows 操作系统
- Visual Basic 6.0 或 TwinBasic 开发环境

### 安装

1. 克隆或下载本仓库
2. 将 `src` 目录下的类文件导入到你的 VB6/TwinBasic 项目中
3. 引用 `cVBMAN` 作为入口点即可开始使用

## 许可证说明

本项目采用 **GNU General Public License v3.0 (GPL-3.0)** 开源协议。

### 使用规则

1. **二进制分发**：本项目编译后的二进制文件永久免费，无任何使用限制

2. **源代码使用**：
   - 个人用户可以免费使用源代码（需附带 LICENSE 文件）
   - 任何商业或个人项目如果**再次发行**（包括但不限于分发、销售、作为服务提供），必须**开源**并采用 GPL 兼容协议
   - 如需闭源使用，必须联系作者购买商业授权

3. **GPL 协议遵守要点**：
   - 分发时必须包含源代码或提供获取源代码的方式
   - 修改后的作品也必须采用 GPL 协议开源
   - 保留版权声明和免责声明
   - 详细条款请参见 [LICENSE](LICENSE) 文件

### 商业授权

如果您希望在闭源商业项目中使用本项目的源代码，请联系作者获取商业授权：

- 作者：邓伟
- 网站：https://a-vi.com

## 致谢

本项目使用了以下开源项目：

| 项目名称                                                        | 许可证 | 用途                     |
| --------------------------------------------------------------- | ------ | ------------------------ |
| [wqweto/VbAsyncSocket](https://github.com/wqweto/VbAsyncSocket) | MIT    | 所有 `Socket` 对象基于它 |
| [Tim Hall/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)      | MIT    | `cJson` 对象后端使用它   |
| [Jason Peter Brown/HttpMimeType](mailto://jason@bitspaces.com)  | MIT    | 使用了 `HttpMimeType`    |
| [David Zimmer/cTimer](http://sandsprite.com)                    | /      | `cTimer` 对象来源于它    |

完整致谢列表请参见 [docs/CREDITS.md](docs/CREDITS.md)

## 项目结构

```
vbman/
├── src/           # 源代码目录
│   ├── StaticClass/   # 静态类定义
│   ├── Tools/         # 工具类
│   ├── Json/          # JSON 处理
│   ├── HttpClient/    # HTTP 客户端
│   └── ...
├── docs/          # 文档目录
│   ├── global/    # API 文档
│   └── CREDITS.md   # 致谢列表
├── dist/          # 编译输出
└── test/          # 测试代码
```

## 相关项目

- **ASPMAN** - ASP 框架子项目
- **TBMAN** - TwinBasic 版本 (源码发布，直接在tb包里引用)
- **VBMAN2** - 下一代版本，基于 TwinBasic 的高性能框架

## 开发说明

### 未完成的功能

以下对象尚未完成开发，将在 **VBMAN2** 中继续完成：

- `cAI` - AI 对象
- `modbus` - Modbus 协议支持
- `mqtt` - MQTT 协议支持

### 关于 VBMAN2

VBMAN2 早期是基于 TwinBasic 的 WebView2 控件（感谢 TwinBasic wayen）封装给 VB6/VBA 使用的纯 WebView2 控件库，相对于 VBMAN 是个独立库，编译产物是 OCX 控件文件，存在不少问题。

**2026.05.25 重大更新**：将 TwinBasic WebView2 控件重写为 DLL 对象，通过渲染到任何有句柄的原生控件（如 `Form1.hWnd`、`Picture1.hWnd`）即可显示网页，解决了 OCX 控件文件的烦恼，并且为升级 VBMAN 打下了最好的基础。vbman2 不再和 vbman 相对独立，而是会成为最好的 vbman 升级版。（vbman2没有开源计划，但同样的，永久免费使用二进制dll）

#### VBMAN2 核心特性

##### 双向数据绑定

VBMAN2 提供类似 Vue 的双向数据绑定能力，实现 VB6/VBA 宿主与 WebView2 网页 UI 的无缝联动：

| 方向 | API | 说明 |
|------|-----|------|
| UI → VB6/VBA | `BindUI` / `UnbindUI` | DOM 事件触发 → 回调宿主方法 |
| VB6/VBA → UI | `BindData` / `SetData` | 宿主设值 → 自动更新 DOM 属性 |

**核心设计**：显式组合而非隐式劫持。不同于 Vue 依赖 ES6 Proxy 的自动数据劫持，VBMAN2 采用 `BindUI` + `BindData` 显式组合，更适合跨进程 WebView2 场景，且完全兼容 VB6/VBA：

```vb
' 单向绑定：VB6/VBA → UI
wv.BindData "username", "#user-name", "textContent"
wv.SetData "username", "张三"   ' UI 自动更新

' 双向绑定：输入框 ↔ VB6/VBA
wv.BindData "search", "#search-input", "value"   ' 数据 → UI
wv.BindUI Me, "OnSearch", "#search-input", EventName:="input"   ' UI 事件 → VB6/VBA

Public Sub OnSearch(ByVal EventName As String, ByVal Detail As String)
    wv.SetData "search", JsonParser.GetValue(Detail, "value")   ' 回写数据
End Sub
```

**支持的 DOM 属性**：`textContent` / `innerHTML` / `value` / `checked` / `disabled` / `visible` / `className` / `src` / `href` / `style` 等，同时支持批量更新 `SetDataBatch` 减少 IPC 调用。

#### VBMAN2 未来规划

- 高性能的 IOCP 网络库
- 基于 IOCP 的 HTTP 服务器
- 真正可调试的多线程池
- 真正可用的 AI 对象
- 集合各种数据库驱动
- ... more

## 联系方式

- 项目作者：邓伟
- 个人网站：https://a-vi.com
- 开发文档：https://doc.vb6.pro
- 项目仓库：https://gitcode.com/woeoio/vbman

---

**声明**：本项目按"原样"提供，不提供任何明示或暗示的担保。详见 LICENSE 文件。
