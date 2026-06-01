# VBMAN - BASIC Network Application Development Framework

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-orange.svg)](https://www.microsoft.com/windows)

## Project Overview

VBMAN is a network application development framework built with BASIC language, designed to provide developers with concise and efficient server-side and client-side development tools.

## Project History
### Origin (2017)

This project originated from an idea in 2017 - to build a series of network application development frameworks using BASIC language. The project was initially named **BSMAN**, short for **Basic Server Man**.

BSMAN planned a series of sub-projects:

- **ASPMAN** - Active Server Pages framework
- **VBMAN** - Visual Basic component library (this project)
- **VBSMAN** - VBScript toolkit
- **VBAMAN** - VBA (Visual Basic for Applications) extension library
- **TBMAN** - TwinBasic version migration

### Development Timeline

| Date       | Milestone                                                                  |
| ---------- | -------------------------------------------------------------------------- |
| 2023-10-26 | First **ASPMAN** sub-project written                                       |
| 2024-09-30 | Officially started developing the **VBMAN** sub-project                    |
| 2025       | Began migrating VBMAN to **TwinBasic** platform, named **TBMAN**           |
| 2026-06-01 | **VBMAN officially open-sourced**                                          |

## Features

- **Rich Tool Library** - String processing, date/time, file operations, HTTP requests, and more
- **Database Support** - Simplified database connection and operation interfaces
- **JSON Processing** - Built-in JSON serialization/deserialization
- **HTTP Client** - Convenient HTTP request encapsulation
- **HTTP Server** - Lightweight HTTP server functionality
- **Logging System** - Hierarchical logging functionality
- **Registry Operations** - Simplified Windows registry read/write interface
- **More Features** - Continuously evolving...

## Quick Start

### Requirements

- Windows operating system
- Visual Basic 6.0 or TwinBasic development environment

### Installation

1. Clone or download this repository
2. Import class files from the `src` directory into your VB6/TwinBasic project
3. Reference `cVBMAN` as the entry point to start using

## License

This project is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

### Usage Rules

1. **Binary Distribution**: Compiled binary files are permanently free with no usage restrictions

2. **Source Code Usage**:
   - Individual users can use the source code for free (must include LICENSE file)
   - Any commercial or personal project that **redistributes** (including but not limited to distribution, sales, or service provision) must **open source** and adopt GPL-compatible licenses
   - For closed-source usage, commercial license must be purchased from the author

3. **GPL Compliance**:
   - Distribution must include source code or provide access to source code
   - Modified works must also be open-sourced under GPL license
   - Retain copyright notices and disclaimers
   - See [LICENSE](LICENSE) file for detailed terms

### Commercial License

If you wish to use this project's source code in closed-source commercial projects, please contact the author for commercial licensing:

- Author: Deng Wei (邓伟)
- Website: https://a-vi.com

## Acknowledgments

This project uses the following open-source projects:

| Project                                                          | License | Purpose                              |
| ---------------------------------------------------------------- | ------- | ------------------------------------ |
| [wqweto/VbAsyncSocket](https://github.com/wqweto/VbAsyncSocket) | MIT     | All `Socket` objects are based on it |
| [Tim Hall/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)      | MIT     | `cJson` object backend               |
| [Jason Peter Brown/HttpMimeType](mailto://jason@bitspaces.com)  | MIT     | Uses `HttpMimeType`                  |
| [David Zimmer/cTimer](http://sandsprite.com)                    | /       | `cTimer` object source               |

Full acknowledgment list see [docs/CREDITS.md](docs/CREDITS.md)

## Project Structure

```
vbman/
├── src/           # Source code directory
│   ├── StaticClass/   # Static class definitions
│   ├── Tools/         # Tool classes
│   ├── Json/          # JSON processing
│   ├── HttpClient/    # HTTP client
│   └── ...
├── docs/          # Documentation directory
│   ├── global/    # API documentation
│   └── CREDITS.md   # Acknowledgments
├── dist/          # Compiled output
└── test/          # Test code
```

## Related Projects

- **ASPMAN** - ASP framework sub-project
- **TBMAN** - TwinBasic version (source released, reference directly in tb package)
- **VBMAN2** - Next-generation version, high-performance framework based on TwinBasic

## Development Notes

### Unfinished Features

The following objects are not yet fully developed and will be completed in **VBMAN2**:

- `cAI` - AI object
- `modbus` - Modbus protocol support
- `mqtt` - MQTT protocol support

### About VBMAN2

VBMAN2 was originally a pure WebView2 control library based on TwinBasic's WebView2 control (thanks to TwinBasic wayen) for VB6/VBA, compiled as an OCX control file with various issues.

**Major Update on 2026.05.25**: Rewrote the TwinBasic WebView2 control as a DLL object, rendering to any native control with a handle (such as `Form1.hWnd`, `Picture1.hWnd`) to display web pages, solving the OCX control file issues and laying the best foundation for upgrading VBMAN. VBMAN2 will no longer be relatively independent from VBMAN, but will become the best upgraded version of VBMAN. (VBMAN2 has no open-source plan, but likewise, the binary DLL is permanently free to use.)

**Important**: VBMAN2 **includes all capabilities of VBMAN**, plus advanced features like WebView2. VBMAN users can seamlessly upgrade to VBMAN2.

#### VBMAN2 Core Features

##### Two-way Data Binding

VBMAN2 provides Vue-like two-way data binding, enabling seamless interaction between the VB6/VBA host and WebView2 UI:

| Direction | API | Description |
|-----------|-----|-------------|
| UI → VB6/VBA | `BindUI` / `UnbindUI` | DOM events trigger callbacks to host methods |
| VB6/VBA → UI | `BindData` / `SetData` | Host sets values → DOM properties auto-update |

**Core Design**: Explicit composition over implicit hijacking. Unlike Vue's automatic data hijacking via ES6 Proxy, VBMAN2 uses explicit `BindUI` + `BindData` composition, better suited for cross-process WebView2 scenarios, and is fully compatible with VB6/VBA:

```vb
' One-way binding: VB6/VBA → UI
wv.BindData "username", "#user-name", "textContent"
wv.SetData "username", "John"   ' UI updates automatically

' Two-way binding: input ↔ VB6/VBA
wv.BindData "search", "#search-input", "value"   ' data → UI
wv.BindUI Me, "OnSearch", "#search-input", EventName:="input"   ' UI event → VB6/VBA

Public Sub OnSearch(ByVal EventName As String, ByVal Detail As String)
    wv.SetData "search", JsonParser.GetValue(Detail, "value")   ' write back
End Sub
```

**Supported DOM Properties**: `textContent` / `innerHTML` / `value` / `checked` / `disabled` / `visible` / `className` / `src` / `href` / `style`, plus batch updates via `SetDataBatch` to reduce IPC calls.

#### VBMAN2 Future Roadmap

- High-performance IOCP network library
- IOCP-based HTTP server
- Truly debuggable multi-threading pool
- Truly usable AI object
- Collection of various database drivers
- ... more

## Contact

- Project Author: Deng Wei (邓伟)
- Personal Website: https://a-vi.com
- Documentation: https://doc.vb6.pro
- Repository: https://gitcode.com/woeoio/vbman

---

**Disclaimer**: This project is provided "as is" without any express or implied warranties. See LICENSE file for details.
