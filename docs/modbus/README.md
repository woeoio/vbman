# Modbus 文档目录

本目录包含 Modbus 类库的完整文档。

## 文档列表

### 1. [总览文档 (overview.md)](./overview.md)
Modbus 类库的整体介绍和设计理念，包括：
- 概述和主要特性
- 核心亮点（职责分离、双协议支持、完整功能码等）
- 架构设计（类层次结构、对象关系图、通信流程）
- 文档索引
- 依赖关系
- 兼容性说明

**适合人群**: 所有用户，特别是初次接触 Modbus 类库的开发者

---

### 2. [主站类详细文档 (master.md)](./master.md)
cModbusMaster 类的详细说明，包括：
- 类概述
- 事件列表（OnConnect、OnDisconnect、OnError、OnDataReceived）
- 属性参考（ProtocolType、State、SlaveID、RTU/TCP 配置等）
- 方法参考（Connect、Disconnect、读写操作）
- 事件详解
- 完整示例（基本主站、带重连的主站、数据采集）

**适合人群**: 需要作为 Modbus 主站（客户端）的开发者

---

### 3. [从站类详细文档 (slave.md)](./slave.md)
cModbusSlave 类的详细说明，包括：
- 类概述
- 事件列表（OnStarted、OnStopped、OnClientConnect、OnReadRequest 等）
- 属性参考（ProtocolType、State、SlaveID、数据存储等）
- 方法参考（Start、Stop、数据读写、数据管理等）
- 事件详解
- 完整示例（基本从站、动态数据更新、多客户端处理）

**适合人群**: 需要作为 Modbus 从站（服务器）的开发者

---

### 4. [快速开始 (quickstart.md)](./quickstart.md)
快速入门指南，帮助用户快速上手 Modbus 类库：
- 前置准备（必需文件、项目配置）
- 主站快速入门（TCP 模式）
- 从站快速入门（TCP 模式）
- 主从通信示例
- RTU 模式快速开始
- 完整功能示例
- 常见问题解答

**适合人群**: 初学者、快速原型开发

---

### 5. [进阶应用 (advanced.md)](./advanced.md)
高级功能和最佳实践，涵盖复杂应用场景：
- 高级主题（主从站合一模式、事务管理、异步操作模式）
- 性能优化（批量读取、频率控制、预缓存、连接池管理）
- 错误处理（综合错误处理、异常码处理）
- 多从站管理（设备配置管理、统一轮询管理）
- 数据缓存策略（多级缓存、写入同步）
- 日志与调试（详细日志记录、数据包调试）
- 安全考虑（连接认证、数据加密）
- 实际应用场景（工业数据采集、设备控制、数据网关）
- 常见问题（大数据量处理、热备份、断线重连）

**适合人群**: 有经验的开发者、需要实现复杂功能的用户

---

## 阅读建议

### 初次使用者
建议按以下顺序阅读：
1. [总览文档](./overview.md) - 了解整体架构和设计理念
2. [快速开始](./quickstart.md) - 快速上手，创建第一个应用
3. 根据需求选择：
   - 需要作为主站 → [主站类详细文档](./master.md)
   - 需要作为从站 → [从站类详细文档](./slave.md)

### 有经验的开发者
可以直接查阅相关文档：
- 主站开发 → [master.md](./master.md)
- 从站开发 → [slave.md](./slave.md)
- 性能优化 → [advanced.md](./advanced.md) 性能优化章节
- 高级功能 → [advanced.md](./advanced.md) 相关章节

### 快速参考
- 属性和方法 → 查看 [master.md](./master.md) 和 [slave.md) 的属性/方法参考部分
- 示例代码 → 各文档的"完整示例"章节
- 常见问题 → [quickstart.md](./quickstart.md) 和 [advanced.md](./advanced.md) 的常见问题章节

---

## 代码示例位置

除了文档中的示例代码外，还提供了完整的演示程序：

### 演示程序目录
- 位置: `src/Demos/Modbus/`
- 包含:
  - `Master/` - 主站演示程序
  - `Slave/` - 从站演示程序
  - `DEMO_README.md` - 演示程序使用说明

### 快速测试指南
- 位置: `src/Demos/Modbus/快速测试指南.md`
- 内容: TCP 和 RTU 模式的详细测试步骤

---

## 技术支持

如有问题或建议，请参考：
- 演示程序: `src/Demos/Modbus/`
- 示例代码: `src/Demos/Modbus/src/`
- 类源代码: `src/Modbus/`

---

## 版本信息

- **文档版本**: 1.0.0
- **最后更新**: 2026-01-16
- **作者**: woeoio@qq.com
- **基础库作者**: wqweto@gmail.com

---

**提示**: 建议从 [总览文档](./overview.md) 开始阅读，了解类库的整体设计理念和使用方法。
