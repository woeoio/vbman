# 2026-06-09 修复 Session 和 Cookie 设计缺陷

## P0 严重问题修复

1. **内存模式 Session 不工作**：原代码每次请求 `Dim Session As New cHttpServerSession` 创建全新对象，无全局存储，请求间数据完全丢失。新增 `m_Sessions As Scripting.Dictionary` 全局字典，内存模式直接引用存取。

2. **请求/响应 Cookie 混淆**：原 `Cookies.Decode()` 将请求 Cookie 存入 `Data`，`Encode()` 又把所有 Cookie（含请求带入的）回写为 `Set-Cookie`，导致每次请求重发所有 Cookie。拆分为 `m_RequestCookies`（只读，来自浏览器）和 `Data`（响应 Cookie，写回浏览器）。新增 `GetValue(Key)` 读取请求 Cookie，`Cookie(Key)` 仅用于设置响应 Cookie。

3. **首次写入 Session 数据丢失**：原代码仅 `SessionID <> ""` 时才保存，但 `SessionID` 是懒加载的——用户只写 `Session.Item("k")="v"` 而不读 `SessionID` 时，ID 未生成，数据丢失。现 `Item Let/Set` 时自动生成 SessionID。

## P1 中等问题修复

4. **每次请求都持久化**：新增 `Dirty` 标记，`Item Let/Set`、`Remove`、`Clear`、`Abandon` 时置 True，仅 `Dirty=True` 时才执行文件 I/O 或数据库写入。

5. **Session 过期清理缺失**：新增 `CleanupExpiredSessions` 方法，支持内存/文件/数据库三种模式清理。加载 Session 时发现过期立即删除，同时 `SessionAutoCleanup=True` 时约 1% 概率自动触发清理。

6. **Cookie 删除功能缺失**：新增 `ExpireCookie(Key)` 方法，设置 `Expires = 昨天` 通知浏览器删除指定 Cookie。

## P2 安全与优化

7. **SameSite 默认值**：Session Cookie 默认设置 `SameSite=Lax`，防止 CSRF。

8. **Session 序列化类型安全**：`Serialize()` 跳过 `IsObject` 类型的值，避免序列化对象引用报错。

9. **SessionID 去花括号**：`GenerateSessionID` 去掉 GUID 的 `{}`，Cookie 值更紧凑。

10. **Clear vs Abandon 语义**：`Clear()` 只清数据保留 SessionID（Cookie 不失效），`Abandon()` 才重新生成 ID 并标记 Dirty。

## API 变更提示

- 读取请求 Cookie：`ctx.Cookies.GetValue("name")`（不创建响应 Cookie）
- 设置响应 Cookie：`ctx.Cookies.Cookie("name").Value = "x"`（同前）
- 删除浏览器 Cookie：`ctx.Cookies.ExpireCookie("name")`
