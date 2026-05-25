# 2025-05-25 修复 cHttpClient

## Bug 修复

1. **异步 `OnError` 事件未触发**：`Inst_OnError` 只写 `LastError`，没有 `RaiseEvent`，导致上层（cSSEClient）永远收不到错误通知。新增 `Public Event OnError(ErrorNumber, ErrorDescription)` 并在事件中触发。

2. **异步模式无 HTTP 状态码检查**：4xx/5xx 响应在异步模式下静默通过，调用方无感知。现在 `Inst_OnResponseFinished` 中检查状态码，4xx/5xx 触发 `OnError` 而非 `OnResponseFinished`。

3. **`RequestTimeOut < 30` 会修改用户设置值**：原代码直接 `RequestTimeOut = 30`，改用局部变量 `WaitTimeout`，不污染用户设置。

4. **`ParseSetCookie` 用逗号分割错误**：Cookie 的 `Expires` 值含逗号（如 `Thu, 01 Jan 2026`），按逗号分割会误拆。改用 `vbCrLf` 分割（WinHttp 多个 Set-Cookie 用换行分隔）。

5. **`Fetch` 每次不清除上次响应状态**：请求开始时清除 `m_StatusCode`/`m_StatusText`。

## 功能增强

6. **新增 `StatusCode`/`StatusText` 属性**：异步/同步均可用，不再需要直接访问 `Inst`。

7. **`SetCookies` 实现 Cookie 解析**：支持 `"name=value; name2=value2"` 格式自动解析到 `Cookies` 字典；新增 `BuildCookieHeader()` 从字典构建请求头；请求时自动带上已有 Cookie。

## 清理

8. **移除废弃的 `RequestDataBody`**：删除声明、初始化（`CompareMode`）、清理（`RemoveAll`）。
