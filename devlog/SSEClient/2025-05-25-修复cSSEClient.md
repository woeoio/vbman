# 2025-05-25 修复 cSSEClient

## Bug 修复

1. **只支持 GET，不支持 POST**：AI 流式聊天等场景需要 POST 请求体。新增 `ConnectPost(Url, Body)` 方法，自动禁用重连。

## 功能增强

2. **新增 `SetHeader()`/`ClearHeaders()`**：链式调用设置自定义请求头（如 Authorization、Content-Type），在 Connect/ConnectPost 时合并到 HttpClient。

3. **新增 `RequestTimeOut` 属性**：可配置请求超时（默认 60 秒），不再被 cHttpClient 的默认值覆盖。

4. **`Connect()` 改进**：清空 `LastEventData`/`LastEventType` 防止脏数据；合并自定义头；设置超时。
