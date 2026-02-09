# 阿里云验证码 2.0 - 短信登录演示

## 项目概述

本项目演示如何在短信登录场景中集成阿里云验证码 2.0，有效防止短信接口被恶意刷取。

## 界面预览

- **背景**: 淡黄色渐变背景，配有浮动气泡动画效果
- **登录框**: 简洁的手机号 + 短信验证码表单
- **交互**: 点击"获取验证码"前触发阿里云人机验证

## 流程说明

```
┌─────────────────────────────────────────────────────────────────┐
│                         用户访问登录页                            │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 1. 用户输入手机号                                                │
│    - 前端验证手机号格式（1[3-9]开头的11位数字）                     │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 2. 点击"获取验证码"按钮                                          │
│    - 验证手机号格式是否正确                                       │
│    - 格式正确则弹出阿里云验证组件                                  │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 3. 用户完成阿里云人机验证                                        │
│    - 可能是滑动验证、拼图验证或点击验证                             │
│    - 验证成功后获取 CaptchaVerifyParam 参数                        │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 4. 发送请求到 VB6 后端                                           │
│    POST /api/send-sms                                           │
│    {                                                            │
│        "phone": "13800138000",                                  │
│        "captchaVerifyParam": "eyJjZXJ0...",                     │
│        "sceneId": "xxxxxx"                                      │
│    }                                                            │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 5. VB6 后端处理流程                                              │
│    ├─ 接收请求参数                                               │
│    ├─ 调用 cAliyunCaptcha.Verify() 验证 CaptchaVerifyParam       │
│    ├─ 阿里云返回验证结果                                          │
│    │   ├─ 验证通过 (VerifyCode = T001/T005)                     │
│    │   │   └─ 调用短信服务商接口发送短信验证码                      │
│    │   │   └─ 返回 {success: true} 给前端                        │
│    │   └─ 验证失败                                               │
│    │       └─ 返回 {success: false, message: "验证失败"}         │
└─────────────────────────────┬───────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│ 6. 前端处理结果                                                  │
│    ├─ 后端返回成功                                                │
│    │   └─ 显示"验证码已发送"，开始60秒倒计时                      │
│    └─ 后端返回失败                                                │
│        └─ 显示错误信息，允许用户重新获取                           │
└─────────────────────────────────────────────────────────────────┘
```

## 后端 API 设计

### 1. 发送短信验证码

**请求**
```http
POST /api/send-sms
Content-Type: application/json

{
    "phone": "13800138000",
    "captchaVerifyParam": "eyJjZXJ0aWZ5SWQiOiIxMjM0NTYiLC...",
    "sceneId": "xxxxxxxx"
}
```

**响应成功**
```json
{
    "success": true,
    "message": "验证码已发送",
    "requestId": "xxxxx"
}
```

**响应失败**
```json
{
    "success": false,
    "message": "阿里云验证失败：疑似攻击请求",
    "verifyCode": "F001"
}
```

### 2. 登录验证

**请求**
```http
POST /api/login
Content-Type: application/json

{
    "phone": "13800138000",
    "smsCode": "123456"
}
```

**响应**
```json
{
    "success": true,
    "message": "登录成功",
    "token": "xxxxx"
}
```

## VB6 后端代码示例

```vb
' 发送短信验证码接口
Public Sub SendSmsCode(ByVal RequestBody As String, Response As String)
    On Error GoTo ErrorHandler
    
    ' 解析请求
    Dim Json As Object
    Set Json = ParseJson(RequestBody)
    
    Dim phone As String
    Dim captchaVerifyParam As String
    Dim sceneId As String
    
    phone = Json("phone")
    captchaVerifyParam = Json("captchaVerifyParam")
    sceneId = Json("sceneId")
    
    ' 第一步：验证阿里云验证码
    Dim Captcha As New cAliyunCaptcha
    Dim VerifyResult As AliyunCaptchaVerifyResult
    
    With Captcha
        .AccessKeyId("your-access-key-id") _
         .AccessKeySecret("your-access-key-secret") _
         .Region(ALIYUN_CAPTCHA_REGION_CN) _
         .Timeout(30000)
        
        VerifyResult = .Verify(captchaVerifyParam, sceneId)
    End With
    
    ' 检查阿里云验证结果
    If Not VerifyResult.Result Then
        Response = "{ ""success"": false, ""message"": """ & _
                   GetVerifyCodeDescription(VerifyResult.VerifyCode) & """, ""verifyCode"": """ & _
                   VerifyResult.VerifyCode & """ }"
        Exit Sub
    End If
    
    ' 第二步：阿里云验证通过，发送短信验证码
    Dim SmsCode As String
    SmsCode = GenerateRandomCode(6)  ' 生成6位随机码
    
    ' 保存到数据库/缓存（用于后续登录验证）
    SaveSmsCodeToCache phone, SmsCode, 300  ' 5分钟有效
    
    ' 调用短信服务商接口发送短信
    Dim SendResult As Boolean
    SendResult = SendSmsViaProvider(phone, "您的验证码是：" & SmsCode & "，5分钟内有效。")
    
    If SendResult Then
        Response = "{ ""success"": true, ""message"": ""验证码已发送"", ""requestId"": """ & _
                   VerifyResult.RequestId & """ }"
    Else
        Response = "{ ""success"": false, ""message"": ""短信发送失败，请重试"" }"
    End If
    
    Exit Sub
    
ErrorHandler:
    Response = "{ ""success"": false, ""message"": ""服务器错误：" & Err.Description & """ }"
End Sub

' 登录验证接口
Public Sub Login(ByVal RequestBody As String, Response As String)
    On Error GoTo ErrorHandler
    
    Dim Json As Object
    Set Json = ParseJson(RequestBody)
    
    Dim phone As String
    Dim smsCode As String
    
    phone = Json("phone")
    smsCode = Json("smsCode")
    
    ' 从缓存获取之前发送的验证码
    Dim CachedCode As String
    CachedCode = GetSmsCodeFromCache(phone)
    
    ' 验证短信验证码
    If CachedCode = "" Then
        Response = "{ ""success"": false, ""message"": ""验证码已过期，请重新获取"" }"
        Exit Sub
    End If
    
    If CachedCode <> smsCode Then
        Response = "{ ""success"": false, ""message"": ""验证码错误"" }"
        Exit Sub
    End If
    
    ' 验证成功，删除缓存的验证码
    RemoveSmsCodeFromCache phone
    
    ' 生成登录Token
    Dim Token As String
    Token = GenerateLoginToken(phone)
    
    Response = "{ ""success"": true, ""message"": ""登录成功"", ""token"": """ & Token & """ }"
    Exit Sub
    
ErrorHandler:
    Response = "{ ""success"": false, ""message"": ""服务器错误"" }"
End Sub
```

## 前端关键代码

### 触发阿里云验证
```javascript
document.getElementById('btn-send').addEventListener('click', function() {
    const phone = document.getElementById('phone').value;
    
    // 验证手机号格式
    if (!validatePhone(phone)) {
        showPhoneError(true);
        return;
    }
    
    // 显示阿里云验证弹窗
    document.getElementById('captcha-container').classList.add('show');
});
```

### 验证成功后发送请求
```javascript
window.initAliyunCaptcha({
    SceneId: "xxxxxx",
    mode: "popup",
    
    success: function(captchaVerifyParam) {
        // 验证成功，发送到后端
        fetch('/api/send-sms', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                phone: phone,
                captchaVerifyParam: captchaVerifyParam,
                sceneId: "xxxxxx"
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                startCountdown();  // 开始倒计时
            } else {
                showError(data.message);
            }
        });
    }
});
```

## 安全说明

1. **阿里云验证必须在发送短信前完成**
   - 防止机器人直接调用短信接口
   - 每次获取短信都需要重新验证

2. **短信验证码有效期**
   - 建议设置为 5 分钟
   - 错误次数限制（如连续错误3次需重新获取）

3. **手机号频率限制**
   - 同一手机号每分钟只能获取1次
   - 同一手机号每小时最多获取5次
   - 同一手机号每天最多获取10次

4. **IP 频率限制**
   - 同一 IP 每分钟最多请求10次
   - 超出限制需进行更严格验证或封禁

## 文件结构

```
www/
├── index.html      # 登录页面（主入口）
└── README.md       # 本文档
```

## 配置说明

### 方式一：硬编码配置（推荐用于开发/演示）

在 `index.html` 的 `GLOBAL_CONFIG` 中直接配置：

```javascript
const GLOBAL_CONFIG = {
    prefix: "your-prefix",              // 阿里云身份标识 Prefix
    sceneId: "your-scene-id",           // 阿里云场景ID SceneId
    region: "cn",                       // 地域: cn=中国内地, sgp=新加坡
    apiBaseUrl: "http://localhost:8080", // 后端API基础地址（可选）
    apiSendSms: "/api/send-sms",        // 发送短信API端点（可选，默认 /api/send-sms）
    apiLogin: "/api/login"              // 登录API端点（可选，默认 /api/login）
};
```

**优点：**
- 配置好以后直接打开页面即可使用，无需每次都输入
- 适合开发调试和固定环境部署

### 方式二：运行时配置（适合多用户环境）

将 `GLOBAL_CONFIG` 中的值留空：

```javascript
const GLOBAL_CONFIG = {
    prefix: "",              // 空字符串表示使用运行时配置
    sceneId: "",             // 空字符串表示使用运行时配置
    region: "cn",
    apiBaseUrl: "",
    apiSendSms: "/api/send-sms",  // API端点路径
    apiLogin: "/api/login"
};
```

页面打开后会弹出配置表单，用户输入后保存到浏览器 LocalStorage。

### 配置优先级

1. **GLOBAL_CONFIG** - 如果 `prefix` 和 `sceneId` 都有值，优先使用
2. **LocalStorage** - 如果 GLOBAL_CONFIG 为空，尝试读取之前保存的配置
3. **配置弹窗** - 如果都没有，显示配置弹窗让用户输入

### 获取配置值

登录 [阿里云验证码控制台](https://captcha.console.aliyun.com/) 获取：
- **Prefix** (身份标识) - 在控制台的基本配置中
- **SceneId** (场景ID) - 创建场景后获得
- **Region** - 根据您的服务器地域选择

## 依赖

- 阿里云验证码 2.0 JS SDK
- VB6 后端：`cAliyunCaptcha.cls` 和 `mAliyunCaptcha.bas`

## 浏览器兼容性

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+
- IE 11（需 polyfill）
