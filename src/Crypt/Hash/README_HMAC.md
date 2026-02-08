# cCryptoHMAC - HMAC (基于哈希的消息认证码) 实现类

## 概述

`cCryptoHMAC` 是一个基于 RFC 2104 标准的 HMAC 实现类，支持 HMAC-SHA1 和 HMAC-SHA256 算法。

HMAC 是一种使用密钥的消息认证码算法，用于验证消息的完整性和真实性。

## 支持的算法

- **HMAC-SHA1** - 160位输出，兼容性好
- **HMAC-SHA256** - 256位输出，安全性更高（推荐）

## 支持的编码

- **HMAC_ENCODING_UTF8** - UTF-8 编码（默认）
- **HMAC_ENCODING_ANSI** - ANSI 编码

## 使用方式

### 方式一：链式调用（推荐）

链式调用提供流畅的 API 设计，一行代码完成 HMAC 计算。

#### 基本语法

```vb
HMAC.Secret([密钥]).DataString/Bytes([数据]).ReturnHex/Base64/Bytes()
```

#### 快速开始

```vb
Dim HMAC As New cCryptoHMAC

' 最简单的用法（默认 HMAC-SHA256）
Dim sHMAC As String
sHMAC = HMAC.Secret("my-secret-key").DataString("Hello World").ReturnHex()
' 输出: a3c5... (64位十六进制字符串)
```

#### 详细用法

**1. 指定算法**

```vb
' HMAC-SHA256（默认）
sHMAC = HMAC.Secret("secret").DataString("data").ReturnHex()

' HMAC-SHA1
sHMAC = HMAC.Mode(HMAC_ALG_SHA1).Secret("secret").DataString("data").ReturnHex()
```

**2. 输入数据**

```vb
' 字符串输入（默认 UTF-8）
sHMAC = HMAC.Secret("secret").DataString("Hello World").ReturnHex()

' 字符串输入（ANSI 编码）
sHMAC = HMAC.Secret("secret").DataString("Hello", HMAC_ENCODING_ANSI).ReturnHex()

' 字节数组输入
Dim baData() As Byte
baData = StrConv("Hello World", vbFromUnicode)
sHMAC = HMAC.Secret("secret").DataBytes(baData).ReturnHex()
```

**3. 密钥设置**

```vb
' 字符串密钥（默认 UTF-8）
sHMAC = HMAC.Secret("my-secret-key").DataString("data").ReturnHex()

' 字节数组密钥
Dim baKey() As Byte
baKey = StrConv("my-secret-key", vbFromUnicode)
sHMAC = HMAC.SecretBytes(baKey).DataString("data").ReturnHex()
```

**4. 输出格式**

```vb
' 十六进制小写（默认）
sHMAC = HMAC.Secret("secret").DataString("data").ReturnHex()

' 十六进制大写
sHMAC = HMAC.Secret("secret").DataString("data").ReturnHex(True)

' Base64 编码
sHMAC = HMAC.Secret("secret").DataString("data").ReturnBase64()

' 字节数组
Dim baHMAC() As Byte
baHMAC = HMAC.Secret("secret").DataString("data").ReturnBytes()
```

### 方式二：传统方法

传统方法适合需要多次复用密钥的场景。

```vb
Dim HMAC As New cCryptoHMAC
Dim sHMAC As String

' 方法1：一行代码计算
sHMAC = HMAC.Compute("Hello World", "my-secret-key")
sHMAC = HMAC.Compute("Hello World", "my-secret-key", HMAC_ALG_SHA1)

' 方法2：先设置密钥，再计算
HMAC.SetKey "my-secret-key"
sHMAC = HMAC.Compute("Hello World", "my-secret-key")

' 字节数组计算
Dim baData() As Byte, baKey() As Byte
baData = StrConv("Hello World", vbFromUnicode)
baKey = StrConv("secret", vbFromUnicode)
sHMAC = HMAC.ComputeBytesToHex(baData, baKey, HMAC_ALG_SHA256)
```

## 完整示例

### 示例1：链式调用

```vb
Private Sub TestHMACChainCall()
    Dim HMAC As New cCryptoHMAC
    Dim sText As String
    sText = "Hello World"

    Debug.Print "HMAC Chain Call Demo"
    Debug.Print "====================="

    ' 1. 最简单的用法（默认 SHA256）
    Debug.Print "SHA256: " & HMAC.Secret("secret").DataString(sText).ReturnHex()

    ' 2. HMAC-SHA1
    Debug.Print "SHA1:   " & HMAC.Mode(HMAC_ALG_SHA1).Secret("secret").DataString(sText).ReturnHex()

    ' 3. 不同输出格式
    Debug.Print "Hex:    " & HMAC.Secret("secret").DataString(sText).ReturnHex()
    Debug.Print "Base64: " & HMAC.Secret("secret").DataString(sText).ReturnBase64()
    Debug.Print "Upper:  " & HMAC.Secret("secret").DataString(sText).ReturnHex(True)

    ' 4. 中文文本
    Debug.Print "Chinese UTF8: " & HMAC.Secret("密钥").DataString("你好世界", HMAC_ENCODING_UTF8).ReturnHex()

    ' 5. 重复使用
    HMAC.Mode(HMAC_ALG_SHA256)
    Debug.Print "Data 1: " & HMAC.Secret("key1").DataString("Hello").ReturnHex()
    Debug.Print "Data 2: " & HMAC.Secret("key2").DataString("World").ReturnHex()
End Sub
```

### 示例2：API 签名（阿里云/AWS等）

```vb
Private Function SignRequest(ByVal StringToSign As String, ByVal AccessKeySecret As String) As String
    Dim HMAC As New cCryptoHMAC

    ' 阿里云签名通常使用 HMAC-SHA1
    SignRequest = HMAC _
        .Mode(HMAC_ALG_SHA1) _
        .Secret(AccessKeySecret & "&") _
        .DataString(StringToSign, HMAC_ENCODING_UTF8) _
        .ReturnBase64()
End Function

' 使用
Dim sSignature As String
sSignature = SignRequest("GET&%2F&...", "your-access-key-secret")
Debug.Print "Signature: " & sSignature
```

### 示例3：JWT 签名验证

```vb
Private Function CreateJWTSignature(ByVal HeaderPayload As String, ByVal Secret As String) As String
    Dim HMAC As New cCryptoHMAC

    ' JWT 通常使用 HMAC-SHA256
    CreateJWTSignature = HMAC _
        .Mode(HMAC_ALG_SHA256) _
        .Secret(Secret) _
        .DataString(HeaderPayload, HMAC_ENCODING_UTF8) _
        .ReturnBase64()
End Function
```

### 示例4：传统方法

```vb
Private Sub TestHMACTraditional()
    Dim HMAC As New cCryptoHMAC
    Dim sData As String, sKey As String

    sData = "Hello World"
    sKey = "my-secret-key"

    Debug.Print "HMAC Traditional Demo"
    Debug.Print "====================="

    ' 使用默认算法（SHA256）
    Debug.Print "SHA256: " & HMAC.Compute(sData, sKey)

    ' 指定算法
    Debug.Print "SHA1:   " & HMAC.Compute(sData, sKey, HMAC_ALG_SHA1)

    ' 字节数组
    Dim baData() As Byte, baKey() As Byte
    baData = StrConv(sData, vbFromUnicode)
    baKey = StrConv(sKey, vbFromUnicode)

    Debug.Print "Bytes:  " & HMAC.ComputeBytesToHex(baData, baKey, HMAC_ALG_SHA256)
End Sub
```

## API 参考

### 枚举

| 枚举                 | 值    | 说明                     |
| -------------------- | ----- | ------------------------ |
| `HMAC_ALG_SHA1`      | 32772 | HMAC-SHA1 算法           |
| `HMAC_ALG_SHA256`    | 32780 | HMAC-SHA256 算法（默认） |
| `HMAC_ENCODING_UTF8` | 1     | UTF-8 编码（默认）       |
| `HMAC_ENCODING_ANSI` | 0     | ANSI 编码                |

### 属性

| 属性        | 说明                    |
| ----------- | ----------------------- |
| `Algorithm` | 获取/设置默认 HMAC 算法 |

### 传统方法

| 方法                                            | 说明                                    |
| ----------------------------------------------- | --------------------------------------- |
| `SetKey(KeyString, [Encoding])`                 | 设置密钥（字符串）                      |
| `SetKeyBytes(KeyBytes())`                       | 设置密钥（字节数组）                    |
| `Compute(Data, Key, [Algorithm], [Encoding])`   | 计算 HMAC（返回十六进制）               |
| `ComputeBytesToHex(Data(), Key(), [Algorithm])` | 计算 HMAC（字节数组输入，十六进制输出） |
| `ComputeBytes(Data(), Key(), [Algorithm])`      | 计算 HMAC（返回字节数组）               |

### 链式调用方法

| 方法                           | 说明                  |
| ------------------------------ | --------------------- |
| `Mode(Algorithm)`              | 设置算法              |
| `Secret(KeyString, [Encoding])` | 设置字符串密钥        |
| `SecretBytes(KeyBytes())`      | 设置字节数组密钥      |
| `DataString(Text, [Encoding])` | 输入字符串数据        |
| `DataBytes(Data())`            | 输入字节数组数据      |
| `ReturnHex([UpperCase])`       | 返回十六进制 HMAC     |
| `ReturnBase64()`               | 返回 Base64 编码 HMAC |
| `ReturnBytes()`                | 返回字节数组 HMAC     |

## 注意事项

1. **密钥安全**：
   - 密钥应安全存储，不要硬编码在代码中
   - 生产环境建议使用配置文件或密钥管理服务

2. **编码一致性**：
   - 与外部系统对接时，确保双方使用相同的字符编码
   - 大多数现代 API 使用 UTF-8 编码

3. **算法选择**：
   - 优先使用 HMAC-SHA256（更安全）
   - HMAC-SHA1 仅用于兼容旧系统

4. **性能考虑**：
   - 复用 HMAC 对象，避免频繁创建
   - 链式调用内部会缓存数据，注意及时清理

5. **错误处理**：
   - 所有方法都包含错误处理
   - 未设置密钥或数据时会抛出错误

## 兼容性

- **操作系统**：Windows XP 及以上版本
- **依赖**：
  - `cToolsUtf8` - UTF-8 编码处理
  - `cToolsBase64`（可选）- Base64 编码

## 与 cCryptoHash 的区别

| 特性     | cCryptoHash                       | cCryptoHMAC             |
| -------- | --------------------------------- | ----------------------- |
| 用途     | 数据完整性校验                    | 消息认证 + 完整性       |
| 密钥     | 不需要                            | 需要密钥                |
| 算法     | MD5, SHA1, SHA256, SHA384, SHA512 | HMAC-SHA1, HMAC-SHA256  |
| 安全性   | 可被碰撞攻击                      | 密钥保护下更安全        |
| 典型应用 | 文件校验、数据指纹                | API 签名、JWT、消息认证 |

## 作者

- 创建日期: 2026-02-08
- 基于项目编码规范: `.coding-conventions.md`

