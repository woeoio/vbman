# cCryptoHash - Windows CryptoAPI 哈希算法类

## 概述

`cCryptoHash` 是一个基于 Windows CryptoAPI 的哈希算法实现类，提供多种哈希算法的计算功能。

## 特性

- ✅ **支持多种哈希算法**：MD5、SHA1、SHA256、SHA384、SHA512
- ✅ **链式调用支持**：流畅的 API 设计，一行代码完成哈希计算
- ✅ **状态复用机制**：算法和数据可以复用，一次设置多次使用
- ✅ **多种输出格式**：支持十六进制（大小写）、Base64、字节数组
- ✅ **自动 Provider 选择**：根据算法自动选择支持该算法的加密服务提供程序
- ✅ **跨版本兼容**：支持 Windows XP 及以上版本
- ✅ **UTF-8 编码**：使用经过验证的 `cToolsUtf8` 类确保编码一致性
- ✅ **多种输入方式**：支持字符串、字节数组、文件

## 支持的算法

- **MD5** - 128位哈希
- **SHA1** - 160位哈希
- **SHA256** - 256位哈希
- **SHA384** - 384位哈希
- **SHA512** - 512位哈希

## 支持的编码

- **ENCODING_UTF8** - UTF-8 编码（默认，使用 `cToolsUtf8` 类）
- **ENCODING_ANSI** - ANSI 编码

## 公共方法

### 传统方法

- `ComputeHash(Text, [Algorithm], [Encoding])` - 计算字符串哈希（返回十六进制）
- `ComputeHashBytes(Data, [Algorithm])` - 计算字节数组哈希（返回字节数组）
- `ComputeHashBytesToHex(Data, [Algorithm])` - 计算字节数组哈希（返回十六进制）
- `ComputeFileHash(FilePath, [Algorithm])` - 计算文件哈希（返回十六进制）
- `Algorithm` - 属性，设置/获取默认哈希算法
- `ProviderName` - 属性，设置/获取加密服务提供程序名称

### 链式调用方法

- `Mode([Algorithm])` - 设置哈希算法（可选，默认使用 SHA256）
- `DataString(Text, [Encoding])` - 输入字符串数据
- `DataBytes(Data())` - 输入字节数组数据
- `ReturnHex([UpperCase])` - 返回十六进制格式哈希值
- `ReturnBase64()` - 返回 Base64 格式哈希值
- `ReturnBytes()` - 返回字节数组格式哈希值

## 链式调用（推荐）

链式调用提供了一种更加简洁和直观的方式来计算哈希值，支持流畅的编码风格。

### 基本语法

```vb
Hash.Mode([算法]).DataString/Bytes([数据]).ReturnHex/Base64/Bytes()
```

### 快速开始

```vb
Dim Hash As New cCryptoHash

' 最简单的用法（默认 SHA256）
Dim sHash As String
sHash = Hash.DataString("Hello World").ReturnHex()
```

### 详细用法

#### 1. 设置算法（可选）

```vb
' 方式1：省略 Mode，使用默认 SHA256
sHash = Hash.DataString("Hello").ReturnHex()

' 方式2：显式指定算法
sHash = Hash.Mode(HASH_ALG_MD5).DataString("Hello").ReturnHex()
sHash = Hash.Mode(HASH_ALG_SHA256).DataString("Hello").ReturnHex()
sHash = Hash.Mode(HASH_ALG_SHA512).DataString("Hello").ReturnHex()
```

#### 2. 输入数据

**字符串输入：**
```vb
' 默认 UTF8 编码
sHash = Hash.DataString("Hello World").ReturnHex()

' 指定编码
sHash = Hash.DataString("你好世界", ENCODING_UTF8).ReturnHex()
sHash = Hash.DataString("Hello", ENCODING_ANSI).ReturnHex()
```

**字节数组输入：**
```vb
Dim baData() As Byte
baData = StrConv("Hello World", vbFromUnicode)

sHash = Hash.DataBytes(baData).ReturnHex()
```

#### 3. 输出格式

**十六进制输出（小写）：**
```vb
sHash = Hash.DataString("Hello").ReturnHex()
' 输出: "185f8db32271fe25f561a6fc938b2e264306ec304eda518007d1764826381969"
```

**十六进制输出（大写）：**
```vb
sHash = Hash.DataString("Hello").ReturnHex(True)
' 输出: "185F8DB32271FE25F561A6FC938B2E264306EC304EDA518007D1764826381969"
```

**Base64 输出：**
```vb
sHash = Hash.DataString("Hello").ReturnBase64()
' 输出: "GF/Nsicn4+X1aWvyTi4iZkMG7DQ7pUYAfhdkhijgZaQ="
```

**字节数组输出：**
```vb
Dim baHash() As Byte
baHash = Hash.DataString("Hello").ReturnBytes()
Debug.Print "Hash size: " & (UBound(baHash) + 1) & " bytes"
```

### 完整示例

```vb
Private Sub TestChainCall()
    Dim Hash As New cCryptoHash
    
    ' 1. 简单示例
    Debug.Print Hash.DataString("Hello").ReturnHex()
    
    ' 2. 指定算法和编码
    Debug.Print Hash.Mode(HASH_ALG_SHA256).DataString("你好", ENCODING_UTF8).ReturnHex()
    
    ' 3. 不同输出格式
    Debug.Print "Hex: " & Hash.DataString("Test").ReturnHex()
    Debug.Print "Base64: " & Hash.DataString("Test").ReturnBase64()
    Debug.Print "Upper: " & Hash.DataString("Test").ReturnHex(True)
    
    ' 4. 字节数组输入
    Dim baData() As Byte
    baData = StrConv("Hello World", vbFromUnicode)
    Debug.Print Hash.Mode(HASH_ALG_MD5).DataBytes(baData).ReturnHex()
End Sub
```

### 链式调用特性

#### 1. 状态复用（推荐）

算法和数据可以复用，一次设置多次使用：

```vb
Dim Hash As New cCryptoHash

' 设置一次算法，多次使用
Hash.Mode(HASH_ALG_SHA256)
Debug.Print Hash.DataString("Hello").ReturnHex()
Debug.Print Hash.DataString("World").ReturnHex()
Debug.Print Hash.DataString("Test").ReturnHex()
```

```vb
' 设置一次数据，获取多种格式
Hash.DataString("Hello")
Debug.Print Hash.ReturnHex()
Debug.Print Hash.ReturnBase64()
Debug.Print Hash.ReturnHex(True)
```

#### 2. 自动覆盖

新的设置会自动覆盖旧的：

```vb
' 数据覆盖
Hash.DataString("Hello")
Hash.DataString("World")  // 自动覆盖 Hello
```

```vb
' 算法覆盖
Hash.Mode(HASH_ALG_MD5)
Hash.Mode(HASH_ALG_SHA256)  // 自动覆盖 MD5
```

#### 3. 默认算法

Mode 方法是可选的，如果不调用，默认使用 SHA256：

```vb
' 两者等效
sHash = Hash.Mode(HASH_ALG_SHA256).DataString("Hello").ReturnHex()
sHash = Hash.DataString("Hello").ReturnHex()
```

### 链式调用 vs 传统方法

**链式调用（推荐）：**
```vb
' 一行代码完成计算
sHash = New cCryptoHash.DataString("Hello").ReturnHex(True)
```

**传统方法：**
```vb
' 需要多个步骤
Dim Hash As New cCryptoHash
Hash.Algorithm = HASH_ALG_SHA256
sHash = Hash.ComputeHash("Hello")
```

### 最佳实践

1. **简单场景**：优先使用链式调用，代码更简洁
2. **性能关键**：复用 Hash 对象，避免频繁创建
3. **多次计算**：链式调用自动重置，无需手动管理状态
4. **代码清晰**：链式调用更符合自然语言习惯

## 基本用法

### 1. 一行代码计算哈希（推荐）

```vb
Dim Hash As New cCryptoHash

' 使用默认算法 (SHA256) 和 UTF8 编码
Dim sHash As String
sHash = Hash.ComputeHash("Hello World")

' 一行代码计算 MD5
sHash = Hash.ComputeHash("Hello World", HASH_ALG_MD5)

' 一行代码计算 SHA512
sHash = Hash.ComputeHash("Hello World", HASH_ALG_SHA512)

' 指定算法和编码
sHash = Hash.ComputeHash("你好世界", HASH_ALG_SHA256, ENCODING_UTF8)
sHash = Hash.ComputeHash("Hello World", HASH_ALG_SHA256, ENCODING_ANSI)
```

### 2. 使用类属性设置算法（传统方式）

```vb
Dim Hash As New cCryptoHash

' 使用默认算法 (SHA256)
Dim sHash As String
sHash = Hash.ComputeHash("Hello World")
' 输出: "a591a6d40bf420404a011733cfb7b190d62c65bf0bcda32b57b277d9ad9f146e"

' 使用 MD5 算法
Hash.Algorithm = HASH_ALG_MD5
sHash = Hash.ComputeHash("Hello World")
' 输出: "b10a8db164e0754105b7a99be72e3fe5"

' 使用 SHA512 算法
Hash.Algorithm = HASH_ALG_SHA512
sHash = Hash.ComputeHash("Hello World")
' 输出: "2c74fd17edafd80e8447b0d46741ee243b7eb74dd2149a0ab1b9246fb30382f27e853d8585719e0e67cbda0daa8f51671064615d645ae27acb15bfb1447f459b"
```

### 2. 计算字节数组哈希（返回十六进制字符串）

```vb
Dim Hash As New cCryptoHash
Dim baData() As Byte
baData = StrConv("Hello World", vbFromUnicode)

Dim sHash As String
sHash = Hash.ComputeHashBytesToHex(baData)
Debug.Print sHash

' 一行代码指定算法
sHash = Hash.ComputeHashBytesToHex(baData, HASH_ALG_MD5)
```

### 3. 计算字节数组哈希（返回原始字节数组）

```vb
Dim Hash As New cCryptoHash
Dim baData() As Byte
baData = StrConv("Hello World", vbFromUnicode)

Dim baHash() As Byte
baHash = Hash.ComputeHashBytes(baData)

' 获取哈希值长度
Debug.Print "Hash size: " & (UBound(baHash) + 1) * 8 & " bits"

' 一行代码指定算法
baHash = Hash.ComputeHashBytes(baData, HASH_ALG_SHA512)
```

### 4. 字符串编码示例

```vb
Dim Hash As New cCryptoHash
Dim sText As String
sText = "你好世界 Hello World"

' UTF8 编码（默认）
Dim sHashUTF8 As String
sHashUTF8 = Hash.ComputeHash(sText, HASH_ALG_SHA256, ENCODING_UTF8)
Debug.Print "UTF8: " & sHashUTF8

' ANSI 编码
Dim sHashANSI As String
sHashANSI = Hash.ComputeHash(sText, HASH_ALG_SHA256, ENCODING_ANSI)
Debug.Print "ANSI: " & sHashANSI
```

### 5. 自动 Provider 选择

类会根据所选算法自动选择最合适的加密服务提供程序（CSP）：

- **SHA256/SHA384/SHA512**：优先使用支持这些算法的高级 CSP
  - "Microsoft Enhanced RSA and AES Cryptographic Provider" (Windows Vista+)
  - "Microsoft Enhanced Cryptographic Provider v1.0"
  - "Microsoft Strong Cryptographic Provider"
- **MD5/SHA1**：使用基础 CSP（更好的兼容性）
  - "Microsoft Base Cryptographic Provider v1.0"
  - "Microsoft Enhanced Cryptographic Provider v1.0"
  - "Microsoft Strong Cryptographic Provider"

如果需要强制使用特定 Provider，可以设置 `ProviderName` 属性：

```vb
Dim Hash As New cCryptoHash

' 强制使用特定 Provider（不推荐，可能导致算法不支持）
Hash.ProviderName = "Microsoft Base Cryptographic Provider v1.0"
Dim sHash As String
sHash = Hash.ComputeHash("Hello World", HASH_ALG_MD5)  ' 只能使用 MD5/SHA1
```

**注意**：如果不支持的算法被强制使用不合适的 Provider，会抛出错误。

## 完整示例

### 传统方法示例

```vb
Private Sub TestCryptoHash()
    Dim Hash As New cCryptoHash

    ' 测试不同算法（一行代码）
    Dim sText As String
    sText = "VBMAN CryptoAPI Hash Test"

    Debug.Print "Original text: " & sText
    Debug.Print "--------------------------------"

    ' MD5
    Debug.Print "MD5: " & Hash.ComputeHash(sText, HASH_ALG_MD5)

    ' SHA1
    Debug.Print "SHA1: " & Hash.ComputeHash(sText, HASH_ALG_SHA1)

    ' SHA256 (默认)
    Debug.Print "SHA256: " & Hash.ComputeHash(sText, HASH_ALG_SHA256)

    ' SHA384
    Debug.Print "SHA384: " & Hash.ComputeHash(sText, HASH_ALG_SHA384)

    ' SHA512
    Debug.Print "SHA512: " & Hash.ComputeHash(sText, HASH_ALG_SHA512)

    ' 测试不同编码
    Debug.Print "--------------------------------"
    Dim sChinese As String
    sChinese = "你好世界"

    Debug.Print "UTF8: " & Hash.ComputeHash(sChinese, HASH_ALG_SHA256, ENCODING_UTF8)
    Debug.Print "ANSI: " & Hash.ComputeHash(sChinese, HASH_ALG_SHA256, ENCODING_ANSI)
End Sub
```

### 链式调用示例

```vb
Private Sub TestChainCall()
    Dim Hash As New cCryptoHash
    Dim sText As String
    sText = "VBMAN CryptoAPI Hash Test"

    Debug.Print "Chain Call Demo"
    Debug.Print "================================"

    ' 1. 最简单的用法（默认 SHA256）
    Debug.Print "Default SHA256:"
    Debug.Print "   " & Hash.DataString(sText).ReturnHex()
    Debug.Print ""

    ' 2. 不同算法
    Debug.Print "Different algorithms:"
    Debug.Print "   MD5:    " & Hash.Mode(HASH_ALG_MD5).DataString(sText).ReturnHex()
    Debug.Print "   SHA1:   " & Hash.Mode(HASH_ALG_SHA1).DataString(sText).ReturnHex()
    Debug.Print "   SHA256: " & Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    Debug.Print "   SHA512: " & Hash.Mode(HASH_ALG_SHA512).DataString(sText).ReturnHex()
    Debug.Print ""

    ' 3. 不同输出格式
    Debug.Print "Different output formats:"
    Debug.Print "   Hex (lower):  " & Hash.DataString(sText).ReturnHex()
    Debug.Print "   Hex (upper):  " & Hash.DataString(sText).ReturnHex(True)
    Debug.Print "   Base64:       " & Hash.DataString(sText).ReturnBase64()
    Debug.Print ""

    ' 4. 字节数组输入
    Debug.Print "Byte array input:"
    Dim baData() As Byte
    baData = StrConv(sText, vbFromUnicode)
    Debug.Print "   " & Hash.Mode(HASH_ALG_SHA256).DataBytes(baData).ReturnHex()
    Debug.Print ""

    ' 5. 中文文本
    Debug.Print "Chinese text:"
    Dim sChinese As String
    sChinese = "你好世界"
    Debug.Print "   UTF8:  " & Hash.DataString(sChinese, ENCODING_UTF8).ReturnHex()
    Debug.Print "   ANSI:  " & Hash.DataString(sChinese, ENCODING_ANSI).ReturnHex()
    Debug.Print ""

    ' 6. 重复使用（数据自动覆盖）
    Debug.Print "Reuse Hash object:"
    Debug.Print "   First:  " & Hash.DataString("Hello").ReturnHex()
    Debug.Print "   Second: " & Hash.DataString("World").ReturnHex()
    Debug.Print ""

    ' 7. 算法复用
    Debug.Print "Algorithm reuse:"
    Hash.Mode(HASH_ALG_SHA256)  // 设置一次
    Debug.Print "   Data 1: " & Hash.DataString("Hello").ReturnHex()
    Debug.Print "   Data 2: " & Hash.DataString("World").ReturnHex()
    Debug.Print ""

    ' 8. 同一数据多种格式
    Debug.Print "Same data, multiple formats:"
    Hash.DataString("Test")  // 设置一次
    Debug.Print "   Hex:    " & Hash.ReturnHex()
    Debug.Print "   Base64: " & Hash.ReturnBase64()
    Debug.Print "   Upper:  " & Hash.ReturnHex(True)
    Debug.Print ""

    ' 9. 一行代码
    Debug.Print "One-liners:"
    Debug.Print "   " & New cCryptoHash.DataString("Hello").ReturnHex()
    Debug.Print "   " & New cCryptoHash.Mode(HASH_ALG_MD5).DataString("Test").ReturnBase64()
End Sub
```

### 5. 文件哈希计算

```vb
Dim Hash As New cCryptoHash

' 计算文件哈希（使用默认算法 SHA256）
Dim sHash As String
sHash = Hash.ComputeFileHash("C:\path\to\file.txt")
Debug.Print "SHA256: " & sHash

' 一行代码计算文件 MD5
sHash = New cCryptoHash.ComputeFileHash("C:\path\to\file.txt", HASH_ALG_MD5)

' 计算文件 SHA512
sHash = Hash.ComputeFileHash("C:\path\to\file.exe", HASH_ALG_SHA512)
```

**说明**：`ComputeFileHash` 直接以二进制模式读取文件内容计算哈希，不进行任何字符编码处理，适用于所有类型的文件（文本、图片、二进制文件等）。

## 密码哈希示例（简单示例）

```vb
Private Sub HashPassword(ByVal Password As String)
    ' 在实际应用中，应该使用盐值（salt）和多次迭代

    ' 传统方法
    Dim sHash As String
    sHash = New cCryptoHash.ComputeHash(Password, HASH_ALG_SHA256, ENCODING_UTF8)
    Debug.Print "Password hash (traditional): " & sHash

    ' 链式调用方法
    sHash = New cCryptoHash.DataString(Password, ENCODING_UTF8).ReturnHex()
    Debug.Print "Password hash (chain): " & sHash
End Sub
```

## API 说明

本类使用以下 Windows CryptoAPI 函数：

- `CryptAcquireContext` - 获取加密服务提供程序句柄
- `CryptReleaseContext` - 释放加密服务提供程序句柄
- `CryptCreateHash` - 创建哈希对象
- `CryptDestroyHash` - 销毁哈希对象
- `CryptHashData` - 计算数据哈希值
- `CryptGetHashParam` - 获取哈希参数

字符编码处理使用项目中提供的 `cToolsUtf8` 类：

- `Encode` - 将 Unicode 字符串转换为 UTF-8 字节数组
- `Decode` - 将 UTF-8 字节数组转换为 Unicode 字符串

## 注意事项

1. **编码问题**：
   - 默认使用 UTF-8 编码，适合现代应用，与其他语言（如 JavaScript、Python）保持一致
   - 可以通过 `Encoding` 参数指定 ANSI 编码（兼容旧系统）
   - UTF-8 编码使用经过验证的 `cToolsUtf8` 类，确保编码转换的准确性

2. **算法参数**：
   - `Algorithm` 参数为可选参数，不指定时使用类属性 `Algorithm` 的值（默认 SHA256）
   - 可以在函数调用中直接指定算法，实现一行代码完成计算
   - 支持的算法：`HASH_ALG_MD5`, `HASH_ALG_SHA1`, `HASH_ALG_SHA256`, `HASH_ALG_SHA384`, `HASH_ALG_SHA512`

3. **Provider 选择**：
   - 类会自动根据算法选择最合适的加密服务提供程序
   - 自动处理不同 Windows 版本的兼容性
   - 不建议手动指定 Provider，除非有特殊需求

4. **链式调用**：
   - Mode 方法是可选的，默认使用 SHA256 算法
   - 算法和数据可以复用，一次设置多次使用
   - 新的设置会自动覆盖旧的设置
   - ReturnXxx 方法不会重置状态，允许重复获取不同格式

5. **大文件处理**：对于大文件，建议分块读取和哈希计算

6. **安全性**：
   - SHA1 和 MD5 已被证明存在碰撞攻击，不应在安全敏感的场景中使用
   - 对于密码存储，建议使用专门的密码哈希函数（如 bcrypt, PBKDF2, Argon2）

7. **错误处理**：所有方法都包含错误处理，遇到错误会抛出异常

8. **中文支持**：
   - 使用 UTF-8 编码可以正确处理中文字符
   - ANSI 编码可能会因系统代码页不同而产生不同结果

## 性能

- 使用系统原生 CryptoAPI，性能优秀
- 调用完成后自动清理资源，无内存泄漏

## 兼容性

- **操作系统**：Windows XP 及以上版本
- **加密服务**：系统自动选择最合适的加密服务提供程序（CSP）
  - Windows Vista+：支持 SHA2 系列算法（SHA256/384/512）
  - Windows XP/2000：支持 MD5、SHA1 等基础算法
- **编码依赖**：使用项目中的 `cToolsUtf8` 类处理 UTF-8 编码

## 作者

- 创建日期: 2026-02-08
- 基于项目编码规范: `.coding-conventions.md`

