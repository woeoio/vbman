# cFormater - 通用格式化器类

## 简介

`cFormater` 是一个支持**链式调用**的通用格式化器类，提供文件大小、时间日期、数字、文本、掩码等多种格式化功能。

兼容 VB6/VBA，使用 `Currency` 类型处理大数值（最大支持约 922 PB）。

---

## 快速开始

```vb
Dim F As New cFormater

' 文件大小格式化
Debug.Print F.Data(1536000@).ReturnFileSize()   ' 输出: 1.46 MB

' 相对时间
Debug.Print F.Data(Now - 0.5).ReturnTimeAgo()   ' 输出: 12小时前

' 手机号掩码
Debug.Print F.Data("13812345678").ReturnMaskedPhone()   ' 输出: 138****5678
```

---

## 核心方法

### Data(Source) - 链式入口

设置要格式化的数据源，返回 `Me` 支持链式调用。

| 参数 | 类型 | 说明 |
|------|------|------|
| `Source` | `Variant` | 任意类型的数据 |

```vb
' 基本用法
F.Data(123456).ReturnFileSize()

' With 语句支持
With F.Data("hello_world")
    Debug.Print .ReturnCamelCase()      ' helloWorld
    Debug.Print .ReturnPascalCase()     ' HelloWorld
End With
```

---

## 文件大小格式化

### ReturnFileSize([DecimalPlaces])

将字节数转换为人类可读的文件大小格式 (B/KB/MB/GB/TB/PB)。

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `DecimalPlaces` | `Long` | `2` | 小数位数 |

```vb
Debug.Print F.Data(512).ReturnFileSize()                 ' 512 B
Debug.Print F.Data(1536000@).ReturnFileSize()            ' 1.46 MB
Debug.Print F.Data(1073741824@).ReturnFileSize()         ' 1.00 GB
Debug.Print F.Data(1099511627776@).ReturnFileSize(0)     ' 1 TB
```

**注意**：VB6 中使用 `@` 后缀表示 `Currency` 类型，可精确表示大文件大小。

---

## 时间日期格式化

### ReturnTimeAgo()

返回相对时间描述（刚刚/n秒前/n分钟前/n小时前/n天前/n月前/n年前）。

```vb
Debug.Print F.Data(Now).ReturnTimeAgo()                  ' 刚刚
Debug.Print F.Data(Now - 0.0001).ReturnTimeAgo()         ' 8秒前
Debug.Print F.Data(Now - 0.02).ReturnTimeAgo()           ' 28分钟前
Debug.Print F.Data(Now - 0.5).ReturnTimeAgo()            ' 12小时前
Debug.Print F.Data(#1/1/2024#).ReturnTimeAgo()           ' 4月前 (假设当前5月)
```

### ReturnDateTime([FormatStr])

自定义格式输出日期时间。

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `FormatStr` | `String` | `"yyyy-MM-dd HH:mm:ss"` | VB 日期格式字符串 |

```vb
Debug.Print F.Data(Now).ReturnDateTime()                           ' 2024-01-15 09:30:00
Debug.Print F.Data(Now).ReturnDateTime("yyyy年MM月dd日")            ' 2024年01月15日
Debug.Print F.Data(Now).ReturnDateTime("HH:mm")                    ' 09:30
```

### ReturnShortDate() / ReturnShortTime()

快捷方法，分别返回日期和时间部分。

```vb
Debug.Print F.Data(Now).ReturnShortDate()    ' 2024-01-15
Debug.Print F.Data(Now).ReturnShortTime()    ' 09:30:00
```

---

## 数字格式化

### ReturnNumber([DecimalPlaces])

返回带千分位的格式化数字。

```vb
Debug.Print F.Data(1234567.89).ReturnNumber()       ' 1,234,567.89
Debug.Print F.Data(1234567.89).ReturnNumber(0)      ' 1,234,568
```

### ReturnCurrency([Symbol])

返回货币格式。

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `Symbol` | `String` | `"¥"` | 货币符号 |

```vb
Debug.Print F.Data(1999.99).ReturnCurrency()        ' ¥1,999.99
Debug.Print F.Data(1999.99).ReturnCurrency("$")     ' $1,999.99
```

### ReturnPercent([DecimalPlaces])

返回百分比格式（自动乘以100）。

```vb
Debug.Print F.Data(0.856).ReturnPercent()           ' 85.6%
Debug.Print F.Data(0.856).ReturnPercent(0)          ' 86%
```

### ReturnRoman()

返回罗马数字（支持 1-3999）。

```vb
Debug.Print F.Data(2024).ReturnRoman()              ' MMXXIV
Debug.Print F.Data(1999).ReturnRoman()              ' MCMXCIX
```

---

## 文本格式化

### ReturnTruncate(MaxLength, [Suffix])

截断超长文本，显示省略号。

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `MaxLength` | `Long` | - | 最大长度 |
| `Suffix` | `String` | `"..."` | 后缀 |

```vb
Debug.Print F.Data("这是一段很长的文本").ReturnTruncate(5)      ' 这是一...
Debug.Print F.Data("这是一段很长的文本").ReturnTruncate(5, "→") ' 这是一→
```

### ReturnPadLeft(TotalWidth, [PadChar])

左侧填充至指定宽度。

```vb
Debug.Print F.Data("123").ReturnPadLeft(6, "0")     ' 000123
Debug.Print F.Data("ABC").ReturnPadLeft(8)          '      ABC
```

### ReturnPadRight(TotalWidth, [PadChar])

右侧填充至指定宽度。

```vb
Debug.Print F.Data("Hello").ReturnPadRight(10, "-")  ' Hello-----
```

### ReturnProperCase()

首字母大写（每个单词）。

```vb
Debug.Print F.Data("hello world").ReturnProperCase()    ' Hello World
```

### ReturnCamelCase()

驼峰命名（camelCase）。

```vb
Debug.Print F.Data("hello_world-example").ReturnCamelCase()     ' helloWorldExample
Debug.Print F.Data("user name").ReturnCamelCase()               ' userName
```

### ReturnPascalCase()

帕斯卡命名（PascalCase）。

```vb
Debug.Print F.Data("hello_world-example").ReturnPascalCase()    ' HelloWorldExample
Debug.Print F.Data("user name").ReturnPascalCase()              ' UserName
```

### ReturnSnakeCase()

下划线命名（snake_case）。

```vb
Debug.Print F.Data("Hello World").ReturnSnakeCase()             ' hello_world
Debug.Print F.Data("user-name").ReturnSnakeCase()               ' user_name
```

### ReturnKebabCase()

短横线命名（kebab-case）。

```vb
Debug.Print F.Data("Hello World").ReturnKebabCase()             ' hello-world
Debug.Print F.Data("user_name").ReturnKebabCase()               ' user-name
```

### ReturnHtmlEncode()

HTML 特殊字符转义。

```vb
Debug.Print F.Data("<div>Hello & World</div>").ReturnHtmlEncode()
' &lt;div&gt;Hello &amp; World&lt;/div&gt;
```

### ReturnUrlEncode()

URL 编码（保留字符不编码）。

```vb
Debug.Print F.Data("Hello World!").ReturnUrlEncode()    ' Hello%20World%21
```

---

## 掩码/隐私保护

### ReturnMaskedPhone()

手机号掩码（保留前3后4位）。

```vb
Debug.Print F.Data("13812345678").ReturnMaskedPhone()   ' 138****5678
Debug.Print F.Data("138-1234-5678").ReturnMaskedPhone() ' 138****5678
```

### ReturnMaskedIDCard()

身份证号掩码（18位或15位）。

```vb
Debug.Print F.Data("110101199001011234").ReturnMaskedIDCard()   ' 110101********1234
Debug.Print F.Data("110101900101123").ReturnMaskedIDCard()      ' 110101******123
```

### ReturnMaskedEmail()

邮箱掩码（保留首字母和域名）。

```vb
Debug.Print F.Data("test@gmail.com").ReturnMaskedEmail()      ' t***@gmail.com
Debug.Print F.Data("admin@company.com").ReturnMaskedEmail()   ' a***@company.com
```

---

## 文件路径处理

### ReturnFileName()

返回文件名（含扩展名）。

```vb
Debug.Print F.Data("C:\Users\Test\file.txt").ReturnFileName()   ' file.txt
```

### ReturnFileNameWithoutExt()

返回文件名（不含扩展名）。

```vb
Debug.Print F.Data("C:\Users\Test\file.txt").ReturnFileNameWithoutExt()   ' file
```

### ReturnFileExt()

返回扩展名（小写）。

```vb
Debug.Print F.Data("C:\Users\Test\file.TXT").ReturnFileExt()    ' txt
```

### ReturnFilePath()

返回目录路径。

```vb
Debug.Print F.Data("C:\Users\Test\file.txt").ReturnFilePath()   ' C:\Users\Test
```

---

## 其他格式化

### ReturnHex()

返回十六进制字符串。

```vb
' 字符串转十六进制
Debug.Print F.Data("Hello").ReturnHex()             ' 48656C6C6F

' 字节数组转十六进制
Dim Bytes(0 To 2) As Byte
Bytes(0) = &HAB: Bytes(1) = &HCD: Bytes(2) = &HEF
Debug.Print F.Data(Bytes).ReturnHex()               ' ABCDEF
```

### ReturnBase64()

返回 Base64 编码。

```vb
Debug.Print F.Data("Hello World").ReturnBase64()    ' SGVsbG8gV29ybGQ=
```

### ReturnBooleanText([TrueText], [FalseText])

布尔值转自定义文本。

```vb
Debug.Print F.Data(True).ReturnBooleanText()                    ' 是
Debug.Print F.Data(False).ReturnBooleanText("开启", "关闭")      ' 关闭
Debug.Print F.Data(1).ReturnBooleanText("Yes", "No")            ' Yes
```

---

## 辅助属性和方法

### HasData (Property Get)

检查是否已设置数据源。

```vb
If F.HasData Then Debug.Print "有数据"
```

### RawValue (Property Get)

获取原始数据源值。

```vb
F.Data("test")
Debug.Print F.RawValue      ' test
```

### Clear()

清空数据源。

```vb
F.Clear()
Debug.Print F.HasData       ' False
```

---

## 完整示例

```vb
Sub Demo()
    Dim F As New cFormater
    
    ' 文件大小示例
    Debug.Print "=== 文件大小 ==="
    Debug.Print F.Data(512).ReturnFileSize()                    ' 512 B
    Debug.Print F.Data(1536000@).ReturnFileSize()               ' 1.46 MB
    Debug.Print F.Data(1073741824@).ReturnFileSize(1)           ' 1.0 GB
    
    ' 时间示例
    Debug.Print vbCrLf & "=== 相对时间 ==="
    Debug.Print F.Data(Now).ReturnTimeAgo()                     ' 刚刚
    Debug.Print F.Data(Now - 0.02).ReturnTimeAgo()              ' 28分钟前
    Debug.Print F.Data(#1/1/2024#).ReturnDateTime("yyyy/MM/dd") ' 2024/01/01
    
    ' 数字示例
    Debug.Print vbCrLf & "=== 数字格式化 ==="
    Debug.Print F.Data(1234567.89).ReturnNumber(2)              ' 1,234,567.89
    Debug.Print F.Data(0.856).ReturnPercent(1)                  ' 85.6%
    Debug.Print F.Data(1999.99).ReturnCurrency("$")             ' $1,999.99
    
    ' 文本示例
    Debug.Print vbCrLf & "=== 文本格式化 ==="
    Debug.Print F.Data("hello_world").ReturnCamelCase()         ' helloWorld
    Debug.Print F.Data("hello world").ReturnPascalCase()        ' HelloWorld
    Debug.Print F.Data("test").ReturnPadLeft(6, "0")            ' 000test
    
    ' 掩码示例
    Debug.Print vbCrLf & "=== 隐私掩码 ==="
    Debug.Print F.Data("13812345678").ReturnMaskedPhone()       ' 138****5678
    Debug.Print F.Data("admin@test.com").ReturnMaskedEmail()    ' a***@test.com
    
    ' 文件路径示例
    Debug.Print vbCrLf & "=== 文件路径 ==="
    With F.Data("C:\Users\Test\document.txt")
        Debug.Print .ReturnFileName()                           ' document.txt
        Debug.Print .ReturnFileNameWithoutExt()                 ' document
        Debug.Print .ReturnFileExt()                            ' txt
        Debug.Print .ReturnFilePath()                           ' C:\Users\Test
    End With
End Sub
```

---

## 技术说明

### 数据类型支持

- **文件大小**：使用 `Currency` 类型（VB6 最大精确数值类型）
  - 最大值：约 922 PB
  - 后缀 `@` 表示 Currency 类型，如 `1073741824@`

- **日期时间**：内部使用 `Date` 类型，支持 VB 标准日期格式

- **文本处理**：自动使用 `CStr()` 转换，支持任意可转换类型

### 错误处理

所有方法都包含基本的错误处理：
- 无效数据返回空字符串或默认值
- 数值转换失败返回 0 或 "0"
- 不会抛出运行时错误

### 性能考虑

- 链式调用创建临时对象，建议在循环外复用实例
- 大量数据处理时，考虑直接使用底层函数
