# cDialog - Windows API 对话框对象类

## 概述

`cDialog` 是一个基于 Windows API 的 VB/VBA 对话框对象类，提供了文件打开/保存对话框、文件夹选择对话框等常用对话框功能的封装。

**作者**: 215879458@qq.com
**版本**：v1.0.2.0

## 主要功能

- ✅ **打开文件对话框** (`ShowOpen`)
- ✅ **保存文件对话框** (`ShowSave`)
- ✅ **文件夹选择对话框** (`ShowBrowseForFolder`)
- ✅ **文件扩展名过滤**
- ✅ **多文件选择支持**
- ✅ **常用过滤器预设** (图片、文档、代码等)
- ✅ **文件路径处理工具方法**
- ✅ **文件打开功能** (`OpenFile`, `SelectAndOpenFile`)

## 属性说明

### 基础属性

| 属性          | 类型   | 说明           | 默认值 |
| ------------- | ------ | -------------- | ------ |
| `DialogTitle` | String | 对话框标题     | `""`   |
| `InitialDir`  | String | 初始目录       | `""`   |
| `DefaultExt`  | String | 默认文件扩展名 | `""`   |
| `FileName`    | String | 默认文件名     | `""`   |
| `Filter`      | String | 文件过滤器     | `""`   |

### 行为属性

| 属性              | 类型    | 说明               | 默认值  |
| ----------------- | ------- | ------------------ | ------- |
| `MultiSelect`     | Boolean | 是否允许多选       | `False` |
| `OverwritePrompt` | Boolean | 覆盖文件时是否提示 | `True`  |
| `PathMustExist`   | Boolean | 路径必须存在       | `True`  |
| `FileMustExist`   | Boolean | 文件必须存在       | `True`  |
| `HideReadOnly`    | Boolean | 隐藏只读复选框     | `True`  |
| `ReadOnly`        | Boolean | 以只读方式打开     | `False` |
| `NoChangeDir`     | Boolean | 不改变当前目录     | `False` |
| `CreatePrompt`    | Boolean | 创建文件时提示     | `False` |
| `NewDialogStyle`  | Boolean | 使用新对话框样式   | `True`  |

## 方法说明

### 核心方法

#### `ShowOpen()` As Variant

显示打开文件对话框。返回单个文件路径（字符串）或文件路径数组（多选模式）。

#### `ShowSave()` As String

显示保存文件对话框。返回用户选择的文件路径。

#### `ShowBrowseForFolder()` As String

显示文件夹选择对话框。返回用户选择的文件夹路径。

#### `SelectFiles()` As Collection

显示打开文件对话框，返回 `Collection` 集合对象，统一处理单选和多选场景。

**特点：**

- 单选和多选统一返回集合对象
- 使用 `Collection.Count` 判断是否选择了文件
- 使用 `For Each` 遍历所有文件
- 未选择文件时返回空集合（Count = 0）

### 过滤器方法

#### `AddFilter(strDescription, strExtension)`

添加文件过滤器。

- `strDescription`: 过滤器描述，如 `"文本文件 (*.txt)"`
- `strExtension`: 扩展名模式，如 `"*.txt"`

#### `ClearFilter()`

清除所有过滤器。

#### `SetCommonFilters()`

设置常用文件过滤器（文本文件、所有文件）。

#### `SetImageFilters()`

设置图片文件过滤器（BMP、JPEG、PNG、GIF）。

#### `SetDocumentFilters()`

设置文档文件过滤器（Word、Excel、PowerPoint、PDF、文本）。

#### `SetCodeFilters()`

设置代码文件过滤器（Basic、C/C++、C#、Java、Python、JavaScript、HTML、XML）。

### 文件操作方法

#### `OpenFile(strFilePath, [strOperation])` As Long

使用系统默认程序打开指定文件。

- `strFilePath`: 文件路径
- `strOperation`: 操作类型，默认为 `"open"`

#### `SelectAndOpenFile()` As Long

显示文件选择对话框并打开选中的文件。

#### `Reset()`

重置所有属性为默认值。

### 路径处理方法

#### `GetFilePath(strFullPath)` As String

从完整路径中提取文件路径（不带文件名）。

#### `GetFileName(strFullPath)` As String

从完整路径中提取文件名（不带路径）。

#### `GetFileExtension(strFullPath)` As String

从完整路径中提取文件扩展名（包括点号）。

#### `GetFileNameWithoutExt(strFullPath)` As String

从完整路径中提取文件名（不带扩展名）。

## 使用示例

### 示例 1: 基本打开文件对话框

```vb
Sub BasicOpenFileDialog()
    Dim dlg As New cDialog

    ' 设置对话框标题
    dlg.DialogTitle = "选择一个文件"

    ' 设置初始目录
    dlg.InitialDir = "C:\Users\Public\Documents"

    ' 设置默认文件扩展名
    dlg.DefaultExt = "txt"

    ' 显示对话框
    Dim strFile As String
    strFile = dlg.ShowOpen()

    If strFile <> "" Then
        MsgBox "您选择了: " & strFile
    Else
        MsgBox "您取消了选择"
    End If
End Sub
```

### 示例 2: 使用过滤器打开文件

```vb
Sub OpenFileWithFilter()
    Dim dlg As New cDialog

    ' 添加过滤器
    dlg.AddFilter "文本文件 (*.txt)", "*.txt"
    dlg.AddFilter "Word 文档 (*.doc;*.docx)", "*.doc;*.docx"
    dlg.AddFilter "所有文件 (*.*)", "*.*"

    ' 设置对话框标题
    dlg.DialogTitle = "选择要打开的文件"

    ' 显示对话框
    Dim strFile As String
    strFile = dlg.ShowOpen()

    If strFile <> "" Then
        MsgBox "选择文件: " & strFile & vbCrLf & _
               "文件名: " & dlg.GetFileName(strFile) & vbCrLf & _
               "扩展名: " & dlg.GetFileExtension(strFile)
    End If
End Sub
```

### 示例 3: 多文件选择

```vb
Sub OpenMultipleFiles()
    Dim dlg As New cDialog

    ' 启用多选
    dlg.MultiSelect = True

    ' 添加过滤器
    dlg.AddFilter "图片文件 (*.bmp;*.jpg;*.png)", "*.bmp;*.jpg;*.png"
    dlg.AddFilter "所有文件 (*.*)", "*.*"

    ' 设置对话框标题
    dlg.DialogTitle = "选择多个图片文件"

    ' 显示对话框
    Dim varFiles As Variant
    varFiles = dlg.ShowOpen()

    If IsArray(varFiles) Then
        Dim i As Long
        Dim strMsg As String

        strMsg = "您选择了 " & UBound(varFiles) - LBound(varFiles) + 1 & " 个文件:" & vbCrLf & vbCrLf
        For i = LBound(varFiles) To UBound(varFiles)
            strMsg = strMsg & (i + 1) & ". " & varFiles(i) & vbCrLf
        Next i

        MsgBox strMsg
    ElseIf varFiles <> "" Then
        MsgBox "您选择了: " & varFiles
    Else
        MsgBox "您取消了选择"
    End If
End Sub
```

### 示例 4: 使用预设过滤器

```vb
Sub UsePresetFilters()
    Dim dlg As New cDialog

    ' 使用预设的图片过滤器
    dlg.SetImageFilters
    dlg.DialogTitle = "选择图片文件"
    Dim strImage As String
    strImage = dlg.ShowOpen()

    ' 重置并使用文档过滤器
    dlg.Reset
    dlg.SetDocumentFilters
    dlg.DialogTitle = "选择文档文件"
    Dim strDoc As String
    strDoc = dlg.ShowOpen()

    ' 重置并使用代码过滤器
    dlg.Reset
    dlg.SetCodeFilters
    dlg.DialogTitle = "选择代码文件"
    Dim strCode As String
    strCode = dlg.ShowOpen()
End Sub
```

### 示例 5: 保存文件对话框

```vb
Sub SaveFileDialog()
    Dim dlg As New cDialog

    ' 设置对话框标题
    dlg.DialogTitle = "保存文件"

    ' 设置初始目录和默认文件名
    dlg.InitialDir = "C:\Users\Public\Documents"
    dlg.FileName = "新建文档"

    ' 设置默认扩展名
    dlg.DefaultExt = "txt"

    ' 添加过滤器
    dlg.AddFilter "文本文件 (*.txt)", "*.txt"
    dlg.AddFilter "所有文件 (*.*)", "*.*"

    ' 关闭覆盖提示（可选）
    dlg.OverwritePrompt = False

    ' 显示对话框
    Dim strFile As String
    strFile = dlg.ShowSave()

    If strFile <> "" Then
        MsgBox "保存到: " & strFile
        ' 这里可以添加保存文件的代码
    End If
End Sub
```

### 示例 6: 文件夹选择对话框

```vb
Sub BrowseForFolderDialog()
    Dim dlg As New cDialog

    ' 设置对话框标题
    dlg.DialogTitle = "选择一个文件夹"

    ' 显示对话框
    Dim strFolder As String
    strFolder = dlg.ShowBrowseForFolder()

    If strFolder <> "" Then
        MsgBox "您选择了文件夹: " & strFolder
    End If
End Sub
```

### 示例 7: 选择并打开文件

```vb
Sub SelectAndOpenFileExample()
    Dim dlg As New cDialog

    ' 设置过滤器
    dlg.AddFilter "文本文件 (*.txt)", "*.txt"
    dlg.AddFilter "所有文件 (*.*)", "*.*"

    ' 选择并打开文件
    Dim lngResult As Long
    lngResult = dlg.SelectAndOpenFile()

    If lngResult > 0 Then
        MsgBox "文件已打开"
    Else
        MsgBox "您取消了选择或打开失败"
    End If
End Sub
```

### 示例 8: 直接打开指定文件

```vb
Sub OpenSpecificFile()
    Dim dlg As New cDialog
    Dim strFile As String

    strFile = "C:\Users\Public\Documents\example.txt"

    ' 检查文件是否存在
    If Dir(strFile) <> "" Then
        Dim lngResult As Long
        lngResult = dlg.OpenFile(strFile)

        If lngResult > 32 Then
            MsgBox "文件已成功打开"
        Else
            MsgBox "打开文件失败"
        End If
    Else
        MsgBox "文件不存在: " & strFile
    End If
End Sub
```

### 示例 9: 路径处理工具

```vb
Sub PathProcessingExample()
    Dim dlg As New cDialog
    Dim strFullPath As String

    strFullPath = "C:\Users\Public\Documents\Report.xlsx"

    Debug.Print "完整路径: " & strFullPath
    Debug.Print "文件夹路径: " & dlg.GetFilePath(strFullPath)
    Debug.Print "文件名: " & dlg.GetFileName(strFullPath)
    Debug.Print "文件名(无扩展名): " & dlg.GetFileNameWithoutExt(strFullPath)
    Debug.Print "扩展名: " & dlg.GetFileExtension(strFullPath)

    ' 输出:
    ' 完整路径: C:\Users\Public\Documents\Report.xlsx
    ' 文件夹路径: C:\Users\Public\Documents
    ' 文件名: Report.xlsx
    ' 文件名(无扩展名): Report
    ' 扩展名: .xlsx
End Sub
```

### 示例 10: 使用 SelectFiles 统一处理文件选择

```vb
Sub UseSelectFilesCollection()
    Dim dlg As New cDialog
    Dim files As Collection
    Dim file As Variant
    Dim strMsg As String

    ' 设置对话框
    dlg.DialogTitle = "选择文件"
    dlg.SetImageFilters
    dlg.MultiSelect = True  ' 可以是 True 或 False

    ' 获取文件集合
    Set files = dlg.SelectFiles()

    ' 判断是否选择了文件
    If files.Count > 0 Then
        strMsg = "您选择了 " & files.Count & " 个文件:" & vbCrLf & vbCrLf
        For Each file In files
            strMsg = strMsg & "- " & file & vbCrLf
        Next file
        MsgBox strMsg
    Else
        MsgBox "您取消了选择"
    End If
End Sub
```

**说明：**

- 无论 `MultiSelect` 是 `True` 还是 `False`，都返回 `Collection` 对象
- 使用 `files.Count` 判断是否选择了文件
- 使用 `For Each` 遍历文件，无需区分单选和多选
- 代码更简洁、更统一

### 示例 11: 完整的文件处理流程

```vb
Sub CompleteFileWorkflow()
    Dim dlg As New cDialog
    Dim strFile As String
    Dim strContent As String
    Dim fso As Object
    Dim ts As Object

    ' 步骤 1: 选择文件
    dlg.SetCommonFilters
    dlg.DialogTitle = "选择要读取的文本文件"
    strFile = dlg.ShowOpen()

    If strFile = "" Then
        MsgBox "您取消了选择"
        Exit Sub
    End If

    ' 步骤 2: 读取文件内容
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(strFile) Then
        Set ts = fso.OpenTextFile(strFile, 1) ' 1 = ForReading
        strContent = ts.ReadAll
        ts.Close

        MsgBox "文件内容: " & vbCrLf & vbCrLf & strContent
    End If

    ' 步骤 3: 修改内容并保存
    strContent = strContent & vbCrLf & vbCrLf & "--- 已在 " & Now() & " 编辑 ---"

    ' 步骤 4: 选择保存位置
    dlg.Reset
    dlg.DialogTitle = "保存修改后的文件"
    dlg.FileName = dlg.GetFileNameWithoutExt(strFile) & "_modified"
    dlg.DefaultExt = "txt"

    Dim strSavePath As String
    strSavePath = dlg.ShowSave()

    If strSavePath <> "" Then
        Set ts = fso.CreateTextFile(strSavePath, True) ' True = overwrite
        ts.Write strContent
        ts.Close

        MsgBox "文件已保存到: " & strSavePath
    End If
End Sub
```

## 注意事项

1. **缓冲区大小**: 多选模式下会分配 32KB 的缓冲区，足够处理大量文件选择。

2. **返回值类型**:
   - `ShowOpen()` 返回 `Variant`，可能是字符串（单选）或数组（多选）
   - 使用前应使用 `IsArray()` 或 `VarType()` 检查返回值类型

3. **文件存在性**: 设置 `FileMustExist = True` 时，用户只能选择已存在的文件。

4. **路径处理**: 所有路径处理方法都使用反斜杠 `\` 作为路径分隔符。

5. **多文件选择**: 启用 `MultiSelect` 时，返回值为数组，第一个元素是路径，后续元素是文件名。

## 技术细节

本类使用以下 Windows API:

- `GetOpenFileNameW` / `GetSaveFileNameW` - 文件对话框
- `SHBrowseForFolder` - 文件夹选择对话框
- `SHGetPathFromIDList` - 路径转换
- `ShellExecute` - 文件打开

支持的 Windows 版本: Windows XP 及以上

---

## 更多对话框类

本库还提供以下增强对话框类，满足更多应用场景：

### cDialogEx - 扩展文件对话框

基于现代 IFileOpenDialog/IFileSaveDialog 接口（Vista+）的扩展文件对话框类，提供更现代的文件选择体验。

**主要特性：**
- 支持 Vista+ 现代文件对话框界面
- 支持文件夹选择（左右结构对话框）
- 支持自定义图标设置
- 兼容传统 cDialog 的属性接口

---

### cColorDialog - 颜色选择对话框

提供 Windows 标准颜色选择对话框功能。

**主要特性：**
- 标准颜色选择器
- RGB 分量独立访问（Red、Green、Blue 属性）
- HTML 颜色格式转换（`ToHtmlColor`、`FromHtmlColor`）
- 自定义 16 色调色板支持
- 全展开模式（显示完整颜色面板）

---

### cFontDialog - 字体选择对话框

提供 Windows 标准字体选择对话框功能。

**主要特性：**
- 字体名称、大小、样式选择
- 支持粗体、斜体、下划线、删除线效果
- 字体颜色选择
- 字体大小范围限制（MinSize、MaxSize）
- 等宽字体过滤选项
- CSS 字体样式输出（`ToCssFont`）
- 字体描述输出（`ToDescription`）

---

### cPrintDialog - 打印对话框

提供 Windows 标准打印对话框功能。

**主要特性：**
- 打印范围选择（全部/选定内容/页码范围）
- 打印份数设置
- 打印排序（Collate）
- 打印到文件选项
- 打印机设备上下文获取（hDC 属性）
- 打印设置描述输出

---

### cPageSetupDialog - 页面设置对话框

提供 Windows 标准页面设置对话框功能。

**主要特性：**
- 纸张大小设置
- 页边距设置
- 纸张方向（纵向/横向）
- 预设纸张尺寸（Letter、A4、Legal）
- 预设边距（默认边距、窄边距）
- 公制/英制单位切换
- 可打印区域计算

---

### cFindReplaceDialog - 查找替换对话框

提供 Windows 标准查找替换对话框功能（无模式对话框）。

**主要特性：**
- 无模式对话框设计
- 查找/替换模式切换
- 区分大小写选项
- 全字匹配选项
- 搜索方向选择
- 替换/全部替换功能
- 注册消息通知（`MessageId` 属性）

---

### cPasswordDialog - 密码输入对话框

提供安全的密码输入对话框，支持密码强度检测和复杂度验证。

**主要特性：**
- 密码强度实时检测（VeryWeak 到 VeryStrong 五级）
- 密码复杂度要求验证
  - 最小/最大长度限制
  - 大写字母要求
  - 小写字母要求
  - 数字要求
  - 特殊字符要求
- 密码确认模式
- 密码修改模式（`ShowChangeDialog`）
- 最大尝试次数限制
- 密码要求描述输出

---

### cCountdownMsgBox - 倒计时消息框

提供带倒计时计时器的消息框，支持自动关闭。

**主要特性：**
- 自动关闭功能（可配置倒计时秒数）
- 支持所有标准 MsgBox 按钮样式
  - OK / OKCancel / YesNo / YesNoCancel 等
- 支持所有标准图标样式
  - Information / Warning / Error / Question
- 超时检测（`IsTimedOut` 方法）
- 结果描述输出
- 便捷方法：`ShowInfo`、`ShowWarning`、`ShowError`、`ShowQuestion`、`ShowConfirm`

---

### cInputBox - 增强版输入对话框

提供增强版输入对话框，支持多种输入验证。

**主要特性：**
- 多种输入验证模式
  - 无验证（None）
  - 数值验证（Numeric）
  - 整数验证（Integer）
  - 字母验证（Alpha）
  - 字母数字验证（AlphaNumeric）
  - 邮箱验证（Email）
  - 电话验证（Phone）
  - URL 验证（URL）
  - 自定义验证（Custom）
- 输入长度限制（最小/最大）
- 占位符提示
- 空值允许控制
- 取消状态检测
- 便捷方法：`ShowSimple`、`ShowNumeric`、`ShowEmail`、`ShowURL`、`ShowPhone`
- 数值转换：`GetValueAsNumber`、`GetValueAsInteger`

---

### cProgressDialog - 进度对话框

提供带进度条的对话框，支持时间估算和取消操作。

**主要特性：**
- 标准进度条模式（0-100 或自定义范围）
- 步进式进度更新
- 已用时间显示
- 剩余时间估算
- 进度描述输出
- 取消操作支持
- 进度变化事件

---

### cAboutDialog - 关于对话框

提供应用程序关于对话框，支持系统信息显示。

**主要特性：**
- 应用信息展示（名称、版本、版权、描述）
- 网站/邮箱链接（可点击打开）
- 许可证信息
- 致谢信息
- 系统信息显示
  - Windows 版本
  - 屏幕分辨率
  - 内存信息（总量、可用、使用率）
- 快捷方法：`ShowSystemInfo`、`ShowLicense`、`ShowCredits`

---

### cDateTimeDialog - 日期时间选择对话框

提供日期时间选择对话框。

**主要特性：**
- 三种模式：日期、时间、日期时间
- 日期格式化输出
  - 短格式（MM/DD/YYYY）
  - 长格式（Monday, January 01, 2026）
  - 自定义格式
- 时间格式化输出（24小时制/12小时制）
- 秒数显示控制
- 年份范围限制
- 日期计算（`AddDays`、`AddHours`）
- 星期名称获取（`GetDayOfWeekName`）
- 月份名称获取（`GetMonthName`）
- 闰年自动处理
- 月份天数自动计算

---

## mDialog 公共模块

提供公共类型定义和工具函数。

**类型定义：**
- `COMDLG_FILTERSPEC` - 过滤器规格结构
- `RECT` - 矩形结构
- `POINT` - 点结构

**常量：**
- 纸张尺寸常量（Letter、Legal、A4、A3）
- 单位转换常量

**工具函数：**
- `InchesToMM` / `MMToInches` - 英寸/毫米转换
- `InchesToPoints` / `PointsToInches` - 英寸/点转换
- `TwipsToInches` / `InchesToTwips` - 缇/英寸转换
- `RGBToHtml` / `HtmlToRGB` - RGB/HTML颜色转换

## 版本历史

- **v1.0.2.0** - 新增 `SelectFiles()` 方法，返回 `Collection` 集合对象，统一处理单选和多选场景
- **v1.0** - 初始版本，支持基本的文件打开/保存和文件夹选择功能

## 许可

请根据项目需求添加适当的许可信息。

---

**需要帮助?** 联系作者: 215879458@qq.com

