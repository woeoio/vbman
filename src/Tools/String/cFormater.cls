VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' 通用格式化器类 - 支持链式调用
'   - 文件大小格式化
'   - 时间格式化（相对时间、日期时间）
'   - 数字格式化
'   - 文本格式化
' ============================================================

' --- 成员变量 ---
Private m_Source        As Variant                                              ' 原始数据源
Private m_HasData       As Boolean                                              ' 是否已设置数据

' --- 字节单位常量 (VB6 使用 Currency 类型支持大数值) ---
Private Const KB_C      As Currency = 1024@                                     ' 1 KB
Private Const MB_C      As Currency = 1048576@                                  ' 1024 * 1024
Private Const GB_C      As Currency = 1073741824@                               ' 1024^3
Private Const TB_C      As Currency = 1099511627776@                            ' 1024^4
'Private Const PB_C      As Currency = 1125899906842624@                         ' 1024^5

' --- 时间单位常量 (毫秒) ---
Private Const MS_SECOND As Long = 1000
Private Const MS_MINUTE As Long = 60000                                         ' 1000 * 60
Private Const MS_HOUR   As Long = 3600000                                       ' 1000 * 60 * 60
Private Const MS_DAY    As Long = 86400000                                      ' 1000 * 60 * 60 * 24

' ============================================================
' 链式入口
' ============================================================

' 设置数据源，返回 Me 支持链式调用
Public Function Data(ByVal Source As Variant) As cFormater
Attribute Data.VB_UserMemId = 0
    m_Source = Source
    m_HasData = True
    Set Data = Me
End Function

' ============================================================
' 文件大小格式化
' ============================================================

' 返回格式化后的文件大小字符串 (KB/MB/GB/TB/PB)
' 参数:
'   DecimalPlaces - 小数位数 (默认2)
'   ForceUnit     - 强制显示单位: "B"/"KB"/"MB"/"GB"/"TB"/"PB" (默认自动)
' 示例:
'   ReturnFileSize()              -> "1.5 MB"  (自动)
'   ReturnFileSize(2, "MB")       -> "1536.00 MB" (强制MB)
Public Function ReturnFileSize(Optional ByVal DecimalPlaces As Long = 2, Optional ByVal ForceUnit As String = "") As String
    Dim SizeValue As Variant
    
    If Not m_HasData Then
        ReturnFileSize = "0 B"
        Exit Function
    End If
    
    ' 转换为数值
    On Error Resume Next
    SizeValue = CDec(m_Source)
    On Error GoTo 0
    
    If Not IsNumeric(SizeValue) Or SizeValue < 0 Then
        ReturnFileSize = "0 B"
        Exit Function
    End If
    
    Dim Result As String
    Result = FormatBytes(CCur(SizeValue), DecimalPlaces, ForceUnit)
    ReturnFileSize = Result
End Function

' 内部函数：格式化字节
' 注意: PB 数值超出 Currency 范围(922万亿)，通过 TB 转换处理
Private Function FormatBytes(ByVal Bytes As Currency, ByVal Decimals As Long, ByVal ForceUnit As String) As String
    Dim Value As Double
    Dim Suffix As String
    Dim TB_Value As Double
    
    ' 处理强制单位
    ForceUnit = UCase$(Trim$(ForceUnit))
    
    Select Case ForceUnit
        Case "B"
            FormatBytes = FormatNumber(Bytes, Decimals) & " B"
            Exit Function
        Case "KB"
            Value = Bytes / KB_C
            FormatBytes = FormatNumber(Value, Decimals) & " KB"
            Exit Function
        Case "MB"
            Value = Bytes / MB_C
            FormatBytes = FormatNumber(Value, Decimals) & " MB"
            Exit Function
        Case "GB"
            Value = Bytes / GB_C
            FormatBytes = FormatNumber(Value, Decimals) & " GB"
            Exit Function
        Case "TB"
            Value = Bytes / TB_C
            FormatBytes = FormatNumber(Value, Decimals) & " TB"
            Exit Function
        Case "PB"
            ' PB 超出 Currency 范围，转为 Double 计算
            TB_Value = Bytes / TB_C  ' 先得到 TB 值
            Value = TB_Value / 1024  ' TB 转 PB
            FormatBytes = FormatNumber(Value, Decimals) & " PB"
            Exit Function
    End Select
    
    ' 自动选择单位
    Select Case Bytes
        Case Is >= TB_C
            ' 检查是否达到 PB 级别 (>= 1024 TB)
            TB_Value = Bytes / TB_C
            If TB_Value >= 1024 Then
                ' 使用 PB
                Value = TB_Value / 1024
                Suffix = " PB"
            Else
                ' 使用 TB
                Value = TB_Value
                Suffix = " TB"
            End If
        Case Is >= GB_C
            Value = Bytes / GB_C
            Suffix = " GB"
        Case Is >= MB_C
            Value = Bytes / MB_C
            Suffix = " MB"
        Case Is >= KB_C
            Value = Bytes / KB_C
            Suffix = " KB"
        Case Else
            FormatBytes = FormatNumber(Bytes, Decimals) & " B"
            Exit Function
    End Select
    
    FormatBytes = FormatNumber(Value, Decimals) & Suffix
End Function

' ============================================================
' 时间格式化
' ============================================================

' 返回相对时间描述
' 示例: ReturnTimeAgo() -> "刚刚", "5秒前", "3分钟前", "2小时前", "1天前"
Public Function ReturnTimeAgo() As String
    Dim TargetTime As Date
    Dim NowTime As Date
    Dim DiffMs As Double
    
    If Not m_HasData Then
        ReturnTimeAgo = ""
        Exit Function
    End If
    
    ' 转换为日期时间
    On Error Resume Next
    TargetTime = CDate(m_Source)
    On Error GoTo 0
    
    If TargetTime = 0 Then
        ReturnTimeAgo = ""
        Exit Function
    End If
    
    NowTime = Now
    DiffMs = DateDiff("s", TargetTime, NowTime) * 1000
    
    ' 未来时间
    If DiffMs < 0 Then
        ReturnTimeAgo = "未来"
        Exit Function
    End If
    
    ' 计算各时间单位
    Dim Seconds As Long
    Dim Minutes As Long
    Dim Hours As Long
    Dim Days As Long
    Dim Months As Long
    Dim Years As Long
    
    Seconds = Int(DiffMs / MS_SECOND)
    Minutes = Int(DiffMs / MS_MINUTE)
    Hours = Int(DiffMs / MS_HOUR)
    Days = Int(DiffMs / MS_DAY)
    
    ' 计算月和年（近似值）
    Months = DateDiff("m", TargetTime, NowTime)
    Years = DateDiff("yyyy", TargetTime, NowTime)
    
    ' 返回最匹配的描述
    Select Case True
        Case DiffMs < MS_SECOND
            ReturnTimeAgo = "刚刚"
        Case Seconds < 60
            ReturnTimeAgo = Seconds & "秒前"
        Case Minutes < 60
            ReturnTimeAgo = Minutes & "分钟前"
        Case Hours < 24
            ReturnTimeAgo = Hours & "小时前"
        Case Days < 30
            ReturnTimeAgo = Days & "天前"
        Case Months < 12
            ReturnTimeAgo = Months & "月前"
        Case Else
            ReturnTimeAgo = Years & "年前"
    End Select
End Function

' 返回格式化日期时间
' 示例: ReturnDateTime("yyyy-MM-dd HH:mm:ss") -> "2024-01-15 09:30:00"
Public Function ReturnDateTime(Optional ByVal FormatStr As String = "yyyy-MM-dd HH:mm:ss") As String
    Dim TargetTime As Date
    
    If Not m_HasData Then
        ReturnDateTime = ""
        Exit Function
    End If
    
    On Error Resume Next
    TargetTime = CDate(m_Source)
    On Error GoTo 0
    
    If TargetTime = 0 Then
        ReturnDateTime = ""
        Exit Function
    End If
    
    ReturnDateTime = Format(TargetTime, FormatStr)
End Function

' 返回日期（短格式）
Public Function ReturnShortDate() As String
    ReturnShortDate = ReturnDateTime("yyyy-MM-dd")
End Function

' 返回时间（短格式）
Public Function ReturnShortTime() As String
    ReturnShortTime = ReturnDateTime("HH:mm:ss")
End Function

' ============================================================
' 数字格式化
' ============================================================

' 返回格式化数字（千分位）
' 示例: ReturnNumber(2) -> "1,234,567.89"
Public Function ReturnNumber(Optional ByVal DecimalPlaces As Long = 0) As String
    Dim NumValue As Double
    
    If Not m_HasData Then
        ReturnNumber = "0"
        Exit Function
    End If
    
    On Error Resume Next
    NumValue = CDbl(m_Source)
    On Error GoTo 0
    
    If DecimalPlaces > 0 Then
        ReturnNumber = FormatNumber(NumValue, DecimalPlaces, , , vbTrue)
    Else
        ReturnNumber = FormatNumber(NumValue, 0, , , vbTrue)
    End If
End Function

' 返回货币格式
' 示例: ReturnCurrency() -> "?1,234,567.89"
Public Function ReturnCurrency(Optional ByVal Symbol As String = "?") As String
    Dim NumValue As Double
    
    If Not m_HasData Then
        ReturnCurrency = Symbol & "0.00"
        Exit Function
    End If
    
    On Error Resume Next
    NumValue = CDbl(m_Source)
    On Error GoTo 0
    
    ReturnCurrency = Symbol & FormatNumber(NumValue, 2, , , vbTrue)
End Function

' 返回百分比
' 示例: ReturnPercent(2) -> "12.50%"
Public Function ReturnPercent(Optional ByVal DecimalPlaces As Long = 0) As String
    Dim NumValue As Double
    
    If Not m_HasData Then
        ReturnPercent = "0%"
        Exit Function
    End If
    
    On Error Resume Next
    NumValue = CDbl(m_Source)
    On Error GoTo 0
    
    ReturnPercent = FormatNumber(NumValue * 100, DecimalPlaces) & "%"
End Function

' 返回罗马数字（支持 1-3999）
Public Function ReturnRoman() As String
    Dim NumValue As Long
    
    If Not m_HasData Then
        ReturnRoman = ""
        Exit Function
    End If
    
    On Error Resume Next
    NumValue = CLng(m_Source)
    On Error GoTo 0
    
    If NumValue < 1 Or NumValue > 3999 Then
        ReturnRoman = CStr(m_Source)
        Exit Function
    End If
    
    Dim Rom As String
    Dim Nums() As Integer
    Dim Romans() As String
    Dim i As Integer
    
    Nums = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    Romans = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    
    Rom = ""
    For i = 0 To UBound(Nums)
        Do While NumValue >= Nums(i)
            Rom = Rom & Romans(i)
            NumValue = NumValue - Nums(i)
        Loop
    Next i
    
    ReturnRoman = Rom
End Function

' ============================================================
' 文本格式化
' ============================================================

' 返回限制长度的文本，超长时显示省略号
' 示例: ReturnTruncate(10) -> "这是一段很长..."
Public Function ReturnTruncate(ByVal MaxLength As Long, Optional ByVal Suffix As String = "...") As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnTruncate = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    If Len(Text) <= MaxLength Then
        ReturnTruncate = Text
    Else
        ReturnTruncate = Left(Text, MaxLength - Len(Suffix)) & Suffix
    End If
End Function

' 返回填充后的文本
' 示例: ReturnPadLeft(10, "0") -> "0000123456"
Public Function ReturnPadLeft(ByVal TotalWidth As Long, Optional ByVal PadChar As String = " ") As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnPadLeft = String(TotalWidth, Left(PadChar, 1))
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    If Len(Text) >= TotalWidth Then
        ReturnPadLeft = Text
    Else
        ReturnPadLeft = String(TotalWidth - Len(Text), Left(PadChar, 1)) & Text
    End If
End Function

' 返回右侧填充后的文本
' 示例: ReturnPadRight(10, "-") -> "Hello-----"
Public Function ReturnPadRight(ByVal TotalWidth As Long, Optional ByVal PadChar As String = " ") As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnPadRight = String(TotalWidth, Left(PadChar, 1))
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    If Len(Text) >= TotalWidth Then
        ReturnPadRight = Text
    Else
        ReturnPadRight = Text & String(TotalWidth - Len(Text), Left(PadChar, 1))
    End If
End Function

' 返回首字母大写
' 示例: ReturnProperCase() -> "Hello World Example"
Public Function ReturnProperCase() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnProperCase = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    ReturnProperCase = StrConv(Text, vbProperCase)
End Function

' 返回驼峰命名
' 示例: ReturnCamelCase() -> "helloWorldExample"
Public Function ReturnCamelCase() As String
    Dim Text As String
    Dim Words() As String
    Dim i As Long
    Dim Result As String
    
    If Not m_HasData Then
        ReturnCamelCase = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    ' 替换常见分隔符为空格
    Text = Replace(Replace(Replace(Replace(Text, "-", " "), "_", " "), ".", " "), "/", " ")
    
    Words = Split(Text, " ")
    
    If UBound(Words) >= 0 Then
        ' 第一个单词小写
        Result = LCase(Words(0))
        ' 后续单词首字母大写
        For i = 1 To UBound(Words)
            If Len(Words(i)) > 0 Then
                Result = Result & UCase(Left(Words(i), 1)) & LCase(Mid(Words(i), 2))
            End If
        Next i
    End If
    
    ReturnCamelCase = Result
End Function

' 返回帕斯卡命名
' 示例: ReturnPascalCase() -> "HelloWorldExample"
Public Function ReturnPascalCase() As String
    Dim Text As String
    Dim Words() As String
    Dim i As Long
    Dim Result As String
    
    If Not m_HasData Then
        ReturnPascalCase = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    ' 替换常见分隔符为空格
    Text = Replace(Replace(Replace(Replace(Text, "-", " "), "_", " "), ".", " "), "/", " ")
    
    Words = Split(Text, " ")
    
    For i = 0 To UBound(Words)
        If Len(Words(i)) > 0 Then
            Result = Result & UCase(Left(Words(i), 1)) & LCase(Mid(Words(i), 2))
        End If
    Next i
    
    ReturnPascalCase = Result
End Function

' 返回下划线命名
' 示例: ReturnSnakeCase() -> "hello_world_example"
Public Function ReturnSnakeCase() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnSnakeCase = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    ' 替换常见分隔符为下划线
    Text = Replace(Replace(Replace(Replace(Text, " ", "_"), "-", "_"), ".", "_"), "/", "_")
    
    ReturnSnakeCase = LCase(Text)
End Function

' 返回短横线命名 (kebab-case)
' 示例: ReturnKebabCase() -> "hello-world-example"
Public Function ReturnKebabCase() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnKebabCase = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    
    ' 替换常见分隔符为短横线
    Text = Replace(Replace(Replace(Replace(Text, " ", "-"), "_", "-"), ".", "-"), "/", "-")
    
    ReturnKebabCase = LCase(Text)
End Function

' 返回HTML转义文本
Public Function ReturnHtmlEncode() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnHtmlEncode = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    Text = Replace(Text, "&", "&amp;")
    Text = Replace(Text, "<", "&lt;")
    Text = Replace(Text, ">", "&gt;")
    Text = Replace(Text, """", "&quot;")
    Text = Replace(Text, "'", "&#39;")
    
    ReturnHtmlEncode = Text
End Function

' 返回URL编码文本
Public Function ReturnUrlEncode() As String
    Dim Text As String
    Dim i As Long
    Dim Char As String
    Dim Result As String
    Dim CharCode As Integer
    
    If Not m_HasData Then
        ReturnUrlEncode = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    Result = ""
    
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        CharCode = Asc(Char)
        
        If (CharCode >= 48 And CharCode <= 57) Or _
           (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Or _
           Char = "-" Or Char = "_" Or Char = "." Or Char = "~" Then
            Result = Result & Char
        Else
            Result = Result & "%" & Hex(CharCode)
        End If
    Next i
    
    ReturnUrlEncode = Result
End Function

' ============================================================
' 掩码格式化
' ============================================================

' 返回手机号掩码 (138****8888)
Public Function ReturnMaskedPhone() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnMaskedPhone = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    Text = ExtractDigits(Text)
    
    If Len(Text) >= 11 Then
        ReturnMaskedPhone = Left(Text, 3) & "****" & Right(Text, 4)
    ElseIf Len(Text) >= 7 Then
        ReturnMaskedPhone = Left(Text, 3) & "****" & Right(Text, Len(Text) - 3)
    Else
        ReturnMaskedPhone = Text
    End If
End Function

' 返回身份证号掩码 (110101********1234)
Public Function ReturnMaskedIDCard() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnMaskedIDCard = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    Text = ExtractDigits(Text)
    
    If Len(Text) = 18 Then
        ReturnMaskedIDCard = Left(Text, 6) & "********" & Right(Text, 4)
    ElseIf Len(Text) = 15 Then
        ReturnMaskedIDCard = Left(Text, 6) & "******" & Right(Text, 3)
    Else
        ReturnMaskedIDCard = Text
    End If
End Function

' 返回邮箱掩码 (a***@gmail.com)
Public Function ReturnMaskedEmail() As String
    Dim Text As String
    Dim AtPos As Long
    Dim LocalPart As String
    Dim DomainPart As String
    
    If Not m_HasData Then
        ReturnMaskedEmail = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    AtPos = InStr(Text, "@")
    
    If AtPos > 1 Then
        LocalPart = Left(Text, AtPos - 1)
        DomainPart = Mid(Text, AtPos)
        
        If Len(LocalPart) > 1 Then
            ReturnMaskedEmail = Left(LocalPart, 1) & String(Len(LocalPart) - 1, "*") & DomainPart
        Else
            ReturnMaskedEmail = "*" & DomainPart
        End If
    Else
        ReturnMaskedEmail = Text
    End If
End Function

' 辅助函数：提取数字
Private Function ExtractDigits(ByVal Text As String) As String
    Dim i As Long
    Dim Result As String
    Dim Char As String
    
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        If Char >= "0" And Char <= "9" Then
            Result = Result & Char
        End If
    Next i
    
    ExtractDigits = Result
End Function

' ============================================================
' 其他实用格式化
' ============================================================

' 返回字节数组的十六进制字符串
' 示例: ReturnHex() -> "48656C6C6F"
Public Function ReturnHex() As String
    Dim Bytes() As Byte
    Dim i As Long
    Dim Result As String
    
    If Not m_HasData Then
        ReturnHex = ""
        Exit Function
    End If
    
    ' 如果是字节数组
    If IsArray(m_Source) Then
        On Error Resume Next
        Bytes = m_Source
        On Error GoTo 0
        
        If UBound(Bytes) > 0 Then
            For i = LBound(Bytes) To UBound(Bytes)
                Result = Result & Right("0" & Hex(Bytes(i)), 2)
            Next i
        End If
    Else
        ' 如果是字符串，转换其字节
        Dim Text As String
        Text = CStr(m_Source)
        For i = 1 To Len(Text)
            Result = Result & Right("0" & Hex(Asc(Mid(Text, i, 1))), 2)
        Next i
    End If
    
    ReturnHex = Result
End Function

' 返回Base64编码
Public Function ReturnBase64() As String
    Dim Text As String
    
    If Not m_HasData Then
        ReturnBase64 = ""
        Exit Function
    End If
    
    Text = CStr(m_Source)
    ReturnBase64 = Base64Encode(Text)
End Function

' Base64编码辅助函数
Private Function Base64Encode(ByVal Text As String) As String
    Dim arrData() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    arrData = StrConv(Text, vbFromUnicode)
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    
    Base64Encode = objNode.Text
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' 返回布尔值的友好文本
Public Function ReturnBooleanText(Optional ByVal TrueText As String = "是", Optional ByVal FalseText As String = "否") As String
    Dim BoolValue As Boolean
    
    If Not m_HasData Then
        ReturnBooleanText = FalseText
        Exit Function
    End If
    
    On Error Resume Next
    BoolValue = CBool(m_Source)
    On Error GoTo 0
    
    ReturnBooleanText = IIf(BoolValue, TrueText, FalseText)
End Function

' 返回文件路径的各个部分
Public Function ReturnFileName() As String
    Dim Path As String
    Dim Pos As Long
    
    If Not m_HasData Then
        ReturnFileName = ""
        Exit Function
    End If
    
    Path = CStr(m_Source)
    Pos = InStrRev(Path, "\")
    
    If Pos > 0 Then
        ReturnFileName = Mid(Path, Pos + 1)
    Else
        ReturnFileName = Path
    End If
End Function

Public Function ReturnFileNameWithoutExt() As String
    Dim FileName As String
    Dim Pos As Long
    
    FileName = ReturnFileName()
    Pos = InStrRev(FileName, ".")
    
    If Pos > 0 Then
        ReturnFileNameWithoutExt = Left(FileName, Pos - 1)
    Else
        ReturnFileNameWithoutExt = FileName
    End If
End Function

Public Function ReturnFileExt() As String
    Dim FileName As String
    Dim Pos As Long
    
    FileName = ReturnFileName()
    Pos = InStrRev(FileName, ".")
    
    If Pos > 0 Then
        ReturnFileExt = LCase(Mid(FileName, Pos + 1))
    Else
        ReturnFileExt = ""
    End If
End Function

Public Function ReturnFilePath() As String
    Dim Path As String
    Dim Pos As Long
    
    If Not m_HasData Then
        ReturnFilePath = ""
        Exit Function
    End If
    
    Path = CStr(m_Source)
    Pos = InStrRev(Path, "\")
    
    If Pos > 0 Then
        ReturnFilePath = Left(Path, Pos - 1)
    Else
        ReturnFilePath = ""
    End If
End Function

' ============================================================
' 清空数据源
' ============================================================
Public Sub Clear()
    m_Source = Empty
    m_HasData = False
End Sub

' 属性：是否有数据
Public Property Get HasData() As Boolean
    HasData = m_HasData
End Property

' 属性：原始值
Public Property Get RawValue() As Variant
    RawValue = m_Source
End Property
