Attribute VB_Name = "ToolsStr"
Option Explicit

'[desc:从字符串中提取数字的函数，支持可选参数来控制提取的个数和保留的小数点位数]
Function ParseNumbers(inputString As String, Optional decimalPlaces As Integer = -1, Optional count As Integer) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As New Collection
    Dim Num As Double
    Dim i As Integer
    
    ' 创建正则表达式对象
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = True
    regex.Pattern = "\d+(\.\d+)?"                                               ' 匹配整数或小数
    
    ' 使用正则表达式提取匹配的数字
    Set matches = regex.Execute(inputString)
    
    ' 遍历匹配结果并处理
    For i = 0 To matches.count - 1
        Set match = matches.Item(i)
        Num = CDbl(match.value)
        
        ' 如果指定了小数位数，进行四舍五入
        If decimalPlaces >= 0 Then
            Num = Round(Num, decimalPlaces)
        End If
        
        ' 将数字添加到集合
        result.Add Num
        
        ' 如果指定了个数，达到个数时停止
        If count > 0 And result.count >= count Then Exit For
    Next i
    
    ' 返回结果集合
    Set ParseNumbers = result
End Function



Public Function ToArray(Text As String) As String()
    Dim inputString As String
    Dim charArray() As String
    Dim i As Integer
    
    inputString = Text                                                          '"Hello"                                                       ' 要分割的字符串
    ReDim charArray(Len(inputString) - 1)                                       ' 为字符数组分配空间
    
    For i = 1 To Len(inputString)
        charArray(i - 1) = Mid(inputString, i, 1)                               ' 获取每个字符并存入数组
    Next i
    
    '    ' 输出数组内容
    '    For i = 0 To UBound(charArray)
    '        Debug.Print charArray(i)
    '    Next i
    ToArray = charArray
End Function

Public Function IsEmptyEx(Text As String) As Boolean
    Text = TrimEx(Text)
    IsEmptyEx = Text = ""
End Function
Public Function RightEx(Text As Variant, Length As Long) As String
    RightEx = Right(Trim(Text), Length)
End Function
Public Function LeftEx(Text As Variant, Length As Long) As String
    LeftEx = Left(Trim(Text), Length)
End Function
Public Function InsertSpan(ByRef inputStr As String, ByVal span As String, ByVal SetpNum As Long, Optional HeadFoot As Boolean) As String
    Dim resultStr As String
    Dim i As Integer
    Dim chunkLength As Integer
    chunkLength = SetpNum                                                       ' 每隔590字符插入一个[RTXREG]
    
    ' 初始化结果字符串
    resultStr = ""
    
    ' 遍历整个字符串，每次处理590个字符
    For i = 1 To Len(inputStr) Step chunkLength
        ' 将当前块添加到结果字符串
        resultStr = resultStr & Mid(inputStr, i, chunkLength)
        
        ' 如果当前块不是最后一块，插入[RTXREG]
        If i + chunkLength <= Len(inputStr) Then
            resultStr = resultStr & span
        End If
    Next i
    
    ' 返回处理后的字符串
    If HeadFoot = True Then
        InsertSpan = span & resultStr & span
    Else
        InsertSpan = resultStr
    End If
End Function


Public Function TrimEx(ByRef Text As String, Optional IsLeft As Boolean = True, Optional IsRight As Boolean = True) As String
    '确保移除不可见字符, 回车, 空格
    Dim i As Integer
    Dim startIdx As Integer
    Dim endIdx As Integer
    
    ' 去除不可见字符（如 ASCII 0 到 31 范围内的字符）
    Dim TSZF As String: TSZF = " " & vbTab & vbCrLf & vbCr & vbLf
    For i = 0 To 31
        TSZF = TSZF & Chr(i)
    Next i
    
    ' 去除前面的空格、回车、换行及不可见字符
    If IsLeft = True Then
        startIdx = 1
        Do While startIdx <= Len(Text) And InStr(TSZF, Mid(Text, startIdx, 1)) > 0
            startIdx = startIdx + 1
        Loop
    End If
    
    ' 去除末尾的空格、回车、换行及不可见字符
    If IsLeft = True Then
        endIdx = Len(Text)
        If endIdx > 0 Then
            Do While endIdx >= startIdx And InStr(TSZF, Mid(Text, endIdx, 1)) > 0
                endIdx = endIdx - 1
            Loop
        End If
    End If
    ' 截取并返回清理后的字符串
    TrimEx = Mid(Text, startIdx, endIdx - startIdx + 1)
End Function



'' test

'Print PercentDecode("%34%33%30%36%38%31%31%39%38%39%31%31%32%39%33%32%31%58")
'43068119891129321X
'43068119891129321X
'
'Print PercentEncode("43068119891129321X")
'%34%33%30%36%38%31%31%39%38%39%31%31%32%39%33%32%31X

Public Function PercentEncode(inputStr As String) As String
    Dim result As String
    Dim i As Integer
    Dim currentChar As String
    Dim hexValue As String
    
    ' 初始化结果字符串
    result = ""
    
    ' 遍历输入字符串的每个字符
    For i = 1 To Len(inputStr)
        currentChar = Mid(inputStr, i, 1)
        ' 判断是否为数字
        If IsNumeric(currentChar) Then
            ' 获取字符的 ASCII 值并转换为十六进制
            hexValue = Hex(Asc(currentChar))
            ' 拼接为 % 十六进制格式
            result = result & "%" & hexValue
        Else
            ' 非数字字符直接拼接到结果
            result = result & currentChar
        End If
    Next i
    
    ' 返回结果
    PercentEncode = result
End Function



Function PercentDecode(encodedStr As String) As String
    Dim result As String
    Dim i As Integer
    Dim hexValue As String
    Dim asciiValue As Integer
    
    ' 初始化结果字符串
    result = ""
    i = 1
    
    ' 遍历编码字符串
    Do While i <= Len(encodedStr)
        If Mid(encodedStr, i, 1) = "%" Then
            ' 提取 % 后的两位十六进制值
            hexValue = Mid(encodedStr, i + 1, 2)
            ' 将十六进制值转换为十进制
            asciiValue = CLng("&H" & hexValue)
            ' 转换为字符并拼接到结果
            result = result & Chr(asciiValue)
            ' 跳过当前的 % 和两位十六进制值
            i = i + 3
        Else
            ' 非 % 开头的字符直接拼接到结果
            result = result & Mid(encodedStr, i, 1)
            i = i + 1
        End If
    Loop
    
    ' 返回结果
    PercentDecode = result
End Function



Public Function UnicodeEncode(ByVal inputString As String, Optional PreFix As String = "\u") As String
    Dim i As Integer
    Dim unicodeString As String
    unicodeString = ""
    
    ' 遍历每个字符
    For i = 1 To Len(inputString)
        Dim charCode As Long
        Dim currentChar As String
        currentChar = Mid$(inputString, i, 1)
        charCode = AscW(currentChar)                                            ' 获取 Unicode 编码
        
        If charCode >= 0 And charCode <= 127 Then
            ' 如果是 ASCII 范围内的字符，直接添加
            unicodeString = unicodeString & currentChar
        Else
            ' 否则转为 \u 格式
            unicodeString = unicodeString & PreFix & Right$("0000" & Hex(charCode), 4)
        End If
    Next i
    
    UnicodeEncode = unicodeString
End Function


Public Function UnicodeDecode(Text As String, Optional PreFix As String = "\u") As String
    Dim i As Integer
    Dim result As String
    Dim unicodeStr As String
    Dim unicodeCode As Long
    
    result = ""
    i = 1
    
    While i <= Len(Text)
        If Mid(Text, i, 2) = PreFix Then
            unicodeStr = Mid(Text, i + 2, 4)                                    ' 获取 \u 后的四个字符
            unicodeCode = CLng("&H" & unicodeStr)                               ' 将四个字符转换为十六进制数
            result = result & ChrW(unicodeCode)                                 ' 将 Unicode 码转换为字符并添加到结果字符串
            i = i + 6                                                           ' 跳过 \uXXXX
        Else
            result = result & Mid(Text, i, 1)                                   ' 如果不是 Unicode 转义序列，直接添加字符
            i = i + 1
        End If
    Wend
    
    UnicodeDecode = result
End Function



' 处理 String 类型的切片
Public Function SliceString(ByVal Arr As String, ByVal StartPos As Long, Optional ByVal EndPos As Long = -1) As String
    Dim slicedStr As String
    
    If EndPos = -1 Then
        slicedStr = Mid$(Arr, StartPos + 1)                                     ' 从 startPos 到字符串末尾
    Else
        slicedStr = Mid$(Arr, StartPos + 1, EndPos - StartPos + 1)              ' 从 startPos 到 endPos
    End If
    
    SliceString = slicedStr
End Function

Public Function JoinStr(span As String, ParamArray Strings() As Variant) As String
    JoinStr = Join(Strings, span)
End Function


Public Function UniVbCrLf(Text As String) As String
    Dim ttt As String
    If Text = "" Then Exit Function
    ttt = Replace(Text, vbCrLf, "@@@")
    If ToolsStr.HasStr(vbLf, ttt) > 0 Then ttt = Replace(ttt, vbLf, vbCrLf)
    If ToolsStr.HasStr(vbCr, ttt) > 0 Then ttt = Replace(ttt, vbCr, vbCrLf)
    UniVbCrLf = Replace(ttt, "@@@", vbCrLf)
End Function

Public Function IsString(var As Variant) As Boolean
    IsString = (varType(var) = vbString)
End Function

Public Function HasStr(ByVal FindStr As String, FullStr As String, Optional StartPos As Long = 1, Optional CompType As VbCompareMethod = vbTextCompare) As Long
    'InStr 的参数顺序太难记了，因为其他语言也是这样的问题：命名模糊，不知道哪个参数是条件，哪个是操作对象。
    '因此增加了这个函数，专门用于判断字符串是否存在，并返回所在位置，
    '0 表示不存在，大于0 表示存在且标识位置
    HasStr = InStr(StartPos, FullStr, FindStr, CompType)
End Function

Public Function HasStrFromRight(ByVal FindStr As String, FullStr As String, Optional StartPos As Long = -1, Optional CompType As VbCompareMethod = vbTextCompare) As Long
    'InStr 的参数顺序太难记了，因为其他语言也是这样的问题：命名模糊，不知道哪个参数是条件，哪个是操作对象。
    '因此增加了这个函数，专门用于判断字符串是否存在，并返回所在位置，
    '0 表示不存在，大于0 表示存在且标识位置
    HasStrFromRight = InStrRev(FullStr, FindStr, StartPos, CompType)
End Function

Public Function SubStr( _
    ByVal Txt As String, _
    ByVal txtFirst As String, _
    Optional ByVal txtEnd As String, _
    Optional RetInt As Boolean, _
    Optional Method As VbCompareMethod = vbBinaryCompare, _
    Optional FindFromEnd As Boolean _
    ) As String
    'vb6截取任意位置字符串.md
    '其实应该使用正则提取更好，或者改进为 doloop 循环获取，返回数组或者集合的形式，
    Dim POSF&, POSE&, hasLeft&
    hasLeft = InStr(1, Txt, txtFirst, Method)
    If hasLeft = 0 Then GoTo ret_nothing
    POSF = hasLeft + Len(txtFirst)
    If txtEnd = "" Then
        POSE = Len(Txt) + 1
    Else
        If FindFromEnd = True Then
            POSE = InStrRev(Txt, txtEnd, -1, Method)
        Else
            POSE = InStr(POSF, Txt, txtEnd, Method)
        End If
    End If
    If POSE > POSF Then
        SubStr = Trim(Mid$(Txt, POSF, POSE - POSF))
    Else
ret_nothing:
        SubStr = IIf(RetInt, "0", "")
    End If
End Function

Public Function MidEx(FullStr As String, Lstr As String, Rstr As String, Optional starindex As Long) As String
    '截取任意位置字符串.函数，支持起始位置
    Dim i As Long
    Dim j As Long
    If starindex = 0 Then
        i = InStr(FullStr, Lstr)
    Else
        i = InStr(starindex, FullStr, Lstr, vbTextCompare)
    End If
    If i = 0 Then Exit Function
    j = InStr(i + 1, FullStr, Rstr)
    If j = 0 Then Exit Function
    starindex = j + Len(Rstr)
    If j - i + Len(Lstr) = 1 Then Exit Function
    Dim tmp As String
    tmp = Mid(FullStr, i + Len(Lstr), j - i - Len(Lstr))
    MidEx = Mid(tmp, 2, (Len(tmp) - 2))
End Function


Public Function LenBytes(inputArray() As Byte) As Long
    ' 计算字节数组长度（Unicode）
    LenBytes = UBound(inputArray) - LBound(inputArray) + 1
End Function
Public Function ToBytes(inputString As String) As Byte()
    ' 将字符串转换为字节数组（Unicode）
    ToBytes = StrConv(inputString, vbFromUnicode)
End Function
Public Function FromByteArray(inputArray() As Byte, Optional CharSet As String = "UTF-8") As String
    ' 将字节数组转换为字符串（Unicode）
    FromByteArray = StrConv(inputArray, vbUnicode)
End Function
Public Function FromHex(hexStr As String, Optional CharSet As String = "UTF-8") As String
    Dim ByteArray() As Byte
    Dim i As Long
    Dim strResult As String
    
    ' 去掉可能的空格
    hexStr = Replace(hexStr, " ", "")
    
    ' 验证字符串是否是偶数长度
    If Len(hexStr) Mod 2 <> 0 Then
        ERR.Raise vbObjectError + 513, "HexToString", "Hex string must have an even length."
        Exit Function
    End If
    
    ' 初始化字节数组
    ReDim ByteArray((Len(hexStr) \ 2) - 1)
    
    ' 转换为字节数组
    For i = 1 To Len(hexStr) Step 2
        ByteArray((i \ 2)) = CByte("&H" & Mid(hexStr, i, 2))
    Next i
    
    ' 将字节数组转为字符串
    strResult = StrConv(ByteArray, vbUnicode)
    
    ' 返回结果
    FromHex = strResult
End Function


Public Function ToHex(InputData As Variant, Optional CharSet As String = "UTF-8") As String
    ' 打印字节数组（以十六进制形式显示）
    Dim inputArray() As Byte
    If IsArray(InputData) = True Then
        inputArray = InputData
    Else
        inputArray = StrConv(InputData, vbFromUnicode)
    End If
    Dim tmp As String, i As Long
    For i = 0 To UBound(inputArray)
        tmp = tmp & Hex(inputArray(i))
    Next i
    ToHex = tmp
End Function

Public Function GetGUID(Optional isFull As Boolean) As String
    Dim TypeLib As Object
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GetGUID = IIf(isFull = True, Left(TypeLib.Guid, 38), Mid(TypeLib.Guid, 2, 36))
End Function

Public Function GetNewIds(Optional ByVal Bath As Long) As String
    GetNewIds = "开发中..."
End Function

Public Function GetFirstChar(Txt As String, Optional Length As Long = 1) As String
    Dim T As String: T = TrimEx(Txt, True, False)
    GetFirstChar = Left(T, Length)
    '    Dim Pos As Long, CharStr As String, CharDec As Long
    '    Do Until (CharDec >= 33 And CharDec <= 126)
    '        Pos = Pos + 1
    '        CharStr = Mid$(Txt, Pos, 1)
    '        CharDec = Asc(CharStr)
    '    Loop
    '    If Length = 1 Then
    '        GetFirstChar = CharStr
    '    Else
    '        GetFirstChar = Mid$(Txt, Pos, Length)
    '    End If
End Function

Public Function GetLastChar(Txt As String, Optional Length As Long = 1) As String
    Dim T As String: T = TrimEx(Txt, False, True)
    GetLastChar = Right(T, Length)
    '    Dim Pos As Long, CharStr As String, CharDec As Long
    '    Dim TxtLength As Long
    '
    '    ' 获取字符串长度
    '    TxtLength = Len(Txt)
    '
    '    ' 从字符串的最后一个字符开始向前查找有效字符
    '    Do Until (CharDec >= 33 And CharDec <= 126) Or Pos = TxtLength
    '        Pos = TxtLength - Pos
    '        CharStr = Mid$(Txt, Pos, 1)
    '        CharDec = Asc(CharStr)
    '    Loop
    '
    '    ' 返回找到的字符或字符段
    '    If Length = 1 Then
    '        GetLastChar = CharStr
    '    Else
    '        GetLastChar = Mid$(Txt, Pos - Length + 1, Length)
    '    End If
End Function


Public Function GetRandStr(Optional ByVal Lens As Integer = 32, Optional Zuhe As String = "1aA") As String
    Rem 先顶用下，后期优化
    '取的随机字符串，组合参数支持  1aA@
    Rem 有时间可以开1000个线程并行计算，看看跑多长时间会出现重复，
    Rem 以 8 位 1aA 参数测试，32位字符重复率视为零
    '制作：2016-01-21   邓伟
    Dim Chars As String, i As Long, Max As Long, c As String, Pos As Long
    Dim Char As String
    Dim C1 As String: C1 = "0123456789"
    Dim xA As String: xA = "abcdefghijklmnopqrstuvwxyz"
    Dim dA As String: dA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim CO As String: CO = "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
    If Zuhe = "" Then
        Chars = C1 & xA & dA & CO
    Else
        If InStr(Zuhe, "1") > 0 Then Chars = Chars & C1
        If InStr(Zuhe, "a") > 0 Then Chars = Chars & xA
        If InStr(Zuhe, "A") > 0 Then Chars = Chars & dA
        If InStr(Zuhe, "@") > 0 Then Chars = Chars & CO
    End If
    If Lens < 1 Then Lens = 8
    Max = Len(Chars)
    For i = 1 To Lens
        Pos = ToolsMath.GetRandRange(1, Max)
        c = Mid$(Chars, Pos, 1)
        Char = Char & c
        '        Debug.Print i, C, Max, Pos
    Next
    GetRandStr = Char
End Function

Public Function GetRandByte()
    Dim ByteArray() As Byte
    Dim byteSize As Integer
    Dim i As Integer
    
    ' 指定字节数组的大小
    byteSize = 10                                                               ' 例如，生成100个字节的数组
    
    ' 初始化随机数生成器
    Randomize
    
    ' 分配字节数组的大小
    ReDim ByteArray(byteSize - 1)
    
    ' 生成随机字节
    For i = 0 To byteSize - 1
        ByteArray(i) = Int(Rnd * 256)                                           ' 生成0到255之间的随机字节
    Next i
    
    ' 打印字节数组（以十六进制形式显示）
    For i = 0 To byteSize - 1
        Debug.Print Hex(ByteArray(i));
    Next i
End Function
