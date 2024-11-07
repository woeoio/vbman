Attribute VB_Name = "ToolsStr"
Option Explicit


Function DecodeUnicode(Text As String) As String
    Dim i As Integer
    Dim result As String
    Dim unicodeStr As String
    Dim unicodeCode As Long
    
    result = ""
    i = 1
    
    While i <= Len(Text)
        If Mid(Text, i, 2) = "\u" Then
            unicodeStr = Mid(Text, i + 2, 4)                                    ' 获取 \u 后的四个字符
            unicodeCode = CLng("&H" & unicodeStr)                               ' 将四个字符转换为十六进制数
            result = result & ChrW(unicodeCode)                                 ' 将 Unicode 码转换为字符并添加到结果字符串
            i = i + 6                                                           ' 跳过 \uXXXX
        Else
            result = result & Mid(Text, i, 1)                                   ' 如果不是 Unicode 转义序列，直接添加字符
            i = i + 1
        End If
    Wend
    
    DecodeUnicode = result
End Function



' 处理 String 类型的切片
Function SliceString(ByVal Arr As String, ByVal StartPos As Long, Optional ByVal EndPos As Long = -1) As String
    Dim slicedStr As String
    
    If EndPos = -1 Then
        slicedStr = Mid$(Arr, StartPos + 1)                                     ' 从 startPos 到字符串末尾
    Else
        slicedStr = Mid$(Arr, StartPos + 1, EndPos - StartPos + 1)              ' 从 startPos 到 endPos
    End If
    
    SliceString = slicedStr
End Function

Public Function JoinStr(Span As String, ParamArray Strings() As Variant) As String
    JoinStr = Join(Strings, Span)
End Function


Public Function UniVbCrLf(Text As String) As String
    Dim ttt As String
    If Text = "" Then Exit Function
    ttt = Replace(Text, vbCrLf, "@@@")
    If ToolsStr.HasStr(vbLf, ttt) > 0 Then ttt = Replace(ttt, vbLf, vbCrLf)
    If ToolsStr.HasStr(vbCr, ttt) > 0 Then ttt = Replace(ttt, vbCr, vbCrLf)
    UniVbCrLf = Replace(ttt, "@@@", vbCrLf)
End Function

Function IsString(var As Variant) As Boolean
    IsString = (varType(var) = vbString)
End Function

Public Function HasStr(FindStr As String, FullStr As String, Optional StartPos As Long = 1, Optional CompType As VbCompareMethod = vbTextCompare) As Long
    'InStr 的参数顺序太难记了，因为其他语言也是这样的问题：命名模糊，不知道哪个参数是条件，哪个是操作对象。
    '因此增加了这个函数，专门用于判断字符串是否存在，并返回所在位置，
    '0 表示不存在，大于0 表示存在且标识位置
    HasStr = InStr(StartPos, FullStr, FindStr, CompType)
End Function


Public Function SubStr(ByVal Txt As String, ByVal txtF$, Optional ByVal txtE$, Optional RetInt As Boolean) As String
    'vb6截取任意位置字符串.md
    '其实应该使用正则提取更好，或者改进为 doloop 循环获取，返回数组或者集合的形式，
    Dim POSF&, POSE&, hasLeft&
    hasLeft = InStr(Txt, txtF)
    If hasLeft = 0 Then GoTo ret_nothing
    POSF = hasLeft + Len(txtF)
    If txtE = "" Then
        POSE = Len(Txt) + 1
    Else
        POSE = InStr(POSF, Txt, txtE)
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
Public Function ToString(inputArray() As Byte) As String
    ' 将字节数组转换为字符串（Unicode）
    ToString = StrConv(inputArray, vbUnicode)
End Function

Public Function ToHex(inputArray() As Byte) As String
    ' 打印字节数组（以十六进制形式显示）
    Dim tmp As String, i
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

Public Function GetNewIds(Optional ByVal Bath As Long)
    GetNewIds = "开发中..."
End Function

Public Function GetFirstChar(Txt As String, Optional Length As Long = 1) As String
    Dim Pos As Long, CharStr As String, CharDec As Long
    Do Until (CharDec >= 33 And CharDec <= 126)
        Pos = Pos + 1
        CharStr = Mid$(Txt, Pos, 1)
        CharDec = Asc(CharStr)
    Loop
    If Length = 1 Then
        GetFirstChar = CharStr
    Else
        GetFirstChar = Mid$(Txt, Pos, Length)
    End If
End Function

Public Function GetLastChar(Txt As String, Optional Length As Long = 1) As String
    Dim Pos As Long, CharStr As String, CharDec As Long
    Do Until (CharDec >= 33 And CharDec <= 126)
        Pos = Pos + 1
        CharStr = Mid$(Txt, Pos, 1)
        CharDec = Asc(CharStr)
    Loop
    If Length = 1 Then
        GetLastChar = CharStr
    Else
        GetLastChar = Mid$(Txt, Pos, Length)
    End If
End Function

Public Function GetRandStr(Optional ByVal Lens As Integer = 32, Optional Zuhe As String = "1aA") As String
    Rem 先顶用下，后期优化
    '取的随机字符串，组合参数支持  1aA@
    Rem 有时间可以开1000个线程并行计算，看看跑多长时间会出现重复，
    Rem 以 8 位 1aA 参数测试，32位字符重复率视为零
    '制作：2016-01-21   邓伟
    Dim Chars As String, i As Long, Max As Long, C As String, Pos As Long
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
        C = Mid$(Chars, Pos, 1)
        Char = Char & C
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
