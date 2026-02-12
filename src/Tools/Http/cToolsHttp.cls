VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




'————————————————VB URL的编解码源码 GB2312 UTF-8编解码
'
'                            版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。
'
'原文链接：https://blog.csdn.net/gs1069405343/article/details/50471825
'Public Function UrlDecodeUtf8(ByVal Url As String) As String
'    Dim i As Long
'    Dim S As String
'    Dim b As String
'    Dim hexStr As String
'    Dim decodedChar As Long
'    Dim utf8Bytes() As Byte
'    Dim utf8Char As String
'    Dim utf8Length As Byte
'
'    S = ""
'    i = 1
'    While i <= Len(Url)
'        b = Mid(Url, i, 1)
'
'        Select Case b
'        Case "+"                                                                ' '+' 转为空格
'            S = S & " "
'            i = i + 1
'
'        Case "%"                                                                ' URL 编码格式
'            ' 获取后两位作为十六进制字符
'            hexStr = Mid(Url, i + 1, 2)
'            If Len(hexStr) = 2 Then
'                decodedChar = CInt("&H" & hexStr)
'
'                ' 检查是否是单字节 ASCII 字符
'                If decodedChar < 128 Then
'                    S = S & ChrW(decodedChar)
'                    i = i + 3                                                   ' 跳过 "％" 和两位十六进制字符
'                Else
'                    ' 处理 UTF-8 编码字符（多字节）
'                    utf8Bytes = DecodeUtf8Bytes(Url, i + 1)
'                    '                    utf8Char = ChrW(CInt(utf8Bytes(0)))
'                    '                    Dim LLL As Long: LLL = UBound(utf8Bytes) + 1
'                    '                    If LLL > 1 Then
'                    '                        utf8Char = ChrW(CInt(utf8Bytes(1)))
'                    '                    End If
'                    '                    If LLL > 2 Then
'                    '                        utf8Char = ChrW(CInt(utf8Bytes(2)))
'                    '                    End If
'                    utf8Length = UBound(utf8Bytes) + 1
'                    utf8Char = ToolsUtf8.Decode(utf8Bytes)
'                    S = S & utf8Char
'                    i = i + utf8Length                                          ' 跳过 UTF-8 字节
'                End If
'            End If
'
'        Case Else                                                               ' 其他字符直接添加
'            S = S & b
'            i = i + 1
'        End Select
'    Wend
'
'    UrlDecodeUtf8 = S
'End Function

' 辅助函数：解码 UTF-8 字节
'Private Function DecodeUtf8Bytes(ByVal Url As String, ByVal startIndex As Long) As Byte()
'    Dim bytes() As Byte
'    Dim i As Long
'    Dim utf8Byte As Long
'    Dim hexStr As String
'
'    i = startIndex
'    While i <= Len(Url)
'        hexStr = Mid(Url, i, 2)
'        utf8Byte = CInt("&H" & hexStr)
'        If utf8Byte < &H80 Then
'            ReDim bytes(0)
'            bytes(0) = utf8Byte
'            DecodeUtf8Bytes = bytes
'            Exit Function
'        ElseIf utf8Byte >= &HC0 And utf8Byte <= &HDF Then
'            ' 处理 2 字节 UTF-8
'            ReDim bytes(1)
'            bytes(0) = utf8Byte
'            hexStr = Mid(Url, i + 3, 2)
'            bytes(1) = CInt("&H" & hexStr)
'            DecodeUtf8Bytes = bytes
'            Exit Function
'        ElseIf utf8Byte >= &HE0 And utf8Byte <= &HEF Then
'            ' 处理 3 字节 UTF-8
'            ReDim bytes(2)
'            bytes(0) = utf8Byte
'            hexStr = Mid(Url, i + 3, 2)
'            bytes(1) = CInt("&H" & hexStr)
'            hexStr = Mid(Url, i + 6, 2)
'            bytes(2) = CInt("&H" & hexStr)
'            DecodeUtf8Bytes = bytes
'            Exit Function
'        End If
'    Wend
'    DecodeUtf8Bytes = bytes
'End Function

Public Function UrlDecodeUtf8(ByVal url As String) As String
    Dim b As Variant, ub As Variant                                             ''中文字的Unicode码(2字节)
    Dim aa As Variant, BB As Variant
    Dim UtfB As Variant                                                         ''Utf-8单个字节
    Dim UtfB1 As Variant, UtfB2 As Variant, UtfB3 As Variant                    ''Utf-8码的三个字节
    Dim i As Long, n As Long, S As String
    Dim str1 As String
    Dim str2 As String
    '    n = 0
    '    ub = 0
    For i = 1 To Len(url)
        b = Mid(url, i, 1)
        Select Case b
        Case "+"
            S = S & " "
        Case "%"
            ub = Mid(url, i + 1, 2)
            If InStr(ub, vbLf) <= 0 And ub <> "" Then
                aa = Mid(ub, 1, 1)
                BB = Mid(ub, 2, 1)
                If aa < "g" And aa < "G" And BB < "g" And BB < "G" And aa <> "%" And BB <> "%" Then
                    UtfB = CInt("&H" & ub)
                End If
            End If
            
            If UtfB < 128 Then
                i = i + 2
                S = S & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000                                 ''取第1个Utf-8字节的二进制后4位
                str1 = Mid(url, i + 4, 2)
                If InStr(str1, vbLf) <= 0 And str1 <> "" Then
                    
                    aa = Mid(str1, 1, 1)
                    BB = Mid(str1, 2, 1)
                    If aa < "g" And aa < "G" And BB < "g" And BB < "G" And aa <> "%" And BB <> "%" Then
                        UtfB2 = (CInt("&H" & str1) And &H3F) * &H40             ''取第2个Utf-8字节的二进制后6位
                    End If
                    
                    str2 = Mid(url, i + 7, 2)
                    If InStr(str2, vbLf) <= 0 And str2 <> "" Then
                        aa = Mid(str2, 1, 1)
                        BB = Mid(str2, 2, 1)
                        If aa < "g" And aa < "G" And BB < "g" And BB < "G" And aa <> "%" And BB <> "%" Then
                            UtfB3 = CInt("&H" & str2) And &H3F                  ''取第3个Utf-8字节的二进制后6位
                        End If
                    End If
                End If
                S = S & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                i = i + 8
            End If
            
        Case Else                                                               ''Ascii码
            S = S & b
        End Select
    Next
    UrlDecodeUtf8 = S
End Function
'UTF-8编码
'Public Function UrlEncodeUtf8(ByVal szInput As Variant) As String
'    Dim wch As Variant, uch As Variant, szRet As Variant
'    Dim x As Variant
'    Dim nAsc As Variant, nAsc2 As Variant, nAsc3 As Variant
'    Dim szSafeChars As String
'
'    ' 定义 URL 安全字符
'    szSafeChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"
'
'    If szInput = "" Then
'        UrlEncodeUtf8 = szInput
'        Exit Function
'    End If
'
'    For x = 1 To Len(szInput)
'        wch = Mid(szInput, x, 1)
'        nAsc = AscW(wch)
'
'        ' 检查是否是安全字符
'        If InStr(szSafeChars, wch) > 0 Then
'            szRet = szRet & wch
'        Else
'            ' 对非安全字符进行编码
'            If (nAsc < 128) Then
'                ' 处理 ASCII 范围内的字符
'                szRet = szRet & "%" & Right("00" & Hex(nAsc), 2)
'            Else
'                '            ElseIf (nAsc And &HF800) = &HF800 Then
'                ' 处理多字节字符 (UTF-8 编码)
'                If (nAsc And &HF000) = 0 Then
'                    uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
'                    szRet = szRet & uch
'                Else
'                    uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
'                    Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
'                    Hex(nAsc And &H3F Or &H80)
'                    szRet = szRet & uch
'                End If
'                '            Else
'                '                ' 处理其他非 ASCII 字符 (UTF-8 编码)
'                '                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
'                '                szRet = szRet & uch
'            End If
'        End If
'    Next
'
'    UrlEncodeUtf8 = szRet
'End Function

Public Function UrlEncodeUtf8(ByVal szInput As Variant) As String
    Dim wch As Variant, uch As Variant, szRet As Variant
    Dim x As Variant
    Dim nAsc As Variant, nAsc2 As Variant, nAsc3 As Variant
    Dim szSafeChars As String
    
    ' RFC3986 标准的安全字符（不需要编码）
    szSafeChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"
    
    If szInput = "" Then
        UrlEncodeUtf8 = szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        
        If nAsc < 0 Then nAsc = nAsc + 65536
        
        ' 检查是否是安全字符（RFC3986）
        If InStr(szSafeChars, wch) > 0 Then
            ' 安全字符，直接追加
            szRet = szRet & wch
        ElseIf (nAsc And &HFF80) = 0 Then
            ' 单字节非安全字符（如 / = & : 等），进行 URL 编码
            szRet = szRet & "%" & Right("00" & Hex(nAsc), 2)
        Else
            ' 多字节 UTF-8 字符
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UrlEncodeUtf8 = szRet
End Function


'GB2312 URL解码
Public Function UrlDecode(ByVal url As String) As String
    Dim i As Long, c As String, d As Long, GB_UrlDecode As String
    i = 1
    While i <= Len(url)
        c = Mid$(url, i, 1)
        i = i + 1
        If c = "%" Then
            d = Val("&H" & Mid$(url, i, 2))
            If d >= 128 Then
                d = d * 256 + Val("&H" & Mid$(url, i + 3, 2))
                i = i + 5
            Else
                i = i + 2
            End If
            GB_UrlDecode = GB_UrlDecode + Chr$(d)
        Else
            GB_UrlDecode = GB_UrlDecode + c
        End If
    Wend
    UrlDecode = GB_UrlDecode
End Function
'GB2312 URL编码
Public Function UrlEncode(ByRef strURL As String) As String
    Dim i As Long, GB_URLEncode As String
    Dim tempStr As Variant
    For i = 1 To Len(strURL)
        If InStr("-,.0123456789/", Mid(strURL, i, 1)) Then
            GB_URLEncode = GB_URLEncode & Mid(strURL, i, 1)
        Else
            If Asc(Mid(strURL, i, 1)) < 0 Then
                tempStr = "%" & Right(CStr(Hex(Asc(Mid(strURL, i, 1)))), 2)
                tempStr = "%" & Left(CStr(Hex(Asc(Mid(strURL, i, 1)))), Len(CStr(Hex(Asc(Mid(strURL, i, 1))))) - 2) & tempStr
                GB_URLEncode = GB_URLEncode & tempStr
            ElseIf (Asc(Mid(strURL, i, 1)) >= 65 And Asc(Mid(strURL, i, 1)) <= 90) Or (Asc(Mid(strURL, i, 1)) >= 97 And Asc(Mid(strURL, i, 1)) <= 122) Then
                GB_URLEncode = GB_URLEncode & Mid(strURL, i, 1)
            Else
                GB_URLEncode = GB_URLEncode & "%" & Hex(Asc(Mid(strURL, i, 1)))
            End If
        End If
    Next
    UrlEncode = GB_URLEncode
End Function


'————————————————VB URL的编解码源码 GB2312 UTF-8编解码
'
'                            版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。
'
'原文链接：https://blog.csdn.net/gs1069405343/article/details/50471825


Sub test()
    ' Usage
    Dim sUrl As String
    Dim sResult As String
    
    sUrl = "POST /a/%E6%B5%8B%E8%AF%95%20%E7%9A%84%E9%82%93%E4%BC%9F/c?d=123&e=%E5%9B%9B%E7%89%A9"
    sResult = UrlDecodeUtf8(sUrl)
    Debug.Print sResult
End Sub

Public Function AddToQueryString(ByVal url As String, ByVal QS As String) As String
    If InStr(url, "?") > 0 Then
        AddToQueryString = url & "&" & QS
    Else
        AddToQueryString = url & "?" & QS
    End If
End Function

Public Function MakeContent(Dic As Scripting.Dictionary, Optional IsUrlEncode As Boolean = True) As String
    If Dic.Count > 0 Then
        Dim Arr() As String
        ReDim Arr(Dic.Count - 1)
        Dim i As Long, x As Variant, d As String
        For Each x In Dic.Keys()
            d = Dic(x)
            If IsUrlEncode = True Then d = UrlEncodeUtf8(d)
            Arr(i) = x & "=" & d
            i = i + 1
        Next
        MakeContent = Join(Arr, "&")
    End If
End Function

Rem 解析Request内容，即原始键值对
Public Function ParseContent(Content As String, Obj As Scripting.Dictionary, Optional IsUrlDecode As Boolean = True) As Boolean
    Rem 需要增加对数组类型的处理
    Dim a As Variant, b As Variant, k As Variant, v As Variant, x As Variant: a = Split(Content, "&")
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            b = Split(a(x) & "==", "=")
            k = Trim(b(0))
            v = Trim(b(1))
            If IsUrlDecode = True Then v = UrlDecodeUtf8(v)
            If k <> "" Then
                Obj(k) = v
            End If
        Next
    End If
End Function
Rem 解析Request.Header内容，即标头键值对
Public Function ParseKeyValue(Content As String, Obj As Scripting.Dictionary) As Boolean
    Rem 需要增加对数组类型的处理
    Dim a As Variant, b As Variant, k As Variant, v As Variant, x As Variant: a = Split(Content, vbCrLf)
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            b = Split(a(x) & "::", ":")
            k = Trim(b(0))
            v = Trim(b(1))
            '            If k <> "" Then
            Obj(k) = v
            '        End If
        Next
    End If
End Function

Public Function MapMethod(Name As String) As Long
    Dim Arr As Variant: Arr = Array("ANY", "POST", "GET", "PUT", "DELETE", "OPTIONS")
    MapMethod = ToolsArray.GetIndexByValue(Arr, Name)
End Function
Public Function MapMethodName(Index As Long) As String
    Dim Arr As Variant: Arr = Array("ANY", "POST", "GET", "PUT", "DELETE", "OPTIONS")
    MapMethodName = Arr(Index)
End Function

'=========================================================================
' cCollection 与 Form URL-Encoded 格式转换
'=========================================================================

' 将集合对象转换为 application/x-www-form-urlencoded 格式字符串
' @param Obj 集合对象（支持 cCollection 或 Scripting.Dictionary）
' @param IsUrlEncode 是否对值进行 URL 编码（默认 True）
' @return Form URL-Encoded 格式字符串，如 "key1=value1&key2=value2"
' @note 对于 Dictionary 按 Keys 插入顺序遍历；对于 cCollection 按当前顺序遍历
Public Function ToWwwFormUrlencoded(ByRef Obj As Object, Optional ByVal IsUrlEncode As Boolean = True) As String
    Dim Result As String
    Dim Key As Variant
    Dim Value As Variant
    Dim EncodedValue As String
    Dim Keys As Variant
    
    If Obj Is Nothing Then
        ToWwwFormUrlencoded = ""
        Exit Function
    End If
    
    If Obj.Count = 0 Then
        ToWwwFormUrlencoded = ""
        Exit Function
    End If
    
    On Error Resume Next
    
    ' 获取 Keys 数组（cCollection 和 Scripting.Dictionary 都支持 Keys() 方法）
    Keys = Obj.Keys()
    If Err.Number <> 0 Then
        ToWwwFormUrlencoded = ""
        Exit Function
    End If
    
    ' 遍历所有 Keys
    For Each Key In Keys
        ' 获取值（两者都支持 Item(Key)）
        Value = Obj.Item(Key)
        
        ' 对值进行 URL 编码（如果需要）
        If IsUrlEncode Then
            EncodedValue = UrlEncodeUtf8(CStr(Value))
        Else
            EncodedValue = CStr(Value)
        End If
        
        ' 拼接键值对
        If Len(Result) > 0 Then
            Result = Result & "&"
        End If
        Result = Result & Key & "=" & EncodedValue
    Next Key
    
    On Error GoTo 0
    ToWwwFormUrlencoded = Result
End Function

' 将 application/x-www-form-urlencoded 格式字符串解析到集合对象
' @param Content Form URL-Encoded 格式字符串
' @param Obj 集合对象（支持 cCollection 或 Scripting.Dictionary，会被清空后填充）
' @param IsUrlDecode 是否对值进行 URL 解码（默认 True）
' @return 是否解析成功
Public Function FromWwwFormUrlencoded(ByVal Content As String, ByRef Obj As Object, Optional ByVal IsUrlDecode As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    If Obj Is Nothing Then
        FromWwwFormUrlencoded = False
        Exit Function
    End If
    
    ' 清空现有内容
    Obj.RemoveAll
    
    If Trim(Content) = "" Then
        FromWwwFormUrlencoded = True
        Exit Function
    End If
    
    Dim Pairs() As String
    Dim Pair As Variant
    Dim KeyValue() As String
    Dim Key As String
    Dim Value As String
    Dim DecodedValue As String
    
    ' 按 & 分割键值对
    Pairs = Split(Content, "&")
    
    For Each Pair In Pairs
        If Trim(Pair) <> "" Then
            ' 按 = 分割键和值
            KeyValue = Split(Pair & "=", "=")
            Key = Trim(KeyValue(0))
            Value = KeyValue(1)
            
            If Key <> "" Then
                ' 对值进行 URL 解码（如果需要）
                If IsUrlDecode Then
                    DecodedValue = UrlDecodeUtf8(Value)
                Else
                    DecodedValue = Value
                End If
                
                ' 添加到集合（两者都支持 Item(Key)=Value）
                Obj.Item(Key) = DecodedValue
            End If
        End If
    Next Pair
    
    FromWwwFormUrlencoded = True
    Exit Function
    
ErrorHandler:
    FromWwwFormUrlencoded = False
End Function





