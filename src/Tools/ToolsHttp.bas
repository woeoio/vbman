Attribute VB_Name = "ToolsHttp"
Option Explicit

Public Const C_CONTENT_TYPE_FORM_URLENCODED As String = "application/x-www-form-urlencoded"
Public Const C_CONTENT_TYPE_FORM_MULTIPART As String = "multipart/form-data"
Public Const C_CONTENT_TYPE_JSON As String = "application/json"
Public Const C_CONTENT_TYPE_STREAM As String = "application/octet-stream"
Public Const C_CONTENT_TYPE_TEXT_PLAIN As String = "text/plain"
Public Const C_CONTENT_TYPE_TEXT_HTML As String = "text/html"
Public Const C_HEADER_FIELD_CONTENT_TYPE As String = "Content-Type"
Public Const C_HEADER_FIELD_CONTENT_LENGHT As String = "Content-Lenght"


'————————————————VB URL的编解码源码 GB2312 UTF-8编解码
'
'                            版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。
'
'原文链接：https://blog.csdn.net/gs1069405343/article/details/50471825

Public Function UrlDecodeUtf8(ByVal Url As String) As String
    Dim b, ub                                                                   ''中文字的Unicode码(2字节)
    Dim aa, BB
    Dim UtfB                                                                    ''Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3                                                     ''Utf-8码的三个字节
    Dim i, n, S
    Dim str1 As String
    Dim str2 As String
    n = 0
    ub = 0
    For i = 1 To Len(Url)
        b = Mid(Url, i, 1)
        Select Case b
        Case "+"
            S = S & " "
        Case "%"
            ub = Mid(Url, i + 1, 2)
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
                str1 = Mid(Url, i + 4, 2)
                If InStr(str1, vbLf) <= 0 And str1 <> "" Then
                    
                    aa = Mid(str1, 1, 1)
                    BB = Mid(str1, 2, 1)
                    If aa < "g" And aa < "G" And BB < "g" And BB < "G" And aa <> "%" And BB <> "%" Then
                        UtfB2 = (CInt("&H" & str1) And &H3F) * &H40             ''取第2个Utf-8字节的二进制后6位
                    End If
                    
                    str2 = Mid(Url, i + 7, 2)
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
Public Function UrlEncodUtf8(ByVal szInput) As String
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If szInput = "" Then
        UrlEncodUtf8 = szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        
        If nAsc < 0 Then nAsc = nAsc + 65536
        
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
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
    UrlEncodUtf8 = szRet
End Function


'GB2312 URL解码
Public Function UrlDecode(ByVal Url As String) As String
    Dim i As Long, C As String, d As Long, GB_UrlDecode As String
    i = 1
    While i <= Len(Url)
        C = Mid$(Url, i, 1)
        i = i + 1
        If C = "%" Then
            d = Val("&H" & Mid$(Url, i, 2))
            If d >= 128 Then
                d = d * 256 + Val("&H" & Mid$(Url, i + 3, 2))
                i = i + 5
            Else
                i = i + 2
            End If
            GB_UrlDecode = GB_UrlDecode + Chr$(d)
        Else
            GB_UrlDecode = GB_UrlDecode + C
        End If
    Wend
    UrlDecode = GB_UrlDecode
End Function
'GB2312 URL编码
Public Function UrlEncode(ByRef strURL As String) As String
    Dim i, GB_URLEncode As String
    Dim tempStr
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


Sub Test()
    ' Usage
    Dim sUrl As String
    Dim sResult As String
    
    sUrl = "POST /a/%E6%B5%8B%E8%AF%95%20%E7%9A%84%E9%82%93%E4%BC%9F/c?d=123&e=%E5%9B%9B%E7%89%A9"
    sResult = UrlDecodeUtf8(sUrl)
    Debug.Print sResult
End Sub

Public Function AddToQueryString(ByVal Url As String, ByVal QS As String) As String
    If InStr(Url, "?") > 0 Then
        AddToQueryString = Url & "&ver=" & QS
    Else
        AddToQueryString = Url & "?ver=" & QS
    End If
End Function

Public Function MakeContent(Dic As Scripting.Dictionary) As String
    If Dic.Count > 0 Then
        Dim Arr() As String
        ReDim Arr(Dic.Count - 1)
        Dim i As Long, x As Variant
        For Each x In Dic.Keys()
            Arr(i) = x & "=" & Dic(x)
            i = i + 1
        Next
        MakeContent = Join(Arr, "&")
    End If
End Function

Rem 解析Request内容，即原始键值对
Public Function ParseContent(Content As String, Obj As Scripting.Dictionary) As Boolean
    Rem 需要增加对数组类型的处理
    Dim a, b, k, v, x: a = Split(Content, "&")
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            b = Split(a(x) & "==", "=")
            k = Trim(b(0))
            v = Trim(b(1))
            If k <> "" Then
                Obj.Add k, v
            End If
        Next
    End If
End Function
Rem 解析Request.Header内容，即标头键值对
Public Function ParseKeyValue(Content As String, Obj As Scripting.Dictionary) As Boolean
    Rem 需要增加对数组类型的处理
    Dim a, b, k, v, x: a = Split(Content, vbCrLf)
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            b = Split(a(x) & "::", ":")
            k = Trim(b(0))
            v = Trim(b(1))
            '            If k <> "" Then
            Obj.Add k, v
            '        End If
        Next
    End If
End Function

Public Function MapMethod(Name As String) As Long
    Dim Arr: Arr = Array("ANY", "POST", "GET", "PUT", "DELETE", "OPTIONS")
    MapMethod = ToolsArray.GetIndexByValue(Arr, Name)
End Function
Public Function MapMethodName(Index As Long) As String
    Dim Arr: Arr = Array("ANY", "POST", "GET", "PUT", "DELETE", "OPTIONS")
    MapMethodName = Arr(Index)
End Function


