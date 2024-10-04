Attribute VB_Name = "ToolsHttp"
Option Explicit

Public Enum EnumMethod
    Any_ = 0
    Post = 1
    Get_ = 2
    Put_ = 3
    Delete = 4
    Option_ = 5
    
    
End Enum

'————————————————VB URL的编解码源码 GB2312 UTF-8编解码
'
'                            版权声明：本文为博主原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接和本声明。
'
'原文链接：https://blog.csdn.net/gs1069405343/article/details/50471825

Public Function UrlDecodeUtf8(ByVal URL As String) As String
    Dim B, ub                                                                   ''中文字的Unicode码(2字节)
    Dim AA, BB
    Dim UtfB                                                                    ''Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3                                                     ''Utf-8码的三个字节
    Dim i, n, s
    Dim str1 As String
    Dim str2 As String
    n = 0
    ub = 0
    For i = 1 To Len(URL)
        B = Mid(URL, i, 1)
        Select Case B
        Case "+"
            s = s & " "
        Case "%"
            ub = Mid(URL, i + 1, 2)
            If InStr(ub, vbLf) <= 0 And ub <> "" Then
                AA = Mid(ub, 1, 1)
                BB = Mid(ub, 2, 1)
                If AA < "g" And AA < "G" And BB < "g" And BB < "G" And AA <> "%" And BB <> "%" Then
                    UtfB = CInt("&H" & ub)
                End If
            End If
            
            If UtfB < 128 Then
                i = i + 2
                s = s & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000                                 ''取第1个Utf-8字节的二进制后4位
                str1 = Mid(URL, i + 4, 2)
                If InStr(str1, vbLf) <= 0 And str1 <> "" Then
                    
                    AA = Mid(str1, 1, 1)
                    BB = Mid(str1, 2, 1)
                    If AA < "g" And AA < "G" And BB < "g" And BB < "G" And AA <> "%" And BB <> "%" Then
                        UtfB2 = (CInt("&H" & str1) And &H3F) * &H40             ''取第2个Utf-8字节的二进制后6位
                    End If
                    
                    str2 = Mid(URL, i + 7, 2)
                    If InStr(str2, vbLf) <= 0 And str2 <> "" Then
                        AA = Mid(str2, 1, 1)
                        BB = Mid(str2, 2, 1)
                        If AA < "g" And AA < "G" And BB < "g" And BB < "G" And AA <> "%" And BB <> "%" Then
                            UtfB3 = CInt("&H" & str2) And &H3F                  ''取第3个Utf-8字节的二进制后6位
                        End If
                    End If
                End If
                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                i = i + 8
            End If
            
        Case Else                                                               ''Ascii码
            s = s & B
        End Select
    Next
    UrlDecodeUtf8 = s
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
                uch = "%" & hex(((nAsc \ 2 ^ 6)) Or &HC0) & hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UrlEncodUtf8 = szRet
End Function


'GB2312 URL解码
Public Function UrlDecode(ByVal URL As String) As String
    Dim i As Long, c As String, d As Long, GB_UrlDecode As String
    i = 1
    While i <= Len(URL)
        c = Mid$(URL, i, 1)
        i = i + 1
        If c = "%" Then
            d = Val("&H" & Mid$(URL, i, 2))
            If d >= 128 Then
                d = d * 256 + Val("&H" & Mid$(URL, i + 3, 2))
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
    Dim i, GB_URLEncode As String
    Dim tempStr
    For i = 1 To Len(strURL)
        If InStr("-,.0123456789/", Mid(strURL, i, 1)) Then
            GB_URLEncode = GB_URLEncode & Mid(strURL, i, 1)
        Else
            If Asc(Mid(strURL, i, 1)) < 0 Then
                tempStr = "%" & Right(CStr(hex(Asc(Mid(strURL, i, 1)))), 2)
                tempStr = "%" & Left(CStr(hex(Asc(Mid(strURL, i, 1)))), Len(CStr(hex(Asc(Mid(strURL, i, 1))))) - 2) & tempStr
                GB_URLEncode = GB_URLEncode & tempStr
            ElseIf (Asc(Mid(strURL, i, 1)) >= 65 And Asc(Mid(strURL, i, 1)) <= 90) Or (Asc(Mid(strURL, i, 1)) >= 97 And Asc(Mid(strURL, i, 1)) <= 122) Then
                GB_URLEncode = GB_URLEncode & Mid(strURL, i, 1)
            Else
                GB_URLEncode = GB_URLEncode & "%" & hex(Asc(Mid(strURL, i, 1)))
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



Rem 解析Request内容，即原始键值对
Public Function ParseContent(Content As String, Obj As Scripting.Dictionary) As Boolean
    Rem 需要增加对数组类型的处理
    Dim a, B, k, v, x: a = Split(Content, "&")
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            B = Split(a(x) & "==", "=")
            k = Trim(B(0))
            v = Trim(B(1))
            If k <> "" Then
                Obj.Add k, v
            End If
        Next
    End If
End Function
