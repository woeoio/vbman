VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 响应 Cookie 存储（用于生成 Set-Cookie 头）
' 需要在工程中引用 Microsoft Scripting Runtime
Dim Data As New Scripting.Dictionary

' 请求 Cookie 存储（来自 Cookie 头，只读键值对）
Private m_RequestCookies As New Scripting.Dictionary

' 获取或创建响应 Cookie（用于 Set-Cookie）
Public Property Get Cookie(ByVal Key As String) As cHttpServerCookieAttr
Attribute Cookie.VB_UserMemId = 0
    If Data.Exists(Key) = False Then
        Dim CK As New cHttpServerCookieAttr
        Set Cookie = CK
        Data.Add Key, CK
        Set CK = Nothing
    Else
        Set Cookie = Data(Key)
    End If
End Property

' 读取请求 Cookie 的值（浏览器发送的 Cookie）
Public Function GetValue(ByVal Key As String) As String
    If m_RequestCookies.Exists(Key) Then
        GetValue = m_RequestCookies(Key)
    End If
End Function

' 检查 Cookie 是否存在（同时检查请求和响应）
Public Function Exists(ByVal Key As String) As Boolean
    Exists = Data.Exists(Key) Or m_RequestCookies.Exists(Key)
End Function

' 检查请求 Cookie 是否存在（浏览器发送的）
Public Function RequestExists(ByVal Key As String) As Boolean
    RequestExists = m_RequestCookies.Exists(Key)
End Function

' 检查响应 Cookie 是否存在（已设置 Set-Cookie 的）
Public Function ResponseExists(ByVal Key As String) As Boolean
    ResponseExists = Data.Exists(Key)
End Function

' 删除指定的响应 Cookie
Public Sub Remove(ByVal Key As String)
    If Data.Exists(Key) Then
        Data.Remove Key
    End If
End Sub

' 通知浏览器删除 Cookie（设置过期时间为过去）
Public Sub ExpireCookie(ByVal Key As String)
    With Cookie(Key)
        .Value = ""
        .Expires = DateAdd("d", -1, Now)
        .Path = "/"
    End With
End Sub

' 清空所有响应 Cookie
Public Sub Clear()
    Data.RemoveAll
End Sub

' 从 Cookie 头字符串解析请求 Cookie
' 用于解析请求头中的 Cookie: name1=value1; name2=value2
' 解析结果存入 m_RequestCookies，不影响响应 Cookie（Data）
Public Function Decode(ByVal cookieString As String) As Boolean
    On Error GoTo EH
    Dim keyValuePairs() As String
    Dim pair As Variant
    Dim keyValue() As String
    Dim Key As String
    Dim value As String
    
    m_RequestCookies.RemoveAll
    If cookieString = "" Then
        Decode = True
        Exit Function
    End If
    
    ' 用分号分隔各个键值对
    keyValuePairs = Split(cookieString, ";")
    
    For Each pair In keyValuePairs
        ' 去除前后空格
        pair = Trim(pair)
        If pair <> "" Then
            ' 用等号分隔键值
            keyValue = Split(pair & "=", "=", 2)
            
            If UBound(keyValue) >= 1 Then
                Key = Trim(keyValue(0))
                value = Trim(keyValue(1))
                ' URL 解码
                value = ToolsHttp.UrlDecodeUtf8(value)
                ' 添加到请求字典
                If Key <> "" Then
                    m_RequestCookies(Key) = value
                End If
            End If
        End If
    Next
    
    Decode = True
    Exit Function
EH:
    Decode = False
End Function

' 将响应 Cookie 编码为 Set-Cookie 头字符串
' 每个 Cookie 单独一行，每行对应一个 Set-Cookie 头
' 仅输出 Data（响应 Cookie），不包含 m_RequestCookies（请求 Cookie）
Public Function Encode() As String
    Dim cookieString As String
    Dim Key As Variant
    Dim CK As cHttpServerCookieAttr
    
    cookieString = ""
    
    ' 仅遍历响应 Cookie（Data），不包含请求 Cookie
    For Each Key In Data.Keys
        Set CK = Data(Key)
        If cookieString <> "" Then
            cookieString = cookieString & vbCrLf
        End If
        ' 基本格式：name=value
        cookieString = cookieString & Key & "=" & ToolsHttp.UrlEncodeUtf8(CK.Value)
        
        ' 添加路径
        If CK.Path <> "" Then
            cookieString = cookieString & "; Path=" & CK.Path
        End If
        
        ' 添加域名
        If CK.Domain <> "" Then
            cookieString = cookieString & "; Domain=" & CK.Domain
        End If
        
        ' 添加过期时间
        If Not IsEmpty(CK.Expires) Then
            cookieString = cookieString & "; Expires=" & FormatHttpDate(CK.Expires)
        End If
        
        ' 添加 Secure 标志
        If CK.Secure = True Then
            cookieString = cookieString & "; Secure"
        End If
        
        ' 添加 HttpOnly 标志
        If CK.HttpOnly = True Then
            cookieString = cookieString & "; HttpOnly"
        End If
        
        ' 添加 SameSite 属性
        If CK.SameSite <> "" Then
            cookieString = cookieString & "; SameSite=" & CK.SameSite
        End If
    Next
    
    Encode = cookieString
End Function

' 格式化 HTTP 日期（RFC 1123 格式）
Private Function FormatHttpDate(dt As Variant) As String
    Dim dtDate As Date
    Dim sWeekDay As String, sMonth As String
    Dim arrWeekDay As Variant, arrMonth As Variant
    
    arrWeekDay = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    arrMonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    dtDate = CDate(dt)
    sWeekDay = arrWeekDay(Weekday(dtDate, vbSunday) - 1)
    sMonth = arrMonth(Month(dtDate) - 1)
    
    FormatHttpDate = sWeekDay & ", " & _
                     Right("0" & Day(dtDate), 2) & " " & _
                     sMonth & " " & _
                     Year(dtDate) & " " & _
                     Right("0" & Hour(dtDate), 2) & ":" & _
                     Right("0" & Minute(dtDate), 2) & ":" & _
                     Right("0" & Second(dtDate), 2) & " GMT"
End Function

Private Sub Class_Initialize()
    Data.CompareMode = TextCompare
    m_RequestCookies.CompareMode = TextCompare
End Sub
