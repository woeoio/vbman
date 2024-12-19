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


' 引入 Scripting.Dictionary 组件
' 需要在工程中引用 Microsoft Scripting Runtime
Dim Data As New Scripting.Dictionary

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

Public Sub Clear()
    Data.RemoveAll
End Sub
' 解析 Cookie 字符串到字典对象的函数
Public Function Decode(ByVal cookieString As String) As Boolean
    '    Dim cookieDict As Dictionary
    '    Set cookieDict = New Dictionary
    '
    '    Dim keyValuePairs() As String
    '    Dim pair As Variant
    '    Dim keyValue() As String
    '    Dim Key As String
    '    Dim value As String
    '
    '    Data.RemoveAll
    '    ' 按分号分隔各个键值对
    '    keyValuePairs = Split(cookieString, ";")
    '
    '    For Each pair In keyValuePairs
    '        ' 去除前后空格
    '        pair = Trim(pair)
    '        ' 按等号分割键和值
    '        keyValue = Split(pair, "=", 2)
    '
    '        If UBound(keyValue) = 1 Then
    '            Key = Trim(keyValue(0))
    '            value = Trim(keyValue(1))
    '            ' 添加到字典
    '            cookieDict.Add Key, value
    '        End If
    '    Next
    '
    '    Decode = True
End Function

' 从字典对象生成 Cookie 字符串的函数
Public Function Encode() As String
'    Dim cookieString As String
'    Dim Key As Variant
'
'    cookieString = ""
'
'    ' 遍历字典中的键值对
'    For Each Key In Data.Keys
'        If cookieString <> "" Then
'            cookieString = cookieString & "; "
'        End If
'        cookieString = cookieString & Key & "=" & cookieDict(Key)
'    Next
'
'    Encode = cookieString
End Function

Private Sub Class_Initialize()
    Data.CompareMode = TextCompare
End Sub
