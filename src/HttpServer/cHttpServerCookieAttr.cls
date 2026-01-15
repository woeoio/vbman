VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerCookieAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Cookie 值
Private pvValue As String
' Cookie 域名
Private pvDomain As String
' Cookie 路径
Private pvPath As String
' Cookie 过期时间
Private pvExpires As Variant
' 是否仅通过 HTTPS 传输
Private pvSecure As Boolean
' 是否禁止 JavaScript 访问
Private pvHttpOnly As Boolean
' SameSite 属性 (Strict, Lax, None)
Private pvSameSite As String

' Cookie 值
Public Property Get Value() As String
    Value = pvValue
End Property

Public Property Let Value(v As String)
    pvValue = v
End Property

' Cookie 域名
Public Property Get Domain() As String
    Domain = pvDomain
End Property

Public Property Let Domain(v As String)
    pvDomain = v
End Property

' Cookie 路径
Public Property Get Path() As String
    Path = pvPath
End Property

Public Property Let Path(v As String)
    If v = "" Then
        pvPath = "/"
    Else
        pvPath = v
    End If
End Property

' Cookie 过期时间
Public Property Get Expires() As Variant
    Expires = pvExpires
End Property

Public Property Let Expires(v As Variant)
    pvExpires = v
End Property

' 是否仅通过 HTTPS 传输
Public Property Get Secure() As Boolean
    Secure = pvSecure
End Property

Public Property Let Secure(v As Boolean)
    pvSecure = v
End Property

' 是否禁止 JavaScript 访问
Public Property Get HttpOnly() As Boolean
    HttpOnly = pvHttpOnly
End Property

Public Property Let HttpOnly(v As Boolean)
    pvHttpOnly = v
End Property

' SameSite 属性
Public Property Get SameSite() As String
    SameSite = pvSameSite
End Property

Public Property Let SameSite(v As String)
    pvSameSite = v
End Property

Private Sub Class_Initialize()
    pvPath = "/"
    pvHttpOnly = False
    pvSecure = False
    pvSameSite = ""
End Sub
