VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpCrossDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'    Request.Response.AddHeaderEntry "Access-Control-Allow-Origin", "*"
'    Request.Response.AddHeaderEntry "Access-Control-Allow-Methods", "*"
'    Request.Response.AddHeaderEntry "Access-Control-Allow-Headers", "*"
'    Request.Response.AddHeaderEntry "Access-Control-Allow-Credentials", "true"
'    Request.Response.AddHeaderEntry "Access-Control-Max-Age", 86400000

'Dim DicOrigin As New Scripting.Dictionary
'Dim DicMethods As New Scripting.Dictionary
'Dim DicHeaders As New Scripting.Dictionary
'Dim DicCredentials As New Scripting.Dictionary

Public AllowOrigin As String
Public AllowMethods As String
Public AllowHeaders As String
Public AllowCredentials As Boolean
Public MaxAge As Long

Dim IsEnable As Boolean
Dim IsInit As Boolean

Public Property Get Enable() As Boolean
    Enable = IsEnable
End Property
Public Property Let Enable(v As Boolean)
    If v = True Then
        If IsInit = False Then
            If AllowOrigin = "" Then AllowOrigin = "*"
            If AllowMethods = "" Then AllowMethods = "*"
            If AllowHeaders = "" Then AllowHeaders = "*"
            If AllowCredentials = False Then AllowCredentials = True
            If MaxAge = 0 Then MaxAge = 86400 * 1000
            IsInit = True
        End If
    End If
    IsEnable = v
End Property

Friend Function AddTo(Headers As Scripting.Dictionary) As Boolean
    If Enable = False Then Exit Function
    Headers("Access-Control-Allow-Origin") = AllowOrigin
    Headers("Access-Control-Allow-Methods") = AllowMethods
    Headers("Access-Control-Allow-Headers") = AllowHeaders
    Headers("Access-Control-Allow-Credentials") = CStr(AllowCredentials)
    Headers("Access-Control-Allow-Max-Age") = MaxAge
End Function
'
'Public Function Use( _
'    Optional AllowOrigin As String = "*", _
'    Optional AllowMethods As String = "*", _
'    Optional AllowHeaders As String = "*", _
'    Optional AllowCredentials As Boolean = True, _
'    Optional MaxAge As String = 86400 * 1000 _
'    ) As Boolean
'    With AccessControl
'        .AllowOrigin = AllowOrigin
'        .AllowMethods = AllowMethods
'        .AllowHeaders = AllowHeaders
'        .AllowCredentials = AllowCredentials
'        .MaxAge = MaxAge
'    End With
'    Enable = True
'    Use = Enable
'End Function

'Private Sub Class_Initialize()
'    MaxAge = 86400 * 1000
'End Sub
