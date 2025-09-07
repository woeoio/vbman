VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAesCBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Html As New MSHTML.HTMLDocument, JsCode As String

Public Function Encode(Data As Variant, Optional ByVal Password As String = "nzgi2g9gi7odbvxw") As String
    
    With Html
        
        .parentWindow.execScript ("document.title=encryptFn('" & Data & "', '" & Password & "');")
        
        Encode = .Title
        
    End With
End Function

Private Sub Class_Initialize()
    With Html
        
        ' 构建完整的HTML文档，包含你的JavaScript代码
        '        Dim htmlContent As String
        '        JsCode = StrConv(LoadResData("AES.CBC", "JSCRIPT"), vbUnicode)
        '        htmlContent = "<script language='javascript'>" & JsCode & "</script>"
        JsCode = StrConv(LoadResData("AES.CBC", "JSCRIPT"), vbUnicode)
        ' 创建script元素
        Dim scriptElement As MSHTML.IHTMLScriptElement
        Set scriptElement = .createElement("script")
        
        ' 设置script属性
        scriptElement.Type = "text/javascript"
        scriptElement.Text = JsCode
        
        ' 将script添加到head中
        .Head.appendChild scriptElement
        .Close
        
    End With
End Sub
