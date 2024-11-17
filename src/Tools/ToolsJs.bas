Attribute VB_Name = "ToolsJs"
Option Explicit

'引用：

'Library MSScriptControl
'    C:\Windows\SysWOW64\msscript.ocx
'    Microsoft Script Control 1.0

Dim MSSC As New MSScriptControl.ScriptControl
Dim IsInit As Boolean

Public Function NewArr(ParamArray Item() As Variant) As Object
    Call Init
    Set NewArr = MSSC.Eval("[];")
End Function

Public Function NewObj() As Object
    Call Init
    Set NewArr = MSSC.Eval("{};")
End Function



Private Sub Init()
    If IsInit = True Then Exit Sub
    MSSC.Language = "JScript"
    Rem ====================================兼容ide大小写问题
    Rem 定义一些内部方法，以 _ 结尾，性能有损耗， 尽可能直接存取属性
    MSSC.Eval ("Object.prototype.get_ = function(x) { return this[x]; };")
    MSSC.Eval ("Array.prototype.get_= function(x) { return this[x]; };")
    Rem 解决vbide自动大小写导致的函数错误
    MSSC.Eval ("Array.prototype.Push = function(x) { this.push(x); };")
    MSSC.Eval ("Array.prototype.Join = function(x) { this.join(x); };")
    Rem ====================================兼容ide大小写问题
    
    Rem ====================================对象属性判断
    MSSC.Eval ("Object.prototype.has_ = function(x) { return this.hasOwnProperty(x); };")
    Rem ====================================对象属性判断
    IsInit = True
End Sub
