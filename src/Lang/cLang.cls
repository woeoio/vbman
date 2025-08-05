VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'作者：邓伟，邮箱：215879458@qq.com

'使用 json 作为多国语言文件，在开发阶段直接从 json文件加载，方便修改调试，编译之后从 资源文件加载
'制作了 cLang.cls 类对象，让整个语言切换可以做到实时生效，并记住用户选择的语言，下次启动默认该语言。
'最终的语言切换，对项目代码的侵入性非常少，只需要在 formload 事件里面 lang.render me 渲染一次即可。
'反正最终多国语言都是直接拿  1033.json 文件丢给 ai 翻译就行了
'除了一次性渲染界面，还支持 Lang("Fmain.HelpAbout.Content")  这样的 单个 语言值 调用
'最大化的保留了语义，，而不是用什么资源 id  101  ，看到懵逼
'这个方案比资源文件好太多了
'编译前，是直接加载的 json 文件，方便自己实时修改


Public Type TypeLang
    ID As Long
    Code As String
    Title As String
End Type

Public Enum EnumLang
    LANG_ID_EN = 1033
    LANG_ID_ZH_CH = 2052
End Enum

Public LangData As New cJson
Public LangInfo As New cJson

Public CurLangID As Long
Public CurLangIndex As Long


Public Function Selector(Items As Object, Index As Integer) As Boolean
    Dim M As Menu: Set M = Items(Index)
    Dim ID As Long: ID = LangInfo("Map")("IndexToID")(Index + 1)
    '
    LoadData ID
    SaveSetting "VBMAN", "Lang", "CurLangIndex", Index
    '
    Dim i As Long
    For i = 0 To Items.Count - 1
        Items(i).Checked = False
    Next
    M.Checked = True
    '
    Render M.Parent
End Function

Public Function Render(Frm As Object) As Boolean
    Dim O As Form: Set O = Frm
    Dim Ctl As Variant, Key As String, LN As String
    On Error Resume Next
    LN = Item(O.Name & ".Caption")
    If LN <> "" Then O.Caption = LN
    For Each Ctl In O
        Key = O.Name & "." & Ctl.Name
        If IsControlArray(Ctl.Index) = True Then
            Key = Key & "." & Ctl.Index
        End If
        Key = Key & ".Caption"
        LN = Item(Key)
        If LN <> "" Then Ctl.Caption = LN
        'Debug.Print """" & Key & """: """ & Ctl.Caption & """, "
    Next
End Function

Public Function Item(Key As String) As String
Attribute Item.VB_UserMemId = 0
    Item = LangData.Item(Key)
End Function

Private Function LoadData(ID As Long) As Boolean
    If IsIDE = True Then
        LangData.LoadFrom App.Path & "\lang\data\" & ID & ".json", "GBK"
    Else
        Dim D() As Byte: D = LoadResData(ID, "CUSTOM")
        Dim S As String: S = StrConv(D, vbUnicode)
        LangData.Decode S
    End If
    LoadData = True
End Function

Private Sub Class_Initialize()
    Call LoadInfo
    CurLangIndex = GetSetting("VBMAN", "Lang", "CurLangIndex", 0)
    CurLangID = LangInfo("Map")("IndexToID")(CurLangIndex + 1)
    Call LoadData(CurLangID)
End Sub


Private Function LoadInfo() As Boolean
    If IsIDE = True Then
        LangInfo.LoadFrom App.Path & "\lang\info.json", "GBK"
    Else
        Dim D() As Byte: D = LoadResData(6000, "CUSTOM")
        Dim S As String: S = StrConv(D, vbUnicode)
        LangInfo.Decode S
    End If
    LoadInfo = True
End Function
