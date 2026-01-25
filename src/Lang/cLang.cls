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

' API声明
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer

Public Type TypeLang
    Index As Long
    ID As String
    Caption As String
End Type

Public Enum EnumLang
    LANG_ID_EN = 1033
    LANG_ID_ZH_CH = 2052
End Enum

Public LangData As New cJson
Public LangInfo As New cJson

Dim pvCurLang As TypeLang
Dim MenuSelector As Object

Public Property Get CurLang() As TypeLang
    CurLang = pvCurLang
End Property

Public Function SetSelector(Items As Object) As Boolean
    Set MenuSelector = Items
    '
    Dim ItemIndex As Long, M As Menu, x As Variant
    For Each x In LangInfo("LangList")
        If ItemIndex = 0 Then
            Set M = Items(0)
        Else
            Load Items(ItemIndex)
            Set M = Items(ItemIndex)
        End If
        M.Caption = x("Caption")
        M.Visible = True
        ItemIndex = ItemIndex + 1
    Next
    '
    If Items.Count > 0 Then
        If Items.Count <= CurLang.Index Then CurLang.Index = 0
        Items(CurLang.Index).Checked = True
    End If
    '
    SetSelector = True
End Function
Public Function Selected(ByVal Index As Long) As Boolean
    Dim M As Menu: Set M = MenuSelector(Index)
    '
    LoadData Index
    SaveSetting "VBMAN", "Lang", "CurLang.Index", Index
    '
    Dim i As Long
    For i = 0 To MenuSelector.Count - 1
        MenuSelector(i).Checked = False
    Next
    M.Checked = True
    '
    Render M.Parent
    Selected = True
End Function

Public Function Render(Frm As Object) As Boolean
    Dim o As Form: Set o = Frm
    Dim Ctl As Variant, Key As String, lN As String
    On Error Resume Next
    lN = Item(o.Name & ".Caption")
    If lN <> "" Then o.Caption = lN
    For Each Ctl In o
        Key = o.Name & "." & Ctl.Name
        If Tools.ToolsArray.IsControlArray(Ctl.Index) = True Then
            Key = Key & "." & Ctl.Index
        End If
        Key = Key & ".Caption"
        lN = Item(Key)
        If lN <> "" Then Ctl.Caption = lN
        'Debug.Print """" & Key & """: """ & Ctl.Caption & """, "
    Next
End Function

Public Function Item(Key As String) As String
Attribute Item.VB_UserMemId = 0
    Item = LangData.Item(Key)
End Function
Private Function LoadData(ByVal Idx As Long) As Boolean
    With LangInfo("LangList")(Idx + 1)
        pvCurLang.Index = Idx
        pvCurLang.ID = .Item("ID")
        pvCurLang.Caption = .Item("Caption")
        '
        If IsIDE = True Then
            LangData.LoadFrom App.Path & "\lang\data\" & pvCurLang.ID & ".json", "GBK"
        Else
            Dim D() As Byte: D = LoadResData(pvCurLang.ID, "LANGUAGE")
            Dim S As String: S = StrConv(D, vbUnicode)
            LangData.Decode S
        End If
    End With
    LoadData = True
End Function

Private Sub Class_Initialize()
    Call LoadInfo
    Dim Idx As Long: Idx = GetSetting("VBMAN", "Lang", "CurLang.Index", -1)
    If Idx = -1 Then
        Dim langID As Integer
        ' 获取系统默认语言ID
        langID = GetSystemDefaultLangID()
        ' 或者使用用户默认语言ID
        ' langID = GetUserDefaultLangID()
        '如果当前不是中文系统，则显示默认语言：en-US
        If langID = EnumLang.LANG_ID_ZH_CH Then
            Idx = 0
        Else
            Idx = LangInfo("DefaultIndex")
        End If
    End If
    Call LoadData(Idx)
End Sub


Private Function LoadInfo() As Boolean
    If IsIDE = True Then
        LangInfo.LoadFrom App.Path & "\lang\info.json", "GBK"
    Else
        Dim D() As Byte: D = LoadResData("-", "LANGUAGE")
        Dim S As String: S = StrConv(D, vbUnicode)
        LangInfo.Decode S
    End If
    LoadInfo = True
End Function
