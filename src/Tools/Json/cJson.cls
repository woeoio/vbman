VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'作者: 邓伟, QQ:215879458

Public RootText As String
Public RootItem As New Scripting.Dictionary
Public RootItems As New Collection
'    【20241014】
'           由于 ToolsJsonVba.ConvertToJson 函数对数字类型值没有判断
'           导致返回的数字值可能超过 js 可识别长度, 无法正常使用,
'           因此应该修改源代码, 在 数字类型 的处理过程中把溢出值转字符串

Dim AllNewItem As New Scripting.Dictionary

Public WhiteSpaceSet As Variant

Dim Fso As New Scripting.FileSystemObject

Public LastError As String
Public Property Get LastSuccess() As Boolean
    LastSuccess = LastError <> ""
End Property

'---------------------------------Property
Public Property Get Root() As Object
    If RootItems.Count > 0 Then Set Root = RootItems: Exit Property
    Set Root = RootItem
End Property

Public Property Get Item(ByVal Key As String) As Variant
    If RootItem.Exists(Key) = False Then Exit Property
    If IsObject(RootItem(Key)) = True Then Set Item = RootItem(Key): Exit Property
    Item = RootItem(Key)
End Property
Public Property Get Items(ByVal Index As Long) As Variant
    If RootItems.Count < Index Then Exit Property
    If IsObject(RootItems(Index)) = True Then Set Items = RootItems(Index): Exit Property
    Items = RootItems(Index)
End Property
Public Property Set Item(ByVal Key As String, Dat As Variant)
    Set RootItem(Key) = Dat
End Property
Public Property Set Items(ByVal Index As Long, Dat As Variant)
    Set RootItems(Index) = Dat
End Property
Public Property Let Item(ByVal Key As String, Dat As Variant)
    RootItem(Key) = Dat
End Property
Public Property Let uuItems(ByVal Index As Long, Dat As Variant)
    RootItems(Index) = Dat
End Property

'--------------------------------Property end

Public Function NewItem(Optional ParentKey As String) As cJson
    Set NewItem = New cJson
    If ParentKey = "" Then
        RootItems.Add NewItem.RootItem
    Else
        Set RootItem(ParentKey) = NewItem.RootItem
    End If
End Function
Public Function NewItems(Optional ParentKey As String) As cJson
    Set NewItems = New cJson
    If ParentKey = "" Then
        RootItems.Add NewItems.RootItems
    Else
        Set RootItem(ParentKey) = NewItems.RootItems
    End If
End Function
'Public Function PushNewItemTo(Name As String, Obj As Object, Optional ByVal Key As String = "") As Boolean
'    'if Obj is Dictionary then have Key, or else is Collection
'    If Key = "" Then
'        Obj.Add AllNewItem.RootItem(Name)
'    Else
'        Set Obj(Key) = AllNewItem.RootItem(Name)
'    End If
'    Set AllNewItem.RootItem(Name) = Nothing
'    AllNewItem.Remove Name
'    PushNewItemTo = True
'End Function

Public Function Encode(Optional Obj As Variant, Optional Whitespace As Variant) As String
    ' [using]: ToolsJsonVba.ConvertToJson from
    ' VBA-JSON v2.3.1
    ' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
    LastError = ""
    If IsMissing(Obj) = True Then
        If RootItems.Count > 0 Then Set Obj = RootItems
        If RootItem.Count > 0 Then Set Obj = RootItem
    End If
    If IsMissing(Obj) = True Then GoTo RET
    If IsObject(Obj) = False Then RootText = Obj: GoTo RET
    '    If TypeOf Obj Is ADODB.Recordset Then
    '        Set Obj = ToolsList.RsToCollection(Obj)
    '    End If
    If TypeOf Obj Is cJson Then
        Set Obj = Obj.Root
    End If
    If Obj.Count = 0 Then GoTo RET
    If IsEmpty(WhiteSpaceSet) = False Then
        If IsMissing(Whitespace) = True Then Whitespace = WhiteSpaceSet
    End If
    RootText = ConvertToJson(Obj, Whitespace)
    Set RootItem = Nothing
    Set RootItems = Nothing
    AllNewItem.RemoveAll
RET:
    Encode = RootText
End Function

Public Function Decode(ByRef Text As String) As Boolean
    ' [using]: ToolsJsonVba.ParseJson from
    ' VBA-JSON v2.3.1
    ' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
    On Error GoTo EH
    LastError = ""
    RootText = Trim(Text)
    If Left(RootText, 1) = "[" Then
        Set RootItems = ParseJson(RootText)
    Else
        Set RootItem = ParseJson(RootText)
    End If
    Decode = True
    Exit Function
EH:
    LastError = "[DECODE ERROR]" & ERR.Number & "#" & ERR.Description & "#" & Text
End Function


Public Function LoadFrom(ByVal Path As String) As Boolean
    Dim TS As Scripting.TextStream
    Set TS = Fso.OpenTextFile(Path)
    Dim Txt As String: Txt = TS.ReadAll()
    LoadFrom = Decode(Txt)
    TS.Close
End Function
'
Public Function SaveTo(ByVal FileName As String, Optional Whitespace As Variant) As Variant
    FileName = ToolsFso.AutoCompleteFullPath(FileName, True)
    Call ToolsFso.AutoMakeDir(FileName, True)
    Dim TS As Scripting.TextStream
    Set TS = Fso.OpenTextFile(FileName, ForWriting, True)
    TS.Write Encode(Root, Whitespace)
    TS.Close
End Function

Public Function ToArray() As Variant
    If RootItems.Count > 0 Then
        
    End If
    If RootItem.Count > 0 Then ToArray = Array(RootItem)
End Function


Private Sub Class_Initialize()
    AllNewItem.CompareMode = TextCompare
End Sub

