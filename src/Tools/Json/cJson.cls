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

'    引用了Microsoft Scripting Runtime
'创建了Microsoft VBScript Regular Expressions 5.5
    
'修改：邓伟
'    【20190702】
'            加入字典复制函数，便于数组层的字典赋值的时候拷贝副本
'            内置字典对象 Dic ，便于外部直接使用。
Public Item As New Scripting.Dictionary
Attribute Item.VB_VarUserMemId = 0
Public Items As New Collection

'    【20241014】
'           由于 ToolsJsonVba.ConvertToJson 函数对数字类型值没有判断
'           导致返回的数字值可能超过 js 可识别长度, 无法正常使用,
'           因此应该修改源代码, 在 数字类型 的处理过程中把溢出值转字符串

Dim AllNewItem As New Scripting.Dictionary

Public WhiteSpaceSet As Variant

Dim Fso As New Scripting.FileSystemObject

'---------------------------------Property
'Public Property Get NewItem(ByVal Index As Long, ByVal Key As String) As Scripting.Dictionary
'    Set NewItem = pvNewItem(Index)
'End Property
'Public Property Get NewItems(ByVal Index As Long, ByVal Key) As Collection
'    Set NewItems = pvNewItems(Index)
'End Property
'Public Property Set NewItem(ByVal Index As Long, ByVal Key As String, Value As Variant)
'    Set pvNewItem(Index).Item(Key) = Value
'End Property
'Public Property Set NewItems(ByVal Index As Long, Value As Variant)
'    Set pvNewItems(Index) = Value
'End Property
'Public Property Let NewItem(ByVal Index As Long, ByVal Key As String, Value As Variant)
'    pvNewItem(Index).Item(Key) = Value
'End Property
'Public Property Let NewItems(ByVal Index As Long, ByVal Key As String, Value As Variant)
'    pvNewItems(Index).Item = Value
'End Property

'--------------------------------Property end
Public Function NewItem(Name As String) As Scripting.Dictionary
    If AllNewItem.Exists(Name) = False Then
        AllNewItem.Add Name, New Scripting.Dictionary
    End If
    Set NewItem = AllNewItem.Item(Name)
End Function
Public Function NewItems(Name As String) As Collection
    If AllNewItem.Exists(Name) = False Then
        AllNewItem.Add Name, New Collection
    End If
    Set NewItems = AllNewItem(Name)
End Function
Public Function PushNewItemTo(Name As String, Obj As Object, Optional ByVal Key As String = "") As Boolean
    'by邓伟，带 key 参数表示塞入到字典, 否则塞入到集合
    If Key = "" Then
        Obj.Add AllNewItem.Item(Name)
    Else
        Set Obj(Key) = AllNewItem.Item(Name)
    End If
    Set AllNewItem.Item(Name) = Nothing
    AllNewItem.Remove Name
    PushNewItemTo = True
End Function

Public Function Encode(Optional Obj As Variant, Optional Whitespace As Variant) As String
    If IsMissing(Obj) Then
        If Items.Count > 0 Then Set Obj = Items
        If Item.Count > 0 Then Set Obj = Item
    End If
    If IsObject(Obj) = False Then Encode = Obj: Exit Function
    '判断是否 Recordset 对象，是就先转换为集合
    If TypeOf Obj Is ADODB.Recordset Then
        Set Obj = ToolsList.RsToCollection(Obj)
    End If
    If IsEmpty(WhiteSpaceSet) = False Then
        If IsMissing(Whitespace) = True Then Whitespace = WhiteSpaceSet
    End If
    Encode = ToolsJsonVba.ConvertToJson(Obj, Whitespace)
    Set Item = Nothing
    Set Items = Nothing
    AllNewItem.RemoveAll
End Function

Public Function Decode(ByRef Text As String) As Boolean                         'woeoio#20220112 修复数组json对象
    Text = Trim(Text)
    If Left(Text, 1) = "[" Then
        Set Items = ToolsJsonVba.ParseJson(Text)
        '        If Items.Count > 0 Then
        '            Set Item = Items(1)
        '        End If
    Else
        Set Item = ToolsJsonVba.ParseJson(Text)
    End If
    Decode = True
End Function


Public Function LoadFrom(ByVal Path As String) As Boolean
    'by邓伟
    Dim TS As Scripting.TextStream
    Set TS = Fso.OpenTextFile(Path)
    Dim Txt As String: Txt = TS.ReadAll()
    LoadFrom = Decode(Txt)
    TS.Close
End Function
'
Public Function SaveTo(ByVal FileName As String, Optional IsJsonArray As Boolean) As Variant
    'by邓伟
    Dim TS As Scripting.TextStream
    Set TS = Fso.OpenTextFile(FileName, ForWriting, True)
    If IsJsonArray = True Then
        TS.Write Encode(Items)
    Else
        TS.Write Encode(Item)
    End If
    TS.Close
End Function

Public Function ToArray() As Variant
    If Items.Count > 0 Then
        
    End If
    If Item.Count > 0 Then ToArray = Array(Item)
End Function


'Public Function NewJsonObject() As Scripting.Dictionary
'    'by邓伟，
'    Set pvNewItem = New Scripting.Dictionary
'    Set NewJsonObject = pvNewItem
'End Function
'Public Function NewJsonArray() As Collection
'    'by邓伟，
'    Set pvNewItems = New Collection
'    Set NewJsonArray = pvNewItems
'End Function

'Public Function pushDic(srcDic As Scripting.Dictionary, dstDic As Scripting.Dictionary, dstDicArrayKey$) As Boolean
'    'by邓伟，push字典对象到某个数组字典里面，成为子级
'    Dim Arr: Arr = dstDic(dstDicArrayKey)
'    Dim Max&: Max = UBound(Arr) + 1
'    '    If Max > 0 Then Max = Max + 1
'    ReDim Preserve Arr(Max)
'    Set Arr(Max) = srcDic
'    dstDic(dstDicArrayKey) = Arr
'    '    Set srcDic = Nothing
'    pushDic = True
'End Function
'
'Public Function copyDic(srcDic As Scripting.Dictionary, dstDic As Scripting.Dictionary) As Boolean
'    'by邓伟，复制字典对象
'    Set dstDic = dicCopy(srcDic)
'    copyDic = True
'End Function
'
'Private Function dicCopy(ByVal Val) As Scripting.Dictionary
'    'by邓伟，通过递归方法实现的字典副本
'    Dim objDic As New Scripting.Dictionary
'    If TypeName(Val) = "Dictionary" Then
'        For Each x In Val
'            If TypeName(Val(x)) = "Dictionary" Then
'                objDic.Add x, dicCopy(Val(x))
'            Else
'                objDic.Add x, Val(x)
'            End If
'        Next
'    End If
'    Set dicCopy = objDic
'End Function

'Public Function jsDecode(ByVal JSONPath As String, ByVal JSONString As String, Optional isObj As Boolean) As Variant
'    'by邓伟，基于js脚本解析器实现的json解析，传入需要访问的json路径，返回json节点值。路径语法同js
'    '    Dim a, b, j As New cJson
'    '    a = "{""a"":{""b"":1}}"
'    '    Set b = j.jsDecode("a", a, True)
'    '    MsgBox b.b
'    On Error GoTo eee
'    Dim Json As Object
'    Set Json = CreateObject("MSScriptControl.ScriptControl")
'    Json.Language = "JScript"
'    If isObj = True Then
'        If JSONPath = "" Then
'            Set jsDecode = Json.Eval(JSONString)
'        Else
'            Set jsDecode = Json.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
'        End If
'    Else
'        jsDecode = Json.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
'    End If
'    Set Json = Nothing
'    Exit Function
'eee:
'    jsDecode = "vbNull"
'End Function

'Public Function jsArrayItem(Arr, Index, Optional isObj As Boolean) As Variant
'    'by 邓伟，由于vb不支持 . 符号直接读取数字索引的数组对象，因此使用该函数来取得
'    If isObj = True Then
'        Set jsArrayItem = CallByName(Arr, Index, VbGet)
'    Else
'        jsArrayItem = CallByName(Arr, Index, VbGet)
'    End If
'End Function


'Private Sub Command1_Click()
'    Dim json As New cJson
'    Dim s$: s = "{""a"": {""arr"": [1,2,3,4,5] }, ""obj"": {""b"": 1} }"
'    Dim Arr: Set Arr = json.jsDecode("a.arr", s, True)
'
'    Dim b&: b = json.jsDecode("obj.b", s, False)
'    Dim obj: Set obj = json.jsDecode("obj", s, True)
'
'    Dim x
'    For Each x In Arr
'        Debug.Print x
'    Next
'
'End Sub

'---------------------------------------------------------------------------

Private Sub Class_Initialize()
    '    MaxNewItemLength = 99
    '    ReDim pvNewItem(MaxNewItemLength)
    '    ReDim pvNewItems(MaxNewItemLength)
    AllNewItem.CompareMode = TextCompare
End Sub
