Attribute VB_Name = "Demo"
Option Explicit

Public Sub Test()
    '    Debug.Print VBMAN.Version()
    '    Debug.Print VBMAN.Version(App)
    '    Debug.Print VBMAN.Csv.LoadFrom("D:\code\上海vb\csv处理器\data\1.csv").Value(4, 2)
    '    Debug.Print VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/home/hello").ReturnText()
End Sub


Public Sub Fetch()
    Dim c As New cHttpClient
    c.DebugStart = True
    c.SetCookies "a=1"
    Debug.Print c.Fetch(ReqGet, "http://a-vi.com/home/hello?name=邓伟").ReturnJson().Encode(, 2, True)
    Debug.Print c.DebugInfo.Encode(, 2, True)
End Sub


Public Sub yyy()
    Dim Json As Object
    Set Json = ToolsJsonVba.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    Json("c")("e") = 789
    
    Debug.Print ToolsJsonVba.ConvertToJson(Json)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print ToolsJsonVba.ConvertToJson(Json, Whitespace:=2)
    ' -> "{
    '       "a": 123,
    '       "b": [
    '         1,
    '         2,
    '         3,
    '         4
    '       ],
    '       "c": {
    '         "d": 456,
    '         "e": 789
    '       }
    '     }"
End Sub
'Public Sub jj()
'    Dim vJson As Variant
'
'    JsonParse "{""num"":123,""test"":[""den"",""wei"",""这是直接赋值""]}", vJson
'    Debug.Print vJson("test")(1)
'    Debug.Print JsonDump(vJson)
'End Sub

'Public Sub tJson(Optional acc As Long)
'
'    '注意,解析数组到集合之后, 下标是从 1  开始的
'    Dim Json As New cJson, sss As String
'    With Json
'        .Item.Add "num", 123
'        '
'        With .NewItems("one")
'            .Add "den"
'            .Add "wei"
'        End With
'        .PushNewItemTo "one", .Item, "test"
'        '
'        .Item("test").Add "这是直接赋值"
'
'        'show
'        '.WhiteSpaceSet = 2
'        sss = .Encode(.Item)
'    End With
'    Debug.Print sss
'    '    Dim a: ToolsJsonUcs.JsonParse sss, a
'    '    Dim a: Set a = ToolsJsonVba.ParseJson(sss)
'    '    Debug.Print a(1)
'    '    Debug.Print a("test")(3)
'    '    Dim newJson As New cJson
'    '    newJson.Decode sss
'    '    Debug.Print newJson.Items(1)
'End Sub

'Sub qiantao()
'
'    Dim Data As New VBMAN.cJson, i As Long
'    With Data.Item
'        .Item("ReportTime") = Format(Date, "yyyy-mm-dd")
'        .Item("Sample model") = "strSN"
'        .Item("Product model") = "StrProduct"
'        .Item("Test number") = " strNumTest"
'        For i = 1 To 2
'            'json对象内置了新建 对象 NewItem(99)  和 数组  NewItmes(99)  的变量数组
'            With Data.NewItem("Lv1")
'                .Item("Test chart number") = "ExcelstrTestNum(i)"
'                .Item("T1") = "ExcelstrT1CH2(i)"
'                .Item("T2") = "ExcelstrT2CH2(i)"
'                .Item("Tp") = "ExcelstrTp(i)"
'                .Item("Tc") = "ExcelstrTc(i)"
'                .Item("S") = "ExcelstrS(i)"
'                With Data.NewItem("Lv2")
'                    .Item("newobj") = 123
'                    .Item("newobj2") = 456
'                End With
'                Data.PushNewItemTo "Lv2", Data.NewItem("Lv1"), "TTT"
'                .Item("IpMax") = "ExcelstrIpMax(i)"
'                .Item("Ipin") = "ExcelstrIpMin(i)"
'                .Item("UpMax") = "ExcelstrUpMax(i)"
'                .Item("UpMin") = "ExcelstrUpMin(i)"
'            End With                                                            'NewItem
'            Data.PushNewItemTo "Lv1", Data.NewItems("Arr1")                     '把新对象push到新数组
'        Next i
'        '把新数组push到根节点的 ExcelstrIntNum 下
'        Data.PushNewItemTo "Arr1", Data.Item, "ExcelstrIntNum"
'    End With
'    Debug.Print Data.Encode(Data.Item)
'End Sub


