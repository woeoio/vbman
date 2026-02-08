Attribute VB_Name = "Demo"

Option Explicit


Public Sub TestHMACChainCall()
    Dim HMAC As New cCryptoHMAC
    Dim sText As String
    sText = "Hello World"
    
    Debug.Print "HMAC Chain Call Demo"
    Debug.Print "====================="
    
    ' 1. 最简单的用法（默认 SHA256）
    Debug.Print "SHA256: " & HMAC.Secret("secret").DataString(sText).ReturnHex()
    
    ' 2. HMAC-SHA1
    Debug.Print "SHA1:   " & HMAC.Mode(HMAC_ALG_SHA1).Secret("secret").DataString(sText).ReturnHex()
    
    ' 3. 不同输出格式
    Debug.Print "Hex:    " & HMAC.Secret("secret").DataString(sText).ReturnHex()
    Debug.Print "Base64: " & HMAC.Secret("secret").DataString(sText).ReturnBase64()
    Debug.Print "Upper:  " & HMAC.Secret("secret").DataString(sText).ReturnHex(True)
    
    ' 4. 中文文本
    Debug.Print "Chinese UTF8: " & HMAC.Secret("密钥").DataString("你好世界", HMAC_ENCODING_UTF8).ReturnHex()
    
    ' 5. 重复使用
    HMAC.Mode (HMAC_ALG_SHA256)
    Debug.Print "Data 1: " & HMAC.Secret("key1").DataString("Hello").ReturnHex()
    Debug.Print "Data 2: " & HMAC.Secret("key2").DataString("World").ReturnHex()
End Sub

'演示链式调用功能
Public Sub DemoChainedCall()
    Debug.Print "=== Chained Call Demo ==="
    Debug.Print ""
    
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Dim sText As String
    sText = "Hello World"
    
    ' 1. 基本：字符串输入 + 十六进制输出（小写）
    Debug.Print "1. String input + Hex output (lowercase):"
    Dim sResult1 As String
    sResult1 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    Debug.Print "   " & sResult1
    Debug.Print ""
    
    ' 2. 字符串输入 + 十六进制输出（大写）
    Debug.Print "2. String input + Hex output (uppercase):"
    Dim sResult2 As String
    sResult2 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex(True)
    Debug.Print "   " & sResult2
    Debug.Print ""
    
    ' 3. 字符串输入 + Base64 输出
    Debug.Print "3. String input + Base64 output:"
    Dim sResult3 As String
    sResult3 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnBase64()
    Debug.Print "   " & sResult3
    Debug.Print ""
    
    ' 4. 字符串输入 + 字节数组输出
    Debug.Print "4. String input + Bytes output:"
    Dim baResult() As Byte
    baResult = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnBytes()
    Debug.Print "   Hash size: " & (UBound(baResult) + 1) & " bytes"
    Debug.Print ""
    
    ' 5. 字节数组输入 + 十六进制输出
    Debug.Print "5. Byte array input + Hex output:"
    Dim baData() As Byte
    baData = StrConv(sText, vbFromUnicode)
    Dim sResult4 As String
    sResult4 = Hash.Mode(HASH_ALG_MD5).DataBytes(baData).ReturnHex()
    Debug.Print "   " & sResult4
    Debug.Print ""
    
    ' 6. 字节数组输入 + Base64 输出
    Debug.Print "6. Byte array input + Base64 output:"
    Dim sResult5 As String
    sResult5 = Hash.Mode(HASH_ALG_SHA512).DataBytes(baData).ReturnBase64()
    Debug.Print "   " & sResult5
    Debug.Print ""
    
    ' 7. 不同算法比较
    Debug.Print "7. Different algorithms comparison:"
    Dim sMD5 As String, sSHA1 As String, sSHA256 As String
    sMD5 = Hash.Mode(HASH_ALG_MD5).DataString(sText).ReturnHex()
    sSHA1 = Hash.Mode(HASH_ALG_SHA1).DataString(sText).ReturnHex()
    sSHA256 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    Debug.Print "   MD5:    " & sMD5
    Debug.Print "   SHA1:   " & sSHA1
    Debug.Print "   SHA256: " & sSHA256
    Debug.Print ""
    
    ' 8. 中文字符测试
    Debug.Print "8. Chinese text with different encodings:"
    Dim sChinese As String
    sChinese = "你好世界"
    
    Debug.Print "   UTF8 (SHA256): " & Hash.Mode(HASH_ALG_SHA256).DataString(sChinese, ENCODING_UTF8).ReturnHex()
    Debug.Print "   ANSI (SHA256): " & Hash.Mode(HASH_ALG_SHA256).DataString(sChinese, ENCODING_ANSI).ReturnHex()
    Debug.Print ""
    
    ' 9. 链式调用验证一致性
    Debug.Print "9. Verify chain call consistency:"
    Dim sChain1 As String, sChain2 As String, sChain3 As String
    sChain1 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sChain2 = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sChain3 = Hash.ComputeHash(sText, HASH_ALG_SHA256)
    Debug.Print "   Chain 1: " & sChain1
    Debug.Print "   Chain 2: " & sChain2
    Debug.Print "   Direct:  " & sChain3
    Debug.Print "   All equal: " & (sChain1 = sChain2 And sChain2 = sChain3)
    Debug.Print ""
    
    ' 10. 省略 Mode 的用法（默认使用 SHA256）
    Debug.Print "10. Omit Mode (default SHA256):"
    Dim sWithMode As String, sWithoutMode As String
    sWithMode = Hash.Mode(HASH_ALG_SHA256).DataString(sText).ReturnHex()
    sWithoutMode = Hash.DataString(sText).ReturnHex()
    Debug.Print "   With Mode:    " & sWithMode
    Debug.Print "   Without Mode: " & sWithoutMode
    Debug.Print "   Equal: " & (sWithMode = sWithoutMode)
    Debug.Print ""
    
    ' 11. 简洁的链式调用（省略 Mode）
    Debug.Print "11. Simplified chain calls (no Mode):"
    Debug.Print "   Hex (default SHA256): " & Hash.DataString(sText).ReturnHex()
    Debug.Print "   Base64 (default SHA256): " & Hash.DataString(sText).ReturnBase64()
    Debug.Print "   Uppercase: " & Hash.DataString(sText).ReturnHex(True)
    Debug.Print ""
    
    ' 12. 重复使用 Hash 对象（自动重置）
    Debug.Print "12. Reuse Hash object (auto reset):"
    Dim sFirst As String, sSecond As String
    sFirst = Hash.DataString("Hello").ReturnHex()
    sSecond = Hash.DataString("World").ReturnHex()
    Debug.Print "   First hash (Hello):  " & sFirst
    Debug.Print "   Second hash (World): " & sSecond
    Debug.Print "   Are different: " & (sFirst <> sSecond)
    Debug.Print ""
    
    ' 13. 连续调用多个输出格式（基于同一数据）
    Debug.Print "13. Multiple output formats (same data):"
    Dim sHex As String, sBase64 As String, sUpper As String
    ' 重复设置同一数据
    sHex = Hash.DataString("Test").ReturnHex()
    sBase64 = Hash.DataString("Test").ReturnBase64()
    sUpper = Hash.DataString("Test").ReturnHex(True)
    Debug.Print "   Hex:      " & sHex
    Debug.Print "   Base64:   " & sBase64
    Debug.Print "   Uppercase: " & sUpper
    Debug.Print ""
    
    ' 14. 手动重置链式调用（使用 Mode）
    Debug.Print "14. Manual reset (using Mode):"
    Dim sBeforeReset As String, sAfterReset As String
    ' 设置数据但不调用 ReturnXxx
    Hash.Mode(HASH_ALG_MD5).DataString ("Test1")
    ' 重新开始新的链式调用
    sAfterReset = Hash.Mode(HASH_ALG_SHA256).DataString("Test2").ReturnHex()
    Debug.Print "   After reset (SHA256 of Test2): " & sAfterReset
    Debug.Print ""
    
    Set Hash = Nothing
End Sub
'演示基本哈希计算
Public Sub DemoBasicHash()
    Dim Hash As cCryptoHash
    Set Hash = New cCryptoHash
    
    Dim sText As String
    sText = "admin"
    
    Debug.Print "=== Basic Hash Demo ==="
    Debug.Print "Text: " & sText
    Debug.Print ""
    
    ' MD5
    Hash.Algorithm = HASH_ALG_MD5
    Debug.Print "MD5: " & Hash.ComputeHash(sText, HASH_ALG_MD5)
    
    ' SHA1
    Hash.Algorithm = HASH_ALG_SHA1
    Debug.Print "SHA1: " & Hash.ComputeHash(sText)
    
    ' SHA256
    Hash.Algorithm = HASH_ALG_SHA256
    Debug.Print "SHA256: " & Hash.ComputeHash(sText)
    
    ' SHA384
    Hash.Algorithm = HASH_ALG_SHA384
    Debug.Print "SHA384: " & Hash.ComputeHash(sText)
    
    ' SHA512
    Hash.Algorithm = HASH_ALG_SHA512
    Debug.Print "SHA512: " & Hash.ComputeHash(sText)
    Debug.Print ""
    
    Set Hash = Nothing
End Sub


Sub Redis()
    With New cRedisClient
        .Connect "127.0.0.1", 6379
        .Set_ "hello", "vbman redis client"
        Debug.Print .Get_("hello")
    End With
End Sub

Sub AesCBC()
    With New cAes
        MsgBox .CBC.Encode("fdsfds")
    End With
End Sub
Sub Ini()
    With New cIni
        .LoadFrom "D:\pro\ToDesk\config.ini"
        With .Section("App")
            .Item("abc") = 123
        End With
        .SaveTo "D:\pro\ToDesk\config2.ini"
    End With
End Sub


Sub StartUp()
    With New cStartUp
        .Toggle "vbm", App, "-a", "-888"
        '        .Toggle "vbm", App.Path & "\" & App.EXEName & ".exe", "-a", "-888"
    End With
End Sub

Sub Reg()
    With New cRegedit
        Dim A() As TypeRegData
        A = .FindItem("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "")
        If A(0).HasName = False Then MsgBox A(0).RegValue
        
        If .SaveItem("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "vbman2", "c:\abc.exe", Array(-1, -2, -3)) = False Then MsgBox .LastError
        If .DeleteItem("HKCU\Software\Microsoft\Windows\CurrentVersion\Run", "vbman2123") = False Then MsgBox .LastError
    End With
End Sub

Sub Pip()
    With New cSTDIO
        Debug.Print .ExecuteCommand("D:\code\vi\vbmanlib\vbman\dist\BIN\VBMAN.PLI -v", , "D:\code\vi\vbmanlib\vbman\dist\BIN\")
        Debug.Print .ExecuteCommand("ping 127.0.0.1")
    End With
End Sub

Sub CsvTest()
    With New cCsv
        .LoadFrom "D:\code\vb6\csv-data\data\test均分.csv"
        MsgBox .Value(1, 6)
    End With
End Sub
Sub cToolsStreamExample()
    Dim fs As New cToolsStream
    Dim FileName As String
    
    FileName = "d:\a\new.csv"
    
    
    With New ADODB.Stream
        .Type = adTypeText
        .CharSet = "UTF-8"
        .Open
        
        .LoadFromFile FileName
        Debug.Print .ReadText()
        .Close
    End With
    
    Exit Sub
    '1. 加载文件
    Dim Content As String
    Content = fs.LoadFileAsText(FileName)
    
    '2. 读取行（行号从1开始）
    Debug.Print "第1行: " & fs.ReadLine(1)
    Debug.Print "第3行: " & fs.ReadLine(3)
    '4. 写入行
    fs.WriteLine "新内容1", 2                                                   '写入到第2行
    fs.WriteLine "新内容2"                                                      '写入到当前行
    fs.WriteLine "新内容3", 5                                                   '写入到第5行
    '3. 顺序读取
    
    Dim lineText As String
    Do
        lineText = fs.ReadLine
        If fs.LastError <> "" Then Exit Do
        Debug.Print "当前行: " & lineText
    Loop
    
    
    
    '5. 保存
    
    '    fs.SaveFileAsText "C:\test_new.txt", "123"
End Sub

Sub FileBase64()
    Debug.Print Tools.ToolsBase64.FileToBase64("D:\a\docx.png", "png")
End Sub


Sub ArrayFor()
    Dim A: A = Array(1, 2, 3)
    Dim x
    For Each x In A
        Debug.Print x
    Next
End Sub


Sub JsonStr()
    Const A As String = "130405089908358152"
    With New cJson
        .Item("a") = A
        Debug.Print .Encode()
    End With
End Sub

Sub jsonbig()
    Dim j As New cJson
    j.LoadFrom "d:\temp\1.json"
    
End Sub
'


Sub Db()
    With New cDataBase
        If .Connect(Mysql, "127.0.0.1,3306", "root", "root", "mysql") = False Then MsgBox .LastErr: Exit Sub
        With .Sql("select * from users")
            .Fetch
            Debug.Print .Async
        End With
    End With
End Sub


Public Function JsonVar()
    With New cJson
        .Decode "var dataSK={""nameen"":""zhengzhou"",""cityname"":""郑州"",""city"":""101180101"",""temp"":""20.8"",""tempf"":""69.4"",""WD"":""北风"",""wde"":""N"",""WS"":""2级"",""wse"":""9km\/h"",""SD"":""50%"",""sd"":""50%"",""qy"":""997"",""njd"":""20km"",""time"":""15:05"",""rain"":""0"",""rain24h"":""0"",""aqi"":""83"",""aqi_pm25"":""83"",""weather"":""阴"",""weathere"":""Overcast"",""weathercode"":""d02"",""limitnumber"":""4和9"",""date"":""05月22日(星期四)""}"
        JsonVar = .Encode()
    End With
End Function

Public Sub DbgJson()
    Dim i As Long
    With New cJson
        For i = 0 To 300                                                        ' 然后就可以用 for 循环来创建无数个数组成员了
            With .NewItem()                                                     '如果数组成员也是对象,则先进性 NewItem 创建个空对象
                .Item("d") = Now()                                              '然后给对象成员赋值
                .Item("e") = 34 + i
                .Item("f") = "进入坦克: " & i
                '根据以上原理, 可以创建无限层级的子孙节点, 比如下面的 g 节点是挂在 c 节点下
                With .NewItem("g")
                    .Item("g1") = 123
                    .Item("g2") = 456
                End With
                With .NewItems("h")                                             '这个节点也挂在 c 节点下, 但这是个数组类型的节点
                    .Items(0) = 456                                             '这个数组节点成员都是普通类型， 所以直接用 items 赋值
                    .Items(0) = "123"                                           '其中 items 参数为0 则表示这个值是新增的数组元素
                    .Items(0) = "456"                                           '如果 items 参数大于0则表示修改指定索引数组元素值
                End With
            End With
        Next
        Debug.Print .Encode(, 2, True)
    End With
End Sub

Public Sub Tlogs()
    With New cLogs
        .ShowLogsViewer = True
        .Data "ghj"
    End With
End Sub


Public Sub test()
    With New cJson
        Debug.Print .Decode("sdhg({""a"":1});").Encode()
        With .NewItems("v")
            .Items(0) = 123
            .Items(0) = 456
        End With
        Debug.Print .Encode()
    End With
    '    Debug.Print VBMAN.Version()
    '    Debug.Print VBMAN.Version(App)
    '    Debug.Print VBMAN.Csv.LoadFrom("D:\code\上海vb\csv处理器\data\1.csv").Value(4, 2)
    '    Debug.Print VBMAN.HttpClient.Fetch(ReqGet, "http://a-vi.com/home/hello").ReturnText()
End Sub


Public Sub Fetch()
    Dim C As New cHttpClient
    C.DebugStart = True
    C.SetCookies "a=1"
    Debug.Print C.Fetch(ReqGet, "http://a-vi.com/home/hello?name=邓伟").ReturnJson().Encode(, 2, True)
    Debug.Print C.DebugInfo.Encode(, 2, True)
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



