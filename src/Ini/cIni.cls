VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'    引用了Microsoft Scripting Runtime
    
'todo: 支持无节点的键值对读写


Dim FileStream As New cToolsStream
Public LastError As String

Dim Aes As New cAes

Dim pvPwd As String
Dim pvPath As String

Dim RootDic As New Dictionary

Public Property Get Root() As Dictionary
Attribute Root.VB_UserMemId = 0
    Set Root = RootDic
End Property

Public Property Get MultiLineText(ByVal SectionName As String, ByVal KeyName As String) As String
    MultiLineText = Replace(Root(SectionName)(KeyName), "\r\n", vbCrLf)
End Property
Public Property Let MultiLineText(ByVal SectionName As String, ByVal KeyName As String, ByVal Text As String)
    Root(SectionName)(KeyName) = Replace(Text, vbCrLf, "\r\n")
End Property

Public Property Get Section(ByVal key As String) As Dictionary
    If RootDic.Exists(key) = False Then
        RootDic.Add key, New Dictionary
    End If
    '这里今天遇到一个奇怪的问题，遇到这个函数居然报错，刚才才发现就是因为字典的特性引起的：只要读取过一个键值，如果字典不存在该项目，也会创建一个空项目，这就导致好多地方会出现诡异的事情
    If IsObject(RootDic(key)) = False Then                                      '所以这里得先判断是否对象，不是就new 一下，同样的问题出现在 save 函数
        Set RootDic(key) = New Dictionary
    End If
    Set Section = RootDic(key)
End Property

Private Sub FixFileExtName(Path As String)
    '自动完善 .ini 后缀名
    If LCase(Right(Path, 4)) = ".ini" Then Exit Sub
    Path = Path & ".ini"
End Sub

Public Function Password(ByVal Pwd As String) As cIni
    Set Password = Me
    pvPwd = Pwd
End Function

Public Function SaveTo(Optional ByVal Path As String, Optional CharSet As String = "UTF-8") As Boolean
    '把字段对象转为 ini 节点数据文本并写入 path 路径的文本文件
    On Error GoTo EH
    Dim Text As String, Value As String
    Dim Section As Variant, key As Variant, Node As Object
    Dim Dic As Dictionary: Set Dic = RootDic
    If Path = "" Then Path = pvPath
    If Path = "" Then LastError = "文件路径不能为空": Exit Function
    '    FixFileExtName Path
    ToolsFso.AutoMakeDir Path, True
    '    If Dic.count = 0 Then LastError = "没有需要写入的数据": Exit Function
    For Each Section In Dic.Keys
        '这里要先判断是否对象，，问题同样出现在 Section函数
        If IsObject(Dic(Section)) = False Then GoTo CT
        Text = Text & "[" & Section & "]" & vbCrLf
        For Each key In Dic(Section).Keys
            Value = Dic(Section)(key)
            Text = Text & key & "=" & Value & vbCrLf
        Next
        Text = Text & vbCrLf
CT:
    Next
    '    If pvPwd <> "" Then Text = Aes.CBC.Encode(Text, pvPwd)
    FileStream.SaveFileAsText Path, Text, CharSet
    FileStream.Inst.Close
    LastError = ""
    SaveTo = True
    Exit Function
EH:
    LastError = "【iniTools.SaveTo错误】" & ERR.Number & "#" & ERR.Description
    ERR.Raise ERR.Number, ERR.Source, LastError
End Function

Public Function LoadFrom(ByVal Path As String, Optional CharSet As String = "UTF-8") As cIni
    Set LoadFrom = Me
    On Error GoTo EH
    Dim Section As String
    Dim TextLine As String, i As Long
    Dim Dic As Dictionary: Set Dic = RootDic
    Call Dic.RemoveAll
    '    FixFileExtName Path
    '  如果是 loadForm 读取的ini文件，saveto 方法可以省略路径自动保存
    pvPath = Path
    FileStream.UseLine = True
    Call FileStream.LoadFileAsText(Path, CharSet)
    For i = 1 To FileStream.LineData.count
        TextLine = Trim(FileStream.ReadLine())
        ' 忽略注释和空行
        If TextLine <> "" And Left(TextLine, 1) <> ";" Then
            ' 判断是否为节（Section）
            If Left(TextLine, 1) = "[" And Right(TextLine, 1) = "]" Then
                Section = Mid(TextLine, 2, Len(TextLine) - 2)
            Else
                ' 解析键值对
                Dim Pos As Long
                Dim key As String
                Dim Value As String
                Pos = InStr(TextLine, "=")
                If Pos > 0 Then
                    key = Trim(Left(TextLine, Pos - 1))
                    Value = Trim(Mid(TextLine, Pos + 1))
                    
                    ' 去除开头的双引号
                    If Left(Value, 1) = """" Then
                        Value = Right(Value, Len(Value) - 1)
                    End If
                    ' 去除结尾的双引号
                    If Right(Value, 1) = """" Then
                        Value = Left(Value, Len(Value) - 1)
                    End If
                    
                    ' 将键值对存储到字典中
                    If Not Dic.Exists(Section) Then
                        Dic.Add Section, New Dictionary
                    End If
                    Dic(Section)(key) = Value
                End If
            End If
        End If
    Next
    FileStream.Inst.Close
    LastError = ""
    Exit Function
EH:
    LastError = "【iniTools.LoadFrom错误】" & ERR.Number & "#" & ERR.Description & ": " & Path
End Function


