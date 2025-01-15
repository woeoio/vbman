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

Dim Fso  As New Scripting.FileSystemObject
Dim FileStream As Scripting.TextStream
Public LastError As String

Public Root As New Scripting.Dictionary
Attribute Root.VB_VarUserMemId = 0

Private Sub FixFileExtName(Path As String)
    '自动完善 .ini 后缀名
    If LCase(Right(Path, 4)) = ".ini" Then Exit Sub
    Path = Path & ".ini"
End Sub

Public Function SaveTo(Optional ByVal Path As String) As Boolean
    '把字段对象转为 ini 节点数据文本并写入 path 路径的文本文件
    On Error GoTo EH
    Dim Text As String, Value As String
    Dim Section As Variant, Key As Variant
    Dim Dic As Scripting.Dictionary: Set Dic = Root
    If Dic.Count = 0 Then LastError = "没有需要写入的数据": Exit Function
    FixFileExtName Path
    ToolsFso.AutoMakeDir Path, True
    Set FileStream = Fso.OpenTextFile(Path, ForWriting, True)
    For Each Section In Dic.Keys
        Text = Text & "[" & Section & "]" & vbCrLf
        For Each Key In Dic(Section).Keys
            Value = Dic(Section)(Key)
            Text = Text & Key & "=" & Value & vbCrLf
        Next
        Text = Text & vbCrLf
    Next
    FileStream.Write Text
    FileStream.Close
    LastError = ""
    SaveTo = True
    Exit Function
EH:
    LastError = "【iniTools.SaveTo错误】" & ERR.Number & "#" & ERR.Description
End Function

Public Function LoadFrom(ByVal Path As String) As cIni
    Set LoadFrom = Me
    On Error GoTo EH
    Dim Section As String
    Dim TextLine As String
    Dim Dic As Scripting.Dictionary: Set Dic = Root
    FixFileExtName Path
    Set FileStream = Fso.OpenTextFile(Path, ForReading)
    Do Until FileStream.AtEndOfStream
        TextLine = Trim(FileStream.ReadLine())
        ' 忽略注释和空行
        If TextLine <> "" And Left(TextLine, 1) <> ";" Then
            ' 判断是否为节（Section）
            If Left(TextLine, 1) = "[" And Right(TextLine, 1) = "]" Then
                Section = Mid(TextLine, 2, Len(TextLine) - 2)
            Else
                ' 解析键值对
                Dim Pos As Long
                Dim Key As String
                Dim Value As String
                Pos = InStr(TextLine, "=")
                If Pos > 0 Then
                    Key = Trim(Left(TextLine, Pos - 1))
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
                        Dic.Add Section, New Scripting.Dictionary
                    End If
                    Dic(Section)(Key) = Value
                End If
            End If
        End If
    Loop
    FileStream.Close
    LastError = ""
    LoadFrom = True
    Exit Function
EH:
    LastError = "【iniTools.LoadFrom错误】" & ERR.Number & "#" & ERR.Description & ": " & Path
End Function


