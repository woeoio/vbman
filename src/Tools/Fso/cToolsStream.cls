VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'作者：邓伟，QQ 215879458

'    Library ADODB
'C:  \Program Files (x86)\Common Files\System\ado\msado28.tlb
'    Microsoft ActiveX Data Objects 2.8 Library

Public Inst As New ADODB.Stream
Public LastError As String, LineNumber As Long                                  'todo 对外需要只读
Public UseLine As Boolean                                                       '如果启用了按行读写，则以 lineData 为数据主体，inst 的内容是原始值
Public LineData As New cCollection
Public Txt As New cToolsStr

Public Function UseLineMode(v As Boolean) As cToolsStream
    Set UseLineMode = Me
    UseLine = v
End Function


Public Function ReadLine(Optional ByVal Line As Long) As String
    '从 LineNumber 读入一行内容，如果提供了 Line 则读入指定行
    If UseLine = False Then ERR.Raise 1, , "未启用按行读写"
    If Line > 0 Then LineNumber = Line
    ReadLine = LineData.Item(LineNumber)
    LineNumber = LineNumber + 1
End Function

Public Function WriteLine(Text As String, Optional ByVal Line As Long) As Boolean
    '向 LineNumber 写入一行内容，如果提供了 Line 则写入指定行
    If UseLine = False Then ERR.Raise 1, , "未启用按行读写"
    If Line > 0 Then LineNumber = Line
    LineData.Add Text, LineNumber
    LineNumber = LineNumber + 1
    WriteLine = True
End Function

Public Function LoadFileAsText(ByVal FileName As String, Optional ByVal CharSet As String = "UTF-8") As String
    With Inst
        If .State <> adStateClosed Then .Close
        .Type = adTypeText
        .Open
        .CharSet = CharSet
        .LoadFromFile FileName
        LoadFileAsText = .ReadText()
        .position = 0
        .Close
    End With
    ' 如果 UseLine 为 true，则将内容按行分割并存储到 LineDict
    If UseLine Then
        Set LineData = Txt.SplitLinesToCollection(LoadFileAsText)
        LineNumber = 1
    End If
    '    Inst.Close
End Function
Public Function SaveFileAsText(ByVal FileName As String, Optional Data As Variant, Optional ByVal CharSet As String = "UTF-8") As Boolean
    Dim Dat As String, isInst As Boolean
    
    If IsMissing(Data) = True Then
        If UseLine = True Then
            Dat = Txt.JoinLinesFromCollection(LineData)
        Else
            isInst = True
        End If
    Else
        Dat = Data
    End If
    
    With Inst
        If .State <> adStateClosed Then .Close
        .Type = adTypeText
        .Open
        .CharSet = CharSet
        If isInst = False Then
            'todo ado流对象不能写入空值，这里只能弄个空格先顶一下，后面有时间改用 cFileEx 对象
            If Len(Dat) = 0 Then Dat = " "
            .WriteText Dat
        End If
        .SaveToFile FileName, adSaveCreateOverWrite
        '        .Close
    End With
End Function

Public Function LoadFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    On Error GoTo EH
    LastError = ""
    ' 修复：每次调用创建独立 Stream 实例，避免共享 Inst 在并重入时被意外关闭
    ' 原代码使用共享 Inst (Public Inst As New ADODB.Stream)，
    ' 在 HTTP 服务器场景下，两个请求交替处理时第二个调用的 .Close 会关闭第一个正在使用的 Stream，
    ' 导致 OutData 为空数组/无效状态，下游 UBound(OutData) 直接 GPF 崩溃
    Dim stm As ADODB.Stream
    Set stm = New ADODB.Stream
    With stm
        .Type = adTypeBinary
        .Open
        .LoadFromFile Path
        If .EOS Then
            ' 0 字节文件：返回安全数组，避免下游 (Not Data) 或 UBound 崩溃
            ' ADODB.Stream 对空文件 .Read() 行为因版本而异，可能返回空/Nothing/异常
            ReDim OutData(0 To 0)
            OutData(0) = 0
        Else
            OutData = .Read()
        End If
        .Close
    End With
    Set stm = Nothing
    LoadFileAsBinary = True
    Exit Function
EH:
    LastError = ERR.Description & "#" & ERR.Description
End Function

Public Function SaveFileAsBinary(ByVal Path As String, OutData() As Byte) As Boolean
    On Error GoTo EH
    LastError = ""
    With Inst
        If .State <> adStateClosed Then .Close
        .Type = adTypeBinary
        .Open
        .Write OutData
        .SaveToFile Path, adSaveCreateOverWrite
        .Close
    End With
    SaveFileAsBinary = True
    Exit Function
EH:
    LastError = ERR.Description & "#" & ERR.Description
End Function

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Inst.Close
End Sub