VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQRcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim pvContent As Variant
Dim pvSize As Long
Dim pvForeColor As OLE_COLOR

Private Sub Class_Initialize()
    pvContent = "hello vbman form www.vb6.pro"
    pvSize = 50
    pvForeColor = vbBlack
End Sub

Public Function Generate(Optional ByVal IsCopyToClipboard As Boolean) As StdPicture
    Set Generate = QRCodegenBarcode(pvContent, pvForeColor, QRCodegenEcc_LOW, 1, 5, , , pvSize)
    If IsCopyToClipboard = True Then
        Clipboard.Clear
        Clipboard.SetData Generate
    End If
End Function

Public Function SetText(ByVal Content As Variant) As cQRcode
    Set SetText = Me
    pvContent = Content
End Function

Public Function SetSize(ByVal Size As Long) As cQRcode
    Set SetSize = Me
    pvSize = Size
End Function

Public Function SetForeColor(ByVal ForeColor As OLE_COLOR) As cQRcode
    Set SetForeColor = Me
    pvForeColor = ForeColor
End Function


'Public Function QRCodegenBarcode(TextOrByteArray As Variant, _
'    Optional ByVal clrFore As OLE_COLOR = vbBlack, _
'    Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
'    Optional ByVal MinVersion As Long = VERSION_MIN, _
'    Optional ByVal MaxVersion As Long = VERSION_MAX, _
'    Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
'    Optional ByVal BoostEcl As Boolean = True, _
'    Optional PicSize As Long = 500) As StdPicture                               ' Saturnian编辑

'Private Sub Command1_Click()
'    On Error Resume Next
'    SavePicture Picture1, App.Path & "\test.jpg"                                ' 将图片保存到文件。
'    MsgBox "图片保存成功！"
'
'End Sub
'
'Private Sub Form_Load()
'    Text1_Change
'
'End Sub
'
'Private Sub HScroll1_Change()                                                   ' Edited by Saturnian
'    Picture1.Picture = QRCodegenBarcode(Text1.Text, 0, QRCodegenEcc_LOW, 1, 5, , , HScroll1.Value)
'    Clipboard.Clear
'    Clipboard.SetData QRCodegenBarcode(Text1.Text, 0, QRCodegenEcc_LOW, 1, 5, , , HScroll1.Value)
'End Sub
'
'Private Sub Text1_Change()
'    Set Image1.Picture = QRCodegenBarcode(Text1.Text)
'    Picture1.Picture = QRCodegenBarcode(Text1.Text, 0, QRCodegenEcc_LOW, 1, 5, , , 5) ' Edited by Saturnian
'    Clipboard.Clear                                                             ' Edited by Saturnian
'    Clipboard.SetData QRCodegenBarcode(Text1.Text, 0, QRCodegenEcc_LOW, 1, 5, , , 5) ' Edited by Saturnian
'
'End Sub

