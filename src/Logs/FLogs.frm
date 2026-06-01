VERSION 5.00
Begin VB.Form FLogs 
   Caption         =   "VBMAN LogViwer"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   14070
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   5295
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FLogs.frx":0000
      Top             =   360
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   6000
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12615
   End
   Begin VB.Menu Menu_Clear 
      Caption         =   "清空"
   End
   Begin VB.Menu Window 
      Caption         =   "窗口"
      Begin VB.Menu WindowTopMost 
         Caption         =   "顶置"
      End
   End
End
Attribute VB_Name = "FLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Layer As New cLayer


Public Sub AddLog(Content As String)
    List1.AddItem Content, 0
End Sub


'===========================================
Private Sub Command1_Click()
    Text1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.Text
    Layer.msg "已复制"
End Sub

Private Sub Form_Load()
    Text1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
    '    Dim i
    '    For i = 0 To 100
    '        List1.AddItem i
    '    Next
    Call WindowTopMost_Click
End Sub

Private Sub Form_Resize()
    List1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight + 100
    Text1.Move Me.ScaleWidth / 2, 0, Me.ScaleWidth / 2, Me.ScaleHeight
    Command1.Move Me.ScaleWidth / 2 - Command1.Width, 100
    Command2.Move Me.ScaleWidth / 2 - Command2.Width, Command1.Height + 100
End Sub

Private Sub List1_Click()
    Text1.Visible = True
    Command1.Visible = True
    Command2.Visible = True
    Text1.Text = List1.Text
End Sub

Private Sub Menu_Clear_Click()
    List1.Clear
    Text1.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End Sub

Private Sub WindowTopMost_Click()
    WindowTopMost.Checked = Not WindowTopMost.Checked
    ToolsWindow.TopMost Me.hWnd, Not WindowTopMost.Checked
End Sub
