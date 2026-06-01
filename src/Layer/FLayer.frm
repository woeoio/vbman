VERSION 5.00
Begin VB.Form FLayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2640
      Top             =   120
   End
   Begin VB.Label LContent 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   855
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "FLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsModal As Boolean
Public IsBlurClose As Boolean

Private Sub Form_LostFocus()
    If IsBlurClose = True Then Unload Me
End Sub

Private Sub Form_Resize()
    LContent.Move 0, Me.ScaleHeight / 2 - LContent.Height / 2, Me.ScaleWidth
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
End Sub

Public Function CloseAt(Optional ByVal Seconds As Long) As FLayer
    Set CloseAt = Me
    If Seconds < 1 Then Unload Me: Exit Function
    Timer1.Interval = Seconds * 1000
    Timer1.Enabled = True
End Function

Public Function ShowTo(Content As String, Optional Owner As Object) As FLayer
    Set ShowTo = Me
    Timer1.Enabled = False
    LContent.Caption = Content
    IsModal = False
    ToolsWindow.TopMost Me.hwnd
    Me.Refresh
    If Owner Is Nothing Then Me.Show: Exit Function
    Me.Show 1, Owner
End Function

Private Sub Timer2_Timer()
    '    ToolsWindow.TopMost Me.hwnd
    '    Me.Refresh
    
End Sub
