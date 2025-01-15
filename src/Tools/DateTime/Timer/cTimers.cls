VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'You are free to use this code in any Commercial/Non-Commercial applications as you
'wish.
'
'Author:  David Zimmer
'  Date:  Oct 16th 2003
'  Site:  http://sandsprite.com


Public Event Timer(ByVal Index As Integer)

Private mCount As Integer
Private mCol As New Collection 'object ref to child cTimer classes created by this class
Private mClassKey As Integer

Property Get Count() As Integer
    Count = mCount
End Property

Function Add() As Integer
    Dim t As New cTimer
    
    mCount = mCount + 1                                                         '1 based indexes (same as collection object)
    
    'these two friend properties were added to base class (cTimer) to support this framework
    t.Index = mCount
    t.ParentsColKey = mClassKey
    
    mCol.Add t
    Add = mCount
    
End Function

Function BulkLoad(num As Integer)
    Dim i As Long
    
    'mabey add check to make sure .Add() not called before, so indexes are auto aligned 1-num ?
    If num > 0 And num < 200 Then
        For i = 0 To num
            Add
        Next
    Else
        ERR.Raise vbObjectError, "cTimers.BulkLoad", "No more than 200 times allowed per class"
    End If
                    
End Function


'default function of class (so we can refer to it as myCTimersClass(index).
Function Item(ByVal Index As Integer) As cTimer
Attribute Item.VB_UserMemId = 0
    If Index > Count Then ERR.Raise vbObjectError, "cTimers.Item", "Index does not Exist"
    Set Item = mCol(Index)
End Function

'called from modTimer directly
Friend Function RaiseTimer_Event(ByVal Index As Integer)
    RaiseEvent Timer(Index)
End Function

Private Sub Class_Initialize()
    mClassKey = RegisterTimerCollection(Me)
End Sub

Private Sub Class_Terminate()
    Dim t As cTimer
    
    For Each t In mCol
        Set t = Nothing
    Next
    
    Set mCol = Nothing
    
    CTimersCol.Remove "key:" & mClassKey
    
End Sub
