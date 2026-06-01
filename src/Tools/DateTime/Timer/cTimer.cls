VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'You are free to use this code in any Commercial/Non-Commercial
'applications as you wish.
'
'Author:  David Zimmer
'  Date:  Oct 16th 2003
'  Site:  http://sandsprite.com

'This class can either be used independantly as single timer object, or it can be
'used through cTimers which will allow for use as an index in an array of timers

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private timerId As Long
Private isActive As Boolean
Private Const DEBUGMODE As Boolean = False

Public Interval As Long
Public Tag As Variant
Public Event Timer()

'*********************************************************************
'this block was added to base timer class to add support for using as
'part of a parent collection with events
'
'friend properties wont be visible if compiled into DLL
'these are for our internal book keeping and do not want them
'to be externally visible (these classes are planned to be
'wrapped in a dll ultimatly but not for this demo for clarity)

Private mIndex As Long                                                          '----- these two added to support using as part
Private mParentsColKey                                                          '__|     of collection with events

Friend Property Let Index(lIndex As Long)
    mIndex = lIndex
End Property

Friend Property Get Index() As Long
    Index = mIndex
End Property

Friend Property Let ParentsColKey(key As Long)
    mParentsColKey = key
End Property

Friend Property Get ParentsColKey() As Long
    ParentsColKey = mParentsColKey
End Property
'
'*******************************************************
'
 
 

Property Get Enabled() As Boolean
    Enabled = isActive
End Property

Property Let Enabled(activateIt As Boolean)
    
    If activateIt And isActive Then
        Reset
    ElseIf activateIt And Not isActive Then
        Enable
    ElseIf isActive And Not activateIt Then
        Disable
    ElseIf Not isActive And Not activateIt Then
        'dumb ass
    End If
    
End Property

Sub EnableTimer(mInterval As Long)
    If Interval = mInterval And isActive Then TimerDestroy
    Interval = mInterval
    Enable
End Sub

Sub Enable()
    If isActive Then TimerDestroy
    TimerCreate
End Sub

Sub Disable()
    TimerDestroy
End Sub

Sub Reset()
    TimerDestroy
    TimerCreate
End Sub

' Must be public so that Timer object can't terminate while client's _Timer
' event is being processed--Friend wouldn't prevent this disaster
Public Sub RaiseTimer_Event()
Attribute RaiseTimer_Event.VB_MemberFlags = "40"
    RaiseEvent Timer
End Sub

Private Sub Class_Initialize()
    Interval = 2000
End Sub

Private Sub Class_Terminate()
    TimerDestroy
End Sub

Private Function TimerCreate() As Boolean
    
    If Interval <= 0 Then Exit Function
    
    timerId = SetTimer(0&, 0&, Interval, AddressOf ToolsTimer.TimerProc)
    If timerId Then
        TimerCreate = True
        timers.Add Me, "id:" & timerId
        isActive = True
    Else
        TimerCreate = False
        timerId = 0
        isActive = False
    End If
    
    dbg "Timer Create: " & timerId & IIf(TimerCreate, " Succedded", "Failed")
    
End Function

Private Function TimerDestroy() As Long
   
    Dim i As Integer, f As Boolean
        
    If TimerExists() Then
        f = KillTimer(0, timerId)
        timers.Remove "id:" & timerId
        TimerDestroy = True
        isActive = False
        dbg "Timer: " & timerId & " Exists Trying to kill"
    End If
     
    dbg "Killing Timer: " & timerId & IIf(TimerDestroy, " Destroyed", "Failed")
    
End Function

Private Function TimerExists() As Boolean
    On Error Resume Next
    Dim c As cTimer
    Set c = timers("id:" & timerId)
    If ERR.Number = 0 Then
        dbg "Timer Exists"
        TimerExists = True
    Else
        dbg "Timer not Found!"
    End If
    Set c = Nothing
End Function

Private Function dbg(msg As String)
    If DEBUGMODE Then Debug.Print msg
End Function
