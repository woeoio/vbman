Attribute VB_Name = "ToolsTimer"
Option Explicit

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'each timer class registers an obj instance in this collection
'key= "id:" & timerID , item = reference to live class object
Public timers As New Collection

'each cTimers class registers itself by its class key here
'key= "key:" & intID , item = reference to live class object
Public CTimersCol As New Collection

Private mTimersColCount As Integer

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal idEvent As Long, ByVal dwTime As Long)
    
    Dim t As cTimer
    Dim c As cTimers
    
    On Error Resume Next
    Set t = timers("id:" & idEvent)
    If t Is Nothing Then
        KillTimer 0&, idEvent
    Else
        If t.ParentsColKey > 0 Then                                             'this timer is an index in cTimers
            Set c = CTimersCol("key:" & t.ParentsColKey)
            If c Is Nothing Then
                KillTimer 0&, idEvent
                Debug.Print "THIS SHOULDNT HAPPEN: parent collection died?"
            Else
                'raise the event in the parent collection class instead of timer class
                c.RaiseTimer_Event t.index
            End If
        Else
            t.RaiseTimer_Event
        End If
    End If
    Set t = Nothing
    
End Sub

'returns key to this class in collection
Function RegisterTimerCollection(c As cTimers) As Integer
    Dim key As String
    
    mTimersColCount = mTimersColCount + 1
    key = "key:" & mTimersColCount                                              'will always be unique because counting
    CTimersCol.Add c, key
    RegisterTimerCollection = mTimersColCount
    
End Function

