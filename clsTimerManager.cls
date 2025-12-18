VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTimerManager
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : requires a reference to the Microsoft scripting dictionary
'             This module contains the routines that do the behind-the-scenes stuff with APIs and populating the scripting dictionary
'---------------------------------------------------------------------------------------

Option Explicit

' member variables for locally holding property values

Private mHWnd As Long
Private mNextID As Long
Private mTimers As Object

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Error

    ' create a dictionary containing the timers by timerID
    Set mTimers = CreateObject("Scripting.Dictionary")
    
    ' Create a hidden timer window returning the window handle
    mHWnd = CreateTimerWindow

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module clsTimerManager"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : StartTimer
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub StartTimer(ByVal t As clsTimer, ByVal IntervalMS As Long)
    On Error GoTo StartTimer_Error

    ' increment the timer ID
    mNextID = mNextID + 1
    t.TimerID = mNextID

    ' add timer to the Dictionary collection
    mTimers.Add t.TimerID, t
    
    ' use the setTimer API to start a timer run
    SetTimer mHWnd, t.TimerID, IntervalMS, 0

    On Error GoTo 0
    Exit Sub

StartTimer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StartTimer of Class Module clsTimerManager"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : StopTimer
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : stop the specified timer
'---------------------------------------------------------------------------------------
'
Public Sub StopTimer(ByVal t As clsTimer)
    On Error GoTo StopTimer_Error

    ' test the collection for timer existence
    If mTimers.Exists(t.TimerID) Then
        ' kill it by ID
        KillTimer mHWnd, t.TimerID
        ' remove the timer from the collection
        mTimers.Remove t.TimerID
    End If

    On Error GoTo 0
    Exit Sub

StopTimer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StopTimer of Class Module clsTimerManager"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Dispatch
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : Called by WindowProc in the BAS module modTimerWnd on captured form timer events
'---------------------------------------------------------------------------------------
'
Friend Sub Dispatch(ByVal TimerID As Long)
    On Error GoTo Dispatch_Error

    ' tests the timer in the timer collection to see if it exists, if so, raise a tick event in the calling form frmTimer
    If mTimers.Exists(TimerID) Then
        ' call the raiseTick subroutine in the clsTimer class that in turn raises a tick event in the calling form frmTimer
        mTimers(TimerID).RaiseTick
    End If

    On Error GoTo 0
    Exit Sub

Dispatch_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Dispatch of Class Module clsTimerManager"
End Sub
