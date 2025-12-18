VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTimer
' Author    : beededea & chatGPT (with corrections/new properties/documentation by me)
' Date      : 18/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

' Member variables IDs for the in-code timers.

Private mTimerID As Long
Private mEnabled As Boolean
Private mRunning As Boolean
Private mInterval As Long

' public tick event for an object of this class

Public Event Tick() ' the tick event appears to be the standard named event for most timers in other languages

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : just set a few base member variables
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Error

    mInterval = 1000 ' one second by default
    mEnabled = False

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module clsTimerManager"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TimerID
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : property to obtain the timer ID number
'---------------------------------------------------------------------------------------
'
Friend Property Get TimerID() As Long
    On Error GoTo TimerID_Error

    TimerID = mTimerID

    On Error GoTo 0
    Exit Property

TimerID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerID of Class Module clsTimer"
End Property

'---------------------------------------------------------------------------------------
' Procedure : TimerID
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : property to set the timer ID number
'---------------------------------------------------------------------------------------
'
Friend Property Let TimerID(ByVal newValue As Long)
    On Error GoTo TimerID_Error

    mTimerID = newValue

    On Error GoTo 0
    Exit Property

TimerID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerID of Class Module clsTimer"
End Property



'---------------------------------------------------------------------------------------
' Procedure : Enabled
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : property to obtain status
'---------------------------------------------------------------------------------------
'
Friend Property Get Enabled() As Boolean
    On Error GoTo Enabled_Error

    Enabled = mEnabled

    On Error GoTo 0
    Exit Property

Enabled_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Enabled of Class Module clsTimer"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Enabled
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : property to start or disable a new timer according to previously supplied interval
'---------------------------------------------------------------------------------------
'
Friend Property Let Enabled(ByVal newValue As Boolean)
    On Error GoTo Enabled_Error
    
    'if we are going to start the timer but the timer is already running then stop the timer
    If mRunning = True Then
        StopTimer
    End If
    
    mEnabled = newValue

    If mEnabled = True Then
        TimerManager.StartTimer Me, mInterval
        mRunning = True
    Else
        StopTimer
    End If

    On Error GoTo 0
    Exit Property

Enabled_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Enabled of Class Module clsTimer"
End Property



'---------------------------------------------------------------------------------------
' Procedure : Interval
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : property to obtain the Interval
'---------------------------------------------------------------------------------------
'
Friend Property Get Interval() As Long
    On Error GoTo Interval_Error

    Interval = mInterval

    On Error GoTo 0
    Exit Property

Interval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Interval of Class Module clsTimer"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Interval
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : property to set the interval
'---------------------------------------------------------------------------------------
'
Friend Property Let Interval(ByVal newValue As Long)
    On Error GoTo Interval_Error
    
    mInterval = newValue
    
    'if the timer is already running then stop and restart the timer with the new interval
    If mRunning = True Then
        StopTimer
        TimerManager.StartTimer Me, mInterval
        mRunning = True
    End If
    
    On Error GoTo 0
    Exit Property

Interval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Interval of Class Module clsTimer"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Start
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : start a new timer according to supplied interval
'---------------------------------------------------------------------------------------
'
Public Sub Start(ByVal IntervalMS As Long)
    On Error GoTo Start_Error

    If mRunning Then StopTimer
    mInterval = IntervalMS
    TimerManager.StartTimer Me, IntervalMS
    mRunning = True

    On Error GoTo 0
    Exit Sub

Start_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Start of Class Module clsTimer"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : StopTimer
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub StopTimer()
    On Error GoTo StopTimer_Error

    If mRunning Then
        TimerManager.StopTimer Me
        mRunning = False
    End If

    On Error GoTo 0
    Exit Sub

StopTimer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StopTimer of Class Module clsTimer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RaiseTick
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : raise a tick event for an object of this class
'---------------------------------------------------------------------------------------
'
Public Sub RaiseTick()
    On Error GoTo RaiseTick_Error

    If mRunning Then
        RaiseEvent Tick
    End If

    On Error GoTo 0
    Exit Sub

RaiseTick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RaiseTick of Class Module clsTimer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : terminate the current object of this class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Error

    StopTimer

    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module clsTimer"
End Sub



