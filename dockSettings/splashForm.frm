VERSION 5.00
Begin VB.Form splashForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer splashShrinkerTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   4680
   End
   Begin VB.CheckBox chkSplashDisable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Suppress this splash pop-up"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1470
      TabIndex        =   0
      ToolTipText     =   "Click here and the splash screen will never show again"
      Top             =   8220
      Width           =   3120
   End
   Begin VB.Timer splashTimer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   705
      Top             =   5250
   End
   Begin VB.PictureBox picSplash 
      BorderStyle     =   0  'None
      Height          =   8550
      Left            =   0
      Picture         =   "splashForm.frx":0000
      ScaleHeight     =   8550
      ScaleWidth      =   6060
      TabIndex        =   1
      ToolTipText     =   "Click anywhere to hide the splash screen"
      Top             =   0
      Width           =   6060
   End
End
Attribute VB_Name = "splashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' .01 DAEB splashForm new splashShrinkerTimer_Timer
' .02 DAEB splashForm new form load subroutine
' .03 DAEB splashForm.frm 09/02/2021 handling any potential divide by zero

Public splashTimerCount As Integer
Public splashWidth As Integer
Public splashHeight As Integer
Public splashFormWidth As Integer
Public pic As Picture


'---------------------------------------------------------------------------------------
' Procedure : chkSplashDisable_Click
' Author    : beededea
' Date      : 02/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSplashDisable_Click()
   On Error GoTo chkSplashDisable_Click_Error

    splashForm.Hide
    
    If chkSplashDisable.Value = 1 Then
        sDSplashStatus = "0"
    Else
        sDSplashStatus = "1"
    End If
    
    PutINISetting "Software\SteamyDock\DockSettings", "SplashStatus", sDSplashStatus, dockSettingsFile

   On Error GoTo 0
   Exit Sub

chkSplashDisable_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkSplashDisable_Click of Form splashForm"
 
End Sub
' .02 DAEB splashForm new form load subroutine
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 23/01/2021
' Purpose   : sets some resizing variables and loads the splash image into a picture object
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    splashTimerCount = 0
    splashFormWidth = splashForm.Width
        
    If FExists(App.Path & "\steamydock-splash.jpg") Then
        Set pic = LoadPicture(App.Path & "\steamydock-splash.jpg")
    End If
    
    splashWidth = ScaleX(pic.Width, vbHimetric, vbTwips)
    splashHeight = ScaleY(pic.Height, vbHimetric, vbTwips)

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form splashForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : picSplash_Click
' Author    : beededea
' Date      : 23/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picSplash_Click()
   On Error GoTo picSplash_Click_Error

    splashForm.Hide
    splashShrinkerTimer.Enabled = False

   On Error GoTo 0
   Exit Sub

picSplash_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picSplash_Click of Form splashForm"
End Sub
' .01 DAEB splashForm new splashShrinkerTimer_Timer STARTS
'---------------------------------------------------------------------------------------
' Procedure : splashShrinkerTimer_Timer
' Author    : beededea
' Date      : 23/01/2021
' Purpose   : resizes the splash picture box x & y as a ratio but resizes the underlying form
'             as a simple decrement, leaving an interesting effect as it does so.
'---------------------------------------------------------------------------------------
'
Private Sub splashShrinkerTimer_Timer()
   On Error GoTo splashShrinkerTimer_Timer_Error
    
    If splashForm.Height > 51 Then splashForm.Height = splashForm.Height - 50
    If splashForm.Width > 51 Then
        splashForm.Width = splashForm.Width - 50
        picSplash.Width = picSplash.Width - 50
    
        If splashWidth = 0 Then splashWidth = 1 ' .03 DAEB splashForm.frm 09/02/2021 handling any potential divide by zero

        sngRatio = picSplash.Width / splashWidth
        If splashHeight * sngRatio > picSplash.Height Then
            If splashHeight = 0 Then splashHeight = 1 ' .03 DAEB splashForm.frm 09/02/2021 handling any potential divide by zero
            sngRatio = picSplash.Height / splashHeight
        End If
        
        picSplash.AutoRedraw = True
        picSplash.PaintPicture pic, 0, 0, splashWidth * sngRatio, splashHeight * sngRatio
    Else
        splashShrinkerTimer.Enabled = False
        splashForm.Hide
        splashForm.Width = 6075
        splashForm.Height = 8535
        picSplash.Width = 6060
        picSplash.Height = 8550
        
        picSplash.Cls
        picSplash.AutoRedraw = True
        picSplash.PaintPicture pic, 0, 0, 6060, 8550
    End If
    
    

    If splashForm.Width <= 1 Or splashForm.Height <= 1 Then
        splashShrinkerTimer.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

splashShrinkerTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure splashShrinkerTimer_Timer of Form splashForm"
End Sub
' .01 DAEB splashForm ENDS

'---------------------------------------------------------------------------------------
' Procedure : splashTimer_Timer
' Author    : beededea
' Date      : 23/01/2021
' Purpose   : does nothing first iteration, then triggers the resizing animation timer
'---------------------------------------------------------------------------------------
'
Private Sub splashTimer_Timer()
   On Error GoTo splashTimer_Timer_Error

    If splashTimerCount = 0 Then ' this prevents the 3.25 second timer doing anything on its first iteration
        splashTimerCount = splashTimerCount + 1
    Else
        splashShrinkerTimer.Enabled = True
        splashTimer.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

splashTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure splashTimer_Timer of Form splashForm"
End Sub
