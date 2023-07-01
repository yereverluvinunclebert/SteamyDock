VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   0  'None
   Caption         =   "About the Settings for Rocketdock"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox aboutPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   -45
      Picture         =   "about.frx":20FA
      ScaleHeight     =   5460
      ScaleWidth      =   5205
      TabIndex        =   0
      Top             =   -15
      Width           =   5205
      Begin VB.Label lblPunklabsLink 
         BackStyle       =   0  'Transparent
         Caption         =   "                                                         "
         Height          =   225
         Left            =   3810
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   2925
         Width           =   915
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2820
         TabIndex        =   5
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label lblVersionText 
         BackStyle       =   0  'Transparent
         Caption         =   "Version Number"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   4
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3345
         TabIndex        =   3
         Top             =   5040
         Width           =   525
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   2
         Top             =   5040
         Width           =   225
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2595
         TabIndex        =   1
         Top             =   5040
         Width           =   225
      End
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : aboutPicture_Click
' Author    : beededea
' Date      : 16/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub aboutPicture_Click()
   On Error GoTo aboutPicture_Click_Error
   If debugflg = 1 Then Debug.Print "%aboutPicture_Click"

    Me.Hide

   On Error GoTo 0
   Exit Sub

aboutPicture_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure aboutPicture_Click of Form about"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblPunklabsLink_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblPunklabsLink_Click()
   On Error GoTo lblPunklabsLink_Click_Error
   If debugflg = 1 Then Debug.Print "%lblPunklabsLink_Click"

        Call ShellExecute(Me.hwnd, "Open", "http://www.punklabs.com", vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

lblPunklabsLink_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblPunklabsLink_Click of Form about"

End Sub
