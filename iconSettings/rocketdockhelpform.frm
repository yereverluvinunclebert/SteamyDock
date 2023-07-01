VERSION 5.00
Begin VB.Form rdHelpForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "RocketDock Settings Help"
   ClientHeight    =   11595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   11595
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBoxHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11595
      Left            =   0
      Picture         =   "rocketdockhelpform.frx":0000
      ScaleHeight     =   11595
      ScaleWidth      =   14910
      TabIndex        =   0
      ToolTipText     =   "click to make me go away"
      Top             =   0
      Width           =   14910
   End
End
Attribute VB_Name = "rdHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub picBoxHelp_Click()
    rdHelpForm.Hide
End Sub
