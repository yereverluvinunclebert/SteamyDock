VERSION 5.00
Begin VB.Form rdIconSelectForm 
   Caption         =   "Rocketdock Icon Selector"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   Icon            =   "rocketdocksettings.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4755
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   4725
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   8235
      Begin VB.PictureBox picLarge 
         BorderStyle     =   0  'None
         Height          =   1980
         Left            =   1530
         ScaleHeight     =   1980
         ScaleWidth      =   2250
         TabIndex        =   13
         Top             =   975
         Visible         =   0   'False
         Width           =   2250
         Begin VB.Label lblPicLarge 
            BackStyle       =   0  'Transparent
            Caption         =   "Rocketdock settings"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   180
            TabIndex        =   14
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rocketdock settings"
            ForeColor       =   &H00808080&
            Height          =   210
            Left            =   195
            TabIndex        =   15
            Top             =   15
            Width           =   1140
         End
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   11
         Left            =   4140
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   10
         Left            =   4695
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   9
         Left            =   5250
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   8
         Left            =   5805
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   6360
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   6
         Left            =   6915
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   1920
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   2475
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   4
         Left            =   3030
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   5
         Left            =   3585
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   810
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   495
      End
      Begin VB.PictureBox picDock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   1365
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         ToolTipText     =   "RocketDock icon number "
         Top             =   210
         Width           =   500
      End
      Begin VB.TextBox txtFilename 
         Height          =   360
         Left            =   135
         TabIndex        =   16
         Text            =   "Text"
         Top             =   825
         Width           =   8010
      End
   End
End
Attribute VB_Name = "rdIconSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim fileName1 As String
    Dim fullName1 As String
    Dim lPic As Picture
    Dim useloop As Integer
    Dim suffix As String
        
    ' display small versions of the icons on the icon selector
    
    For useloop = 0 To 11 ' 11
        'FileName1 = rdAppPath & "\" & sFileName(useloop)
        'If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
        'FileName = FileName & filesIconList.FileName
'        picDock(useloop).AutoRedraw = True
'
'        picLarge.AutoRedraw = True
        
        'Set picDock(useloop).Picture = StdPictureEx.LoadPicture(sFileName(useloop), lpsCustom, , 32, 32)
        'Set picDock(useloop).Picture = StdPictureEx.LoadPicture(FileName1, lpsCustom, , 32, 32)
        
'        suffix = Right(fileName1, Len(fileName1) - InStr(1, fileName1, "."))
'        If InStr("png,jpg,bmp,jpeg", LCase(suffix)) <> 0 Then
'            Set picDock(useloop).Picture = StdPictureEx.LoadPicture(fileName1)
'        Else ' ICO, CUR
'            Set picDock(useloop).Picture = StdPictureEx.LoadPicture(fileName1, lpsLargeShell)
'        End If
'
'        txtFilename.Text = fileName1
    Next useloop

End Sub



Private Sub Frame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picLarge.Visible = False
End Sub

