VERSION 5.00
Begin VB.Form frmRegistry 
   BorderStyle     =   0  'None
   Caption         =   "Config. sources"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConfig.frx":0000
   ScaleHeight     =   1470
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReadConfig 
      Enabled         =   0   'False
      Height          =   1350
      Left            =   75
      TabIndex        =   9
      Tag             =   "this frame must remain disabled"
      ToolTipText     =   "arse"
      Top             =   30
      Width           =   1185
      Begin VB.CheckBox chkReadConfig 
         Height          =   210
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "This tells you that the config. details are being saved to the SteamyDock file in the user data area"
         Top             =   1020
         Width           =   210
      End
      Begin VB.CheckBox chkReadSettings 
         Height          =   240
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock settings file"
         Top             =   720
         Width           =   225
      End
      Begin VB.CheckBox chkReadRegistry 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "This means that Steamydock is reading from Rocketdock's registry"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lblReadReg 
         Caption         =   "Registry"
         Height          =   240
         Left            =   405
         TabIndex        =   13
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock registry"
         Top             =   435
         Width           =   615
      End
      Begin VB.Label lblReadSet 
         Caption         =   "Settings"
         Height          =   240
         Left            =   390
         TabIndex        =   12
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock settings file"
         Top             =   735
         Width           =   570
      End
      Begin VB.Label lblReadConfig 
         Caption         =   "Config"
         Height          =   240
         Left            =   390
         TabIndex        =   11
         ToolTipText     =   "This tells you that the config. details are being saved to the SteamyDock file in the user data area"
         Top             =   1020
         Width           =   570
      End
      Begin VB.Label Label6 
         Caption         =   "Read"
         Height          =   240
         Left            =   165
         TabIndex        =   10
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock registry"
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.CommandButton btnKillIcon 
      Height          =   255
      Left            =   2415
      Picture         =   "frmConfig.frx":177C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close this box!"
      Top             =   90
      Width           =   240
   End
   Begin VB.Frame fraWriteConfig 
      Enabled         =   0   'False
      Height          =   1350
      Left            =   1290
      TabIndex        =   0
      Tag             =   "this frame must remain disabled"
      Top             =   30
      Width           =   1080
      Begin VB.CheckBox chkWriteConfig 
         Height          =   225
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "This tells you whether SteamyDock is saving to the settings.ini file"
         Top             =   1005
         Width           =   210
      End
      Begin VB.CheckBox chkWriteSettings 
         Height          =   225
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "This tells you whether Rocketdock is saving to the settings.ini file"
         Top             =   720
         Width           =   210
      End
      Begin VB.CheckBox chkWriteRegistry 
         Height          =   255
         Left            =   135
         TabIndex        =   1
         ToolTipText     =   "This will tell you whether Rocketdock is saving to the registry"
         Top             =   405
         Width           =   195
      End
      Begin VB.Label lblWriteConfig 
         Caption         =   "Config"
         Height          =   240
         Left            =   375
         TabIndex        =   7
         ToolTipText     =   "This tells you that the config. details are being saved to the SteamyDock file in the user data area"
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label lblWriteSet 
         Caption         =   "Settings"
         Height          =   240
         Left            =   375
         TabIndex        =   6
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock settings file"
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblWriteReg 
         Caption         =   "Registry"
         Height          =   240
         Left            =   390
         TabIndex        =   5
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock registry"
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Write"
         Height          =   240
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "This tells you that the config. details are being saved to the Rocketdock registry"
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.Frame fraToolTips 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   30
      TabIndex        =   17
      ToolTipText     =   "This tells you where the config. details are being read from and saved to."
      Top             =   30
      Width           =   2670
      Begin VB.Label lblTooltip 
         Height          =   1410
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "This tells you where the config. details are being read from and saved to. Best to read the help!"
         Top             =   30
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note: the frames around the check buttons are deliberately disabled. This prevents the checkboxes from being
'clicked or generating events when their values are changed.

Private Sub btnKillIcon_Click()
    frmRegistry.Hide
    rDIconConfigForm.btnSettingsDown.Visible = False
    rDIconConfigForm.btnSettingsUp.Visible = False
End Sub








Private Sub btnKillIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnKillIcon.hwnd, "This window displays the location of the current settings. This tells you where the configuration details are being stored and where they are being read from and saved to. The help has more information.", _
                  TTIconInfo, "Help on the Configuration Settings Location", , , , True
End Sub
