VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form dockSettings 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SteamyDock Settings"
   ClientHeight    =   9315
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8220
   Icon            =   "dockSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   176
      ToolTipText     =   "Click here to open tool's HTML help page in your browser"
      Top             =   8790
      Width           =   1065
   End
   Begin VB.PictureBox picBusy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4785
      Picture         =   "dockSettings.frx":058A
      ScaleHeight     =   795
      ScaleWidth      =   825
      TabIndex        =   175
      ToolTipText     =   "The program is doing something..."
      Top             =   8700
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Timer busyTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4425
      Top             =   8760
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3915
      Top             =   8760
   End
   Begin VB.Timer repaintTimer 
      Interval        =   1000
      Left            =   3405
      Top             =   8760
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7095
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit this utility"
      Top             =   8805
      Width           =   1065
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Save && Restart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5685
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "This will save your changes and restart the dock."
      Top             =   8805
      Width           =   1335
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "De&faults"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Revert ALL settings to the defaults"
      Top             =   8790
      Width           =   1065
   End
   Begin VB.PictureBox iconBox 
      BackColor       =   &H00FFFFFF&
      Height          =   8520
      Left            =   120
      ScaleHeight     =   8460
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   135
      Width           =   1035
      Begin VB.Frame fmeIconAbout 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1425
         Left            =   90
         TabIndex        =   20
         Top             =   6930
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   5
            Left            =   -60
            Picture         =   "dockSettings.frx":1005
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   21
            ToolTipText     =   "About the dock settings"
            Top             =   -75
            Width           =   960
         End
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   22
            Top             =   855
            Width           =   570
         End
      End
      Begin VB.Frame fmeIconGeneral 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   45
         TabIndex        =   19
         Top             =   -30
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   0
            Left            =   0
            Picture         =   "dockSettings.frx":1B6A
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   28
            ToolTipText     =   "General Configuration Options"
            Top             =   0
            Width           =   960
         End
         Begin VB.Frame fmeLblFrame 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   960
            Width           =   735
            Begin VB.Label lblText 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "General"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   -30
               TabIndex        =   26
               ToolTipText     =   "General Configuration Options"
               Top             =   -15
               Width           =   795
            End
         End
      End
      Begin VB.Frame fmeIconPosition 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   15
         TabIndex        =   17
         Top             =   5430
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   4
            Left            =   0
            Picture         =   "dockSettings.frx":2630
            ScaleHeight     =   1020
            ScaleWidth      =   960
            TabIndex        =   18
            ToolTipText     =   "Dock Positioning"
            Top             =   60
            Width           =   960
            Begin VB.Label lblText 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Position"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   150
               TabIndex        =   23
               Top             =   810
               Width           =   765
            End
         End
      End
      Begin VB.Frame fmeIconStyle 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   -15
         TabIndex        =   8
         Top             =   4005
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   3
            Left            =   0
            Picture         =   "dockSettings.frx":352A
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   9
            ToolTipText     =   "Dock theme and text configuration"
            Top             =   15
            Width           =   960
         End
         Begin VB.Label lblText 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Style"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   300
            TabIndex        =   24
            Top             =   945
            Width           =   765
         End
      End
      Begin VB.Frame fmeIconBehaviour 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   30
         TabIndex        =   6
         Top             =   2505
         Width           =   915
         Begin VB.Frame fmeLblFrame 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   30
            TabIndex        =   31
            Top             =   990
            Width           =   930
            Begin VB.Label lblText 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Behaviour"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   2
            Left            =   45
            Picture         =   "dockSettings.frx":412B
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   7
            ToolTipText     =   "Icon bounce and pop up effects"
            Top             =   15
            Width           =   960
         End
      End
      Begin VB.Frame fmeIconIcons 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   30
         TabIndex        =   4
         Top             =   1305
         Width           =   915
         Begin VB.Frame fmeLblFrame 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   915
            Width           =   735
            Begin VB.Label lblText 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Icons"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   -30
               TabIndex        =   30
               Top             =   0
               Width           =   795
            End
         End
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   1
            Left            =   -30
            Picture         =   "dockSettings.frx":4C31
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   27
            ToolTipText     =   "Icon effects and quality"
            Top             =   -30
            Width           =   960
         End
         Begin VB.Label lblIcons 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Icons"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   195
            TabIndex        =   5
            Top             =   900
            Width           =   690
         End
      End
   End
   Begin VB.PictureBox picHiddenPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   6270
      ScaleHeight     =   1605
      ScaleWidth      =   1650
      TabIndex        =   92
      ToolTipText     =   "The icon size in the dock"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Icon && Dock Behaviour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   2
      Left            =   1230
      TabIndex        =   64
      ToolTipText     =   "Here you can control the behaviour of the animation effects"
      Top             =   30
      Width           =   6930
      Begin VB.ComboBox cmbHidingKey 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":5ABC
         Left            =   2190
         List            =   "dockSettings.frx":5AE7
         TabIndex        =   238
         Text            =   "F11"
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4965
         Width           =   2620
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   555
         TabIndex        =   231
         Top             =   4050
         Width           =   6120
         Begin CCRSlider.Slider sliContinuousHide 
            Height          =   315
            Left            =   1500
            TabIndex        =   232
            ToolTipText     =   "Determine how long Steamydock will disappear when told to hide using F11"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   1
            Max             =   120
            Value           =   1
            TickFrequency   =   3
            SelStart        =   1
         End
         Begin VB.Label lblContinuousHideMsLow 
            Caption         =   "1 min"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1095
            TabIndex        =   233
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   600
         End
         Begin VB.Label lblContinuousHide 
            Caption         =   "Continuous Hide"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   236
            ToolTipText     =   "Determine how long Steamydock will disappear when told to hide for the next few minutes"
            Top             =   -30
            Width           =   1350
         End
         Begin VB.Label lblContinuousHideMsCurrent 
            Caption         =   "(30) mins"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4950
            TabIndex        =   235
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblContinuousHideMsHigh 
            Caption         =   "120m"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4440
            TabIndex        =   234
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   405
         End
      End
      Begin VB.Frame fraAutoHideType 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   615
         TabIndex        =   225
         Top             =   465
         Width           =   5100
         Begin VB.ComboBox cmbBehaviourAutoHideType 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":5B28
            Left            =   1590
            List            =   "dockSettings.frx":5B35
            TabIndex        =   230
            Text            =   "Fade"
            ToolTipText     =   "The type of auto-hide, fade, instant or a slide like Rocketdock"
            Top             =   885
            Width           =   2620
         End
         Begin VB.CheckBox chkBehaviourAutoHide 
            Caption         =   "On/Off"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1590
            TabIndex        =   229
            ToolTipText     =   "You can determine whether the dock will auto-hide or not"
            Top             =   480
            Width           =   2235
         End
         Begin VB.ComboBox cmbBehaviourActivationFX 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":5B4F
            Left            =   1590
            List            =   "dockSettings.frx":5B5C
            TabIndex        =   226
            Text            =   "Bounce"
            Top             =   0
            Width           =   2620
         End
         Begin VB.Label Label46 
            Caption         =   "Toggle AutoHide"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   228
            ToolTipText     =   "You can determine whether the dock will auto-hide or not"
            Top             =   495
            Width           =   1440
         End
         Begin VB.Label Label59 
            Caption         =   "Icon Attention Effect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   227
            ToolTipText     =   $"dockSettings.frx":5B80
            Top             =   45
            Width           =   1605
         End
      End
      Begin VB.Frame fraAutoHideDuration 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   555
         TabIndex        =   219
         Top             =   1800
         Width           =   6180
         Begin CCRSlider.Slider sliBehaviourAutoHideDuration 
            Height          =   315
            Left            =   1515
            TabIndex        =   220
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   270
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Min             =   1
            Max             =   5000
            Value           =   1
            TickFrequency   =   100
            SelStart        =   1
         End
         Begin VB.Label lblAutoHideDurationMsLow 
            Caption         =   "1ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1140
            TabIndex        =   224
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblAutoHideDurationMsHigh 
            Caption         =   "5000ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4380
            TabIndex        =   223
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblAutoHideDurationMsCurrent 
            Caption         =   "(250)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5040
            TabIndex        =   222
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblAutoHideDuration 
            Caption         =   "AutoHide Duration"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   221
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   555
         TabIndex        =   213
         Top             =   2565
         Width           =   5805
         Begin CCRSlider.Slider sliBehaviourPopUpDelay 
            Height          =   315
            Left            =   1500
            TabIndex        =   214
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   315
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   1
            Max             =   1000
            Value           =   1
            TickFrequency   =   20
            SelStart        =   1
         End
         Begin VB.Label lblAutoRevealDurationMsLow 
            Caption         =   "1ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1140
            TabIndex        =   218
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   375
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblAutoRevealDuration 
            Caption         =   "AutoReveal Duration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   217
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   0
            Width           =   1965
         End
         Begin VB.Label lblBehaviourPopUpDelayMsCurrrent 
            Caption         =   "(0)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5055
            TabIndex        =   216
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   480
         End
         Begin VB.Label lblAutoRevealDurationMsHigh 
            Caption         =   "1000ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4380
            TabIndex        =   215
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   585
         End
      End
      Begin VB.Frame fraAutoHideDelay 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   555
         TabIndex        =   207
         Top             =   3375
         Width           =   6120
         Begin CCRSlider.Slider sliBehaviourAutoHideDelay 
            Height          =   315
            Left            =   1500
            TabIndex        =   208
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Max             =   2000
            TickFrequency   =   200
         End
         Begin VB.Label lblAutoHideDelayMsLow 
            Caption         =   "3s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1140
            TabIndex        =   212
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   600
         End
         Begin VB.Label lblAutoHideDelayMsHigh 
            Caption         =   "5s"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4440
            TabIndex        =   211
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   405
         End
         Begin VB.Label lblAutoHideDelayMsCurrent 
            Caption         =   "(5) secs"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4950
            TabIndex        =   210
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblAutoHideDelay 
            Caption         =   "AutoHide Delay"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   209
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   -30
            Width           =   1350
         End
      End
      Begin VB.CheckBox chkBehaviourMouseActivate 
         Caption         =   "Pop-up on Mouseover"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4380
         TabIndex        =   206
         ToolTipText     =   "Essential functionality for the dock - pops up when  given focus"
         Top             =   8070
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame fraAnimationInterval 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   555
         TabIndex        =   184
         Top             =   6570
         Width           =   6180
         Begin CCRSlider.Slider sliAnimationInterval 
            Height          =   315
            Left            =   1575
            TabIndex        =   185
            ToolTipText     =   $"dockSettings.frx":5C12
            Top             =   285
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            Enabled         =   0   'False
            Min             =   1
            Max             =   20
            Value           =   10
            TickFrequency   =   5
            SelStart        =   1
         End
         Begin VB.Label lblAnimationIntervalMsLow 
            Caption         =   "1ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1140
            TabIndex        =   189
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   630
         End
         Begin VB.Label lblAnimationIntervalMsHigh 
            Caption         =   "20ms"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4365
            TabIndex        =   188
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblAnimationIntervalMsCurrent 
            Caption         =   "(20)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4950
            TabIndex        =   187
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblAnimationIntervalLabel 
            Caption         =   "Animation Interval"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            LinkItem        =   "150"
            TabIndex        =   186
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.Frame fraIconEffect 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   105
         TabIndex        =   119
         Top             =   945
         Width           =   5025
      End
      Begin VB.Label lblHidingKey 
         Caption         =   "Dock Hiding Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   585
         LinkItem        =   "150"
         TabIndex        =   237
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4995
         Width           =   1440
      End
      Begin VB.Label lblAnimationInformationLabel 
         Caption         =   $"dockSettings.frx":5CA1
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1740
         TabIndex        =   203
         Top             =   7530
         Width           =   4485
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "About SteamyDock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   5
      Left            =   1245
      TabIndex        =   65
      ToolTipText     =   "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
      Top             =   30
      Width           =   6930
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1545
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs used by Rocketdock"
         Top             =   420
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   795
         Width           =   1470
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1170
         Width           =   1470
      End
      Begin VB.Label Label20 
         Caption         =   "(32bit)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2985
         TabIndex        =   241
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "Windows XP, Vista, 7, 8 && 10 + ReactOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   240
         Top             =   1560
         Width           =   2955
      End
      Begin VB.Label Label10 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   239
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblSquiggle 
         Caption         =   "-oOo-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         TabIndex        =   118
         Top             =   6090
         Width           =   675
      End
      Begin VB.Label lblAboutPara6 
         AutoSize        =   -1  'True
         Caption         =   $"dockSettings.frx":5D33
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   116
         Top             =   7245
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPunklabsLink 
         BackStyle       =   0  'Transparent
         Caption         =   "                                                                                                                        "
         Height          =   225
         Index           =   0
         Left            =   2175
         MousePointer    =   2  'Cross
         TabIndex        =   112
         Top             =   870
         Width           =   1710
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
         Left            =   2175
         TabIndex        =   90
         Top             =   510
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
         Left            =   1815
         TabIndex        =   89
         Top             =   510
         Width           =   225
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
         Left            =   2535
         TabIndex        =   88
         Top             =   510
         Width           =   525
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
         Left            =   2010
         TabIndex        =   87
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblAboutPara2 
         Caption         =   $"dockSettings.frx":5E6C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   150
         TabIndex        =   78
         Top             =   2985
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAboutPara5 
         AutoSize        =   -1  'True
         Caption         =   $"dockSettings.frx":6060
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   77
         Top             =   6585
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAboutPara4 
         Caption         =   $"dockSettings.frx":60F4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   150
         TabIndex        =   76
         Top             =   5220
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label63 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   75
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label Label60 
         Caption         =   "Dean Beedell © 2018"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   74
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label Label74 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   73
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblAboutPara3 
         Caption         =   $"dockSettings.frx":61A7
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   150
         TabIndex        =   71
         Top             =   4425
         Width           =   6465
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label65 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   70
         Top             =   855
         Width           =   795
      End
      Begin VB.Label Label61 
         Caption         =   "Punklabs © 2005-2007"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   69
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblAboutPara1 
         Caption         =   $"dockSettings.frx":6291
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   150
         TabIndex        =   72
         Top             =   2250
         Width           =   6465
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Style Themes and Fonts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   3
      Left            =   1245
      TabIndex        =   50
      ToolTipText     =   "This panel allows you to change the styling of the icon labels and the dock background image"
      Top             =   15
      Width           =   6930
      Begin VB.CheckBox chkLabelBackgrounds 
         Caption         =   "Enable Label Backgrounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3525
         TabIndex        =   204
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4065
         Width           =   195
      End
      Begin VB.PictureBox picThemeSample 
         Height          =   2070
         Left            =   630
         Picture         =   "dockSettings.frx":6343
         ScaleHeight     =   2010
         ScaleWidth      =   5265
         TabIndex        =   190
         ToolTipText     =   "An example preview of the chosen theme."
         Top             =   1830
         Width           =   5325
      End
      Begin VB.Frame fraFontOpacity 
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   210
         TabIndex        =   120
         ToolTipText     =   "The theme background "
         Top             =   6750
         Width           =   6525
         Begin CCRSlider.Slider sliStyleShadowOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   121
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   750
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin CCRSlider.Slider sliStyleOutlineOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   122
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1245
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin CCRSlider.Slider sliStyleFontOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   198
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label Label34 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1635
            TabIndex        =   202
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   270
            Width           =   540
         End
         Begin VB.Label Label30 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4680
            TabIndex        =   201
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   270
            Width           =   555
         End
         Begin VB.Label lblStyleFontOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   200
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   255
            Width           =   630
         End
         Begin VB.Label Label22 
            Caption         =   "Font Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   465
            TabIndex        =   199
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   -30
            Width           =   1350
         End
         Begin VB.Label Label23 
            Caption         =   "Outline Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   450
            TabIndex        =   130
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   975
            Width           =   1365
         End
         Begin VB.Label lblStyleOutlineOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   129
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label Label35 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4665
            TabIndex        =   128
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label Label36 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1635
            TabIndex        =   127
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label Label37 
            Caption         =   "Shadow Opacity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   465
            TabIndex        =   126
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label lblStyleShadowOpacityCurrent 
            Caption         =   "(0%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5325
            TabIndex        =   125
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   765
            Width           =   630
         End
         Begin VB.Label Label39 
            Caption         =   "100%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4680
            TabIndex        =   124
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1635
            TabIndex        =   123
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   780
            Width           =   540
         End
      End
      Begin VB.PictureBox picStylePreview 
         Height          =   735
         Left            =   630
         ScaleHeight     =   675
         ScaleWidth      =   5280
         TabIndex        =   62
         ToolTipText     =   $"dockSettings.frx":9B9C
         Top             =   4440
         Width           =   5340
         Begin VB.Label lblPreviewFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2355
            TabIndex        =   63
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2400
            TabIndex        =   169
            Top             =   285
            Width           =   570
         End
         Begin VB.Label lblPreviewLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2340
            TabIndex        =   170
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblPreviewRight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2370
            TabIndex        =   171
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lblPreviewTop 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2355
            TabIndex        =   172
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblPreviewBottom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2355
            TabIndex        =   173
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2415
            TabIndex        =   174
            Top             =   285
            Width           =   570
         End
      End
      Begin VB.CommandButton btnStyleOutline 
         Caption         =   "&Outline Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "The colour of the outline, click the button to change"
         Top             =   6180
         Width           =   1470
      End
      Begin VB.CommandButton btnStyleShadow 
         Caption         =   "&Shadow Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5775
         Width           =   1470
      End
      Begin VB.CommandButton btnStyleFont 
         Caption         =   "Select &Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "The font used in the labels, click the button to change"
         Top             =   5370
         Width           =   1470
      End
      Begin VB.CheckBox chkStyleDisable 
         Caption         =   "Disable Icon Labels"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   630
         TabIndex        =   58
         ToolTipText     =   "You can toggle the icon labels on/off here"
         Top             =   4065
         Width           =   2235
      End
      Begin VB.ComboBox cmbStyleTheme 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":9C26
         Left            =   2205
         List            =   "dockSettings.frx":9C28
         TabIndex        =   51
         ToolTipText     =   "The dock background theme can be selected here"
         Top             =   405
         Width           =   2520
      End
      Begin CCRSlider.Slider sliStyleOpacity 
         Height          =   315
         Left            =   2085
         TabIndex        =   53
         ToolTipText     =   "The theme background opacity is set here"
         Top             =   900
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Max             =   100
         TickFrequency   =   10
      End
      Begin CCRSlider.Slider sliStyleThemeSize 
         Height          =   315
         Left            =   2085
         TabIndex        =   191
         ToolTipText     =   "The theme background overall size is set here"
         Top             =   1335
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   1
         Max             =   177
         Value           =   30
         TickFrequency   =   10
         SelStart        =   50
      End
      Begin VB.Label lblChkLabelBackgrounds 
         Caption         =   "Enable Label Backgrounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3795
         TabIndex        =   205
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblThemeSizeTextLow 
         Caption         =   "1px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1650
         TabIndex        =   192
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblThemeSizeText 
         Caption         =   "Theme Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   195
         ToolTipText     =   "The theme background overall size is set here"
         Top             =   1365
         Width           =   945
      End
      Begin VB.Label lblStyleSizeCurrent 
         Caption         =   "(118px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   194
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblThemeSizeTextHigh 
         Caption         =   "118px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   193
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label Label999 
         Height          =   375
         Left            =   720
         TabIndex        =   168
         Top             =   7560
         Width           =   4215
      End
      Begin VB.Label Label45 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1815
         TabIndex        =   57
         Top             =   945
         Width           =   420
      End
      Begin VB.Label lblStyleOutlineColourDesc 
         Caption         =   "Shadow Colour: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   86
         Top             =   6225
         Width           =   2700
      End
      Begin VB.Label lblStyleFontFontShadowColor 
         Caption         =   "Shadow Colour:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   85
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5820
         Width           =   2490
      End
      Begin VB.Label lblStyleFontOutlineTest 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         TabIndex        =   81
         ToolTipText     =   "The colour of the outline, click the button to change"
         Top             =   6225
         Width           =   390
      End
      Begin VB.Label lblStyleFontFontShadowTest 
         Caption         =   "Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         TabIndex        =   80
         ToolTipText     =   "The colour of the shadow, click the button to change"
         Top             =   5820
         Width           =   450
      End
      Begin VB.Label lblStyleFontName 
         Caption         =   "Font : Open Sans"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2265
         TabIndex        =   79
         ToolTipText     =   "The font used in the labels, click the button to change"
         Top             =   5445
         Width           =   3765
      End
      Begin VB.Label Label44 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   56
         Top             =   945
         Width           =   585
      End
      Begin VB.Label lblStyleOpacityCurrent 
         Caption         =   "(0%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   55
         Top             =   945
         Width           =   630
      End
      Begin VB.Label LblOpacityText 
         Caption         =   "Opacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   54
         ToolTipText     =   "The theme background opacity is set here"
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label Label21 
         Caption         =   "Theme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   645
         TabIndex        =   52
         ToolTipText     =   "The dock background theme can be selected here"
         Top             =   435
         Width           =   795
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Position the Dock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   4
      Left            =   1230
      TabIndex        =   33
      ToolTipText     =   "This panel controls the positioning of the whole dock"
      Top             =   15
      Width           =   6930
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4800
         Left            =   90
         Picture         =   "dockSettings.frx":9C2A
         ScaleHeight     =   4800
         ScaleWidth      =   3555
         TabIndex        =   114
         Top             =   3705
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.ComboBox cmbPositionLayering 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":12A09
         Left            =   2190
         List            =   "dockSettings.frx":12A16
         TabIndex        =   48
         Text            =   "Always Below"
         ToolTipText     =   "Should the dock appear on top of other windows or underneath?"
         Top             =   1905
         Width           =   2595
      End
      Begin VB.ComboBox cmbPositionMonitor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":12A3F
         Left            =   2205
         List            =   "dockSettings.frx":12A55
         TabIndex        =   37
         Text            =   "Monitor 1"
         ToolTipText     =   "Here you can determine upon which monitor the dock will appear"
         Top             =   480
         Width           =   2565
      End
      Begin VB.ComboBox cmbPositionScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":12A9B
         Left            =   2190
         List            =   "dockSettings.frx":12AAB
         TabIndex        =   36
         Text            =   "Bottom"
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1185
         Width           =   2595
      End
      Begin CCRSlider.Slider sliPositionEdgeOffset 
         Height          =   315
         Left            =   2085
         TabIndex        =   34
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3270
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   -15
         Max             =   128
         TickFrequency   =   8
      End
      Begin CCRSlider.Slider sliPositionCentre 
         Height          =   315
         Left            =   2085
         TabIndex        =   35
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2625
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   -100
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.PictureBox piccogs2 
         BorderStyle     =   0  'None
         Height          =   2970
         Left            =   3645
         Picture         =   "dockSettings.frx":12AC9
         ScaleHeight     =   2970
         ScaleWidth      =   3015
         TabIndex        =   91
         Top             =   5400
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label33 
         Caption         =   "Layering"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         TabIndex        =   49
         ToolTipText     =   "Should the dock appear on top of other windows or underneath?"
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label lblPositionMonitor 
         Caption         =   "Monitor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   47
         ToolTipText     =   "Here you can determine upon which monitor the dock will appear"
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label32 
         Caption         =   "Screen Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   675
         TabIndex        =   46
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Centre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   45
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2670
         Width           =   795
      End
      Begin VB.Label lblPositionCentrePercCurrent 
         Caption         =   "(0%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   44
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2670
         Width           =   630
      End
      Begin VB.Label Label29 
         Caption         =   "+100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4905
         TabIndex        =   43
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2670
         Width           =   585
      End
      Begin VB.Label Label28 
         Caption         =   "-100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1590
         TabIndex        =   42
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2670
         Width           =   630
      End
      Begin VB.Label Label27 
         Caption         =   "Edge Offset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   41
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3315
         Width           =   990
      End
      Begin VB.Label lblPositionEdgeOffsetPxCurrent 
         Caption         =   "(5px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5535
         TabIndex        =   40
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3300
         Width           =   630
      End
      Begin VB.Label Label25 
         Caption         =   "128px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4890
         TabIndex        =   39
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3315
         Width           =   555
      End
      Begin VB.Label Label24 
         Caption         =   "-15px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1650
         TabIndex        =   38
         ToolTipText     =   "Position from the bottom/top edge of the screen"
         Top             =   3315
         Width           =   540
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "Icon Characteristics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   1
      Left            =   1245
      TabIndex        =   93
      ToolTipText     =   "This panel allows you to set the icon sizes and hover effects"
      Top             =   30
      Width           =   6930
      Begin VB.PictureBox picSizePreview 
         Height          =   4065
         Left            =   105
         ScaleHeight     =   4005
         ScaleWidth      =   6645
         TabIndex        =   157
         Top             =   4425
         Width           =   6705
         Begin VB.PictureBox picMinSize 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1920
            Left            =   0
            ScaleHeight     =   1920
            ScaleWidth      =   1920
            TabIndex        =   159
            ToolTipText     =   "The icon size in the dock when static"
            Top             =   915
            Width           =   1920
         End
         Begin VB.PictureBox picZoomSize 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3840
            Left            =   2775
            ScaleHeight     =   3840
            ScaleWidth      =   3840
            TabIndex        =   158
            ToolTipText     =   "The maximum icon size of an animated icon"
            Top             =   15
            Width           =   3840
         End
         Begin VB.Label Label1 
            Caption         =   "Icon Sizing Preview"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   167
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Icon size fully zoomed "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4575
            TabIndex        =   166
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   3810
            Width           =   1875
         End
         Begin VB.Label Label13 
            Caption         =   "Size of icon in the dock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   285
            TabIndex        =   165
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   3810
            Width           =   1875
         End
      End
      Begin VB.Frame fraZoomConfigs 
         BorderStyle     =   0  'None
         Height          =   1110
         Left            =   195
         TabIndex        =   131
         Top             =   3165
         Width           =   6495
         Begin CCRSlider.Slider sliIconsDuration 
            Height          =   315
            Left            =   1845
            TabIndex        =   132
            ToolTipText     =   "How long the effect is applied"
            Top             =   735
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   100
            Max             =   500
            Value           =   100
            TickFrequency   =   50
            SelStart        =   100
         End
         Begin CCRSlider.Slider sliIconsZoomWidth 
            Height          =   315
            Left            =   1845
            TabIndex        =   133
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   195
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   2
            Value           =   2
            SelStart        =   2
         End
         Begin VB.Label Label19 
            Caption         =   "100ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1335
            TabIndex        =   141
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   525
         End
         Begin VB.Label Label18 
            Caption         =   "500ms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4650
            TabIndex        =   140
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   555
         End
         Begin VB.Label lblIconsDurationMsCurrent 
            Caption         =   "(200ms)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5265
            TabIndex        =   139
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   630
         End
         Begin VB.Label Label16 
            Caption         =   "Duration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   138
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label15 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1665
            TabIndex        =   137
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label14 
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4650
            TabIndex        =   136
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lblIconsZoomWidth 
            Caption         =   "(5)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5295
            TabIndex        =   135
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label12 
            Caption         =   "Zoom Width"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   465
            TabIndex        =   134
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.CheckBox chkIconsZoomOpaque 
         Caption         =   "Zoom Opaque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   96
         ToolTipText     =   "Should the zoom be opaque too?"
         Top             =   1320
         Width           =   1440
      End
      Begin VB.ComboBox cmbIconsQuality 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":169D4
         Left            =   2160
         List            =   "dockSettings.frx":169E1
         TabIndex        =   95
         Text            =   "Low quality (Faster)"
         ToolTipText     =   $"dockSettings.frx":16A23
         Top             =   390
         Width           =   2520
      End
      Begin CCRSlider.Slider sliIconsZoom 
         Height          =   315
         Left            =   2040
         TabIndex        =   94
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2775
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   1
         Max             =   256
         Value           =   1
         TickFrequency   =   32
         SelStart        =   1
      End
      Begin CCRSlider.Slider sliIconsSize 
         Height          =   315
         Left            =   2040
         TabIndex        =   97
         ToolTipText     =   "The size of each icon in the dock before any effect is applied"
         Top             =   2190
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   16
         Max             =   128
         Value           =   16
         TickFrequency   =   14
         SelStart        =   16
      End
      Begin CCRSlider.Slider sliIconsOpacity 
         Height          =   315
         Left            =   2040
         TabIndex        =   98
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   900
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         Min             =   50
         Max             =   100
         Value           =   50
         TickFrequency   =   7
         SelStart        =   50
      End
      Begin VB.Frame fraHoverEffect 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   150
         TabIndex        =   142
         Top             =   1575
         Width           =   6705
         Begin VB.ComboBox cmbIconsHoverFX 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "dockSettings.frx":16AE1
            Left            =   1995
            List            =   "dockSettings.frx":16AF4
            TabIndex        =   143
            Text            =   "None"
            ToolTipText     =   "The zoom effect to apply"
            Top             =   120
            Width           =   2595
         End
         Begin VB.Label Label7 
            Caption         =   "Hover Effect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   144
            ToolTipText     =   "The zoom effect to apply"
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.Label lblHidText3 
         Caption         =   "Some animation options are unavailable when running SteamyDock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         TabIndex        =   145
         Top             =   3585
         Width           =   5325
      End
      Begin VB.Label lblQuality 
         Caption         =   "Quality"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   111
         ToolTipText     =   "Lower power machines will benefit from the lower quality setting"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblOpacity 
         Caption         =   "Opacity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   110
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   795
      End
      Begin VB.Label lblSize 
         Caption         =   "Icon Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   109
         ToolTipText     =   "The size of each icon in the dock before any effect is applied"
         Top             =   2235
         Width           =   795
      End
      Begin VB.Label lblIconsOpacity 
         Caption         =   "(100%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   108
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label lblIconsSize 
         Caption         =   "(19px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   107
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4785
         TabIndex        =   106
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "50%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1710
         TabIndex        =   105
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label Label5 
         Caption         =   "128px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4845
         TabIndex        =   104
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "16px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   103
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label8 
         Caption         =   "Zoom Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   645
         TabIndex        =   102
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   795
      End
      Begin VB.Label lblIconsZoom 
         Caption         =   "(19px)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5490
         TabIndex        =   101
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   630
      End
      Begin VB.Label lblIconsZoomSizeMax 
         Caption         =   "256px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4845
         TabIndex        =   100
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   585
      End
      Begin VB.Label Label11 
         Caption         =   "1px"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1755
         TabIndex        =   99
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   630
      End
   End
   Begin VB.Frame fmeMain 
      Caption         =   "General Configuration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Index           =   0
      Left            =   1230
      TabIndex        =   1
      ToolTipText     =   "These are the main settings for the dock"
      Top             =   30
      Width           =   6930
      Begin VB.CheckBox genChkShowIconSettings 
         Caption         =   "Automatically display Icon Settings after adding an icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   244
         ToolTipText     =   "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
         Top             =   8145
         Width           =   225
      End
      Begin VB.CheckBox chkSplashStatus 
         Caption         =   "Show Splash Screen at Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   196
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   7815
         Width           =   255
      End
      Begin VB.Frame fraReadOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   540
         TabIndex        =   153
         Top             =   900
         Width           =   6315
         Begin VB.OptionButton optGeneralReadSettings 
            Caption         =   "Read Settings from Rocketdock's portable SETTINGS.INI (single-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   156
            ToolTipText     =   "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
            Top             =   165
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralReadRegistry 
            Caption         =   "Read Settings from RocketDock's Registry (multi-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   155
            ToolTipText     =   $"dockSettings.frx":16B34
            Top             =   465
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralReadConfig 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   154
            ToolTipText     =   $"dockSettings.frx":16BF7
            Top             =   780
            Width           =   225
         End
         Begin VB.Label lblGeneralReadConfig 
            Caption         =   "Read Settings From SteamyDock's Own Configuration Area (modern)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   630
            TabIndex        =   164
            ToolTipText     =   "Reads the configuration data from a new location that is compatible with the methods used by current Windows"
            Top             =   780
            Width           =   6135
         End
      End
      Begin VB.Frame fraRunAppIndicators 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   450
         TabIndex        =   147
         Top             =   4635
         Width           =   5955
         Begin CCRSlider.Slider sliGenRunAppInterval 
            Height          =   315
            Left            =   1020
            TabIndex        =   148
            ToolTipText     =   "The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   450
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Min             =   5
            Max             =   65
            Value           =   5
            TickFrequency   =   3
            SelStart        =   15
            Transparent     =   -1  'True
         End
         Begin VB.Label lblGenRunAppInterval2 
            Caption         =   "5s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            TabIndex        =   152
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lblGenRunAppInterval3 
            Caption         =   "65s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3960
            TabIndex        =   151
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   585
         End
         Begin VB.Label lblGenRunAppIntervalCur 
            Caption         =   "(15 seconds)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4560
            TabIndex        =   150
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label lblGenRunAppInterval1 
            Caption         =   "Running Application Check Interval"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   750
            LinkItem        =   "150"
            TabIndex        =   149
            ToolTipText     =   "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
            Top             =   120
            Width           =   3210
         End
      End
      Begin VB.CheckBox chkGenAlwaysAsk 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   945
         TabIndex        =   146
         ToolTipText     =   $"dockSettings.frx":16C8C
         Top             =   7395
         Width           =   210
      End
      Begin VB.CommandButton btnGeneralRdFolder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   5745
         TabIndex        =   115
         ToolTipText     =   "Select the folder location of Rocketdock here"
         Top             =   6975
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CheckBox chkGenRun 
         Caption         =   "Running Application Indicators"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   14
         ToolTipText     =   $"dockSettings.frx":16D25
         Top             =   4350
         Width           =   2985
      End
      Begin VB.CheckBox chkGenDisableAnim 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "If you dislike the minimise animation, click this"
         Top             =   3915
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkGenOpen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   15
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   5520
         Width           =   240
      End
      Begin VB.TextBox txtGeneralRdLocation 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   915
         TabIndex        =   84
         Text            =   "C:\programs"
         ToolTipText     =   $"dockSettings.frx":16DC4
         Top             =   6960
         Width           =   4710
      End
      Begin VB.ComboBox cmbDefaultDock 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "dockSettings.frx":16E5A
         Left            =   2085
         List            =   "dockSettings.frx":16E64
         TabIndex        =   82
         Text            =   "Rocketdock"
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
         Top             =   6255
         Width           =   2310
      End
      Begin VB.CheckBox chkGenLock 
         Caption         =   "Disable Drag/Drop and Icon Deletion (Lock the Dock)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   930
         TabIndex        =   16
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, click it!"
         Top             =   5865
         Width           =   4620
      End
      Begin VB.CheckBox chkGenMin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   915
         TabIndex        =   12
         ToolTipText     =   "This allows running applications to appear in the dock"
         Top             =   3585
         Width           =   255
      End
      Begin VB.CheckBox chkGenWinStartup 
         Caption         =   "Run at Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   915
         TabIndex        =   2
         ToolTipText     =   "This will cause the current dock to run when Windows starts"
         Top             =   360
         Width           =   1440
      End
      Begin VB.Frame fraWriteOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   510
         TabIndex        =   177
         Top             =   2340
         Width           =   6165
         Begin VB.OptionButton optGeneralWriteSettings 
            Caption         =   "Write Settings to Rocketdock's portable SETTINGS.INI (single-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   390
            TabIndex        =   180
            ToolTipText     =   "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
            Top             =   0
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralWriteRegistry 
            Caption         =   "Write Settings to RocketDock's Registry (multi-user)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   390
            TabIndex        =   179
            ToolTipText     =   $"dockSettings.frx":16E80
            Top             =   300
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralWriteConfig 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   390
            TabIndex        =   178
            ToolTipText     =   $"dockSettings.frx":16F43
            Top             =   615
            Width           =   225
         End
         Begin VB.Label lblGeneralWriteConfig 
            Caption         =   "Write Settings to SteamyDock's Own Configuration Area  (modern)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   660
            TabIndex        =   181
            ToolTipText     =   "Writes the configuration data to a new location that is compatible with the methods used by current Windows"
            Top             =   615
            Width           =   5445
         End
      End
      Begin VB.Label genLblShowIconSettings 
         Caption         =   "Automatically display Icon Settings after adding an icon to the dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1245
         TabIndex        =   243
         ToolTipText     =   "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
         Top             =   8145
         Width           =   4995
      End
      Begin VB.Label lblChkSplashStartup 
         Caption         =   "Show Splash Screen at Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1245
         TabIndex        =   197
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   7815
         Width           =   3870
      End
      Begin VB.Label lblWriteSettings 
         Caption         =   "Write to this location:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   915
         TabIndex        =   183
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label lblReadSettings 
         Caption         =   "Read from this location:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   915
         TabIndex        =   182
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lblChkAlwaysConfirm 
         Caption         =   "Utilities Always Confirm Which Dock to Configure at Startup"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1245
         TabIndex        =   163
         ToolTipText     =   $"dockSettings.frx":16FD8
         Top             =   7455
         Width           =   5310
      End
      Begin VB.Label lblChkOpenRunning 
         Caption         =   "Open Running Application Instance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   162
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   5595
         Width           =   3465
      End
      Begin VB.Label lblChkDisable 
         Caption         =   "Disable Minimise Animations"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   161
         Top             =   3975
         Width           =   2505
      End
      Begin VB.Label lblChkMinimise 
         Caption         =   "Minimise Windows to the Dock"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1185
         TabIndex        =   160
         Top             =   3660
         Width           =   3510
      End
      Begin VB.Label lblRdLocation 
         Caption         =   "Dock Folder Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   915
         TabIndex        =   113
         ToolTipText     =   $"dockSettings.frx":17071
         Top             =   6690
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Default Dock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   915
         TabIndex        =   83
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
         Top             =   6300
         Width           =   1530
      End
   End
   Begin VB.Label Label26 
      Caption         =   "Show Splash Screen at Startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2490
      TabIndex        =   242
      ToolTipText     =   "Show Splash Screen on Start-up"
      Top             =   8205
      Width           =   3870
   End
   Begin VB.Menu mnupopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About this utility"
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font selection for this utility"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with paypal"
         Index           =   2
      End
      Begin VB.Menu mnuSweets 
         Caption         =   "Donate some sweets/candy with Amazon"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuHelpPdf 
            Caption         =   "View Help (HTML)"
         End
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version"
         End
         Begin VB.Menu mnuMoreIcons 
            Caption         =   "Visit Deviantart to download some more Icons"
         End
         Begin VB.Menu mnuWidgets 
            Caption         =   "See the complementary steampunk widgets"
         End
         Begin VB.Menu mnuFacebook 
            Caption         =   "Chat about SteamyDock functionality on Facebook"
         End
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuseparator1 
         Caption         =   ""
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Turn Debugging ON"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close this Program"
      End
   End
End
Attribute VB_Name = "dockSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Changes:
' 25/10/2020 docksettings .01 DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
' 26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
' 26/10/2020 docksettings .04 DAEB added a caption change to autohide toggle checkbox using the IDE only
' 26/10/2020 docksettings .05 DAEB added a manual click to the autohide toggle checkbox
' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
' 23/01/2021 docksettings .07 DAEB Added themeing to two new sliders
' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
' .09 DAEB 01/02/2021 docksettings Make the sample image functionality disabled for rocketdock
' .10 DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
' .12 DAEB 26/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation
' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock
' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD

'
' Status:

' There is a bug with the absence of a settings file when switching between config locations
' this is to do with one option deleting a file and the other expecting it to exist

' dock hiding rdToggle keypress should only contain the rocketdock default keypress when defaultdock =0


' Credits :
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           KillApp code from an unknown, untraceable source, possibly on MSN.
'           Registry reading code from ALLAPI.COM.
'           Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.
'
'           Elroy on VB forums for his Persistent debug window
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
'           Krool on the VBForums for his impressive common control replacements
'           si_the_geek for his special folder code
'           Gary Beene        Get list of drive letters https://www.garybeene.com/code/visual%20basic145.htm
'
' NOTE - Do not END this program within the IDE as GDI will not release memory and usage will grow and grow
' ALWAYS use the QUIT option on the application right click menu.

' NOTE - When adding new slider controls remember to add them to the themeing menu option for light/high contrast

' NOTE - When building the binary, ensure that the ccrslider.ocx is in the docksettings folder
'         The manifest should be modified to incorporate the ocx

' SETTINGS: There are four settings files:
' o The first is the RD settings file SETTINGS.INI that only exists if RD is NOT using the registry

' NOTE: Rocketdock overwrites its own settings.ini when it closes meaning that we have to always work on a copy.
' In addition, when SD determines that RD is using the registry it extracts the data and creates a temporary copy of the settings file that we work on.
' In this manner we are always working on a .ini file in the same manner only writing it back to the registry when the user hits 'save & restart' or 'apply'.

' o The second is our tools copy of RD's settings file, we copy the original or create our own from RD's registry settings
' o The third is the settings file for this tool only, to store its own local preferences.

' origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file in program files.
' tmpSettingsFile = App.path & "\tmpSettings.ini" ' a temporary copy of the settings file that we work on.
' toolSettingsFile = SpecialFolder_AppData & <utilityName> "\settings.ini" the tool's own settings file.

' o The fourth settings file is the dockSettings.ini that sits in this location:
' C:\Users\<username>\AppData\Roaming\steamyDock\
'
' When the flag to write the 3rd settings file is set in the dock settings utility,
' we will write the rocketdock variable values to this file.
'
' docksettings.ini is partitioned as follows:
'
' [Software\SteamyDock\DockSettings] - the dockSettings tool writes here
' [Software\SteamyDock\IconSettings] - the iconSettings tool writes here
' [Software\SteamyDock\SteamyDockSettings] - the dock itself could write here but in reality it will most likely be in the areas above
'
' re: toolSettingsFile - The utilities read their own config files for their own personal set up in their own folders in appdata
' Settings.ini, this is just for local settings that concern only the utility, look and feel
'
' eg.
' C:\Users\<username>\AppData\Roaming\dockSettings\settings.ini
'
' toolSettingsFile - Dock - the following items are currently inserted into the toolSettingsFile for the dockSettings utility
'
' [Software\SteamyDockSettings]
' defaultStrength = 400
' defaultStyle = False
' defaultFont=Centurion Light SF

' toolSettingsFile - Icons - the following items are currently inserted into the toolSettingsFile for the iconSettings utility

' [Software\SteamyDockSettings]
' defaultFolderNodeKey=C:\Program Files (x86)\SteamyDock\iconSettings\my collection ' this could be moved to the docksettings.ini later
' rdMapState = Visible ' as could this
' defaultSize = 8
' defaultStrength = False
' defaultStyle = False
' Quality = 1
' defaultFont=Centurion Light SF
'
' The registry and the original settings.ini that Rocketdock provides for variable storage are
' left-overs from XP days when the registry storage was trendy and encouraged, the use of program files
' for the settings.ini was a left-over from the days before the registry when settings were stored locally
' in the program files folder with no folder security. This program allows access to these to retain
' compatibility with Rocketdock but offers the fourth storage option to be compatible with modern windows requirements.

' Separate labels for checkboxes
' the reason there is a separate label for certain checkboxes is due to the way that VB6 greys out checkbox labels using specific fonts
' causing them to appear 'crinkled'. When a discrete label is created that is unattached to the chkbox then it greys out correctly.

Option Explicit

'Simulate MouseEnter event to reset the icons on one frame
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

'API to test whether the user is running as an administrator account
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer



Private busyCounter As Integer
Private totalBusyCounter As Integer

Private Const COLOR_BTNFACE As Long = 15

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean

' Flag for debug mode
Private mbDebugMode As Boolean  ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without






Private Sub chkLabelBackgrounds_Click()

   sDShowLblBacks = chkLabelBackgrounds.Value ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files

    ' add a background to the icon titles in dock's drawtext function
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbHidingKey_Click ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
' Author    : beededea
' Date      : 26/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbHidingKey_Click()
   On Error GoTo cmbHidingKey_Click_Error

    rDHotKeyToggle = cmbHidingKey.Text

   On Error GoTo 0
   Exit Sub

cmbHidingKey_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbHidingKey_Click of Form dockSettings"
End Sub

Private Sub fraAnimationInterval_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 16/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
   On Error GoTo btnHelp_Click_Error

    Call mnuHelpPdf_click
    
   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSplashStatus_Click
' Author    : beededea
' Date      : 01/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkSplashStatus_Click()
   On Error GoTo chkSplashStatus_Click_Error

    If chkSplashStatus.Value = 1 Then
        sDSplashStatus = "1"
    Else
        sDSplashStatus = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkSplashStatus_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkSplashStatus_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbBehaviourAutoHideType_Click
' Author    : beededea
' Date      : 25/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbBehaviourAutoHideType_Click()


   On Error GoTo cmbBehaviourAutoHideType_Click_Error

    sDAutoHideType = cmbBehaviourAutoHideType.ListIndex
    
    If cmbBehaviourAutoHideType.ListIndex = 2 Then
        lblAutoHideDuration.Enabled = False
        lblAutoHideDurationMsLow.Enabled = False
        sliBehaviourAutoHideDuration.Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblAutoRevealDuration.Enabled = False
        lblAutoRevealDurationMsLow.Enabled = False
        lblAutoRevealDurationMsHigh.Enabled = False
        sliBehaviourPopUpDelay.Enabled = False
        lblBehaviourPopUpDelayMsCurrrent.Enabled = False
        
    Else
        lblAutoHideDuration.Enabled = True
        lblAutoHideDurationMsLow.Enabled = True
        sliBehaviourAutoHideDuration.Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblAutoRevealDuration.Enabled = True
        lblAutoRevealDurationMsLow.Enabled = True
        lblAutoRevealDurationMsHigh.Enabled = True
        sliBehaviourPopUpDelay.Enabled = True
        lblBehaviourPopUpDelayMsCurrrent.Enabled = True
       
    End If
    

   On Error GoTo 0
   Exit Sub

   On Error GoTo 0
   Exit Sub

cmbBehaviourAutoHideType_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbBehaviourAutoHideType_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : Load the dockSettings form
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
    ' variables declared
    Dim NameProcess As String
    Dim AppExists As Boolean
    Dim answer As VbMsgBoxResult
    
    ' initial values assigned
    NameProcess = ""
    AppExists = False
    answer = vbNo
    
    ' other variable assignments
    defaultDock = 0
    debugflg = 0
    startupFlg = True
    rdAppPath = ""
    busyCounter = 1
    totalBusyCounter = 1
    
    mnupopmenu.Visible = False

    On Error GoTo Form_Load_Error
    If debugflg = 1 Then DebugPrint "%Form_Load"
    
    Call ShowDevices(sAllDrives)
                           
    'if the process already exists then kill it
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "docksettings.exe"
        checkAndKill NameProcess, False
        'MsgBox "You now have two instances of this utility running, they will conflict..."
    End If
    
    ' the frames can jump about in the IDE during development, this just places them accurately at runtime
    Call placeFrames
    
    'load the about text
    Call loadAboutText
      
    ' get the location of this tool's settings file
    Call getToolSettingsFile
    
    ' check the Windows version
    Call testWinVer(classicThemeCapable)
    
    'MsgBox ProgramFilesDir
    ' turn on the timer that tests every 10 secs whether the visual theme has changed
    ' only on those o/s versions that need it
    
    If classicThemeCapable = True Then
        dockSettings.mnuAuto.Caption = "Auto Theme Disable"
        dockSettings.themeTimer.Enabled = True
    Else
        dockSettings.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"
        dockSettings.themeTimer.Enabled = False
    End If
    
    ' admin is required to read the registry and access the settings.ini in RD's program folder
    If IsUserAnAdmin() = 0 Then
        MsgBox "This tool requires to be run as administrator on Windows 7 and above in order to function. Admin access is NOT required on Win7 and below. If you aren't entirely happy with that then you'll need to remove the software now. This is a limitation imposed by Windows itself. To enable administrator access find this tool's exe and right-click properties, compatibility - run as administrator. YOU have to do this manually, I can't do it for you."
    End If

    ' check where rocketdock is installed
    Call checkRocketdockInstallation
    If rocketDockInstalled = True Then
        dockAppPath = rdAppPath
        txtGeneralRdLocation.Text = rdAppPath
        defaultDock = 0
    End If
    
    ' we check to see if rocketdock is installed in order to know the location of the settings.ini file used by Rocketdock
    'If rocketdock Is Not installed Then test the registry read
    ' if the registry settings are located then offer them as a choice.
        
    ' check where steamyDock is installed
    Call checkSteamyDockInstallation
    
    'update a filed with the installation details
    txtGeneralRdLocation.Text = dockAppPath
    
    ' if both docks are installed we need to determine which is the default
    Call checkDefaultDock
    
    'load the resizing image into a hidden picbox
    picHiddenPicture.Picture = LoadPicture(App.Path & "\gpu-z-256.gif")
    
    'read the correct config location according to the default selection
    Call readDockConfiguration
    
    ' RD can use the different monitors, SD cannot yet.
    Call GetMonitorCount
    
    ' read the local tool settings file and do some local things for the first and only time
    Call readAndSetUtilityFont
    
    ' display the version number on the general panel
    Call displayVersionNumber
    
    ' click on the panel that is set by default
    Call picIcon_Click(Val(rDOptionsTabIndex) - 1)
    
    ' set the theme on startup
    Call setThemeSkin
    
    startupFlg = False ' now negate the startup flag

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form dockSettings"
     
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   If debugflg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    mnuDebug_Click

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuSweets_Click

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : busyTimer_Timer
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : rotates the hourglass timer
'---------------------------------------------------------------------------------------
'
Private Sub busyTimer_Timer()
        Dim thisWindow As Long
        Dim busyFilename As String
        
        On Error GoTo busyTimer_Timer_Error

        thisWindow = FindWindowHandle("SteamyDock")
        busyFilename = ""
        
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        totalBusyCounter = totalBusyCounter + 1
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        If classicTheme = True Then
            busyFilename = App.Path & "\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.Path & "\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        
        If thisWindow <> 0 And totalBusyCounter >= 50 Then
            busyTimer.Enabled = False
            busyCounter = 1
            totalBusyCounter = 1
            picBusy.Visible = False
        End If

   On Error GoTo 0
   Exit Sub

busyTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure busyTimer_Timer of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenAlwaysAsk_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenAlwaysAsk_Click()
   On Error GoTo chkGenAlwaysAsk_Click_Error

    rDAlwaysAsk = chkGenAlwaysAsk.Value

   On Error GoTo 0
   Exit Sub

chkGenAlwaysAsk_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenAlwaysAsk_Click of Form dockSettings"

End Sub



Private Sub fraAutoHideDelay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHideDuration_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub





'---------------------------------------------------------------------------------------
' Procedure : genChkShowIconSettings_Click
' Author    : beededea
' Date      : 01/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub genChkShowIconSettings_Click()

    On Error GoTo genChkShowIconSettings_Click_Error
    
    If genChkShowIconSettings.Value = 1 Then
        sDShowIconSettings = "1"
    Else
        sDShowIconSettings = "0"
    End If
    
    On Error GoTo 0
    Exit Sub

genChkShowIconSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure genChkShowIconSettings_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : genLblShowIconSettings_Click
' Author    : beededea
' Date      : 01/05/2021
' Purpose   : click on the additional label = we have these as they can be style correctly unlike the default VB6 ones
'---------------------------------------------------------------------------------------
'
Private Sub genLblShowIconSettings_Click()

    On Error GoTo genLblShowIconSettings_Click_Error
    
    If genChkShowIconSettings.Value = 1 Then
        genChkShowIconSettings.Value = 0
    Else
        genChkShowIconSettings.Value = 1
    End If
    
    On Error GoTo 0
    Exit Sub

genLblShowIconSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure genLblShowIconSettings_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblChkAlwaysConfirm_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblChkAlwaysConfirm_Click()
    
   On Error GoTo lblChkAlwaysConfirm_Click_Error

    If chkGenAlwaysAsk.Value = 1 Then
        chkGenAlwaysAsk.Value = 0
    Else
        chkGenAlwaysAsk.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblChkAlwaysConfirm_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblChkAlwaysConfirm_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblChkDisable_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblChkDisable_Click()

   On Error GoTo lblChkDisable_Click_Error

    If chkGenDisableAnim.Value = 1 Then
        chkGenDisableAnim.Value = 0
    Else
        chkGenDisableAnim.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblChkDisable_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblChkDisable_Click of Form dockSettings"
End Sub

Private Sub lblChkLabelBackgrounds_Click()
' the reason there is a separate label for certain checkboxes is due to the way that VB6 greys out checkbox labels using specific fonts causing them to be crinkled. When the label is unattached to the chkbox then it greys out correctly.
    Call chkLabelBackgrounds_Click
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblChkMinimise_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblChkMinimise_Click()
   On Error GoTo lblChkMinimise_Click_Error

    If chkGenMin.Value = 1 Then
        chkGenMin.Value = 0
    Else
        chkGenMin.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblChkMinimise_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblChkMinimise_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblChkOpenRunning_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblChkOpenRunning_Click()

   On Error GoTo lblChkOpenRunning_Click_Error

    If chkGenOpen.Value = 1 Then
        chkGenOpen.Value = 0
    Else
        chkGenOpen.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblChkOpenRunning_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblChkOpenRunning_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblChkSplashStartup_Click
' Author    : beededea
' Date      : 01/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblChkSplashStartup_Click()
   On Error GoTo lblChkSplashStartup_Click_Error

    If chkSplashStatus.Value = 1 Then
        chkSplashStatus.Value = 0
    Else
        chkSplashStatus.Value = 1
    End If

   On Error GoTo 0
   Exit Sub

lblChkSplashStartup_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblChkSplashStartup_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblGeneralReadConfig_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGeneralReadConfig_Click()
   On Error GoTo lblGeneralReadConfig_Click_Error
   
   'optGeneralReadConfig_Click ' no good

    If optGeneralReadConfig.Value = False Then
        optGeneralReadConfig.Value = True
    End If

   On Error GoTo 0
   Exit Sub

lblGeneralReadConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGeneralReadConfig_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblGeneralWriteConfig_Click
' Author    : beededea
' Date      : 30/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGeneralWriteConfig_Click()

   On Error GoTo lblGeneralWriteConfig_Click_Error

    If optGeneralWriteConfig.Value = False Then
        optGeneralWriteConfig.Value = True
    End If

   On Error GoTo 0
   Exit Sub

lblGeneralWriteConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGeneralWriteConfig_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadConfig_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : set a value to a flag that indicates we will read from the 3rd settings file
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadConfig_Click()


   On Error GoTo optGeneralReadConfig_Click_Error

    If startupFlg = True Then '.NET
        ' don't do this on the first startup run
        Exit Sub

    End If
    
    If chkGenRun.Value = 1 Then
        lblGenRunAppInterval1.Enabled = True
        lblGenRunAppInterval2.Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenRunAppInterval3.Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
    If optGeneralReadConfig.Value = True And defaultDock = 1 And steamyDockInstalled = True And rocketDockInstalled = True Then
        chkGenAlwaysAsk.Enabled = True
        lblChkAlwaysConfirm.Enabled = True
    End If
    
    rDGeneralReadConfig = optGeneralReadConfig.Value ' this is the nub
    
    'Call locateDockSettingsFile

   On Error GoTo 0
   Exit Sub

optGeneralReadConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadConfig_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadRegistry_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadRegistry_Click()
   On Error GoTo optGeneralReadRegistry_Click_Error

        If optGeneralReadRegistry.Value = True Then
            ' nothing to do, the checkbox value is used later to determine where to write the data
        End If
        If defaultDock = 0 Then optGeneralWriteRegistry.Value = True ' if running Rocketdock the two must be kept in sync
        lblGenRunAppInterval1.Enabled = False
        lblGenRunAppInterval2.Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenRunAppInterval3.Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
        chkGenAlwaysAsk.Enabled = False
        lblChkAlwaysConfirm.Enabled = False
        
        rDGeneralReadConfig = optGeneralReadConfig.Value ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralReadRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadRegistry_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralReadSettings_Click
' Author    : beededea
' Date      : 04/03/2020
' Purpose   : The existence of a file in the rocketdock program files location is the sole flag that Rocketdocki
'             uses to determine whether it should write the settings to the registry or the settings file.
'             The file is settings.ini.
'
'             I had expected a flag in the registry but none exists... When I created a file in the
'             Rocketdock folder and it used it straight away.
'
'             The changes only come into effect on a click of the 'apply' button.
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralReadSettings_Click()

   On Error GoTo optGeneralReadSettings_Click_Error
   If debugflg = 1 Then Debug.Print "%optGeneralReadSettings_Click"
   
    tmpSettingsFile = rdAppPath & "\tmpSettings.ini" ' temporary copy of Rocketdock 's settings file
    
    If startupFlg = True Then '.NET
        ' don't do this on the first startup run
        Exit Sub
    Else
        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
            ' create a settings.ini file in the rocketdock folder
            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
            Close #1         ' close the file and
             ' test it exists
            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
            End If
        End If
    End If
        
    lblGenRunAppInterval1.Enabled = False
    lblGenRunAppInterval2.Enabled = False
    sliGenRunAppInterval.Enabled = False
    lblGenRunAppInterval3.Enabled = False
    lblGenRunAppIntervalCur.Enabled = False
    chkGenAlwaysAsk.Enabled = False
    lblChkAlwaysConfirm.Enabled = False
    
    rDGeneralReadConfig = optGeneralReadConfig.Value ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralReadSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadSettings_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readIconsWriteSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the registry icon store one line at a time and create a temporary settings file
'---------------------------------------------------------------------------------------
'
Private Sub readIconsWriteSettings(location As String, settingsFile As String)
    
    ' variables declared
    Dim useloop As Integer
    Dim regRocketdockSection As String
    Dim theCount As Integer
    
    ' initial values assigned
     useloop = 0
     regRocketdockSection = ""
     theCount = 0
    
    On Error GoTo readIconsWriteSettings_Error
    If debugflg = 1 Then DebugPrint "%" & "readIconsWriteSettings"
    
    'initialise the dimensioned variables
    useloop = 0
    regRocketdockSection = ""
    theCount = 0
        
    ' get items from the registry that are required to 'default' the dock but aren't controlled by the dock settings utility
    theCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count") 'dean
    rdIconCount = theCount - 1
            
    ' first we read and write the individual icon data values
    For useloop = 0 To rdIconCount
         ' get the relevant entries from the registry
         readRegistryIconValues (useloop)
         ' read the rocketdock alternative settings.ini
         Call writeIconSettingsIni(location & "\Icons", useloop, settingsFile) ' the alternative settings.ini exists when RD is set to use it
     Next useloop
     
    PutINISetting location & "\Icons", "Count", theCount, settingsFile
    
   On Error GoTo 0
   Exit Sub

readIconsWriteSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconsWriteSettings of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : readRegistryIconValues
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read the registry and set obtain the necessary icon data for the specific icon
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryIconValues(ByVal iconNumberToRead As Integer)
    ' read the settings from the registry
   On Error GoTo readRegistryIconValues_Error
   If debugflg = 1 Then DebugPrint "%" & "readRegistryIconValues"
  
    sFilename = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName")
    sFileName2 = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName2")
    sTitle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Title")
    sCommand = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Command")
    sArguments = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Arguments")
    sWorkingDirectory = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-WorkingDirectory")
    sShowCmd = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-ShowCmd")
    sOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-OpenRunning")
    sIsSeparator = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-IsSeparator")
    sUseContext = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseContext")
    sDockletFile = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-DockletFile")

   On Error GoTo 0
   Exit Sub

readRegistryIconValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryIconValues of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : fmeSizePreview_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeSizePreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo fmeSizePreview_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%fmeSizePreview_MouseDown"
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

fmeSizePreview_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeSizePreview_MouseDown of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkDefaultDock
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : if both rocketdock and steamydock are installed, then asks which dock you would like to maintain/configure
'---------------------------------------------------------------------------------------
'
Private Sub checkDefaultDock()

    ' variables declared
    Dim answer As VbMsgBoxResult
        
    ' initial values assigned
     answer = vbNo
    
   On Error GoTo checkDefaultDock_Error
   
    'initialise the dimensioned variables
    answer = vbNo
    
    If steamyDockInstalled = True Then
        ' get the location of the dock's new settings file
        Call locateDockSettingsFile
        chkGenAlwaysAsk.Value = Val(GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile))
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
        If rDGeneralReadConfig <> "" Then
            optGeneralReadConfig.Value = rDGeneralReadConfig
        Else
            optGeneralReadConfig.Value = False
        End If
    End If
    
    If steamyDockInstalled = True And rocketDockInstalled = True Then
        If chkGenAlwaysAsk.Value = 1 Then  ' depends upon being able to read the new configuration file in the user data area
            answer = MsgBox("Both Rocketdock and SteamyDock are installed on this system. Use SteamyDock by default? ", vbYesNo)
            If answer = vbYes Then
                cmbDefaultDock.ListIndex = 1 ' steamy dock
                dockAppPath = sdAppPath
                txtGeneralRdLocation.Text = sdAppPath
                defaultDock = 1
            Else
                cmbDefaultDock.ListIndex = 0 ' rocket dock
                dockAppPath = rdAppPath
                txtGeneralRdLocation.Text = rdAppPath
                defaultDock = 0
            End If
        Else
            ' if the question is not being asked then use the default dock as specified in the docksettings.ini file
            If rDDefaultDock = "steamydock" Then
                cmbDefaultDock.ListIndex = 1
                dockAppPath = sdAppPath
                txtGeneralRdLocation.Text = dockAppPath
                defaultDock = 1
            ElseIf rDDefaultDock = "rocketdock" Then
                cmbDefaultDock.ListIndex = 0 ' rocket dock
                dockAppPath = rdAppPath
                txtGeneralRdLocation.Text = rdAppPath
                defaultDock = 0
            Else
                If cmbDefaultDock.ListIndex = 1 Then  ' depends upon being able to read the new configuration file in the user data area
                    dockAppPath = sdAppPath
                    txtGeneralRdLocation.Text = dockAppPath
                    defaultDock = 1
                Else
                    cmbDefaultDock.ListIndex = 0 ' rocket dock
                    dockAppPath = rdAppPath
                    txtGeneralRdLocation.Text = rdAppPath
                    defaultDock = 0
                End If
            End If
        End If
    ElseIf steamyDockInstalled = True Then ' just steamydock installed
            cmbDefaultDock.ListIndex = 1
            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
            
            dockAppPath = sdAppPath
            txtGeneralRdLocation.Text = dockAppPath
            defaultDock = 1
            ' write the default dock to the SteamyDock settings file
            PutINISetting "Software\SteamyDockSettings", "defaultDock", defaultDock, toolSettingsFile
            
    ElseIf rocketDockInstalled = True Then ' just rocketdock installed
            cmbDefaultDock.ListIndex = 0
            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
            
            dockAppPath = rdAppPath
            txtGeneralRdLocation.Text = rdAppPath
            defaultDock = 0
    End If
    
    ' it is possible to run this program without steamydock being installed
    If steamyDockInstalled = False And rocketDockInstalled = False Then
        answer = MsgBox(" Neither Rocketdock nor SteamyDock has been installed on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
         Dim ofrm As Form
         For Each ofrm In Forms
             Unload ofrm
         Next
         End
    End If

   On Error GoTo 0
   Exit Sub

checkDefaultDock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkDefaultDock of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readDockConfiguration
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : read the configurations, settings.ini, registry and dockSettings.ini
'---------------------------------------------------------------------------------------
'
Private Sub readDockConfiguration()
    ' select the settings source STARTS
            
   On Error GoTo readDockConfiguration_Error

    'final check to be sure that we aren't using an incorrectly set dockSettings.ini file when RD has never been installed
    If rocketDockInstalled = False And RDregistryPresent = False Then
        rDGeneralReadConfig = True
        optGeneralReadConfig.Value = True
    End If

    If steamyDockInstalled = True And defaultDock = 1 And optGeneralReadConfig.Value = True Then ' it will always exist even if not used
        ' read the dock settings from the new configuration file
        Call readDockSettingsFile("Software\SteamyDock\DockSettings", dockSettingsFile)
        Call validateInputs
        Call adjustControls
                            
        If defaultDock = 0 Then
            rDVersion = "1.3.5"
        Else
            rDVersion = App.Major & "." & App.Minor & "." & App.Revision
        End If
    End If
    
    If optGeneralReadConfig.Value = False Then
        ' read the dock settings from INI or from registry
        Call readDockSettings
        Call adjustControls
    End If
    
    'if rocketdock set the automatic startup string to Rocketdock
    If defaultDock = 0 Then ' rocketdock
        rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "RocketDock")
        If rdStartupRunString <> "" Then
            rDStartupRun = "1"
            chkGenWinStartup.Value = 1
        End If
    ElseIf defaultDock = 1 Then 'if rocketdock set the automatic startup string to Steamydock
        rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SteamyDock")
        If rdStartupRunString <> "" Then
            rDStartupRun = "1"
            chkGenWinStartup.Value = 1
        End If
    End If

   On Error GoTo 0
   Exit Sub

readDockConfiguration_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockConfiguration of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadAboutText
' Author    : beededea
' Date      : 12/03/2020
' Purpose   : The text for the abour page is stored here
'---------------------------------------------------------------------------------------
'
Sub loadAboutText()

   On Error GoTo loadAboutText_Error
   If debugflg = 1 Then Debug.Print "%loadAboutText"

    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project to allow easier configuration, bug-fixing and enhancement of Rocketdock and currently underway, a fully open source version of a Rocketdock clone."

    lblAboutPara4.Caption = "The first steps are the two VB6 utilities that replicate the icons settings and dock settings screen. The subsequent step is the dock itself. I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received."

    lblAboutPara1.Caption = "The original Rocketdock was developed by the Apple fanboy and fangirl team at Punklabs. They developed it as a peace offering from the Mac community to the Windows Community."
    lblAboutPara2.Caption = "This new dock, now known as SteamyDock, was developed by a Windows/ReactOS fanboy on Windows 7 using VB6. This utility faithfully reproduces the original as created by Punklabs, originally done solely as a homage to the original as that version is no longer being supported but now it has evolved into a set of tools that has become a replacement for rocketdock itself. It must be said, the initial idea for this dock came from Punklabs and Rocketdock's OS/X dock predecessors. All HAIL to Punklabs!"
    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project. It is open source to allow easier configuration, bug-fixing and enhancement of Rocketdock and community contribution towards this new dock."
    lblAboutPara4.Caption = "The first steps were the two VB6 utilities that replicate the icons settings and dock settings screen (this utility). These are largely complete and the dock itself is now under development and 90% complete. A future step is conversion to VB.NET for future-proofing and 64bit-ness. This next step is 1/3rd underway."
    
    lblAboutPara5.Caption = "I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received if you choose to contribute."
    lblAboutPara6.Caption = "This utility MUST run as administrator in order to access Rocketdock's " & _
                            "registry settings (due to a Windows shadow registry feature/bug that " & _
                            "gives incorrect shadow data). If you run it without admin rights and " & _
                            "you want to change the values in the registry then some of the values may " & _
                            "be incorrect and the resulting dock might look and act rather strange. " & _
                            "You have been warned!"

   On Error GoTo 0
   Exit Sub

loadAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAboutText of Form dockSettings"
    
End Sub
    
    





   
'---------------------------------------------------------------------------------------
' Procedure : InIDE
' Author    : beededea
' Date      : 03/03/2020
' Purpose   : There are ocasions when the program will act differently when running in the IDE
'             We need to know when.  Compatibility mode means that it believes it is running under XP and will return as such.
'---------------------------------------------------------------------------------------
'
Function InIDE() As Boolean
'Returns whether we are running in vb(true), or compiled (false)
 
   On Error GoTo InIDE_Error
   If debugflg = 1 Then Debug.Print "%InIDE"

    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        InIDE = True
    End If

   On Error GoTo 0
   Exit Function

InIDE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InIDE of Form dockSettings"
 
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : InDebugMode
' Author    : beededea
' Date      : 02/03/2021
' Purpose   : ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
'---------------------------------------------------------------------------------------
'
Private Function InDebugMode() As Boolean
   On Error GoTo InDebugMode_Error

    mbDebugMode = True
    InDebugMode = True

   On Error GoTo 0
   Exit Function

InDebugMode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InDebugMode of Form dockSettings"
End Function
'---------------------------------------------------------------------------------------
' Procedure : btnApply_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : Apply the registry or settings.ini
'---------------------------------------------------------------------------------------
'
Private Sub btnApply_Click()
    
    ' variables declared
    Dim NameProcess As String
    Dim ans As Boolean
    Dim answer As VbMsgBoxResult
    Dim positionZeroFail As Boolean
    Dim positionThreeFail As Boolean
    Dim debugPoint As Integer

    On Error GoTo btnApply_Click_Error
    If debugflg = 1 Then DebugPrint "%btnApply_Click"
   
   'initialise the dimensioned variables
    NameProcess = ""
    ans = False
    answer = vbNo
    positionZeroFail = False
    positionThreeFail = False
    debugPoint = 0
   
    If InIDE = True Then
        If optGeneralReadRegistry.Value = True Then
            answer = MsgBox("Running in the IDE. The registry values may corrupt - be warned. Continue?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    Else
        If IsUserAnAdmin() = 0 Then
            answer = MsgBox("This program is not running as admin. Some of the settings may be strange and unwanted - be warned. Continue?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
   
    ' kill the rocketdock /steamydock process first
    
    If defaultDock = 0 Then
        NameProcess = "RocketDock.exe"
    Else
        NameProcess = "steamyDock.exe"
    End If
    ans = checkAndKill(NameProcess, False)
            
    ' if the settings.ini has been chosen as an option then the creation of it will already have occurred,
    ' so, if the temporary settings file exists then it means that the user clicked "use settings.ini file"
    ' in which case we copy it to the main settings.ini file.
    
    debugPoint = 1
    ' Steamydock exists so we shall write to the settings file those additonal items that need to be there regardless of the location of the dock data
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralReadConfig", rDGeneralReadConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralWriteConfig", rDGeneralWriteConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "RunAppInterval", rDRunAppInterval, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "DefaultDock", rDDefaultDock, dockSettingsFile
    
    If optGeneralWriteConfig.Value = True Then ' the 3rd option
        debugPoint = 2

        ' writes to the new config file
        Call writeDockSettings("Software\SteamyDock\DockSettings", dockSettingsFile)
        
        ' the docksettings tool only writes the basic 'dock' configuration
        ' however, if the 'icon' settings do not exist in the 3rd config option then the actual dock will fail to show any icons
        ' (the other icon settings tool is meant to write the icon data but that tool may not yet have been run).
        
        ' in the unlikely event that this program is run before the main dock program, there is a chance that the dockSettings.ini
        ' will not have been created previously and may not contain any icon details. This next bit tests the docksettings.ini
        ' file for valid icon records.
        
        'test the first record
        sFilename = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-FileName", dockSettingsFile)
        sTitle = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-Title", dockSettingsFile)
        sCommand = GetINISetting("Software\SteamyDock\IconSettings\Icons", "0-Command", dockSettingsFile)
        If sFilename = "" And sTitle = "" And sCommand = "" Then positionZeroFail = True

         
        'test the third record - it assumes all docks will have at least three elements and therfore three records
        sFilename = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-FileName", dockSettingsFile)
        sTitle = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-Title", dockSettingsFile)
        sCommand = GetINISetting("Software\SteamyDock\IconSettings\Icons", "3-Command", dockSettingsFile)
        If sFilename = "" And sTitle = "" And sCommand = "" Then positionThreeFail = True
        
        ' the dock icon settings are empty?
        If positionZeroFail = True And positionThreeFail = True Then
            If FExists(dockSettingsFile) Then ' does the temporary settings.ini exist?
                ' read the registry values for each of the icons and write them to the settings.ini
                Call readIconsWriteSettings("Software\SteamyDock\IconSettings", dockSettingsFile)
            End If
        End If

    Else
        debugPoint = 3
        If rocketDockInstalled = True Then
          If optGeneralWriteSettings.Value = True Then ' use the settings file
            debugPoint = 4
            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
                Call writeDockSettings("Software\RocketDock", tmpSettingsFile)
                ' if it exists, read the registry values for each of the icons and write them to the settings.ini
                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
            End If
            If FExists(tmpSettingsFile) Then ' does the tmp settings.ini exist?
                debugPoint = 5
                If FExists(origSettingsFile) Then ' does the tmp settings.ini exist?
                    debugPoint = 6
                    Kill origSettingsFile
                End If
                debugPoint = 7
                Name tmpSettingsFile As origSettingsFile ' rename 'our' settings file to the one used by RD
            End If
          Else ' WRITE THE REGISTRY AND remove the settings file
            Call writeRegistry
            ' this function restarts Rocketdock so that the changes 'take'.
            Sleep (1300) ' this is required as the o/ses' final commit of the data to the registry can be delayed
                         ' and without the pause the restart does not picku p the committed data.
            'see if the settings.ini exists
            ' if it does exist, ensure it no longer does so by deleting it, RD will then use the registry.
            If FExists(origSettingsFile) Then ' does the original settings.ini exist?
                    Kill origSettingsFile
            End If
          End If
        End If
    End If

  
    ' From the general panel
    ' these write to registry areas available to any program not just Rocketdock
    
    picBusy.Visible = True
    busyTimer.Enabled = True
    
    If rDStartupRun = "1" Then
        'if rocketdock set the string to Rocketdock startup
        If defaultDock = 0 Then ' rocketdock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RocketDock", """" & txtGeneralRdLocation.Text & "\" & "RocketDock.exe""")
        End If
        'if steamydock set the string to steamydock startup
        If defaultDock = 1 Then ' steamydock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteamyDock", """" & txtGeneralRdLocation.Text & "\" & "SteamyDock.exe""")
        End If
    Else
        If defaultDock = 0 Then ' rocketdock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RocketDock", "")
        End If
        If defaultDock = 1 Then ' steamydock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteamyDock", "")
        End If
    End If
    
    busyTimer.Enabled = True
    
    ' if the rocketdock process has died then create a new one.
    If ans = True Then
        ' restart Rocketdock
        If FExists(dockAppPath & "\" & NameProcess) Then
            Call ShellExecute(hwnd, "Open", dockAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
        End If
    Else
        answer = MsgBox("Could not find a " & NameProcess & " process, would you like me to restart " & NameProcess & "?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If

        ' restart Rocketdock
        If FExists(dockAppPath & "\" & NameProcess) Then
            Call ShellExecute(hwnd, "Open", dockAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

btnApply_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") " & debugPoint & " in procedure btnApply_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error
   If debugflg = 1 Then DebugPrint "%btnClose_Click"

    Form_Unload 0

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnDefaults_Click
' Author    : beededea
' Date      : 02/03/2020
' Purpose   : The registry is written first and then the settings file is recreated afterwards
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaults_Click()

        
    ' variables declared
    Dim NameProcess As String
    Dim ans As Boolean
    Dim answer As VbMsgBoxResult

   On Error GoTo btnDefaults_Click_Error
   If debugflg = 1 Then DebugPrint "%btnDefaults_Click"

   'initialise the dimensioned variables
    NameProcess = ""
    ans = False
    answer = vbNo
    
    If InIDE = True Then
        answer = MsgBox("Running in the IDE. The registry values may corrupt - be warned. Continue?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    Else
        If IsUserAnAdmin() = 0 Then
            answer = MsgBox("This program is not running as admin. Some of the settings may be strange and unwanted - be warned. Continue?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If

    answer = MsgBox("Are you sure you want to rest Rocketdock to its default settings? Note: this will not lose your icons.?", vbYesNo)
    If answer = vbNo Then
        Exit Sub
    End If
 
    ' kill the rocketdock process
    
    If defaultDock = 0 Then
        NameProcess = "RocketDock.exe"
    Else
        NameProcess = "steamyDock.exe"
    End If
    ans = checkAndKill(NameProcess, False)
    
    If defaultDock = 0 Then
        rDVersion = "1.3.5"
    Else
        rDVersion = App.Major & "." & App.Minor & "." & App.Revision
    End If
    
    rDCustomIconFolder = ""
    
    If defaultDock = 0 Then ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
        rDHotKeyToggle = "Control+Alt+R"
    Else
        rDHotKeyToggle = "F11"
    End If
    
    ' removed
    'cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
    If defaultDock = 1 Then
        cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    Else
        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    End If
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
   
    
    rDtheme = "CrystalXP.net"
    cmbStyleTheme.Text = rDtheme
    
    rDThemeOpacity = "100"
    sliStyleOpacity.Value = Val(rDThemeOpacity)
    
    rDIconOpacity = "100"
    sliIconsOpacity.Value = Val(rDIconOpacity)
    
    rDFontSize = "-8"
    rDFontFlags = "0"
    rDFontName = "Centurion Light SF"
    rDFontColor = "65535"
    rDFontCharSet = "1"
    
    'lblPreviewFont.ForeColor = Convert_Dec2RGB(rDFontColor)  ' converts the stored decimal value to RGB

    lblPreviewFont.FontName = rDFontName
    lblPreviewFont.FontSize = Abs(rDFontSize)
    lblPreviewFont.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewFont.ForeColor = rDFontColor
    
    lblPreviewTop.FontName = rDFontName
    lblPreviewTop.FontSize = Abs(rDFontSize)
    lblPreviewTop.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewBottom.FontName = rDFontName
    lblPreviewBottom.FontSize = Abs(rDFontSize)
    lblPreviewBottom.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewLeft.FontName = rDFontName
    lblPreviewLeft.FontSize = Abs(rDFontSize)
    lblPreviewLeft.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewRight.FontName = rDFontName
    lblPreviewRight.FontSize = Abs(rDFontSize)
    lblPreviewRight.FontBold = rDFontFlags
    'lblPreviewFont.FontItalic = suppliedStyle
    lblPreviewTop.ForeColor = rDFontOutlineColor
    
    lblPreviewFontShadow.FontName = rDFontName
    lblPreviewFontShadow.FontSize = Abs(rDFontSize)
    lblPreviewFontShadow.FontBold = rDFontFlags
    'lblPreviewFontShadow.FontItalic = suppliedStyle
    lblPreviewFontShadow.ForeColor = rDFontShadowColor
    
    lblPreviewFontShadow2.FontName = rDFontName
    lblPreviewFontShadow2.FontSize = Abs(rDFontSize)
    lblPreviewFontShadow2.FontBold = rDFontFlags
    'lblPreviewFontShadow2.FontItalic = suppliedStyle
    lblPreviewFontShadow2.ForeColor = rDFontShadowColor
    
    
    
    
    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Abs(rDFontSize) & "pt"
        
    
    rDFontOutlineColor = "255"
    lblStyleOutlineColourDesc.Caption = "Outline Colour: " & Convert_Dec2RGB(rDFontOutlineColor)
    lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor
    
    rDFontOutlineOpacity = "0"
    sliStyleOutlineOpacity.Value = Val(rDFontOutlineOpacity)
    
    rDFontShadowColor = "12632256"
    lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & Convert_Dec2RGB(rDFontShadowColor)
    lblStyleFontOutlineTest.ForeColor = rDFontShadowColor
    
    rDFontShadowOpacity = "30"
    sliStyleShadowOpacity.Value = Val(rDFontShadowOpacity)
    
    sDFontOpacity = "100"
    sliStyleFontOpacity.Value = Val(sDFontOpacity)
    
    rDIconMin = "16"
    sliIconsSize.Value = Val(rDIconMin)
    
    
    sliIconsZoom.Value = Val(rdIconMax) - 17
    
    Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animatio

    
    rDZoomWidth = "4"
    sliIconsZoomWidth.Value = Val(rDZoomWidth)
    
    rDZoomTicks = "199"
    sliIconsDuration.Value = Val(rDZoomTicks)
    
    rDAutoHideTicks = "186"
    sliBehaviourAutoHideDuration.Value = Val(rDAutoHideTicks)
    
    rDAnimationInterval = "10"
    sliAnimationInterval.Value = Val(rDAnimationInterval)
    
    rDSkinSize = "118"
    sliStyleThemeSize.Value = Val(rDSkinSize)
    
    sDSplashStatus = "1"
    chkSplashStatus.Value = Val(sDSplashStatus)
    
    sDShowIconSettings = "1"
    genChkShowIconSettings.Value = Val(sDShowIconSettings) ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock
    
    rDAutoHideDelay = "174"
    sliBehaviourAutoHideDelay.Value = Val(rDAutoHideDelay)
    
    rDPopupDelay = "68"
    sliBehaviourPopUpDelay.Value = Val(rDPopupDelay)
    
    sDContinuousHide = "10" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    sliContinuousHide.Value = sDContinuousHide
    
    sDBounceZone = "75"
    ' sDBounceZone
    
    rDIconQuality = "2"
    cmbIconsQuality.ListIndex = Val(rDIconQuality)
    
    rDLangID = "1033"
    
    rDHideLabels = "0"
    chkStyleDisable.Value = Val(rDHideLabels)
    
   ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
    sDShowLblBacks = "0"
    chkLabelBackgrounds.Value = Val(sDShowLblBacks)

    
    rDZoomOpaque = "1"
    chkIconsZoomOpaque.Value = Val(rDZoomOpaque)
    
    rDLockIcons = "1"
    chkGenLock.Value = Val(rDLockIcons)
    
    rDAutoHide = "1"
    chkBehaviourAutoHide.Value = Val(rDAutoHide)
' 26/10/2020 docksettings .05 DAEB  added a manual click to the autohide toggle checkbox
' a checkbox value assignment does not trigger a checkbox click for this checkbox (in a frame) as normally occurs and there is no equivalent 'change event' for a checkbox
' so to force it to trigger we need a call to the click event
    Call chkBehaviourAutoHide_Click
    
    rDManageWindows = "0"
    chkGenMin.Value = Val(rDManageWindows)
    
    rDDisableMinAnimation = "1"
    chkGenDisableAnim.Value = Val(rDDisableMinAnimation)
    
    rDShowRunning = "1"
    chkGenRun.Value = Val(rDShowRunning)
    
    rDOpenRunning = "1"
    chkGenOpen.Value = Val(rDOpenRunning)
    
    rDHoverFX = "1"
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)
    
    rDzOrderMode = "0"
    cmbPositionLayering.ListIndex = Val(rDzOrderMode)
    
    rDMouseActivate = "1"
    chkBehaviourMouseActivate.Value = Val(rDMouseActivate)
    
    rDIconActivationFX = "2"
    cmbBehaviourActivationFX.ListIndex = Val(rDIconActivationFX)
    
    sDAutoHideType = "0"
    cmbBehaviourAutoHideType.ListIndex = Val(sDAutoHideType)
    
    rDMonitor = "0" ' ie. monitor 1
    cmbPositionMonitor.ListIndex = Val(rDMonitor)
    
    rDSide = "1"
    cmbPositionScreen.ListIndex = Val(rDSide)
    
    rDOffset = "0"
    sliPositionCentre.Value = Val(rDOffset)
    
    rDvOffset = "0"
    sliPositionEdgeOffset.Value = Val(rDvOffset)
    
    rDOptionsTabIndex = "4"
    rdIconMax = "128" ' 128 rocketdock
    
    ' The following has been commented out as the reversion to defaults should happen using the temporary settings file, not the registry

'    If rocketDockInstalled = True Then ' if RD hasn't been installed then the registry nor the settings.ini file will exist
'
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMax", "128")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LoadError", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Version", "1.3.5")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "CustomIconFolder", "?E:\\dean\\steampunk theme\\icons")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HotKey-Toggle", "Control+Alt+R")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Theme", "CrystalXP.net")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ThemeOpacity", "100")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconOpacity", "100")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontSize", "-8")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontFlags", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontName", "Times New Roman")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontColor", "65535")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontCharSet", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineColor", "255")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineOpacity", "9")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowColor", "12632256")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowOpacity", "30")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMin", "16")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomWidth", "4")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomTicks", "199")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideTicks", "186")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideDelay", "174")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "PopupDelay", "68")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconQuality", "2")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LangID", "1033")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HideLabels", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomOpaque", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHide", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ManageWindows", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "DisableMinAnimation", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ShowRunning", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OpenRunning", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HoverFX", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "zOrderMode", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "MouseActivate", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconActivationFX", "2")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Monitor", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", "1")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Offset", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "vOffset", "0")
'        Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OptionsTabIndex", "4")
'    End If
    
    'regardless of the method used, registry, settings or the new 3rd option, they all use the temporary settings file.
    If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
        Call writeDockSettings("Software\RocketDock", tmpSettingsFile)
        ' if it exists, read the registry values for each of the icons and write them to the settings.ini
        Call readIconsWriteSettings("Software\RocketDock\Icons", tmpSettingsFile)
        
    End If
    
    ' writes directly to the new config file without any intervening temporary file
    If optGeneralReadConfig.Value = True Then
        Call writeDockSettings("Software\SteamyDock\DockSettings", dockSettingsFile)
    End If

    'NOTE: the settings are NOT written to the registry until the apply button is pressed.
    
    ' if the rocketdock process has died then
'    If ans = True Then
'        ' restart Rocketdock
'        Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
'    Else
'
'        answer = MsgBox("Could not find a " & NameProcess & " process, would you like me to restart " & NameProcess & "?", vbYesNo)
'        If answer = vbNo Then
'            Exit Sub
'        End If
'
'        ' restart Rocketdock
'        Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.Path, 1)
'    End If

   On Error GoTo 0
   Exit Sub

btnDefaults_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaults_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGeneralRdFolder_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   : unused disabled
'---------------------------------------------------------------------------------------
'
Private Sub btnGeneralRdFolder_Click()

        
    ' variables declared
    Dim getFolder As String
    Dim dialogInitDir As String
   
    On Error GoTo btnGeneralRdFolder_Click_Error
    If debugflg = 1 Then DebugPrint "%btnGeneralRdFolder_Click"
    
   'initialise the dimensioned variables
    getFolder = ""
    dialogInitDir = ""
    
    If txtGeneralRdLocation.Text <> vbNullString Then
        If DirExists(txtGeneralRdLocation.Text) Then
            dialogInitDir = txtGeneralRdLocation.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder
    'getFolder = ChooseDir_Click ' old method to show the dialog box to select a folder
    If getFolder <> vbNullString Then txtGeneralRdLocation.Text = getFolder

   On Error GoTo 0
   Exit Sub

btnGeneralRdFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGeneralRdFolder_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStyleFont_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   : select the dock's font
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleFont_Click()
        
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedFontSize As Integer
    
    Dim suppliedStyle As Boolean
    Dim suppliedColour As Variant
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    Dim fontSelected As Boolean
    
    On Error GoTo btnStyleFont_Click_Error
    If debugflg = 1 Then DebugPrint "%btnStyleFont_Click"
   
   'initialise the dimensioned variables
    suppliedFont = ""
    suppliedSize = 0
    suppliedWeight = 0
    suppliedBold = False
    suppliedFontSize = 0
    
    suppliedStyle = False
    'suppliedColour =
    suppliedItalics = False
    suppliedUnderline = False
    fontSelected = False
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    
    displayFontSelector rDFontName, suppliedFontSize, suppliedWeight, suppliedStyle, rDFontColor, suppliedItalics, suppliedUnderline, fontSelected
    If fontSelected = False Then Exit Sub
    
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    
    
   On Error GoTo 0
   Exit Sub

btnStyleFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleFont_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : preFontInformation
' Author    : beededea
' Date      : 17/05/2020
' Purpose   : the fontsize used by Rocketdock does not equate to the pt size in the font selector
'             so we have to determine the fontsize to be displayed by the font selector
'---------------------------------------------------------------------------------------
'
Private Sub preFontInformation(suppliedFontSize As Integer, suppliedBold As Boolean, suppliedItalics As Boolean, suppliedUnderline As Boolean, suppliedWeight As Integer)

   On Error GoTo preFontInformation_Error

    If rDFontSize = "-8" Then suppliedFontSize = 6
    If rDFontSize = "-11" Then suppliedFontSize = 8
    If rDFontSize = "-12" Then suppliedFontSize = 9
    If rDFontSize = "-13" Then suppliedFontSize = 10
    If rDFontSize = "-15" Then suppliedFontSize = 11
    If rDFontSize = "-16" Then suppliedFontSize = 12
    If rDFontSize = "-19" Then suppliedFontSize = 14
    If rDFontSize = "-21" Then suppliedFontSize = 16
    If rDFontSize = "-24" Then suppliedFontSize = 18
    If rDFontSize = "-27" Then suppliedFontSize = 20
    If rDFontSize = "-29" Then suppliedFontSize = 22
    
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False
    'suppliedWeight = False
    
    If rDFontFlags = 1 Or rDFontFlags = 3 Or rDFontFlags = 7 Or rDFontFlags = 11 Or rDFontFlags = 15 Then suppliedBold = True
    If rDFontFlags = 2 Or rDFontFlags = 3 Or rDFontFlags = 6 Or rDFontFlags = 7 Or rDFontFlags = 10 Or rDFontFlags = 11 Or rDFontFlags = 13 Or rDFontFlags = 14 Or rDFontFlags = 15 Then suppliedItalics = True
    If rDFontFlags = 6 Or rDFontFlags = 14 Then suppliedUnderline = True


   On Error GoTo 0
   Exit Sub

preFontInformation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure preFontInformation of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontInformation
' Author    : beededea
' Date      : 17/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayFontInformation(suppliedFontSize As Integer, suppliedBold As Boolean, suppliedItalics As Boolean, suppliedUnderline As Boolean, suppliedWeight As Integer)
    
    ' the fontsize used by Rocketdock does not equate to the pt size in the font selector
    ' so we have to calculate the rDFontSize that will be written to rocketdock registry/settings.
    
   On Error GoTo displayFontInformation_Error

    If suppliedFontSize = 6 Then rDFontSize = "-8"
    If suppliedFontSize = 8 Then rDFontSize = "-11"
    If suppliedFontSize = 9 Then rDFontSize = "-12"
    If suppliedFontSize = 10 Then rDFontSize = "-13"
    If suppliedFontSize = 11 Then rDFontSize = "-15"
    If suppliedFontSize = 12 Then rDFontSize = "-16"
    If suppliedFontSize = 14 Then rDFontSize = "-19"
    If suppliedFontSize = 16 Then rDFontSize = "-21"
    If suppliedFontSize = 18 Then rDFontSize = "-24"
    If suppliedFontSize = 20 Then rDFontSize = "-27"
    If suppliedFontSize = 22 Then rDFontSize = "-29"
    
    lblPreviewFont.FontName = rDFontName
    lblPreviewFont.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewTop.FontName = rDFontName
    lblPreviewTop.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewBottom.FontName = rDFontName
    lblPreviewBottom.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewLeft.FontName = rDFontName
    lblPreviewLeft.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewRight.FontName = rDFontName
    lblPreviewRight.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewFontShadow.FontName = rDFontName
    lblPreviewFontShadow.FontSize = Abs(suppliedFontSize) + 4
    
    lblPreviewFontShadow2.FontName = rDFontName
    lblPreviewFontShadow2.FontSize = Abs(suppliedFontSize) + 4
    
    If suppliedWeight > 400 Then
        suppliedBold = True
    Else
        suppliedBold = False
    End If
    
    lblPreviewFont.FontBold = suppliedBold
    lblPreviewFont.FontItalic = suppliedItalics
    lblPreviewFont.ForeColor = rDFontColor
    lblPreviewFont.FontUnderline = suppliedUnderline
    
    lblPreviewTop.FontBold = suppliedBold
    lblPreviewTop.FontItalic = suppliedItalics
    lblPreviewTop.ForeColor = rDFontOutlineColor
    lblPreviewTop.FontUnderline = suppliedUnderline
    
    lblPreviewBottom.FontBold = suppliedBold
    lblPreviewBottom.FontItalic = suppliedItalics
    lblPreviewBottom.ForeColor = rDFontOutlineColor
    lblPreviewBottom.FontUnderline = suppliedUnderline
    
    lblPreviewLeft.FontBold = suppliedBold
    lblPreviewLeft.FontItalic = suppliedItalics
    lblPreviewLeft.ForeColor = rDFontOutlineColor
    lblPreviewLeft.FontUnderline = suppliedUnderline
    
    lblPreviewRight.FontBold = suppliedBold
    lblPreviewRight.FontItalic = suppliedItalics
    lblPreviewRight.ForeColor = rDFontOutlineColor
    lblPreviewRight.FontUnderline = suppliedUnderline
    
    lblPreviewFontShadow.FontBold = suppliedBold
    lblPreviewFontShadow.FontItalic = suppliedItalics
    lblPreviewFontShadow.ForeColor = rDFontShadowColor
    lblPreviewFontShadow.FontUnderline = suppliedUnderline
    
    lblPreviewFontShadow2.FontBold = suppliedBold
    lblPreviewFontShadow2.FontItalic = suppliedItalics
    lblPreviewFontShadow2.ForeColor = rDFontShadowColor
    lblPreviewFontShadow2.FontUnderline = suppliedUnderline
    
    'lblPreviewFontShadow.Visible = False
    
    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Abs(suppliedFontSize) & "pt"
    If suppliedBold = True Then lblStyleFontName.Caption = lblStyleFontName.Caption & " Bold"
    If suppliedItalics = True Then lblStyleFontName.Caption = lblStyleFontName.Caption & " Italic"

    ' now change the rocketdock vars
    
    ' 0 - no qualifiers or alterations
    ' 1 - bold
    ' 2 - light italics
    ' 3 - bold italics
    ' 4 - strikeout & light ' unsupported
    ' 6 - underline and italics
    ' 7 - bold, italics & underline
    ' 10 - strikeout & italics ' unsupported
    ' 11 - bold, italics & strikeout  ' unsupported
    ' 13 - strikeout & italics        ' unsupported
    ' 14 - underline, strikeout and italics ' unsupported
    ' 15 - bold, underline, strikeout and italics ' unsupported
        
    lblPreviewFont.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewFont.Top = (735 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewTop.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewTop.Top = (715 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewBottom.Left = (5340 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewBottom.Top = (755 / 2) - (lblPreviewFont.Height / 2)
    
    lblPreviewLeft.Left = (5320 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewLeft.Top = (735 / 2) - (lblPreviewFont.Height / 2)
        
    lblPreviewRight.Left = (5370 / 2) - (lblPreviewFont.Width / 2)
    lblPreviewRight.Top = (735 / 2) - (lblPreviewFont.Height / 2)
        
    lblPreviewFontShadow.Left = (5440 / 2) - (lblPreviewFontShadow.Width / 2)
    lblPreviewFontShadow.Top = (825 / 2) - (lblPreviewFontShadow.Height / 2)
        
    lblPreviewFontShadow2.Left = (5450 / 2) - (lblPreviewFontShadow2.Width / 2)
    lblPreviewFontShadow2.Top = (835 / 2) - (lblPreviewFontShadow2.Height / 2)
        
    rDFontFlags = 0
    If suppliedBold = True Then rDFontFlags = 1
    If suppliedItalics = True Then rDFontFlags = 2
    If suppliedItalics = True And suppliedBold = True Then rDFontFlags = 3
    If suppliedUnderline = True And suppliedItalics = True Then rDFontFlags = 6
    If suppliedUnderline = True And suppliedItalics = True And suppliedBold = True Then rDFontFlags = 7

   On Error GoTo 0
   Exit Sub

displayFontInformation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontInformation of Form dockSettings"
  End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnStyleShadow_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   : determine the shadow colour
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleShadow_Click()
        
    ' variables declared
    Dim colourResult As Long
    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
   
    'initialise the dimensioned variables
     colourResult = 0
     suppliedFontSize = 0
     suppliedWeight = 0
     suppliedBold = False
     suppliedItalics = False
     suppliedUnderline = False
   
   On Error GoTo btnStyleShadow_Click_Error
   If debugflg = 1 Then DebugPrint "%btnStyleShadow_Click"

    colourResult = ShowColorDialog(Me.hwnd, True, rDFontShadowColor)

    If colourResult <> -1 And colourResult <> 0 Then
        rDFontShadowColor = colourResult
        
        lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & Convert_Dec2RGB(rDFontShadowColor)
        lblStyleFontFontShadowTest.ForeColor = rDFontShadowColor
        
        'rDFontShadowOpacity = str(btnStyleShadow.value)

    End If
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
       

    On Error GoTo 0
    Exit Sub

btnStyleShadow_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleShadow_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnStyleOutline_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnStyleOutline_Click()

       
    ' variables declared
    Dim colourResult As Long
    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    
    'initialise the dimensioned variables
     colourResult = 0
     suppliedFontSize = 0
     suppliedWeight = 0
     suppliedBold = False
     suppliedItalics = False
     suppliedUnderline = False
    
    On Error GoTo btnStyleOutline_Click_Error
   If debugflg = 1 Then DebugPrint "%btnStyleOutline_Click"
    
    ' this will take 255, VBRed,  16711680
    colourResult = ShowColorDialog(Me.hwnd, True, rDFontOutlineColor)
    
    If colourResult <> -1 And colourResult <> 0 Then
        rDFontOutlineColor = (colourResult)
        lblStyleOutlineColourDesc.Caption = "Outline Colour: " & Convert_Dec2RGB(rDFontOutlineColor)
        lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor
    End If
   
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)
   
   On Error GoTo 0
   Exit Sub

btnStyleOutline_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnStyleOutline_Click of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkBehaviourAutoHide_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkBehaviourAutoHide_Click()
   On Error GoTo chkBehaviourAutoHide_Click_Error
   If debugflg = 1 Then DebugPrint "%chkBehaviourAutoHide_Click"

    If chkBehaviourAutoHide.Value = 1 Then
        chkBehaviourAutoHide.Caption = "Autohide Enabled"
        sliBehaviourAutoHideDuration.Enabled = True
  
        lblAutoHideDuration.Enabled = True

        lblAutoHideDurationMsLow.Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblAutoHideDelay.Enabled = True
        lblAutoHideDelayMsLow.Enabled = True
        sliBehaviourAutoHideDelay.Enabled = True
        lblAutoHideDelayMsHigh.Enabled = True
        lblAutoHideDelayMsCurrent.Enabled = True
        
        lblAutoRevealDuration.Enabled = True
        lblAutoRevealDurationMsLow.Enabled = True
        lblAutoRevealDurationMsHigh.Enabled = True
        sliBehaviourPopUpDelay.Enabled = True
        
        lblBehaviourPopUpDelayMsCurrrent.Enabled = True
        
        cmbBehaviourAutoHideType.Enabled = True
        
    Else
        chkBehaviourAutoHide.Caption = "Autohide Disabled"
        sliBehaviourAutoHideDuration.Enabled = False

        lblAutoHideDuration.Enabled = False

        lblAutoHideDurationMsLow.Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblAutoHideDelay.Enabled = False
        lblAutoHideDelayMsLow.Enabled = False
        sliBehaviourAutoHideDelay.Enabled = False
        lblAutoHideDelayMsHigh.Enabled = False
        lblAutoHideDelayMsCurrent.Enabled = False
        
                
        lblAutoRevealDuration.Enabled = False
        lblAutoRevealDurationMsLow.Enabled = False
        sliBehaviourPopUpDelay.Enabled = False
        lblAutoRevealDurationMsHigh.Enabled = False
        lblBehaviourPopUpDelayMsCurrrent.Enabled = False
        
        cmbBehaviourAutoHideType.Enabled = False
    
    End If
    
    rDAutoHide = chkBehaviourAutoHide.Value
    

   On Error GoTo 0
   Exit Sub

chkBehaviourAutoHide_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkBehaviourAutoHide_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkBehaviourMouseActivate_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkBehaviourMouseActivate_Click()
   On Error GoTo chkBehaviourMouseActivate_Click_Error
   If debugflg = 1 Then DebugPrint "%chkBehaviourMouseActivate_Click"

    rDMouseActivate = chkBehaviourMouseActivate.Value

   On Error GoTo 0
   Exit Sub

chkBehaviourMouseActivate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkBehaviourMouseActivate_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenDisableAnim_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenDisableAnim_Click()

   On Error GoTo chkGenDisableAnim_Click_Error
   If debugflg = 1 Then DebugPrint "%chkGenDisableAnim_Click"

   rDDisableMinAnimation = chkGenDisableAnim.Value
   
   rDMouseActivate = chkGenDisableAnim.Value

   On Error GoTo 0
   Exit Sub

chkGenDisableAnim_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenDisableAnim_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenLock_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenLock_Click()

   On Error GoTo chkGenLock_Click_Error
   If debugflg = 1 Then DebugPrint "%chkGenLock_Click"

    rDLockIcons = chkGenLock.Value

   On Error GoTo 0
   Exit Sub

chkGenLock_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenLock_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenMin_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenMin_Click()
   On Error GoTo chkGenMin_Click_Error
   If debugflg = 1 Then DebugPrint "%chkGenMin_Click"

    If chkGenMin.Value = 0 Then
        chkGenDisableAnim.Enabled = False
        lblChkDisable.Enabled = False
    Else
        chkGenDisableAnim.Enabled = True
        lblChkDisable.Enabled = True
    End If
    
    rDManageWindows = chkGenMin.Value

   On Error GoTo 0
   Exit Sub

chkGenMin_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenMin_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenOpen_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenOpen_Click()
   On Error GoTo chkGenOpen_Click_Error
   If debugflg = 1 Then DebugPrint "%chkGenOpen_Click"

    rDOpenRunning = chkGenOpen.Value

   On Error GoTo 0
   Exit Sub

chkGenOpen_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenOpen_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenRun_Click
' Author    : beededea
' Date      : 11/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenRun_Click()
   On Error GoTo chkGenRun_Click_Error
   If debugflg = 1 Then Debug.Print "%chkGenRun_Click"

    rDShowRunning = chkGenRun.Value
    
    If chkGenRun.Value = 0 Then
        lblGenRunAppInterval1.Enabled = False
        lblGenRunAppInterval2.Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenRunAppInterval3.Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else

        If optGeneralReadConfig.Value = True Then ' steamydock
            lblGenRunAppInterval1.Enabled = True
            lblGenRunAppInterval2.Enabled = True
            sliGenRunAppInterval.Enabled = True
            lblGenRunAppInterval3.Enabled = True
            lblGenRunAppIntervalCur.Enabled = True
        End If
    End If

   On Error GoTo 0
   Exit Sub

chkGenRun_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenRun_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGenWinStartup_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGenWinStartup_Click()

   On Error GoTo chkGenWinStartup_Click_Error
   If debugflg = 1 Then DebugPrint "%chkGenWinStartup_Click"

    rDStartupRun = chkGenWinStartup.Value

   On Error GoTo 0
   Exit Sub

chkGenWinStartup_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenWinStartup_Click of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkIconsZoomOpaque_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkIconsZoomOpaque_Click()

   On Error GoTo chkIconsZoomOpaque_Click_Error
   If debugflg = 1 Then DebugPrint "%chkIconsZoomOpaque_Click"

    rDZoomOpaque = chkIconsZoomOpaque.Value

   On Error GoTo 0
   Exit Sub

chkIconsZoomOpaque_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIconsZoomOpaque_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkStyleDisable_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkStyleDisable_Click()

   On Error GoTo chkStyleDisable_Click_Error
   If debugflg = 1 Then DebugPrint "%chkStyleDisable_Click"

   rDHideLabels = chkStyleDisable.Value
   
    If chkStyleDisable.Value = 1 Then
        chkLabelBackgrounds.Enabled = False ' .01 docksettings DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
        lblChkLabelBackgrounds.Enabled = False ' .01
        
        btnStyleFont.Enabled = False
        lblStyleFontName.Enabled = False
        btnStyleShadow.Enabled = False
        lblStyleFontFontShadowColor.Enabled = False
        lblStyleFontFontShadowTest.Enabled = False
        btnStyleOutline.Enabled = False
        lblStyleOutlineColourDesc.Enabled = False
        lblStyleFontOutlineTest.Enabled = False
        Label37.Enabled = False
        Label40.Enabled = False
        sliStyleShadowOpacity.Enabled = False
        Label39.Enabled = False
        lblStyleShadowOpacityCurrent.Enabled = False
        Label23.Enabled = False
        Label36.Enabled = False
        sliStyleOutlineOpacity.Enabled = False
        Label35.Enabled = False
        lblStyleOutlineOpacityCurrent.Enabled = False
        
        sliStyleFontOpacity.Enabled = False
        lblStyleFontOpacityCurrent.Enabled = False
        Label22.Enabled = False
        Label34.Enabled = False
        Label30.Enabled = False
        
    Else
        chkLabelBackgrounds.Enabled = True  ' .01 docksettings DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
        lblChkLabelBackgrounds.Enabled = True ' .01
        
        btnStyleFont.Enabled = True
        lblStyleFontName.Enabled = True
        btnStyleShadow.Enabled = True
        lblStyleFontFontShadowColor.Enabled = True
        lblStyleFontFontShadowTest.Enabled = True
        btnStyleOutline.Enabled = True
        lblStyleOutlineColourDesc.Enabled = True
        lblStyleFontOutlineTest.Enabled = True
        Label37.Enabled = True
        Label40.Enabled = True
        sliStyleShadowOpacity.Enabled = True
        Label39.Enabled = True
        lblStyleShadowOpacityCurrent.Enabled = True
        Label23.Enabled = True
        Label36.Enabled = True
        sliStyleOutlineOpacity.Enabled = True
        Label35.Enabled = True
        lblStyleOutlineOpacityCurrent.Enabled = True
                
        sliStyleFontOpacity.Enabled = True
        lblStyleFontOpacityCurrent.Enabled = True
        Label22.Enabled = True
        Label34.Enabled = True
        Label30.Enabled = True
    End If
   

   On Error GoTo 0
   Exit Sub

chkStyleDisable_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkStyleDisable_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbBehaviourActivationFX_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbBehaviourActivationFX_Click()

   On Error GoTo cmbBehaviourActivationFX_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbBehaviourActivationFX_Click"

    rDIconActivationFX = cmbBehaviourActivationFX.ListIndex

   On Error GoTo 0
   Exit Sub

cmbBehaviourActivationFX_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbBehaviourActivationFX_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbIconsHoverFX_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbIconsHoverFX_Click()

   On Error GoTo cmbIconsHoverFX_Change_Error
   If debugflg = 1 Then DebugPrint "%cmbIconsHoverFX_Change"

    rDHoverFX = cmbIconsHoverFX.ListIndex
    
    'bubble
    'plateau
    'flat
    'bumpy
    
    'Call setMinimumHoverFX    ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation


   On Error GoTo 0
   Exit Sub

cmbIconsHoverFX_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbIconsHoverFX_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setMinimumHoverFX
' Author    : beededea
' Date      : 29/04/2021
' Purpose   : Set the large icon minimum size to 85 pixels when using the bumpy animation
'---------------------------------------------------------------------------------------
'
Private Sub setMinimumHoverFX()
    On Error GoTo setMinimumHoverFX_Error

    If Val(rDHoverFX) = 4 And sliIconsZoom.Value <= 85 Then
        sliIconsZoom.Value = 85
        sliIconsZoom.ToolTipText = "The maximum size after a zoom can be no smaller than 85 pixels when Zoom:Bumpy is chosen"
    Else
        sliIconsZoom.ToolTipText = "The maximum size after a zoom"
    End If

    On Error GoTo 0
    Exit Sub

setMinimumHoverFX_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMinimumHoverFX of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbIconsQuality_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbIconsQuality_Click()

   On Error GoTo cmbIconsQuality_Change_Error
   If debugflg = 1 Then DebugPrint "%cmbIconsQuality_Change"
    
    rDIconQuality = cmbIconsQuality.ListIndex

   On Error GoTo 0
   Exit Sub

cmbIconsQuality_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbIconsQuality_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionLayering_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionLayering_Click()


   On Error GoTo cmbPositionLayering_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbPositionLayering_Click"

   rDzOrderMode = cmbPositionLayering.ListIndex


   On Error GoTo 0
   Exit Sub

cmbPositionLayering_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionLayering_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionMonitor_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : This is called twice during startup by GetMonitorCount adjustControls, only one really required
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionMonitor_Click()

   On Error GoTo cmbPositionMonitor_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbPositionMonitor_Click"
    
    If startupFlg = True Then '.NET
        ' don't do this on the startup run only when actually clicked upon
        Exit Sub
    Else
        rDMonitor = cmbPositionMonitor.ListIndex
    End If

   On Error GoTo 0
   Exit Sub

cmbPositionMonitor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionMonitor_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbPositionScreen_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbPositionScreen_Click()

   On Error GoTo cmbPositionScreen_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbPositionScreen_Click"

   rDSide = cmbPositionScreen.ListIndex
   
   If defaultDock = 1 Then 'disallow left or right under steamydock
        If rDSide = "2" Or rDSide = "3" Then
            rDSide = "1"
            cmbPositionScreen.ListIndex = 1
        End If
   End If

   On Error GoTo 0
   Exit Sub

cmbPositionScreen_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPositionScreen_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbStyleTheme_Change
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : if a theme is selected from the dropdown list then make it the default
'---------------------------------------------------------------------------------------
'
Private Sub cmbStyleTheme_Click()
    Dim themePic As String

    On Error GoTo cmbStyleTheme_Change_Error
    If debugflg = 1 Then DebugPrint "%cmbStyleTheme_Change"
    
    rDtheme = cmbStyleTheme.List(cmbStyleTheme.ListIndex)
    
    ' .09 DAEB 01/02/2021 docksettings Make the sample image functionality disabled for rocketdock
    If defaultDock = 1 Then
        themePic = sdAppPath & "\skins\" & rDtheme & "\sample.jpg"
        
        If FExists(themePic) Then
            picThemeSample.Picture = LoadPicture(sdAppPath & "\skins\" & rDtheme & "\sample.jpg")
        End If
    End If
    
    On Error GoTo 0
    Exit Sub

cmbStyleTheme_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbStyleTheme_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDefaultDock_Click
' Author    : beededea
' Date      : 13/05/2020
' Purpose   : certain options are disabled when selecting Steamydock (and vice versa)
'---------------------------------------------------------------------------------------
'
Private Sub cmbDefaultDock_Click()
   On Error GoTo cmbDefaultDock_Click_Error
   If debugflg = 1 Then Debug.Print "%cmbDefaultDock_Click"

   On Error GoTo cmbDefaultDock_Change_Error
   If debugflg = 1 Then DebugPrint "%cmbDefaultDock_Change"

    If cmbDefaultDock.List(cmbDefaultDock.ListIndex) = "RocketDock" Then
        ' check where rocketdock is installed
        Call checkRocketdockInstallation
        defaultDock = 0 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
        
        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
            optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
            optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
        Else
            optGeneralReadRegistry.Value = True
            optGeneralWriteRegistry.Value = True
        End If
        
        rDDefaultDock = "rocketdock"
        
        ' re-enable all the controls that Rocketdock supports
        chkGenMin.Enabled = True
        lblChkMinimise.Enabled = True
        'cmbBehaviourActivationFX.Enabled = True

        'cmbStyleTheme.Enabled = True
        'cmbPositionMonitor.Enabled = True
        'cmbIconsQuality.Enabled = True
        chkIconsZoomOpaque.Enabled = True
        sliIconsDuration.Enabled = True
        'chkGenOpen.Enabled = True
        
        'lblChkOpenRunning.Enabled = True
        
        If chkGenMin.Value = 0 Then
            chkGenDisableAnim.Enabled = False
            lblChkDisable.Enabled = False
        Else
            chkGenDisableAnim.Enabled = True
            lblChkDisable.Enabled = True
        End If
    
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        sliIconsZoomWidth.Enabled = True
'        sliIconsDuration.Enabled = True

        cmbIconsHoverFX.Enabled = True
        
        Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
        
        Call setBounceTypes
        
        sliBehaviourAutoHideDuration.Enabled = True
        sliAnimationInterval.Enabled = True
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
        'fraZoomConfigs.Visible = True
        
        fraAutoHideDuration.Visible = True
        fraFontOpacity.Visible = True

        optGeneralReadConfig.Enabled = False ' RD does not support storing the configs at the correct location
        lblGeneralReadConfig.Enabled = False
        optGeneralWriteConfig.Enabled = False ' RD does not support storing the configs at the correct location
        lblGeneralWriteConfig.Enabled = False
        
        optGeneralReadSettings.Enabled = True
        optGeneralReadRegistry.Enabled = True
        
        lblGenRunAppInterval1.Enabled = False
        lblGenRunAppInterval2.Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenRunAppInterval3.Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
        
        chkGenAlwaysAsk.Enabled = False
        lblChkAlwaysConfirm.Enabled = False
        
        sliAnimationInterval.Enabled = False
        lblAnimationIntervalLabel.Enabled = False
        lblAnimationIntervalMsLow.Enabled = False
        lblAnimationIntervalMsHigh.Enabled = False
        lblAnimationIntervalMsCurrent.Enabled = False
        lblAnimationInformationLabel.Enabled = False
        
        cmbBehaviourAutoHideType.Enabled = False
        
        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        picThemeSample.Enabled = False
        lblThemeSizeText.Enabled = False
        lblThemeSizeTextLow.Enabled = False
        sliStyleThemeSize.Enabled = False
        lblThemeSizeTextHigh.Enabled = False
        lblStyleSizeCurrent.Enabled = False
        
        lblContinuousHide.Enabled = False
        lblContinuousHideMsLow.Enabled = False
        sliContinuousHide.Enabled = False
        lblContinuousHideMsHigh.Enabled = False
        lblContinuousHideMsCurrent.Enabled = False
        
        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
    Else
        ' check where/if steamydock is installed
        Call checkSteamyDockInstallation 'defaultDock is set here
        
        rDDefaultDock = "steamydock"
        defaultDock = 1 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
        
        'disable all the controls that steamy dock does not support
        
        chkGenMin.Enabled = False ' does not support minimising apps to the dock
        lblChkMinimise.Enabled = False
        'cmbBehaviourActivationFX.Enabled = False
        'cmbStyleTheme.Enabled = False ' does not support themes yet
        'cmbPositionMonitor.Enabled = False
        'cmbIconsQuality.Enabled = False '  does not support enhanced or lower quality icons
        chkIconsZoomOpaque.Enabled = False ' does not support opaque/transparent zoom
        
        sliIconsDuration.Enabled = False ' ' does not support animations at all
        Label16.Enabled = False
        Label19.Enabled = False
        Label18.Enabled = False
        lblIconsDurationMsCurrent.Enabled = False
        
        
        'chkGenOpen.Enabled = False ' does not support showing opening running applications, always opens new apps.
        'lblChkOpenRunning.Enabled = False
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        sliIconsZoomWidth.Enabled = False ' does not support zoomwidth though this is possible later
'        sliIconsDuration.Enabled = False ' does not support animations at all

        '.nn cmbIconsHoverFX.Enabled = False ' does not support hover effects other than the default
        '.nn sliBehaviourAutoHideDuration.Enabled = False ' does not support animation at all
        'sliAnimationInterval.Enabled = False ' does not support animation at all
        chkGenDisableAnim.Enabled = False
        '.nn lblChkDisable.Enabled = False
        
        
        ' Some of the controls have been bundled onto frames so that they can all be hidden entirely for Steamydock users
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
        'fraZoomConfigs.Visible = False
        
        'fraAutoHideDuration.Visible = true
        fraFontOpacity.Visible = True
        
        optGeneralReadConfig.Enabled = True
        lblGeneralReadConfig.Enabled = True
        
        optGeneralWriteConfig.Enabled = True
        lblGeneralWriteConfig.Enabled = True

        lblGenRunAppInterval1.Enabled = True
        lblGenRunAppInterval2.Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenRunAppInterval3.Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
        
        If optGeneralReadConfig.Value = True And steamyDockInstalled = True And rocketDockInstalled = True Then
            chkGenAlwaysAsk.Enabled = True
            lblChkAlwaysConfirm.Enabled = True
        End If

        sliAnimationInterval.Enabled = True
        lblAnimationIntervalLabel.Enabled = True
        lblAnimationIntervalMsLow.Enabled = True
        lblAnimationIntervalMsHigh.Enabled = True
        lblAnimationIntervalMsCurrent.Enabled = True
        lblAnimationInformationLabel.Enabled = True
        
        cmbBehaviourAutoHideType.Enabled = True
        
        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        picThemeSample.Enabled = True
        lblThemeSizeText.Enabled = True
        lblThemeSizeTextLow.Enabled = True
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
        
        lblContinuousHide.Enabled = True
        lblContinuousHideMsLow.Enabled = True
        sliContinuousHide.Enabled = True
        lblContinuousHideMsHigh.Enabled = True
        lblContinuousHideMsCurrent.Enabled = True
        
        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        
    End If
    
   On Error GoTo 0
   Exit Sub

cmbDefaultDock_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDefaultDock_Change of Form dockSettings"

   On Error GoTo 0
   Exit Sub

cmbDefaultDock_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDefaultDock_Click of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   If debugflg = 1 Then DebugPrint "%btnFacebook_Click"

    mnuFacebook_Click

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   If debugflg = 1 Then DebugPrint "%btnUpdate_Click"

    mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeMain_MouseDown
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo fmeMain_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%fmeMain_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    
    ' setting capture of the mouseEnter event on the frame
'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hWnd Then 'MouseEnter
'                Call SetCapture(.hWnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then DebugPrint "%fmeMain_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If
    
   On Error GoTo 0
   Exit Sub

fmeMain_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeMain_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fmeSizePreview_Click
' Author    : beededea
' Date      : 01/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub fmeSizePreview_Click()
   On Error GoTo fmeSizePreview_Click_Error
   If debugflg = 1 Then DebugPrint "%fmeSizePreview_Click"
   
    ' setting capture of the mouseEnter event on the frame
'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hWnd Then 'MouseEnter
'                Call SetCapture(.hWnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then DebugPrint "%fmeMain_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If

   On Error GoTo 0
   Exit Sub

fmeSizePreview_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fmeSizePreview_Click of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Form_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%Form_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_MouseMove
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : If the resizing previews are covered by another window then thet are blanked
'             when a mouse enters the form, if the panel is showing the previews are redrawn
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Form_MouseMove_Error
   'If debugflg = 1 Then DebugPrint "%Form_MouseMove"
   
' setting capture of the mouseEnter event on the form causes weird delays on the whole operation
' of the form controls, so it is now commented out

'    If fmeMain(1).Visible = True Then
'       With Me
'            If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then 'MouseLeave
'                Call ReleaseCapture
'            ElseIf GetCapture() <> .hwnd Then 'MouseEnter
'                Call SetCapture(.hwnd)
'                    Call sliIconsSize_Change
'                    Call sliIconsZoom_Change
'                    If debugflg = 1 Then DebugPrint "%Form_MouseEnter"
'            Else
'                'Normal MouseMove here
'            End If
'        End With
'    End If
'    If fmeMain(1).Visible = True Then
'        Call sliIconsSize_Change
'        Call sliIconsZoom_Change
'    End If
   On Error GoTo 0
   Exit Sub

Form_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 01/03/2020
' Purpose   : the icon sizing images need to redraw when the window returns from being minimised
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error
   If debugflg = 1 Then DebugPrint "%Form_Resize"

    If fmeMain(1).Visible = True Then
        Call sliIconsSize_Change
        Call sliIconsZoom_Change
    End If

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : What to do when unloading the main form
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
        
    ' variables declared
    Dim NameProcess As String
    Dim ofrm As Form
    
    'initialise the dimensioned variables
    NameProcess = ""

    On Error GoTo Form_Unload_Error
    If debugflg = 1 Then DebugPrint "%" & "Form_Unload"
    
    NameProcess = "PersistentDebugPrint.exe"

    If debugflg = 1 Then
        checkAndKill NameProcess, False
    End If

    
    'this was initially commented out as it caused a crash on exit in Win 7 (only) subsequent to the two Krool's
    'controls being added or perhaps it was the failure to close GDI properly
    'then I added it back in as an END is the wrong thing to do supposedly - but I do like a good END.
    
    For Each ofrm In Forms
        'fcount = fcount + 1
        Unload ofrm
    Next
    
'    Sleep 5000
'    MsgBox ("END " & fcount)
    
    'End ' on 32bit Windows this causes a crash and untidy exit so removed.
   
   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form dockSettings"
    
End Sub



Private Sub lblAboutPara4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub


Private Sub lblAboutPara5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblPunklabsLink_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblPunklabsLink_Click(Index As Integer)
   On Error GoTo lblPunklabsLink_Click_Error
   If debugflg = 1 Then Debug.Print "%lblPunklabsLink_Click"

        Call ShellExecute(Me.hwnd, "Open", "http://www.punklabs.com", vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

lblPunklabsLink_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblPunklabsLink_Click of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    ' set the menu checks
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Enable"
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Disable"
            themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    
    rDSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error

    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    rDSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form dockSettings"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(redC As Integer, greenC As Integer, blueC As Integer)
    
        
    ' variables declared
    Dim a As Long
    Dim Ctrl As Control
    Dim useloop As Integer
    
    'initialise the dimensioned variables
     a = 0
     'Ctrl As Control
     useloop = 0
    
    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    ' a method of looping through all the controls that require reversion of any background colouring
    For Each Ctrl In dockSettings.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
    Else
        classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
    End If
    
    ' these elements are normal elements that should have their styling reverted
    ' the loop above changes the background colour and we don't want that for all items
    
    iconBox.BackColor = vbWhite
    
    ' loop through the selection icons and revert these to white
    For useloop = 0 To 5
        
        lblText(useloop).BackColor = vbWhite
        picIcon(useloop).BackColor = vbWhite
        'lblText(useloop).BackColor = vbWhite
    Next useloop
    
    fmeLblFrame(0).BackColor = vbWhite
    fmeLblFrame(1).BackColor = vbWhite
    fmeLblFrame(2).BackColor = vbWhite

    ' now set the frames that underly the selection icons and revert these to white
    fmeIconGeneral.BackColor = vbWhite
    fmeIconBehaviour.BackColor = vbWhite
    fmeIconAbout.BackColor = vbWhite
    fmeIconStyle.BackColor = vbWhite
    fmeIconIcons.BackColor = vbWhite
    fmeIconPosition.BackColor = vbWhite
    
    ' labels within the preview box that must stay the high contrast colours
    Label9.BackColor = RGB(212, 208, 199)
    Label13.BackColor = RGB(212, 208, 199)
    Label1.BackColor = RGB(212, 208, 199)
    
    picStylePreview.BackColor = RGB(212, 208, 199)
    picSizePreview.BackColor = RGB(212, 208, 199)
    picZoomSize.BackColor = RGB(212, 208, 199)
    picMinSize.BackColor = RGB(212, 208, 199)
    
    ' now style the reamining elements by hand to the lighter theme colour RGB(redC, greenC, blueC)
    
    'all other buttons go here
    
    btnGeneralRdFolder.BackColor = RGB(redC, greenC, blueC)
    
    sliBehaviourAutoHideDuration.BackColor = RGB(redC, greenC, blueC)
    sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    sliBehaviourAutoHideDelay.BackColor = RGB(redC, greenC, blueC)
    sliBehaviourPopUpDelay.BackColor = RGB(redC, greenC, blueC)
    
    '.0n DAEB Added themeing to two new sliders
    sliStyleThemeSize.BackColor = RGB(redC, greenC, blueC)
    sliStyleFontOpacity.BackColor = RGB(redC, greenC, blueC)
    
    sliContinuousHide.BackColor = RGB(redC, greenC, blueC)
    
    'general tab slider
    sliGenRunAppInterval.BackColor = RGB(redC, greenC, blueC)

    
    'style tab sliders
    
    sliStyleOpacity.BackColor = RGB(redC, greenC, blueC)
    sliStyleShadowOpacity.BackColor = RGB(redC, greenC, blueC)
    sliStyleOutlineOpacity.BackColor = RGB(redC, greenC, blueC)
    
    'position tab sliders
    
    sliPositionCentre.BackColor = RGB(redC, greenC, blueC)
    sliPositionEdgeOffset.BackColor = RGB(redC, greenC, blueC)
    
    ' icons tab picture and frame elements
        
'    picSizePreview.BackColor = RGB(redC, greenC, blueC)
'    Label9.BackColor = RGB(redC, greenC, blueC)
'    Label13.BackColor = RGB(redC, greenC, blueC)
'    Label1
    
    ' icons tab picboxes
    
'    picZoomSize.BackColor = RGB(redC, greenC, blueC)
'    picMinSize.BackColor = RGB(redC, greenC, blueC)
'    picHiddenPicture.BackColor = RGB(redC, greenC, blueC)
    
    ' icons tab sliders
    
    sliIconsOpacity.BackColor = RGB(redC, greenC, blueC)
    sliIconsSize.BackColor = RGB(redC, greenC, blueC)
    sliIconsZoom.BackColor = RGB(redC, greenC, blueC)
    sliIconsZoomWidth.BackColor = RGB(redC, greenC, blueC)
    sliIconsDuration.BackColor = RGB(redC, greenC, blueC)
    
    PutINISetting "Software\SteamyDockSettings", "SkinTheme", rDSkinTheme, toolSettingsFile ' now saved to the toolsettingsfile

    End Sub
'---------------------------------------------------------------------------------------
' Procedure : picCogs1_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picCogs1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo picCogs1_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%picCogs1_MouseDown"
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

picCogs1_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picCogs1_MouseDown of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteConfig_Click
' Author    : beededea
' Date      : 05/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralWriteConfig_Click()

   On Error GoTo optGeneralWriteConfig_Click_Error

    If startupFlg = True Then '.NET
        ' don't do this on the first startup run
        Exit Sub
    Else
    
        rDGeneralWriteConfig = optGeneralWriteConfig.Value ' this is the nub


'        If optGeneralWriteConfig.Value = True Then
'            rDGeneralWriteConfig = "True"
'
'        Else
'            rDGeneralWriteConfig = "False"
'        End If
    
    End If

   On Error GoTo 0
   Exit Sub

optGeneralWriteConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteConfig_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteRegistry_Click
' Author    : beededea
' Date      : 05/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralWriteRegistry_Click()
   On Error GoTo optGeneralWriteRegistry_Click_Error
   
    If optGeneralWriteRegistry.Value = True Then
        ' nothing to do, the checkbox value is used later to determine where to write the data
    End If
    If defaultDock = 0 Then optGeneralReadRegistry.Value = True ' if running Rocketdock the two must be kept in sync
    
    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralWriteRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteRegistry_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteSettings_Click
' Author    : beededea
' Date      : 01/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optGeneralWriteSettings_Click()

   On Error GoTo optGeneralWriteSettings_Click_Error

    tmpSettingsFile = rdAppPath & "\tmpSettings.ini" ' temporary copy of Rocketdock 's settings file

    If startupFlg = True Then '.NET
        ' don't do this on the first startup run
        Exit Sub
    Else
    
        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
            ' create a settings.ini file in the rocketdock folder
            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
            Close #1         ' close the file and
             ' test it exists
            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
            End If
        End If
    
        If defaultDock = 0 Then ' Rocketdock
            If optGeneralWriteSettings.Value = True Then ' keep the two in synch.
                If optGeneralReadSettings.Value = False Then
                    optGeneralReadSettings.Value = True
                End If
            End If
        End If
    End If
    
    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralWriteSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteSettings_Click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : piccogs2_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub piccogs2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo piccogs2_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%piccogs2_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

piccogs2_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure piccogs2_MouseDown of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picIcon_Click
' Author    : beededea
' Date      : 27/02/2020
' Purpose   : remembering which tab was last clicked
'---------------------------------------------------------------------------------------
'
Private Sub picIcon_Click(Index As Integer)

   On Error GoTo picIcon_Click_Error
   If debugflg = 1 Then DebugPrint "%picIcon_Click"
   
    rDOptionsTabIndex = Index + 1
    
    If rocketDockInstalled = True Then
        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
            PutINISetting "Software\RocketDock", "OptionsTabIndex", rDOptionsTabIndex, origSettingsFile
        Else
            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "OptionsTabIndex", rDOptionsTabIndex)
        End If
    Else
        'CFG - write the current open tab to the 3rd config settings
        PutINISetting "Software\RocketDock", "OptionsTabIndex", rDOptionsTabIndex, origSettingsFile
    End If
    
    fmeMain(0).Visible = False
    fmeMain(1).Visible = False
    fmeMain(2).Visible = False
    fmeMain(3).Visible = False
    fmeMain(4).Visible = False
    fmeMain(5).Visible = False
    
    fmeMain(Index).Visible = True

    fmeMain(Index).Left = 1245
    fmeMain(Index).Top = 30
    
    ' ensure the resizing icons always display.
    ' it seems that when the picturebox is hidden and given focus then the images are lost.
    ' calling these routines restores the images.
        
    picIcon(0).Picture = LoadPicture(App.Path & "\general.gif")
    picIcon(1).Picture = LoadPicture(App.Path & "\icons.gif")
    picIcon(2).Picture = LoadPicture(App.Path & "\behaviour.gif")
    picIcon(3).Picture = LoadPicture(App.Path & "\style.gif")
    picIcon(4).Picture = LoadPicture(App.Path & "\position.gif")
    picIcon(5).Picture = LoadPicture(App.Path & "\about.gif")
    
    If Index = 0 Then
        picIcon(0).Picture = LoadPicture(App.Path & "\generalHighlighted.gif")
    End If
    If Index = 1 Then
        picIcon(1).Picture = LoadPicture(App.Path & "\iconsHighlighted.gif")
    End If
    If Index = 2 Then
        picIcon(2).Picture = LoadPicture(App.Path & "\behaviourHighlighted.gif")
    End If
    If Index = 3 Then
        picIcon(3).Picture = LoadPicture(App.Path & "\styleHighlighted.gif")
    End If
    If Index = 4 Then
        picIcon(4).Picture = LoadPicture(App.Path & "\positionHighlighted.gif")
    End If
    If Index = 5 Then
        picIcon(5).Picture = LoadPicture(App.Path & "\aboutHighlighted.gif")
    End If
    
    
   On Error GoTo 0
   Exit Sub

picIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picIcon_Click of Form Form1"
    
End Sub



    
'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
        
    ' variables declared
    Dim toolSettingsDir As String
    
    'initialise the dimensioned variables
    toolSettingsDir = ""
    
    On Error GoTo getToolSettingsFile_Error
    If debugflg = 1 Then DebugPrint "%getToolSettingsFile"
    
    toolSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\dockSettings" ' just for this user alone
    toolSettingsFile = toolSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not DirExists(toolSettingsDir) Then
        MkDir toolSettingsDir
    End If
    
    'if the settings.ini does not exist then create the file by copying
    If Not FExists(toolSettingsFile) Then
        FileCopy App.Path & "\settings.ini", toolSettingsFile
    End If
    
    'confirm the settings file exists, if not use the version in the app itself
    If Not FExists(toolSettingsFile) Then
        toolSettingsFile = App.Path & "\settings.ini"
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form dockSettings"

End Sub
    


''---------------------------------------------------------------------------------------
'' Procedure : SpecialFolder
'' Author    :  si_the_geek vbforums
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function SpecialFolder(pFolder As eSpecialFolders) As String
''Returns the path to the specified special folder (AppData etc)
'
'    ' variables declared
'    Dim objShell  As Object
'    Dim objFolder As Object
'
'    'initialise the dimensioned variables
'    Set objShell = Nothing
'    Set objFolder = Nothing
'
'   On Error GoTo SpecialFolder_Error
'   If debugflg = 1 Then DebugPrint "%SpecialFolder"
'
'  Set objShell = CreateObject("Shell.Application")
'  Set objFolder = objShell.NameSpace(CLng(pFolder))
'
'  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.path
'
'  Set objFolder = Nothing
'  Set objShell = Nothing
'
'  If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"
'
'   On Error GoTo 0
'   Exit Function
'
'SpecialFolder_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SpecialFolder of Module Module1"
'
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : checkLicenceState
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : 'check the state of the licence
''---------------------------------------------------------------------------------------
''
'Private Sub checkLicenceState()
'
'    ' variables declared
'    Dim slicence As Integer
'
'    'initialise the dimensioned variables
'    slicence = 0
'
'    On Error GoTo checkLicenceState_Error
'
'    'toolSettingsFile = toolSettingsDir & "\settings.ini"
'    'toolSettingsFile = App.Path & "\settings.ini"
'    ' read the tool's own settings file (
'    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        slicence = GetINISetting("Software\SteamyDockSettings", "Licence", toolSettingsFile)
'        ' if the licence state is not already accepted then display the licence form
'        If slicence = 0 Then
'            Call LoadFileToTB(licence.txtLicenceTextBox, App.path & "\licence.txt", False)
'
'            licence.Show vbModal ' show the licence screen in VB modal mode (ie. on its own)
'            ' on the licence box change the state fo the licence acceptance
'        End If
'    End If
'
'    ' show the licence screen if it has never been run before and set it to be in focus
'    If licence.Visible = True Then
'        licence.SetFocus
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'checkLicenceState_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkLicenceState of Form dockSettings"
'
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : checkRocketdockInstallation
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub checkRocketdockInstallation()
'
'    ' variables declared
'    Dim answer As VbMsgBoxResult
'
'    'initialise the dimensioned variables
'    answer = vbNo
'
'    RD86installed = ""
'    RDinstalled = ""
'
'    ' check where rocketdock is installed
'    On Error GoTo checkRocketdockInstallation_Error
'
'    RD86installed = driveCheck("Program Files (x86)\Rocketdock", "RocketDock.exe")
'    RDinstalled = driveCheck("Program Files\Rocketdock", "RocketDock.exe")
'
'    If RDinstalled = "" And RD86installed = "" Then
'        rocketDockInstalled = False
'        'answer = MsgBox(" Rocketdock has not been installed in the program files (x86) folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
'        txtGeneralRdLocation.Text = ""
'        Exit Sub
'
'    Else
'        rocketDockInstalled = True
'        If RDinstalled <> "" Then
'            rdAppPath = RDinstalled
'        End If
'        'the one in the x86 folder has precedence
'        If RD86installed <> "" Then
'            rdAppPath = RD86installed
'        End If
'
'    End If
'
'    ' If rocketdock Is Not installed Then test the registry
'    ' if the registry settings are not located then remove them as a source.
'
'    ' rocketDockInstalled = False ' debug
'
'    ' read selected random entries from the registry, if each are false then the RD registry entries do not exist.
'    If rocketDockInstalled = False Then
'        rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIconsd")
'        rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunnings")
'        rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunnings")
'        rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindowsw")
'        rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimations")
'        If rDLockIcons = "" And rDOpenRunning = "" And rDShowRunning = "" And rDManageWindows = "" And rDDisableMinAnimation = "" Then
'            ' rocketdock registry entries do not exist so RD has never been installed or it has been wiped entirely.
'            RDregistryPresent = False
'        Else
'            RDregistryPresent = True 'rocketdock HAS been installed in the past as the registry entries are still present
'        End If
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'checkRocketdockInstallation_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Form dockSettings"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : readDockSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readDockSettings()

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    
    ' the first is the RD settings file that only exists if RD is NOT using the registry
    ' the second is the settings file for this tool to store its own preferences
        
    ' check to see if the first settings file exists
    
    On Error GoTo readDockSettings_Error
   
    If rocketDockInstalled = True Then
        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
            If optGeneralReadConfig.Value = False Then
                optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
                optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
            End If
            ' here we read from the settings file
            readDockSettingsFile "Software\RocketDock", origSettingsFile
            Call validateInputs
        Else
            If optGeneralReadConfig.Value = False Then
                optGeneralReadRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
                optGeneralWriteRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
            End If
            
            ' read the dock configuration from the registry into variables
            Call readRegistry
        End If
        
    End If
    
    On Error GoTo 0
    Exit Sub

readDockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockSettings of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readAndSetUtilityFont
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : reads the tool's font settings from the local tool file
'---------------------------------------------------------------------------------------
'
Private Sub readAndSetUtilityFont()
  
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedStyle As String
    'Dim suppliedColour As Variant
    
    'initialise the dimensioned variables
    suppliedFont = ""
    suppliedSize = 0
    suppliedWeight = 0
    suppliedStyle = False
    'suppliedColour = Empty

    On Error GoTo readAndSetUtilityFont_Error
    
    ' set the tool's default font
    suppliedFont = GetINISetting("Software\SteamyDockSettings", "defaultFont", toolSettingsFile)
    suppliedSize = Val(GetINISetting("Software\SteamyDockSettings", "defaultSize", toolSettingsFile))
    suppliedWeight = Val(GetINISetting("Software\SteamyDockSettings", "defaultStrength", toolSettingsFile))
    suppliedStyle = GetINISetting("Software\SteamyDockSettings", "defaultStyle", toolSettingsFile)
    rDSkinTheme = GetINISetting("Software\SteamyDockSettings", "SkinTheme", toolSettingsFile)
        
        
    If Not suppliedFont = "" Then
        Call changeFont(suppliedFont, suppliedSize, suppliedWeight, CBool(LCase(suppliedStyle)))
    End If

   On Error GoTo 0
   Exit Sub

readAndSetUtilityFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readAndSetUtilityFont of Form dockSettings on line " & Erl
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPreviewFontColours
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPreviewFontColours(suppliedColour)
   On Error GoTo setPreviewFontColours_Error
   If debugflg = 1 Then DebugPrint "%setPreviewFontColours"

    lblPreviewFont.ForeColor = suppliedColour

   On Error GoTo 0
   Exit Sub

setPreviewFontColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPreviewFontColours of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : setPreviewConvertedFontColours
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPreviewConvertedFontColours(suppliedColour)
        
   On Error GoTo setPreviewConvertedFontColours_Error
   If debugflg = 1 Then DebugPrint "%setPreviewConvertedFontColours"

    lblPreviewFont.ForeColor = Convert_Dec2RGB(suppliedColour)

   On Error GoTo 0
   Exit Sub

setPreviewConvertedFontColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPreviewConvertedFontColours of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : writeRegistry
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : utility needs admin to write to Rocketdock's registry entries
'---------------------------------------------------------------------------------------
'
Private Sub writeRegistry()
   
    On Error GoTo writeRegistry_Error

    ' all tested and working but ONLY when run as admin
    
    'general panel

    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", rDLockIcons)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OpenRunning", rDOpenRunning)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ShowRunning", rDShowRunning)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ManageWindows", rDManageWindows)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "DisableMinAnimation", rDDisableMinAnimation)
    
    'icon panel
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconQuality", Val(rDIconQuality))
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconOpacity", rDIconOpacity)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomOpaque", rDZoomOpaque)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMin", rDIconMin)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HoverFX", rDHoverFX)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMax", rdIconMax)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomWidth", Val(rDZoomWidth))
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomTicks", rDZoomTicks)
    
    'behaviour panel
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconActivationFX", rDIconActivationFX)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHide", rDAutoHide) '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHideTicks", rDAutoHideTicks)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHideDelay", rDAutoHideDelay)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "MouseActivate", rDMouseActivate)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "PopupDelay", rDPopupDelay)
    
    
    'position panel
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Monitor", rDMonitor)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", rDSide)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "zOrderMode", rDzOrderMode)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Offset", rDOffset)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "vOffset", rDvOffset)
        
    'style panel
    'If rDtheme = "blank" Then rDtheme = ""
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "theme", rDtheme)
    'If rDtheme = "" Then rDtheme = "blank"
    
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ThemeOpacity", rDThemeOpacity)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HideLabels", rDHideLabels)

    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontName", rDFontName) '*
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontColor", rDFontColor) '*
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontSize", rDFontSize)
    'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontCharSet", rD)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontFlags", rDFontFlags) '*

    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowColor", rDFontShadowColor)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineColor", rDFontOutlineColor)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineOpacity", rDFontOutlineOpacity)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowOpacity", rDFontShadowOpacity)
    
   On Error GoTo 0
   Exit Sub

writeRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistry of Form dockSettings"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : picStylePreview_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picStylePreview_Click()
       
    ' variables declared
    Dim colourResult As Long
        
    'initialise the dimensioned variables
    colourResult = 0
    
    On Error GoTo picStylePreview_Click_Error

    colourResult = ShowColorDialog(Me.hwnd, True, rDFontShadowColor)

    If colourResult <> -1 And colourResult <> 0 Then
        picStylePreview.BackColor = colourResult
    End If

   On Error GoTo 0
   Exit Sub

picStylePreview_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picStylePreview_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : repaintTimer_Timer
' Author    : beededea
' Date      : 09/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub repaintTimer_Timer()

   On Error GoTo repaintTimer_Timer_Error
   If debugflg = 1 Then Debug.Print "%repaintTimer_Timer"
    If fmeMain(1).Visible = True Then
        Call sliIconsSize_Change
        Call sliIconsZoom_Change
    End If
   On Error GoTo 0
   Exit Sub

repaintTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure repaintTimer_Timer of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliAnimationInterval_Change
' Author    : beededea
' Date      : 10/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliAnimationInterval_Change()

   On Error GoTo sliAnimationInterval_Change_Error
    lblAnimationIntervalMsCurrent.Caption = "(" & sliAnimationInterval.Value & ")"

    rDAnimationInterval = sliAnimationInterval.Value


   On Error GoTo 0
   Exit Sub

sliAnimationInterval_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliAnimationInterval_Change of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliBehaviourAutoHideDelay_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliBehaviourAutoHideDelay_Change()
   On Error GoTo sliBehaviourAutoHideDelay_Click_Error
   If debugflg = 1 Then DebugPrint "%sliBehaviourAutoHideDelay_Click"

    lblAutoHideDelayMsCurrent.Caption = "(" & 3 + (sliBehaviourAutoHideDelay.Value / 1000) & ") secs"
    
    rDAutoHideDelay = sliBehaviourAutoHideDelay.Value

   On Error GoTo 0
   Exit Sub

sliBehaviourAutoHideDelay_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliBehaviourAutoHideDelay_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliBehaviourAutoHideDuration_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliBehaviourAutoHideDuration_Change()
   On Error GoTo sliBehaviourAutoHideDuration_Change_Error
   If debugflg = 1 Then DebugPrint "%sliBehaviourAutoHideDuration_Change"

    lblAutoHideDurationMsCurrent.Caption = "(" & sliBehaviourAutoHideDuration.Value & ")"

    rDAutoHideTicks = sliBehaviourAutoHideDuration.Value

   On Error GoTo 0
   Exit Sub

sliBehaviourAutoHideDuration_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliBehaviourAutoHideDuration_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliBehaviourPopUpDelay_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliBehaviourPopUpDelay_Change()
   On Error GoTo sliBehaviourPopUpDelay_Change_Error
   If debugflg = 1 Then DebugPrint "%sliBehaviourPopUpDelay_Change"

    lblBehaviourPopUpDelayMsCurrrent.Caption = "(" & sliBehaviourPopUpDelay.Value & ")"
    
    rDPopupDelay = sliBehaviourPopUpDelay.Value

   On Error GoTo 0
   Exit Sub

sliBehaviourPopUpDelay_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliBehaviourPopUpDelay_Change of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : sliContinuousHide_Change
' Author    : beededea
' Date      : 25/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliContinuousHide_Change()

   On Error GoTo sliContinuousHide_Change_Error

    If sliContinuousHide.Value = 1 Then
        lblContinuousHideMsCurrent.Caption = "(" & sliContinuousHide.Value & ") min"
    Else
        lblContinuousHideMsCurrent.Caption = "(" & sliContinuousHide.Value & ") mins"
    End If
    sDContinuousHide = sliContinuousHide.Value

   On Error GoTo 0
   Exit Sub

sliContinuousHide_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliContinuousHide_Change of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliGenRunAppInterval_Change
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliGenRunAppInterval_Change()

   On Error GoTo sliGenRunAppInterval_Change_Error

    lblGenRunAppIntervalCur.Caption = "(" & sliGenRunAppInterval.Value & " seconds)"
    rDRunAppInterval = sliGenRunAppInterval.Value

   On Error GoTo 0
   Exit Sub

sliGenRunAppInterval_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGenRunAppInterval_Change of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsDuration_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsDuration_Change()
   On Error GoTo sliIconsDuration_Change_Error
   If debugflg = 1 Then DebugPrint "%sliIconsDuration_Change"

    lblIconsDurationMsCurrent.Caption = "(" & sliIconsDuration.Value & "ms)"
    
    rDZoomTicks = sliIconsDuration.Value

   On Error GoTo 0
   Exit Sub

sliIconsDuration_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsDuration_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsOpacity_Change()
   On Error GoTo sliIconsOpacity_Change_Error
   If debugflg = 1 Then DebugPrint "%sliIconsOpacity_Change"

    lblIconsOpacity.Caption = "(" & sliIconsOpacity.Value & "%)"
    
    rDIconOpacity = sliIconsOpacity.Value

   On Error GoTo 0
   Exit Sub

sliIconsOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsOpacity_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsSize_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsSize_Change()
        
    ' variables declared
    Dim newSize As Integer
        
    'initialise the dimensioned variables
    newSize = 0

    On Error GoTo sliIconsSize_Change_Error
    If debugflg = 1 Then DebugPrint "%sliIconsSize_Change"

    lblIconsSize.Caption = "(" & sliIconsSize.Value & "px)"
          
    newSize = PixelsToTwips(sliIconsSize.Value)
    picMinSize.Cls
    Call picMinSize.PaintPicture(picHiddenPicture, 60 + (1920 / 2) - (newSize / 2), 60 + (1920 / 2) - (newSize / 2), newSize, newSize)

    rDIconMin = sliIconsSize.Value

   On Error GoTo 0
   Exit Sub

sliIconsSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsSize_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliIconsZoom_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsZoom_Change()
       
    ' variables declared
    Dim newSize As Long
        
    'initialise the dimensioned variables
    newSize = 0
    
   On Error GoTo sliIconsZoom_Change_Error
   If debugflg = 1 Then DebugPrint "%sliIconsZoom_Change"

    lblIconsZoom.Caption = "(" & sliIconsZoom.Value & "px)"
    
    
    Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation
    
    newSize = PixelsToTwips(sliIconsZoom.Value)
    picZoomSize.Cls
    Call picZoomSize.PaintPicture(picHiddenPicture, 60 + (3840 / 2) - (newSize / 2), 60 + (3840 / 2) - (newSize / 2), newSize, newSize)
    
    
'    Call picZoomSize.PaintPicture(picHiddenPicture, 60 + (1920 / 2) - (newSize / 2), 60 + (1920 / 2) - (newSize / 2), newSize, newSize)
'    Call picZoomSize.PaintPicture(picHiddenPicture, 60, 60, newSize, newSize)
'
'    'picZoomSize.Left = 2640 + (3840 / 2) - (3840 / 2)
'    picZoomSize.Top = 100 + (3840 - newSize)

    rdIconMax = sliIconsZoom.Value



    
    

   On Error GoTo 0
   Exit Sub

sliIconsZoom_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsZoom_Change of Form Form1"
End Sub

   '---------------------------------------------------------------------------------------
    ' Procedure : vb6TwipsToPixels
    ' Author    : beededea
    ' Date      : 17/10/2019
    ' Purpose   : VB6 polyfills, not using VB6 compatibility mode
    ' doing away with VB6 compatibility mode will remove the 32bit limitation...
    '---------------------------------------------------------------------------------------
    '
    Public Function TwipsToPixels(ByVal intTwips As Integer) As Integer
                
    ' variables declared
    Dim nTwips As Integer
   
    'initialise the dimensioned variables
    nTwips = 0

            'vb6TwipsToPixels = intTwips * g.DpiX / 1440
            nTwips = intTwips / Screen.twipsPerPixelX

            TwipsToPixels = nTwips
    End Function

'---------------------------------------------------------------------------------------
' Procedure : PixelsToTwips
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : RD works with pixels but VB6 works with twips
'---------------------------------------------------------------------------------------
'
Public Function PixelsToTwips(ByVal intPixels As Integer) As Integer

        
    ' variables declared
    Dim nTwips As Integer
        
    'initialise the dimensioned variables
    nTwips = 0
    
   On Error GoTo PixelsToTwips_Error
   If debugflg = 1 Then DebugPrint "%PixelsToTwips"

            'vb6PixelsToTwips = intPixels / g.DpiX * 1440
            nTwips = intPixels * Screen.twipsPerPixelX
            
            PixelsToTwips = nTwips

   On Error GoTo 0
   Exit Function

PixelsToTwips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PixelsToTwips of Form dockSettings"

End Function

'---------------------------------------------------------------------------------------
' Procedure : sliIconsZoomWidth_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliIconsZoomWidth_Change()
   On Error GoTo sliIconsZoomWidth_Change_Error
   If debugflg = 1 Then DebugPrint "%sliIconsZoomWidth_Change"

    lblIconsZoomWidth.Caption = "(" & sliIconsZoomWidth.Value & ")"
    
    rDZoomWidth = sliIconsZoomWidth.Value

   On Error GoTo 0
   Exit Sub

sliIconsZoomWidth_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliIconsZoomWidth_Change of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : sliPositionCentre_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliPositionCentre_Change()
   On Error GoTo sliPositionCentre_Change_Error
   If debugflg = 1 Then DebugPrint "%sliPositionCentre_Change"

    lblPositionCentrePercCurrent.Caption = "(" & Val(sliPositionCentre.Value) & "%)"
    
    rDOffset = sliPositionCentre.Value

   On Error GoTo 0
   Exit Sub

sliPositionCentre_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPositionCentre_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliPositionEdgeOffset_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliPositionEdgeOffset_Click()
   On Error GoTo sliPositionEdgeOffset_Click_Error
   If debugflg = 1 Then DebugPrint "%sliPositionEdgeOffset_Click"

    lblPositionEdgeOffsetPxCurrent.Caption = "(" & Val(sliPositionEdgeOffset.Value) & "px)"
    
    rDvOffset = sliPositionEdgeOffset.Value
    
   On Error GoTo 0
   Exit Sub

sliPositionEdgeOffset_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPositionEdgeOffset_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleFontOpacity_Click
' Author    : beededea
' Date      : 17/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleFontOpacity_Click()

   On Error GoTo sliStyleFontOpacity_Click_Error

    lblStyleFontOpacityCurrent.Caption = "(" & Val(sliStyleFontOpacity.Value) & "%)"
    
    sDFontOpacity = sliStyleFontOpacity.Value
    
   On Error GoTo 0
   Exit Sub

sliStyleFontOpacity_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleFontOpacity_Click of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleOpacity_Change()
   On Error GoTo sliStyleOpacity_Change_Error
   If debugflg = 1 Then DebugPrint "%sliStyleOpacity_Change"

    lblStyleOpacityCurrent.Caption = "(" & Val(sliStyleOpacity.Value) & "%)"
    
    rDThemeOpacity = sliStyleOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleOpacity_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleOutlineOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleOutlineOpacity_Change()
   On Error GoTo sliStyleOutlineOpacity_Change_Error
   If debugflg = 1 Then DebugPrint "%sliStyleOutlineOpacity_Change"

    lblStyleOutlineOpacityCurrent.Caption = "(" & Val(sliStyleOutlineOpacity.Value) & "%)"

    rDFontOutlineOpacity = sliStyleOutlineOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleOutlineOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleOutlineOpacity_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliStyleShadowOpacity_Change
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleShadowOpacity_Change()
   On Error GoTo sliStyleShadowOpacity_Change_Error
   If debugflg = 1 Then DebugPrint "%sliStyleShadowOpacity_Change"

    lblStyleShadowOpacityCurrent.Caption = "(" & Val(sliStyleShadowOpacity.Value) & "%)"
    
    rDFontShadowOpacity = sliStyleShadowOpacity.Value

   On Error GoTo 0
   Exit Sub

sliStyleShadowOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleShadowOpacity_Change of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFont_Click
' Author    : beededea
' Date      : 28/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFont_Click()

        
    ' variables declared
    Dim suppliedFont As String
    Dim suppliedSize As String
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedStyle As Boolean
    Dim suppliedColour As Variant
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    Dim fontSelected As Boolean
        
    'initialise the dimensioned variables
    
    suppliedFont = ""
    suppliedSize = 0
    suppliedWeight = 0
    suppliedStyle = False
    suppliedColour = Empty
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False
    fontSelected = False
    
    On Error GoTo mnuFont_Click_Error
    If debugflg = 1 Then DebugPrint "%mnuFont_Click"

    displayFontSelector suppliedFont, Val(suppliedSize), suppliedWeight, suppliedStyle, suppliedColour, suppliedItalics, suppliedUnderline, fontSelected
    If fontSelected = False Then Exit Sub

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\SteamyDockSettings", "defaultFont", suppliedFont, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultSize", suppliedSize, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultStrength", suppliedWeight, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultStyle", suppliedStyle, toolSettingsFile
    End If

    If suppliedWeight > 700 Then
        suppliedBold = True
    Else
        suppliedBold = False
    End If
    
    If suppliedFont <> vbNullString Then
        Call changeFont(suppliedFont, Val(suppliedSize), suppliedWeight, suppliedStyle)
    End If

   On Error GoTo 0
   Exit Sub

mnuFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFont_Click of Form dockSettings"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub changeFont(suppliedFont As String, suppliedSize As Integer, suppliedWeight As Integer, suppliedStyle As Boolean)
        
    ' variables declared
    Dim useloop As Integer
    Dim Ctrl As Control
        
    'initialise the dimensioned variables
    useloop = 0
    'Ctrl
    
    On Error GoTo changeFont_Error
    
    If debugflg = 1 Then DebugPrint "%" & "changeFont"
      
    ' a method of looping through all the controls and identifying the labels and text boxes
    For Each Ctrl In dockSettings.Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
           If suppliedFont <> "" Then Ctrl.Font.Name = suppliedFont
           If suppliedSize > 0 Then Ctrl.Font.Size = suppliedSize
           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
        End If
    Next
    
    ' The comboboxes all autoselect when the font is changed, we need to reset this afterwards

    cmbIconsQuality.SelLength = 0
    cmbIconsHoverFX.SelLength = 0
    cmbDefaultDock.SelLength = 0
    cmbBehaviourActivationFX.SelLength = 0
    cmbStyleTheme.SelLength = 0
    cmbPositionMonitor.SelLength = 0
    cmbPositionScreen.SelLength = 0
    cmbPositionLayering.SelLength = 0
    cmbBehaviourAutoHideType.SelLength = 0
   
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Form dockSettings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(Optional ByRef currFont As String, Optional ByRef currSize As Integer, Optional ByRef currWeight As Integer, Optional ByRef currStyle As Boolean, Optional ByRef currColour, Optional ByRef currItalics As Boolean, Optional ByRef currUnderline As Boolean, Optional ByRef fontResult As Boolean)

       
    ' variables declared
    Dim f As FormFontInfo
        
    'initialise the dimensioned variables
    'f =
   
   On Error GoTo displayFontSelector_Error
   If debugflg = 1 Then DebugPrint "%displayFontSelector"

    With f
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = DialogFont(f)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If f.Name = "" Then f.Name = "times new roman"
    If f.Name = "" Then Exit Sub
    
    With f
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of Form dockSettings"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(Index As Integer)
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuCoffee_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuCoffee_Click"
    
    answer = MsgBox(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelpPdf_click
' Author    : beededea
' Date      : 30/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpPdf_click()
       
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
   On Error GoTo mnuHelpPdf_click_Error
   If debugflg = 1 Then DebugPrint "%mnuHelpPdf_click"

    answer = MsgBox("This option opens a browser window and displays this tool's help. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        If FExists(App.Path & "\help\SteamyDockSettings.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\SteamyDockSettings.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("The help file - SteamyDockSettings.html- is missing from the help folder.")
        End If
    End If

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpPdf_click of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFacebook_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFacebook_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuFacebook_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuFacebook_Click"

    answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFacebook_Click of Form quartermaster"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuLatest_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLatest_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuLatest_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuLatest_Click"

    answer = MsgBox("Download latest version of the program - this button opens a browser window and connects to the widget download page where you can check and download the latest zipped file). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If


    On Error GoTo 0
    Exit Sub

mnuLatest_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLatest_Click of Form quartermaster"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicence_Click()
    On Error GoTo mnuLicence_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuLicence_Click"
        
    Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
    licence.Show

    On Error GoTo 0
    Exit Sub

mnuLicence_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_Click of Form quartermaster"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    On Error GoTo mnuSupport_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuSupport_Click"

    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSweets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSweets_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    

    On Error GoTo mnuSweets_Click_Error
       If debugflg = 1 Then DebugPrint "%" & "mnuSweets_Click"
    
    
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.amazon.co.uk/gp/registry/registry.html?ie=UTF8&id=A3OBFB6ZN4F7&type=wishlist", vbNullString, App.Path, 1)
    End If
    
    On Error GoTo 0
    Exit Sub

mnuSweets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSweets_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuWidgets_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuWidgets_Click()
        
    ' variables declared
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo

    On Error GoTo mnuWidgets_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuWidgets_Click"
    
    answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuWidgets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuWidgets_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClose_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Close the program from the menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuClose_Click()
    On Error GoTo mnuClose_Click_Error
    If debugflg = 1 Then DebugPrint "mnuClose_Click"
    
    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Clickg_Click of Form dockSettings"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuDebug_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Run the runtime debugging window exectuable
'---------------------------------------------------------------------------------------
'
Private Sub mnuDebug_Click()
        
    ' variables declared
    Dim NameProcess As String
    Dim debugPath As String
        
    'initialise the dimensioned variables
    NameProcess = ""
    debugPath = ""
    
    On Error GoTo mnuDebug_Click_Error
    If debugflg = 1 Then DebugPrint "%mnuDebug_Click"

    NameProcess = "PersistentDebugPrint.exe"
    debugPath = App.Path() & "\" & NameProcess
    
    If debugflg = 0 Then
        debugflg = 1
        mnuDebug.Caption = "Turn Debugging OFF"
        If FExists(debugPath) Then
            Call ShellExecute(hwnd, "Open", debugPath, vbNullString, App.Path, 1)
            Sleep (500) ' a 1/2 sec delay is required before the process is ready to listen to debugprint statements
        End If
    Else
        debugflg = 0
        mnuDebug.Caption = "Turn Debugging ON"
        checkAndKill NameProcess, False
    End If

   On Error GoTo 0
   Exit Sub

mnuDebug_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDebug_Click of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click(Index As Integer)
    
    On Error GoTo mnuAbout_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAbout_Click"
          
     about.lblMajorVersion.Caption = App.Major
     about.lblMinorVersion.Caption = App.Minor
     about.lblRevisionNum.Caption = App.Revision
     
     about.Show
     
     If (about.WindowState = 1) Then
         about.WindowState = 0
     End If


    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayVersionNumber
' Author    : beededea
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub displayVersionNumber()
   On Error GoTo displayVersionNumber_Error
   If debugflg = 1 Then DebugPrint "%displayVersionNumber"

     dockSettings.lblMajorVersion.Caption = App.Major
     dockSettings.lblMinorVersion.Caption = App.Minor
     dockSettings.lblRevisionNum.Caption = App.Revision

   On Error GoTo 0
   Exit Sub

displayVersionNumber_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayVersionNumber of Form dockSettings"
End Sub



'An OLE_COLOR value is a BGR (Blue, Green, Red) value. To determine the BGR value, specify blue, green, or red (each of which has a value from 0 - 255) in the following formula:
'
'
'BGR Value = (blue * 65536) + (green * 256) + red
'
'
'r = 238
'G = 239
'B = 221
'
'
'The formula to convert to OLE_COLOR was:
'BGR Value = (blue * 65536) + (green * 256) + red
'
'
'
'221 * 65536 = 14483456 (Blue)
'239 * 256 = 61184          (Green)
'238                                (Red)
'
'
'14483456 + 61184 + 238 = 14544878 (Decimal) or &HDDEFEE (Hex)

'---------------------------------------------------------------------------------------
' Procedure : IsValidOleColor
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IsValidOleColor(ByVal nColor As Long) As Boolean
   On Error GoTo IsValidOleColor_Error

  Select Case nColor
    Case 0& To &H100FFFF, &H2000000 To &H2FFFFFF
         IsValidOleColor = True
    Case &H80000000 To &H80FF0018
         IsValidOleColor = (nColor And &HFFFF&) <= &H18
  End Select

   On Error GoTo 0
   Exit Function

IsValidOleColor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsValidOleColor of Form dockSettings"
End Function



'---------------------------------------------------------------------------------------
' Procedure : sliStyleThemeSize_Change
' Author    : beededea
' Date      : 14/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliStyleThemeSize_Change()

   On Error GoTo sliStyleThemeSize_Change_Error

    lblStyleSizeCurrent.Caption = "(" & Val(sliStyleThemeSize.Value) & "px)"
    
    rDSkinSize = sliStyleThemeSize.Value

   On Error GoTo 0
   Exit Sub

sliStyleThemeSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStyleThemeSize_Change of Form dockSettings"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0

' This should only be required on a machine that can give the Windows classic theme to the UI
' that excludes windows 8 and 10 so this timer can be switched off on these o/s.

   On Error GoTo themeTimer_Timer_Error

    SysClr = GetSysColor(COLOR_BTNFACE)
    If debugflg = 1 Then DebugPrint "COLOR_BTNFACE = " & SysClr ' generates too many debug statements in the log
    If SysClr <> storeThemeColour Then
    
        Call setThemeColour

    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Public Sub setThemeColour()
    
        
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0
    
   On Error GoTo setThemeColour_Error
   If debugflg = 1 Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        rDSkinTheme = "dark"
    Else
        'MsgBox "Windows Alternate Theme detected"
        SysClr = GetSysColor(COLOR_BTNFACE)
        If SysClr = 13160660 Then
            Call setThemeShade(212, 208, 199)
            rDSkinTheme = "light"
        Else ' 15790320
            Call setThemeShade(240, 240, 240)
            rDSkinTheme = "dark"
        End If

    End If

    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of Form dockSettings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setThemeSkin
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setThemeSkin()
   On Error GoTo setThemeSkin_Error

    If rDSkinTheme = "dark" Then
        Call setThemeShade(212, 208, 199)
    Else
        Call setThemeShade(240, 240, 240)
    End If

   On Error GoTo 0
   Exit Sub

setThemeSkin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeSkin of Form dockSettings"
End Sub





'    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        PutINISetting "Software\RocketDockSettings", "defaultFont", suppliedFont, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultSize", suppliedSize, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultStrength", suppliedStrength, toolSettingsFile
'        PutINISetting "Software\RocketDockSettings", "defaultStyle", suppliedStyle, toolSettingsFile
'    End If


'---------------------------------------------------------------------------------------
' Procedure : placeFrames
' Author    : beededea
' Date      : 09/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub placeFrames()

        
    ' variables declared
    Dim topDown As Integer
        
    'initialise the dimensioned variables
    topDown = 0


   On Error GoTo placeFrames_Error
   If debugflg = 1 Then Debug.Print "%placeFrames"

    topDown = 220
    fmeIconGeneral.Top = 200
    fmeIconIcons.Top = fmeIconGeneral.Top + 1185 + topDown
    fmeIconBehaviour.Top = fmeIconIcons.Top + 1185 + topDown
    fmeIconStyle.Top = fmeIconBehaviour.Top + 1185 + topDown
    fmeIconPosition.Top = fmeIconStyle.Top + 1185 + topDown
    fmeIconAbout.Top = fmeIconPosition.Top + 1185 + topDown

   On Error GoTo 0
   Exit Sub

placeFrames_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure placeFrames of Form dockSettings"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : writeDockSettings
' Author    : beededea
' Date      : 12/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub writeDockSettings(location As String, settingsFile As String)

' Alternative settings.ini file called docksettings.ini
' partitioned as follows:
'
' [Software\SteamyDock\DockSettings]
' [Software\SteamyDock\IconSettings]
' [Software\SteamyDock\SteamyDock\DockSettings]

    On Error GoTo writeDockSettings_Error
    If debugflg = 1 Then Debug.Print "%writeDockSettings"
    
    ' first we save the Steamydock specific settings
    If FExists(dockSettingsFile) Then
        PutINISetting "Software\SteamyDock\DockSettings", "GeneralReadConfig", rDGeneralReadConfig, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "GeneralWriteConfig", rDGeneralWriteConfig, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "RunAppInterval", rDRunAppInterval, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "DefaultDock", rDDefaultDock, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "AnimationInterval", rDAnimationInterval, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "SkinSize", rDSkinSize, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "SplashStatus", sDSplashStatus, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "ShowIconSettings", sDShowIconSettings, dockSettingsFile ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock

        PutINISetting "Software\SteamyDock\DockSettings", "FontOpacity", sDFontOpacity, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "AutoHideType", sDAutoHideType, dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "ShowLblBacks", sDShowLblBacks, dockSettingsFile ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
        PutINISetting "Software\SteamyDock\DockSettings", "ContinuousHide", sDContinuousHide, dockSettingsFile   'nn
        PutINISetting "Software\SteamyDock\DockSettings", "BounceZone", sDBounceZone, dockSettingsFile   'nn
    'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "ContinuousHide", sDContinuousHide)
   ' ContinuousHide
   End If
    
    ' then we save those associated to both Rocketdock and SteamyDock
    PutINISetting location, "Version", rDVersion, settingsFile
    PutINISetting location, "HotKey-Toggle", rDHotKeyToggle, settingsFile
    PutINISetting location, "Theme", rDtheme, settingsFile
    PutINISetting location, "ThemeOpacity", rDThemeOpacity, settingsFile
    PutINISetting location, "IconOpacity", rDIconOpacity, settingsFile
    PutINISetting location, "FontSize", rDFontSize, settingsFile
    PutINISetting location, "FontFlags", rDFontFlags, settingsFile
    PutINISetting location, "FontName", rDFontName, settingsFile
    PutINISetting location, "FontColor", rDFontColor, settingsFile
    PutINISetting location, "FontCharSet", rDFontCharSet, settingsFile
    PutINISetting location, "FontOutlineColor", rDFontOutlineColor, settingsFile
    PutINISetting location, "FontOutlineOpacity", rDFontOutlineOpacity, settingsFile
    PutINISetting location, "FontShadowColor", rDFontShadowColor, settingsFile
    PutINISetting location, "FontShadowOpacity", rDFontShadowOpacity, settingsFile
    PutINISetting location, "IconMin", rDIconMin, settingsFile
    PutINISetting location, "IconMax", rdIconMax, settingsFile
    PutINISetting location, "ZoomWidth", rDZoomWidth, settingsFile
    PutINISetting location, "ZoomTicks", rDZoomTicks, settingsFile
    PutINISetting location, "AutoHide", rDAutoHide, settingsFile '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
    PutINISetting location, "AutoHideTicks", rDAutoHideTicks, settingsFile
    PutINISetting location, "AutoHideDelay", rDAutoHideDelay, settingsFile
    PutINISetting location, "PopupDelay", rDPopupDelay, settingsFile
    PutINISetting location, "IconQuality", rDIconQuality, settingsFile
    PutINISetting location, "LangID", rDLangID, settingsFile
    PutINISetting location, "HideLabels", rDHideLabels, settingsFile
    PutINISetting location, "ZoomOpaque", rDZoomOpaque, settingsFile
    PutINISetting location, "LockIcons", rDLockIcons, settingsFile
    PutINISetting location, "ManageWindows", rDManageWindows, settingsFile
    PutINISetting location, "DisableMinAnimation", rDDisableMinAnimation, settingsFile
    PutINISetting location, "ShowRunning", rDShowRunning, settingsFile
    PutINISetting location, "OpenRunning", rDOpenRunning, settingsFile
    PutINISetting location, "HoverFX", rDHoverFX, settingsFile
    PutINISetting location, "zOrderMode", rDzOrderMode, settingsFile
    PutINISetting location, "MouseActivate", rDMouseActivate, settingsFile
    PutINISetting location, "IconActivationFX", rDIconActivationFX, settingsFile
    PutINISetting location, "Monitor", rDMonitor, settingsFile
    PutINISetting location, "Side", rDSide, settingsFile
    PutINISetting location, "Offset", rDOffset, settingsFile
    PutINISetting location, "vOffset", rDvOffset, settingsFile
    PutINISetting location, "OptionsTabIndex", rDOptionsTabIndex, settingsFile
    PutINISetting location & "\WindowFilters", "Count", 0, settingsFile
    


   On Error GoTo 0
   Exit Sub

writeDockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeDockSettings of Form dockSettings"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : adjustControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustControls()

        
    ' variables declared
    Dim rgbRdFontShadowColor As String
    Dim rgbRdFontOutlineColor As String

    Dim suppliedFontSize As Integer
    Dim suppliedWeight As Integer
    Dim suppliedBold As Boolean
    Dim suppliedItalics As Boolean
    Dim suppliedUnderline As Boolean
    
    'Dim MyFile As String
    Dim MyPath  As String
    Dim themePresent As Boolean
    Dim myName As String
    
    'initialise the dimensioned variables
    rgbRdFontShadowColor = ""
    rgbRdFontOutlineColor = ""
    suppliedFontSize = 0
    suppliedWeight = 0
    suppliedBold = False
    suppliedItalics = False
    suppliedUnderline = False

    
    On Error GoTo adjustControls_Error
    If debugflg = 1 Then Debug.Print "%adjustControls"
    
    
    MyPath = dockAppPath & "\Skins\" '"E:\Program Files (x86)\RocketDock\Skins\"
    themePresent = False

    If Not DirExists(MyPath) Then
        MsgBox "WARNING - The skins folder is not present in the correct location " & rdAppPath
    End If
    
    myName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(MyPath & myName) And vbDirectory) = vbDirectory Then
             'Debug.Print MyName   ' Display entry only if it
          End If   ' it represents a directory.
       End If
       myName = Dir   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> "" Then
        cmbStyleTheme.AddItem myName
        'MsgBox MyName
        Debug.Print myName   ' Display entry only if it
        If myName = rDtheme Then themePresent = True
       End If
    Loop

    ' if the theme is not in the list then make it none to ensure no corruption *1
    If themePresent = False Then rDtheme = "blank"

    If rDtheme = "Program Files" Or rDtheme = "" Then
        cmbStyleTheme.Text = "blank"
    Else
        cmbStyleTheme.Text = rDtheme
    End If

    'optGeneralReadConfig.Value = CBool(LCase(rDGeneralReadConfig))
      If rDGeneralReadConfig = "True" Then
          optGeneralReadConfig.Value = True
      Else
          optGeneralReadConfig.Value = False
      End If
      

    'optGeneralWriteConfig.Value = CBool(LCase(rDGeneralWriteConfig))

      If rDGeneralWriteConfig = "True" Then
          optGeneralWriteConfig.Value = True
      Else
          optGeneralWriteConfig.Value = False
      End If


    ' controls for values that do not appear in Rocketdock
    If defaultDock = 1 Then
        sliGenRunAppInterval.Value = Val(rDRunAppInterval)
        chkGenAlwaysAsk.Value = Val(rDAlwaysAsk)
    End If

    'Rocketdock values also used by Steamydock
    
    chkGenLock.Value = Val(rDLockIcons)
    chkGenOpen.Value = Val(rDOpenRunning)
    chkGenRun.Value = Val(rDShowRunning)
    chkGenMin.Value = Val(rDManageWindows)
    chkGenDisableAnim.Value = Val(rDDisableMinAnimation)

    If chkGenMin.Value = 0 Then
        chkGenDisableAnim.Enabled = False
        lblChkDisable.Enabled = False
    Else
        chkGenDisableAnim.Enabled = True
        lblChkDisable.Enabled = True
    End If
    
    If chkGenRun.Value = 0 Then
        lblGenRunAppInterval1.Enabled = False
        lblGenRunAppInterval2.Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenRunAppInterval3.Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else
        lblGenRunAppInterval1.Enabled = True
        lblGenRunAppInterval2.Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenRunAppInterval3.Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
    ' Icons tab
    
    Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected

    Call setBounceTypes
    
    cmbIconsQuality.ListIndex = Val(rDIconQuality)
    sliIconsOpacity.Value = Val(rDIconOpacity)
    chkIconsZoomOpaque.Value = Val(rDZoomOpaque)
    sliIconsSize.Value = Val(rDIconMin)
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)
    
    sliIconsZoom.Value = Val(rdIconMax)
    
    'Call setMinimumHoverFX     ' .12 DAEB 28/04/2021 docksettings Set the large icon minimum size to 85 pixels when using the bumpy animation

    sliIconsZoomWidth.Value = Val(rDZoomWidth)
    sliIconsDuration.Value = Val(rDZoomTicks)
    
    ' position

    cmbPositionMonitor.ListIndex = Val(rDMonitor)
    cmbPositionScreen.ListIndex = Val(rDSide)
    cmbPositionLayering.ListIndex = Val(rDzOrderMode)
    sliPositionCentre.Value = Val(rDOffset)
    sliPositionEdgeOffset.Value = Val(rDvOffset)
    
    'style panel
    
    sliStyleOpacity.Value = Val(rDThemeOpacity)
    chkStyleDisable.Value = Val(rDHideLabels)
    
    chkLabelBackgrounds.Value = Val(sDShowLblBacks) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files

    lblStyleFontName.Caption = "Font: " & rDFontName & ", size: " & Val(Abs(rDFontSize)) & "pt"

    'the colour data that comes from the registry is RGB decimal
    rgbRdFontShadowColor = Convert_Dec2RGB(rDFontShadowColor)
    lblStyleFontFontShadowColor.Caption = "Shadow Colour: " & rgbRdFontShadowColor
    lblStyleFontOutlineTest.ForeColor = rDFontShadowColor

    rgbRdFontOutlineColor = Convert_Dec2RGB(rDFontOutlineColor)
    lblStyleOutlineColourDesc.Caption = "Outline Colour: " & rgbRdFontOutlineColor
    lblStyleFontOutlineTest.ForeColor = rDFontOutlineColor

    lblPreviewFont.ForeColor = rDFontColor

    sliStyleFontOpacity.Value = Val(sDFontOpacity)
    sliStyleOutlineOpacity.Value = Val(rDFontOutlineOpacity)
    sliStyleShadowOpacity.Value = Val(rDFontShadowOpacity)
    
    Call preFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)

    Call displayFontInformation(suppliedFontSize, suppliedBold, suppliedItalics, suppliedUnderline, suppliedWeight)

    ' behaviour
    
    chkBehaviourAutoHide.Value = Val(rDAutoHide)
' 226/10/2020 docksettings .05 DAEB  added a manual click to the autohide toggle checkbox
' a checkbox value assignment does not trigger a checkbox click for this checkbox (in a frame) as normally occurs and there is no equivalent 'change event' for a checkbox
' so to force it to trigger we need a call to the click event
    Call chkBehaviourAutoHide_Click
    
    sliBehaviourAutoHideDuration.Value = Val(rDAutoHideTicks)
    
    sliContinuousHide.Value = Val(sDContinuousHide) ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    
    'sDBounceZone
    
    sliAnimationInterval.Value = Val(rDAnimationInterval)
    sliStyleThemeSize.Value = Val(rDSkinSize)
    chkSplashStatus.Value = Val(sDSplashStatus)
    
    genChkShowIconSettings.Value = Val(sDShowIconSettings) ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock

    
    sliBehaviourAutoHideDelay.Value = Val(rDAutoHideDelay)
    
    cmbBehaviourAutoHideType.ListIndex = Val(sDAutoHideType)

    chkBehaviourMouseActivate.Value = Val(rDMouseActivate)
    sliBehaviourPopUpDelay.Value = Val(rDPopupDelay)
    
    ' if defaultdock = rocketdock then
    ' add the default key combo for RD
    ' if not then add the key combinations that are allowed for Steamydock
    
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
    If defaultDock = 1 Then
        cmbHidingKey.Clear
        cmbHidingKey.AddItem "F1"
        cmbHidingKey.AddItem "F2"
        cmbHidingKey.AddItem "F3"
        cmbHidingKey.AddItem "F4"
        cmbHidingKey.AddItem "F5"
        cmbHidingKey.AddItem "F6"
        cmbHidingKey.AddItem "F7"
        cmbHidingKey.AddItem "F8"
        cmbHidingKey.AddItem "F9"
        cmbHidingKey.AddItem "F10"
        cmbHidingKey.AddItem "F11"
        cmbHidingKey.AddItem "F12"
        cmbHidingKey.AddItem "Disabled"
        cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    Else
        cmbHidingKey.Clear
        cmbHidingKey.AddItem "Control+Alt+R"
        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    End If
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD ends
    
    ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
    If defaultDock = 1 Then
        picThemeSample.Enabled = True
        lblThemeSizeText.Enabled = True
        lblThemeSizeTextLow.Enabled = True
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
    Else
        picThemeSample.Enabled = False
        lblThemeSizeText.Enabled = False
        lblThemeSizeTextLow.Enabled = False
        sliStyleThemeSize.Enabled = False
        lblThemeSizeTextHigh.Enabled = False
        lblStyleSizeCurrent.Enabled = False
    End If
    
    ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
    
    
   On Error GoTo 0
   Exit Sub

adjustControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustControls of Form dockSettings on line " & Erl

End Sub

Private Sub setBounceTypes()

'None
'UberIcon Effects
'Bounce

cmbBehaviourActivationFX.Clear

    If defaultDock = 0 Then
        cmbBehaviourActivationFX.AddItem "None", 0
        cmbBehaviourActivationFX.AddItem "UberIcon Effects", 1
        cmbBehaviourActivationFX.AddItem "Bounce", 2
        'rDIconActivationFX = "2"
    
    Else
        cmbBehaviourActivationFX.AddItem "None", 0
        cmbBehaviourActivationFX.AddItem "Bounce", 1
        cmbBehaviourActivationFX.AddItem "Miserable", 2
        'rDIconActivationFX = "1"
    End If
    
    cmbBehaviourActivationFX.ListIndex = Val(rDIconActivationFX)
    

End Sub


' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
'---------------------------------------------------------------------------------------
' Procedure : setZoomTypes
' Author    : beededea
' Date      : 29/04/2021
' Purpose   : Set the default zoom types available to the type of dock selected
'---------------------------------------------------------------------------------------
'
Private Sub setZoomTypes()

    On Error GoTo setZoomTypes_Error
    
    cmbIconsHoverFX.Clear

    If defaultDock = 0 Then
        cmbIconsHoverFX.AddItem "None", 0
        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
        rDHoverFX = "1"
    
    Else
        cmbIconsHoverFX.AddItem "None", 0
        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
        cmbIconsHoverFX.AddItem "Zoom: Bumpy", 4
        rDHoverFX = "4"
    End If
    
    cmbIconsHoverFX.ListIndex = Val(rDHoverFX)


    On Error GoTo 0
    Exit Sub

setZoomTypes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setZoomTypes of Form dockSettings"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetMonitorCount
' Author    : beededea
' Date      : 02/03/2020
' Purpose   : populate the monitor dropdown according to the number of monitors available
'---------------------------------------------------------------------------------------
'
Private Sub GetMonitorCount()
    
    ' variables declared
   Dim numberOfMonitors As Integer
   Dim useloop As Integer
    
   'initialise the dimensioned variables
   numberOfMonitors = 1
   useloop = 1
    
   On Error GoTo GetMonitorCount_Error
   If debugflg = 1 Then DebugPrint "%GetMonitorCount"

   numberOfMonitors = GetSystemMetrics(SM_CMONITORS)
   
   If numberOfMonitors <= 1 Then
        cmbPositionMonitor.Clear
        cmbPositionMonitor.AddItem "Monitor 1"
        cmbPositionMonitor.ListIndex = 0
        cmbPositionMonitor.Enabled = False
    Else
        'clear and populate the monitor list
        cmbPositionMonitor.Clear
        For useloop = 1 To numberOfMonitors
            cmbPositionMonitor.AddItem "Monitor " & useloop
        Next useloop
        cmbPositionMonitor.ListIndex = rDMonitor
   End If
   lblPositionMonitor.ToolTipText = "This computer has this many screens - " & numberOfMonitors
   cmbPositionMonitor.ToolTipText = "This computer has this many screens - " & numberOfMonitors

   On Error GoTo 0
   Exit Sub

GetMonitorCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetMonitorCount of Form dockSettings"

End Sub


