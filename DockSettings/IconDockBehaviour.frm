VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form IconDockBehaviour 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Here you can control the behaviour of the animation effects"
      Top             =   0
      Visible         =   0   'False
      Width           =   6930
      Begin VB.Frame fraIconEffect 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   105
         TabIndex        =   39
         Top             =   945
         Width           =   5025
      End
      Begin VB.Frame fraAnimationInterval 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   555
         TabIndex        =   33
         Top             =   6570
         Width           =   6180
         Begin CCRSlider.Slider sliAnimationInterval 
            Height          =   315
            Left            =   1575
            TabIndex        =   34
            ToolTipText     =   $"IconDockBehaviour.frx":0000
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
            TabIndex        =   38
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   0
            Width           =   1605
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
            TabIndex        =   37
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   525
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
            TabIndex        =   36
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   585
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
            TabIndex        =   35
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   630
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
         TabIndex        =   32
         ToolTipText     =   "Essential functionality for the dock - pops up when  given focus"
         Top             =   8070
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame fraAutoHideDelay 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   555
         TabIndex        =   26
         Top             =   3375
         Width           =   6120
         Begin CCRSlider.Slider sliBehaviourAutoHideDelay 
            Height          =   315
            Left            =   1500
            TabIndex        =   27
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Max             =   2000
            TickFrequency   =   200
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
            TabIndex        =   31
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   -30
            Width           =   1350
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
            TabIndex        =   30
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   1185
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
            TabIndex        =   29
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   405
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
            TabIndex        =   28
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   600
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   555
         TabIndex        =   20
         Top             =   2565
         Width           =   5805
         Begin CCRSlider.Slider sliBehaviourPopUpDelay 
            Height          =   315
            Left            =   1500
            TabIndex        =   21
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
            TabIndex        =   25
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   585
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
            TabIndex        =   24
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   480
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
            TabIndex        =   23
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   0
            Width           =   1965
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
            TabIndex        =   22
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   375
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin VB.Frame fraAutoHideDuration 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   555
         TabIndex        =   14
         Top             =   1800
         Width           =   6180
         Begin CCRSlider.Slider sliBehaviourAutoHideDuration 
            Height          =   315
            Left            =   1515
            TabIndex        =   15
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
            TabIndex        =   19
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   0
            Width           =   1605
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
            TabIndex        =   18
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   525
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
            TabIndex        =   17
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   585
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
            TabIndex        =   16
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   630
         End
      End
      Begin VB.Frame fraAutoHideType 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   615
         TabIndex        =   8
         Top             =   465
         Width           =   5100
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
            ItemData        =   "IconDockBehaviour.frx":008F
            Left            =   1590
            List            =   "IconDockBehaviour.frx":009C
            TabIndex        =   11
            Text            =   "Bounce"
            Top             =   0
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
            TabIndex        =   10
            ToolTipText     =   "You can determine whether the dock will auto-hide or not"
            Top             =   480
            Width           =   2235
         End
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
            ItemData        =   "IconDockBehaviour.frx":00C0
            Left            =   1590
            List            =   "IconDockBehaviour.frx":00CD
            TabIndex        =   9
            Text            =   "Fade"
            ToolTipText     =   "The type of auto-hide, fade, instant or a slide like Rocketdock"
            Top             =   885
            Width           =   2620
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
            TabIndex        =   13
            ToolTipText     =   $"IconDockBehaviour.frx":00E7
            Top             =   45
            Width           =   1605
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
            TabIndex        =   12
            ToolTipText     =   "You can determine whether the dock will auto-hide or not"
            Top             =   495
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   555
         TabIndex        =   2
         Top             =   4050
         Width           =   6120
         Begin CCRSlider.Slider sliContinuousHide 
            Height          =   315
            Left            =   1500
            TabIndex        =   3
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
            TabIndex        =   7
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   405
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
            TabIndex        =   6
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   1185
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
            TabIndex        =   5
            ToolTipText     =   "Determine how long Steamydock will disappear when told to hide for the next few minutes"
            Top             =   -30
            Width           =   1350
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
            TabIndex        =   4
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   600
         End
      End
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
         ItemData        =   "IconDockBehaviour.frx":0179
         Left            =   2190
         List            =   "IconDockBehaviour.frx":01A4
         TabIndex        =   1
         Text            =   "F11"
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4965
         Width           =   2620
      End
      Begin VB.Label lblAnimationInformationLabel 
         Caption         =   $"IconDockBehaviour.frx":01E5
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
         TabIndex        =   41
         Top             =   7530
         Width           =   4485
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
         TabIndex        =   40
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4995
         Width           =   1440
      End
   End
End
Attribute VB_Name = "IconDockBehaviour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
