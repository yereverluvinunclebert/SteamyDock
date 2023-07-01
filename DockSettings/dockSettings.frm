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
      TabIndex        =   64
      ToolTipText     =   "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
      Top             =   45
      Width           =   6930
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   6825
         TabIndex        =   236
         Top             =   2175
         Width           =   75
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   6570
         TabIndex        =   235
         Top             =   2175
         Width           =   330
      End
      Begin VB.TextBox lblAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6390
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   234
         Text            =   "dockSettings.frx":058A
         Top             =   2235
         Width           =   6660
      End
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
         TabIndex        =   110
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   222
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
         TabIndex        =   221
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
         TabIndex        =   220
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblPunklabsLink 
         BackStyle       =   0  'Transparent
         Caption         =   "                                                                                                                        "
         Height          =   225
         Index           =   0
         Left            =   2175
         MousePointer    =   1  'Arrow
         TabIndex        =   106
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   510
         Width           =   495
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
         Top             =   495
         Width           =   795
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
         TabIndex        =   69
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
         TabIndex        =   68
         Top             =   855
         Width           =   2175
      End
   End
   Begin VB.Timer positionTimer 
      Interval        =   3000
      Left            =   1455
      Top             =   8790
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
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Revert ALL settings to the defaults"
      Top             =   8790
      Width           =   1065
   End
   Begin VB.CheckBox chkToggleDialogs 
      Caption         =   "Display Info.Dialogs"
      Height          =   225
      Left            =   135
      TabIndex        =   225
      ToolTipText     =   "When checked this toggle will display the information pop-ups and balloon tips "
      Top             =   8880
      Value           =   1  'Checked
      Width           =   1860
   End
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
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   162
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
      Picture         =   "dockSettings.frx":0D02
      ScaleHeight     =   795
      ScaleWidth      =   825
      TabIndex        =   161
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
         TabIndex        =   19
         Top             =   6930
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   5
            Left            =   -60
            Picture         =   "dockSettings.frx":177D
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   20
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
            TabIndex        =   21
            Top             =   855
            Width           =   570
         End
      End
      Begin VB.Frame fmeIconGeneral 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   45
         TabIndex        =   18
         Top             =   -30
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   0
            Left            =   0
            Picture         =   "dockSettings.frx":22E2
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   27
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
            TabIndex        =   24
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
               TabIndex        =   25
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
         TabIndex        =   16
         Top             =   5430
         Width           =   915
         Begin VB.PictureBox picIcon 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   4
            Left            =   0
            Picture         =   "dockSettings.frx":2DA8
            ScaleHeight     =   1020
            ScaleWidth      =   960
            TabIndex        =   17
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
               TabIndex        =   22
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
            Picture         =   "dockSettings.frx":3CA2
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
            TabIndex        =   23
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
            TabIndex        =   30
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
               TabIndex        =   31
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
            Picture         =   "dockSettings.frx":48A3
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
            TabIndex        =   28
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
               TabIndex        =   29
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
            Picture         =   "dockSettings.frx":53A9
            ScaleHeight     =   975
            ScaleWidth      =   960
            TabIndex        =   26
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
      Left            =   6525
      ScaleHeight     =   1605
      ScaleWidth      =   1650
      TabIndex        =   86
      ToolTipText     =   "The icon size in the dock"
      Top             =   240
      Visible         =   0   'False
      Width           =   1650
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
      Left            =   1215
      TabIndex        =   1
      ToolTipText     =   "These are the main settings for the dock"
      Top             =   15
      Width           =   6930
      Begin VB.CheckBox genChkShowIconSettings 
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
         Left            =   945
         TabIndex        =   224
         Top             =   7905
         Width           =   5115
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
         TabIndex        =   179
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   7590
         Width           =   3735
      End
      Begin VB.Frame fraRunAppIndicators 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   450
         TabIndex        =   138
         Top             =   4635
         Width           =   5955
         Begin CCRSlider.Slider sliGenRunAppInterval 
            Height          =   315
            Left            =   1020
            TabIndex        =   139
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label lblGenLabel 
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
            Index           =   2
            Left            =   750
            LinkItem        =   "150"
            TabIndex        =   140
            ToolTipText     =   "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
            Top             =   105
            Width           =   3210
         End
      End
      Begin VB.CommandButton btnGeneralRdFolder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   5745
         TabIndex        =   109
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
         ToolTipText     =   $"dockSettings.frx":6234
         Top             =   4350
         Width           =   2985
      End
      Begin VB.CheckBox chkGenDisableAnim 
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
         Height          =   360
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "If you dislike the minimise animation, click this"
         Top             =   3915
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.CheckBox chkGenOpen 
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
         Height          =   360
         Left            =   930
         TabIndex        =   15
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   5520
         Width           =   3030
      End
      Begin VB.TextBox txtGeneralRdLocation 
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
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "C:\programs"
         ToolTipText     =   "This is the extrapolated location of the currently selected dock. This is for information only."
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
         ItemData        =   "dockSettings.frx":62D3
         Left            =   2085
         List            =   "dockSettings.frx":62DD
         TabIndex        =   76
         Text            =   "Rocketdock"
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
         Top             =   6255
         Width           =   2310
      End
      Begin VB.CheckBox chkGenMin 
         Caption         =   "Minimise Windows to the Dock"
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
         Width           =   3075
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
         Top             =   495
         Width           =   1440
      End
      Begin VB.Frame fraWriteOptionButtons 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   510
         TabIndex        =   163
         Top             =   2475
         Width           =   6165
         Begin VB.OptionButton optGeneralWriteConfig 
            Caption         =   "Write Settings to SteamyDock's Own Configuration Area"
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
            Left            =   390
            TabIndex        =   164
            ToolTipText     =   $"dockSettings.frx":62F9
            Top             =   30
            Width           =   5325
         End
      End
      Begin VB.Frame fraReadOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   540
         TabIndex        =   144
         Top             =   1035
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
            TabIndex        =   147
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
            TabIndex        =   146
            ToolTipText     =   $"dockSettings.frx":638E
            Top             =   465
            Width           =   5500
         End
         Begin VB.OptionButton optGeneralReadConfig 
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
            Height          =   225
            Left            =   360
            TabIndex        =   145
            ToolTipText     =   $"dockSettings.frx":6451
            Top             =   780
            Width           =   5565
         End
         Begin VB.Label lbloptGeneralReadConfig 
            Caption         =   "Read Settings From SteamyDock's Own Configuration Area (modern)"
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
            Left            =   630
            TabIndex        =   226
            Top             =   780
            Width           =   5115
         End
      End
      Begin VB.Label lblSplitter 
         Caption         =   "-o0o-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2640
         TabIndex        =   239
         Top             =   3075
         Width           =   1185
      End
      Begin VB.Label lblChkGenDisableAnim 
         Caption         =   "Disable Minimise Animations"
         Height          =   255
         Left            =   1455
         TabIndex        =   228
         Top             =   3975
         Width           =   2220
      End
      Begin VB.Label lblChkGenMin 
         Caption         =   "Minimise Windows to the Dock"
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
         Left            =   1185
         TabIndex        =   227
         Top             =   3645
         Width           =   2610
      End
      Begin VB.Label lblGenLabel 
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
         Index           =   1
         Left            =   915
         TabIndex        =   166
         Top             =   2175
         Width           =   1800
      End
      Begin VB.Label lblGenLabel 
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
         Index           =   0
         Left            =   915
         TabIndex        =   165
         Top             =   855
         Width           =   1800
      End
      Begin VB.Label lblGenLabel 
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
         Index           =   4
         Left            =   915
         TabIndex        =   107
         ToolTipText     =   $"dockSettings.frx":64E6
         Top             =   6690
         Width           =   1695
      End
      Begin VB.Label lblGenLabel 
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
         Index           =   3
         Left            =   915
         TabIndex        =   77
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
         Top             =   6300
         Width           =   1530
      End
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
      TabIndex        =   63
      ToolTipText     =   "Here you can control the behaviour of the animation effects"
      Top             =   30
      Width           =   6930
      Begin VB.ComboBox cmbBehaviourSoundSelection 
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
         ItemData        =   "dockSettings.frx":65B0
         Left            =   2190
         List            =   "dockSettings.frx":65BD
         TabIndex        =   238
         Text            =   "None"
         Top             =   6150
         Width           =   2620
      End
      Begin VB.CheckBox chkRetainIcons 
         Caption         =   "Retain Original Icons when dragging to the dock"
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
         Left            =   2190
         TabIndex        =   232
         Top             =   5610
         Width           =   4455
      End
      Begin VB.CheckBox chkGenLock 
         Caption         =   "Disable Drag/Drop and Icon Deletion"
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
         Left            =   2190
         TabIndex        =   230
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
         Top             =   5130
         Width           =   4500
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
         ItemData        =   "dockSettings.frx":65E1
         Left            =   2190
         List            =   "dockSettings.frx":660C
         TabIndex        =   219
         Text            =   "F11"
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4515
         Width           =   2620
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   465
         TabIndex        =   212
         Top             =   3675
         Width           =   6120
         Begin CCRSlider.Slider sliContinuousHide 
            Height          =   315
            Left            =   1575
            TabIndex        =   213
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
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   11
            Left            =   1170
            TabIndex        =   214
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   600
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   5
            Left            =   45
            LinkItem        =   "150"
            TabIndex        =   217
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
            TabIndex        =   216
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
            TabIndex        =   215
            ToolTipText     =   "Determine how long Steamydock will disappear when told to go away"
            Top             =   285
            Width           =   405
         End
      End
      Begin VB.Frame fraAutoHideType 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   375
         TabIndex        =   207
         Top             =   465
         Width           =   5325
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
            ItemData        =   "dockSettings.frx":664D
            Left            =   1770
            List            =   "dockSettings.frx":665A
            TabIndex        =   211
            Text            =   "Fade"
            ToolTipText     =   "The type of auto-hide, fade, instant or a slide like Rocketdock"
            Top             =   510
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
            Left            =   90
            TabIndex        =   210
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
            ItemData        =   "dockSettings.frx":6674
            Left            =   1770
            List            =   "dockSettings.frx":6681
            TabIndex        =   208
            Text            =   "Bounce"
            ToolTipText     =   $"dockSettings.frx":66A5
            Top             =   0
            Width           =   2620
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   0
            Left            =   90
            LinkItem        =   "150"
            TabIndex        =   209
            ToolTipText     =   $"dockSettings.frx":6739
            Top             =   45
            Width           =   1605
         End
      End
      Begin VB.Frame fraAutoHideDuration 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   450
         TabIndex        =   201
         Top             =   1500
         Width           =   6180
         Begin CCRSlider.Slider sliBehaviourAutoHideDuration 
            Height          =   315
            Left            =   1590
            TabIndex        =   202
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
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   8
            Left            =   1140
            TabIndex        =   206
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
            Left            =   4425
            TabIndex        =   205
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
            Left            =   5085
            TabIndex        =   204
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   2
            Left            =   45
            LinkItem        =   "150"
            TabIndex        =   203
            ToolTipText     =   "The speed at which the dock auto-hide animation will occur"
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   420
         TabIndex        =   195
         Top             =   2175
         Width           =   5805
         Begin CCRSlider.Slider sliBehaviourPopUpDelay 
            Height          =   315
            Left            =   1620
            TabIndex        =   196
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
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   9
            Left            =   1185
            TabIndex        =   200
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   3
            Left            =   90
            LinkItem        =   "150"
            TabIndex        =   199
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
            Left            =   5100
            TabIndex        =   198
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
            Left            =   4455
            TabIndex        =   197
            ToolTipText     =   "The dock mouse-over delay period"
            Top             =   345
            Width           =   585
         End
      End
      Begin VB.Frame fraAutoHideDelay 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   435
         TabIndex        =   189
         Top             =   2970
         Width           =   6120
         Begin CCRSlider.Slider sliBehaviourAutoHideDelay 
            Height          =   315
            Left            =   1605
            TabIndex        =   190
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Enabled         =   0   'False
            Max             =   2000
            TickFrequency   =   200
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   10
            Left            =   1245
            TabIndex        =   194
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
            TabIndex        =   193
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
            TabIndex        =   192
            ToolTipText     =   "Determine the delay between the last usage of the dock and when it will auto-hide"
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   4
            Left            =   105
            LinkItem        =   "150"
            TabIndex        =   191
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
         TabIndex        =   188
         ToolTipText     =   "Essential functionality for the dock - pops up when  given focus"
         Top             =   8070
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame fraAnimationInterval 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   195
         TabIndex        =   167
         Top             =   6930
         Width           =   6180
         Begin CCRSlider.Slider sliAnimationInterval 
            Height          =   315
            Left            =   1890
            TabIndex        =   168
            ToolTipText     =   $"dockSettings.frx":67CB
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
            Left            =   1500
            TabIndex        =   172
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
            Left            =   4680
            TabIndex        =   171
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
            Left            =   5265
            TabIndex        =   170
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   315
            Width           =   525
         End
         Begin VB.Label lblBehaviourLabel 
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
            Index           =   7
            Left            =   345
            LinkItem        =   "150"
            TabIndex        =   169
            ToolTipText     =   "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
            Top             =   15
            Width           =   1605
         End
      End
      Begin VB.Frame fraIconEffect 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   105
         TabIndex        =   111
         Top             =   945
         Width           =   5025
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Sound Selection"
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
         Index           =   15
         Left            =   540
         TabIndex        =   237
         Top             =   6195
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Icon Origin"
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
         Index           =   14
         Left            =   540
         TabIndex        =   233
         ToolTipText     =   "The original icons may be low quality."
         Top             =   5670
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   "Lock the Dock"
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
         Index           =   13
         Left            =   540
         TabIndex        =   231
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
         Top             =   5190
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
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
         Index           =   6
         Left            =   525
         LinkItem        =   "150"
         TabIndex        =   218
         ToolTipText     =   "This is the key sequence that is used to hide or restore Steamydock"
         Top             =   4545
         Width           =   1440
      End
      Begin VB.Label lblBehaviourLabel 
         Caption         =   $"dockSettings.frx":685A
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
         Index           =   12
         Left            =   1740
         TabIndex        =   185
         Top             =   7755
         Width           =   4485
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
      TabIndex        =   49
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
         TabIndex        =   186
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4065
         Width           =   2490
      End
      Begin VB.PictureBox picThemeSample 
         Height          =   2070
         Left            =   630
         Picture         =   "dockSettings.frx":68EC
         ScaleHeight     =   2010
         ScaleWidth      =   5265
         TabIndex        =   173
         ToolTipText     =   "An example preview of the chosen theme."
         Top             =   1830
         Width           =   5325
      End
      Begin VB.Frame fraFontOpacity 
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   210
         TabIndex        =   112
         ToolTipText     =   "The theme background "
         Top             =   6750
         Width           =   6525
         Begin CCRSlider.Slider sliStyleShadowOpacity 
            Height          =   330
            Left            =   1875
            TabIndex        =   113
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
            TabIndex        =   114
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
            TabIndex        =   180
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   240
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   582
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   8
            Left            =   1635
            TabIndex        =   184
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
            TabIndex        =   183
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
            TabIndex        =   182
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   255
            Width           =   630
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   3
            Left            =   480
            TabIndex        =   181
            ToolTipText     =   "The font transparency can be changed here"
            Top             =   -15
            Width           =   1350
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   5
            Left            =   450
            TabIndex        =   122
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
            TabIndex        =   121
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
            TabIndex        =   120
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   585
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   10
            Left            =   1635
            TabIndex        =   119
            ToolTipText     =   "The label outline transparency, use the slider to change"
            Top             =   1290
            Width           =   630
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   4
            Left            =   465
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
            ToolTipText     =   "The strength of the shadow can be altered here"
            Top             =   780
            Width           =   555
         End
         Begin VB.Label lblStyleLabel 
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
            Index           =   9
            Left            =   1635
            TabIndex        =   115
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
         TabIndex        =   61
         ToolTipText     =   $"dockSettings.frx":A145
         Top             =   4440
         Width           =   5340
         Begin VB.Label lblPreviewFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2355
            TabIndex        =   62
            Top             =   255
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2400
            TabIndex        =   155
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
            TabIndex        =   156
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
            TabIndex        =   157
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
            TabIndex        =   158
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
            TabIndex        =   159
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblPreviewFontShadow2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   195
            Left            =   2415
            TabIndex        =   160
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         ItemData        =   "dockSettings.frx":A1CF
         Left            =   2205
         List            =   "dockSettings.frx":A1D1
         TabIndex        =   50
         ToolTipText     =   "The dock background theme can be selected here"
         Top             =   405
         Width           =   2520
      End
      Begin CCRSlider.Slider sliStyleOpacity 
         Height          =   315
         Left            =   2085
         TabIndex        =   52
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
         TabIndex        =   174
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
         TabIndex        =   187
         ToolTipText     =   "You can toggle the icon label background on/off here"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblStyleLabel 
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
         Index           =   7
         Left            =   1650
         TabIndex        =   175
         Top             =   1380
         Width           =   495
      End
      Begin VB.Label lblStyleLabel 
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
         Index           =   2
         Left            =   660
         TabIndex        =   178
         ToolTipText     =   "The theme background overall size is set here"
         Top             =   1380
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
         TabIndex        =   177
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
         TabIndex        =   176
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label Label999 
         Height          =   375
         Left            =   720
         TabIndex        =   154
         Top             =   7560
         Width           =   4215
      End
      Begin VB.Label lblStyleLabel 
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
         Index           =   6
         Left            =   1815
         TabIndex        =   56
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   945
         Width           =   630
      End
      Begin VB.Label lblStyleLabel 
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
         Index           =   1
         Left            =   675
         TabIndex        =   53
         ToolTipText     =   "The theme background opacity is set here"
         Top             =   945
         Width           =   1050
      End
      Begin VB.Label lblStyleLabel 
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
         Index           =   0
         Left            =   675
         TabIndex        =   51
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
      TabIndex        =   32
      ToolTipText     =   "This panel controls the positioning of the whole dock"
      Top             =   15
      Width           =   6930
      Begin VB.PictureBox picMultipleGears1 
         BorderStyle     =   0  'None
         Height          =   4800
         Left            =   150
         Picture         =   "dockSettings.frx":A1D3
         ScaleHeight     =   4800
         ScaleWidth      =   3495
         TabIndex        =   108
         Top             =   3705
         Width           =   3500
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
         ItemData        =   "dockSettings.frx":11D9E
         Left            =   2190
         List            =   "dockSettings.frx":11DAB
         TabIndex        =   47
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
         ItemData        =   "dockSettings.frx":11DD4
         Left            =   2205
         List            =   "dockSettings.frx":11DEA
         TabIndex        =   36
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
         ItemData        =   "dockSettings.frx":11E30
         Left            =   2190
         List            =   "dockSettings.frx":11E40
         TabIndex        =   35
         Text            =   "Bottom"
         ToolTipText     =   "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
         Top             =   1185
         Width           =   2595
      End
      Begin CCRSlider.Slider sliPositionEdgeOffset 
         Height          =   315
         Left            =   2085
         TabIndex        =   33
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
         TabIndex        =   34
         ToolTipText     =   "You can align the dock so that it is centred or offset as you require"
         Top             =   2625
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Min             =   -100
         Max             =   100
         TickFrequency   =   10
      End
      Begin VB.PictureBox picMultipleGears3 
         BorderStyle     =   0  'None
         Height          =   2970
         Left            =   3645
         Picture         =   "dockSettings.frx":11E5E
         ScaleHeight     =   2970
         ScaleWidth      =   3015
         TabIndex        =   85
         Top             =   5400
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
         TabIndex        =   48
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
      TabIndex        =   87
      ToolTipText     =   "This panel allows you to set the icon sizes and hover effects"
      Top             =   15
      Width           =   6930
      Begin VB.PictureBox picSizePreview 
         Height          =   4065
         Left            =   105
         ScaleHeight     =   4005
         ScaleWidth      =   6645
         TabIndex        =   148
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            TabIndex        =   153
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
            TabIndex        =   152
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
            TabIndex        =   151
            ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
            Top             =   3810
            Width           =   1875
         End
      End
      Begin VB.Frame fraZoomConfigs 
         BorderStyle     =   0  'None
         Height          =   1110
         Left            =   195
         TabIndex        =   123
         Top             =   3165
         Width           =   6495
         Begin CCRSlider.Slider sliIconsDuration 
            Height          =   315
            Left            =   1845
            TabIndex        =   124
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
            TabIndex        =   125
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   195
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   556
            Min             =   2
            Value           =   2
            SelStart        =   2
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   11
            Left            =   1320
            TabIndex        =   133
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   525
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   12
            Left            =   4650
            TabIndex        =   132
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
            TabIndex        =   131
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   6
            Left            =   480
            TabIndex        =   130
            ToolTipText     =   "How long the effect is applied"
            Top             =   780
            Width           =   795
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   10
            Left            =   1665
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
            ToolTipText     =   "How many icons to the left and right are also animated"
            Top             =   225
            Width           =   630
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   5
            Left            =   465
            TabIndex        =   126
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
         TabIndex        =   90
         ToolTipText     =   "Should the zoom be opaque too?"
         Top             =   1320
         Width           =   2685
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
         ItemData        =   "dockSettings.frx":15BDA
         Left            =   2160
         List            =   "dockSettings.frx":15BE7
         TabIndex        =   89
         Text            =   "Low quality (Faster)"
         ToolTipText     =   $"dockSettings.frx":15C29
         Top             =   390
         Width           =   2520
      End
      Begin CCRSlider.Slider sliIconsZoom 
         Height          =   315
         Left            =   2040
         TabIndex        =   88
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
         TabIndex        =   91
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
         TabIndex        =   92
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
         TabIndex        =   134
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
            ItemData        =   "dockSettings.frx":15CE7
            Left            =   1995
            List            =   "dockSettings.frx":15CFA
            TabIndex        =   135
            Text            =   "None"
            ToolTipText     =   "The zoom effect to apply"
            Top             =   105
            Width           =   2595
         End
         Begin VB.Label lblCharacteristicsLabel 
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
            Index           =   2
            Left            =   480
            TabIndex        =   136
            ToolTipText     =   "The zoom effect to apply"
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.Label lblchkIconsZoomOpaque 
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
         Height          =   315
         Left            =   2415
         TabIndex        =   229
         Top             =   1305
         Width           =   2820
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
         TabIndex        =   137
         Top             =   3585
         Width           =   5325
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   0
         Left            =   630
         TabIndex        =   105
         ToolTipText     =   "Lower power machines will benefit from the lower quality setting"
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   1
         Left            =   630
         TabIndex        =   104
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   795
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   3
         Left            =   630
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
         ToolTipText     =   "The icons in the dock can be made transparent here"
         Top             =   915
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   7
         Left            =   1710
         TabIndex        =   99
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
         TabIndex        =   98
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   8
         Left            =   1635
         TabIndex        =   97
         ToolTipText     =   "The size of all the icons in the dock before any effect is applied"
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   4
         Left            =   645
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   585
      End
      Begin VB.Label lblCharacteristicsLabel 
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
         Index           =   9
         Left            =   1755
         TabIndex        =   93
         ToolTipText     =   "The maximum icon size after a zoom"
         Top             =   2820
         Width           =   630
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
      TabIndex        =   223
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
      Begin VB.Menu mnuBringToCentre 
         Caption         =   "Centre Program on Main Monitor"
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
' .16 DAEB 01/07/2022 docksettings added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
' .17 DAEB 07/09/2022 docksettings the dock folder location now changes as it is switched between Rocketdock and Steamy Dock
' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
' .19 DAEB 07/09/2022 docksettings when you select rocketdock it reverts to the registry but when you select steamydock it does not revert to the dock settings file.
' .20 DAEB 07/09/2022 docksettings tab selection fixed
' .21 DAEB 07/09/2022 docksettings moved hiding key definitions to own subroutine
' .22 DAEB 02/10/2022 docksettings added a message pop up on the punklabs link
' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
' add the Steampunk cogs for the light and dark themes
' take the X/Y position and store it, when restarting, set it as per FCW.
' menu option to move the utility to the centre of the main monitor
' for win 11 bottom cut off - need to add another 100 twips
' adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to monitor/resolution changes

' Status/Bugs/Tasks:
' ==================
'
' Define any key to toggle hiding not just function keys - at the moment it is much more sensible to have a single key defined
'   Using this code it can be done - https://www.developerfusion.com/code/271/create-a-hot-key/
'   but this will require subclassing within steamydock. All the solutions I have found require sub-classing.
'   Within the hotkey folder under vb6 there is code that will identify keypresses (dockSettings) and will respond
'   via sub-classing (steamyDock).
'
' The drop-down lists do not support mouseOver events so the balloon tooltips will not work. They will have to be sub-classed
'   to allow the balloon tooltip to function.
'
' update the help files WIP
'
'   test running with a blank tool settings file
'
'   test running with a blank dock settings file
'
'   remove persistent debug and replace with logging to a file as per FCW.
'
'       Elroy's code to add balloon tips to comboBox
'       https://www.vbforums.com/showthread.php?893844-VB6-QUESTION-How-to-capture-the-MouseOver-Event-on-a-comboBox
'
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
' When the main checkbox/radio button is disabled, its width is reduced and the associated label is made visible.
' Note that the balloon tooltips only function on the controls and not on the labels.

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

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


' Flag for debug mode
Private mbDebugMode As Boolean  ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without

Public origSettingsFile As String






Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnAboutDebugInfo.hwnd, "This is the debugging option - Don't use it unless you know what you are doing. This option runs a separate binary, the persistentDebug.exe (an additional binary provided with this tool) is only run when you turn debugging ON. I suggest you do NOT use this utility unless you have a problem that is not easy diagnose. It is a separate exe that my program talks to, sending the program's subroutine entry points and other debug data to that window.When you run it the first time, your anti-malware tool such as malwarebytes will flag it as a possible malware. It is NOT. It only seems that way to anti-malware tools because of the way it operates, ie. one program is talking to another using shared memory.", _
                  TTIconInfo, "Help on the About Button", , , , True
End Sub

Private Sub btnApply_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnApply.hwnd, "Apply your recent changes to the settings and save them.", _
                  TTIconInfo, "Help on the Apply Button", , , , True
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnClose.hwnd, "Close the Dock Settings Utility.", _
                  TTIconInfo, "Help on the Close Button", , , , True
End Sub

Private Sub btnDefaults_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnDefaults.hwnd, "Revert ALL settings to the defaults.", _
                  TTIconInfo, "Help on the Set Defaults Button", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnDonate.hwnd, "Opens a browser window and sends you to the donation page on Amazon.", _
                  TTIconInfo, "Help on the Donate Button", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnFacebook.hwnd, "This will link you to the Rocket/SteamyDock users Group.", _
                  TTIconInfo, "Help on the FaceBook Button", , , , True
End Sub

Private Sub btnGeneralRdFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGeneralRdFolder.hwnd, "Press this button to select the folder location of Rocketdock here. ", _
                  TTIconInfo, "Help on selecting a folder.", , , , True

End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnHelp.hwnd, "This button open the tool's HTML help page in your browser.", _
                  TTIconInfo, "Help on the Help Button", , , , True
End Sub

Private Sub btnStyleFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleFont.hwnd, "This button gives the font selection box. Here you set the font as shown on the icon labels.", _
                  TTIconInfo, "Help on the Font Selection Button.", , , , True
End Sub

Private Sub btnStyleOutline_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleOutline.hwnd, "The colour of the outline, click the button to change.", _
                  TTIconInfo, "Help on the Outline Colour Selection Button.", , , , True
End Sub

Private Sub btnStyleShadow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnStyleShadow.hwnd, "The colour of the shadow, click the button to change.", _
                  TTIconInfo, "Help on the Shadow Colour Selection Button.", , , , True
End Sub

Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnFacebook.hwnd, "Here you can visit the update location where you can download new versions of the programs used by Rocketdock.", _
                  TTIconInfo, "Help on the Update Button", , , , True
End Sub

Private Sub chkBehaviourAutoHide_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkBehaviourAutoHide.hwnd, "This checkbox acts as a toggle. You can determine whether the dock will auto-hide or not and the type of hide that is implemented. using Rocketdock  only supports one type of hide and that is the slide type. Steamydock gives you an additional fade or an instant disappear. The latter is lighter on CPU usage whilst the former two are animated and require a little cpu during the transition.", _
                  TTIconInfo, "Help on the AutoHide Checkbox.", , , , True
End Sub
'
'Private Sub chkGenAlwaysAsk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenAlwaysAsk.hWnd, "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as set above. ", _
'                  TTIconInfo, "Help on Confirming which dock to use.", , , , True
'End Sub

Private Sub chkGenDisableAnim_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenDisableAnim.hwnd, "If you dislike the minimise animation, click this. ", _
                  TTIconInfo, "Help on disabling the minimise animation.", , , , True
End Sub

Private Sub chkGenLock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenLock.hwnd, "This is an essential option that stops you accidentally deleting your dock icons, click it!. ", _
                  TTIconInfo, "Help on Dragging, dropping to or from the dock.", , , , True
                  
End Sub

Private Sub chkGenMin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenMin.hwnd, "This option allows running applications to be minimised, appearing in the dock. Supported by Rocketdock only.", _
                  TTIconInfo, "Help on mimising apps to the dock.", , , , True
End Sub

Private Sub chkGenOpen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenOpen.hwnd, "If you click on an icon that is already running then it can open it or fire up another instance. ", _
                  TTIconInfo, "Help on the Running Application Indicators.", , , , True
End Sub

Private Sub chkGenRun_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenRun.hwnd, "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency set below. ", _
                  TTIconInfo, "Help on Showing Running Applications .", , , , True
End Sub

Private Sub chkGenWinStartup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkGenWinStartup.hwnd, "When this checkbox is ticked it will cause the selected dock to run when Windows starts. ", _
                  TTIconInfo, "Help on the Start with Windows Checkbox", , , , True
End Sub

Private Sub chkIconsZoomOpaque_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkIconsZoomOpaque.hwnd, "Should the zoomed icons be opaque when the others are transparent? Not yet implemented in Steamydock. ", _
                  TTIconInfo, "Help on the Zoom Opacity Checkbox", , , , True
End Sub

Private Sub chkLabelBackgrounds_Click()

   sDShowLblBacks = chkLabelBackgrounds.Value ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
    
    

    ' add a background to the icon titles in dock's drawtext function
End Sub

Private Sub chkLabelBackgrounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkLabelBackgrounds.hwnd, "With this checkbox you can toggle the icon label background on/off.", _
                  TTIconInfo, "Help on Label Background Disable.", , , , True
End Sub


' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
'---------------------------------------------------------------------------------------
' Procedure : chkRetainIcons_Click
' Author    : beededea
' Date      : 07/09/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkRetainIcons_Click()

    On Error GoTo chkRetainIcons_Click_Error

    rDRetainIcons = chkRetainIcons.Value

    On Error GoTo 0
    Exit Sub

chkRetainIcons_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRetainIcons_Click of Form dockSettings"
            Resume Next
          End If
    End With

End Sub

Private Sub chkRetainIcons_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkRetainIcons.hwnd, "When you drag a program binary to the dock it can take an automatically selected icon or you can retain the embedded icon within the binary file. The automatically selected icon will come from our own collection. An embedded icon may well be good enough to display but be aware, older binaries use very small or low quality icons.", _
                  TTIconInfo, "Help on Retaining Original Icons.", , , , True
End Sub

Private Sub chkSplashStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkSplashStatus.hwnd, "When this checkbox is ticked the dock shows a Splash Screen on Start-up.", _
                  TTIconInfo, "Help on the Splash Screen Checkbox", , , , True
End Sub

Private Sub chkStyleDisable_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkStyleDisable.hwnd, "This checkbox disables the labels that appear above the icon in the dock.", _
                  TTIconInfo, "Help on Label Disable.", , , , True

End Sub

Private Sub chkToggleDialogs_Click()
    
    ' .70 DAEB 16/05/2022 rDIConConfig.frm Read the chkToggleDialogs value from a file and save the value for next time
    If chkToggleDialogs.Value = 0 Then
       sdChkToggleDialogs = "0"
    Else
       sdChkToggleDialogs = "1"
    End If
    
    PutINISetting "Software\SteamyDockSettings", "sdChkToggleDialogs", sdChkToggleDialogs, toolSettingsFile

    Call setToolTips
End Sub

Private Sub chkToggleDialogs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkToggleDialogs.hwnd, "This checkbox acts as a toggle to enable/disable the balloon tooltips.", _
                  TTIconInfo, "Help on the Ballooon Tooltip Toggle", , , , True
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

Private Sub fmeMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim descriptiveText As String
    Dim titleText As String
    
    descriptiveText = ""
    titleText = ""
    
    If rDEnableBalloonTooltips = "1" Then
        If Index = 0 Then
            descriptiveText = "Use this panel to configure the general options that apply to the whole dock program. "
            titleText = "Help on the General Pane."
        ElseIf Index = 1 Then
            descriptiveText = "Use this panel to configure the icon characteristics that apply only to the icons themselves. "
            titleText = "Help on the Icon Characteristics Pane."
        ElseIf Index = 2 Then
            descriptiveText = "Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
            titleText = "Help on the Behaviour Pane."
        ElseIf Index = 3 Then
            descriptiveText = "Use this panel to configure the label and font settings."
            titleText = "Help on the Style Themes and Fonts Pane."
        ElseIf Index = 4 Then
            descriptiveText = "This pane is used to control the location of the dock. "
            titleText = "Help on the Position Pane Button."
        ElseIf Index = 5 Then
            ' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
            fraScrollbarCover.Visible = True
            descriptiveText = "The About Panel provides the version number of this utility, useful information when reporting a bug. The text below this gives due credit to Punk labs for being the originator of  and gives thanks to them for coming up with such a useful tool and also to Apple who created the original idea for this whole genre of docks. This pane also gives access to some useful utilities."
            titleText = "Help on the About Pane Button."
        End If
    End If

    CreateToolTip fmeMain(Index).hwnd, descriptiveText, TTIconInfo, titleText, , , , True

End Sub


Private Sub fraAnimationInterval_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHide_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        lblBehaviourLabel(2).Enabled = False
        lblBehaviourLabel(8).Enabled = False
        sliBehaviourAutoHideDuration.Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblBehaviourLabel(3).Enabled = False
        lblBehaviourLabel(9).Enabled = False
        lblAutoRevealDurationMsHigh.Enabled = False
        sliBehaviourPopUpDelay.Enabled = False
        lblBehaviourPopUpDelayMsCurrrent.Enabled = False
        
    Else
        lblBehaviourLabel(2).Enabled = True
        lblBehaviourLabel(8).Enabled = True
        sliBehaviourAutoHideDuration.Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblBehaviourLabel(3).Enabled = True
        lblBehaviourLabel(9).Enabled = True
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


Private Sub Form_Initialize()
    dockSettingsYPos = ""
    dockSettingsXPos = ""
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
    rDEnableBalloonTooltips = "1"
    
    mnupopmenu.Visible = False

    On Error GoTo Form_Load_Error
    If debugflg = 1 Then DebugPrint "%Form_Load"
    
    Call getAllDriveNames(sAllDrives)
                           
    'if the process already exists then kill it
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "docksettings.exe"
        checkAndKill NameProcess, False, False
        'MsgBox "You now have two instances of this utility running, they will conflict..."
    End If
    
    ' the frames can jump about in the IDE during development, this just places them accurately at runtime
    Call placeFrames
    
    'load the about text
    Call loadAboutText
      
    ' get the location of this tool's settings file
    Call getToolSettingsFile
    
    ' check the Windows version
    Call testWindowsVersion(classicThemeCapable)
    
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
'    If IsUserAnAdmin() = 0 Then
'        MsgBox "This tool requires to be run as administrator on Windows 7 and above in order to function. Admin access is NOT required on Win7 and below. If you aren't entirely happy with that then you'll need to remove the software now. This is a limitation imposed by Windows itself. To enable administrator access find this tool's exe and right-click properties, compatibility - run as administrator. YOU have to do this manually, I can't do it for you."
'    End If

    ' check where rocketdock is installed
    Call checkRocketdockInstallation
    'If rocketDockInstalled = True Then
        'dockAppPath = rdAppPath
        'txtGeneralRdLocation.Text = rdAppPath
        'defaultDock = 0
    'End If
    
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
    
    sdChkToggleDialogs = GetINISetting("Software\SteamyDockSettings", "sdChkToggleDialogs", toolSettingsFile)
    
    If sdChkToggleDialogs = "" Then sdChkToggleDialogs = "1" ' validate
    If sdChkToggleDialogs = "1" Then ' set
        chkToggleDialogs.Value = 1
    Else
        chkToggleDialogs.Value = 0
    End If

    
    ' set the tooltips for the utility
    Call setToolTips
    
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties(dockSettings)
    
    Call makeVisibleFormElements
    
    startupFlg = False ' now negate the startup flag

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form dockSettings"
     
End Sub

Private Sub makeVisibleFormElements()

    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0

    screenHeightTwips = GetDeviceCaps(Me.hdc, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(Me.hdc, HORZRES) * screenTwipsPerPixelX ' replaces buggy screen.width

    ' read the form X/Y params from the toolSettings.ini
'    dockSettingsYPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsYPos", toolSettingsFile)
'    dockSettingsXPos = GetINISetting("Software\SteamyDockSettings", "dockSettingsXPos", toolSettingsFile)
'
'    If dockSettingsYPos <> "" Then
'        dockSettings.Top = Val(dockSettingsYPos)
'    Else
'        dockSettings.Top = Screen.Height / 2 - dockSettings.Height / 2
'    End If
'
'    If dockSettingsXPos <> "" Then
'        dockSettings.Left = Val(dockSettingsXPos)
'    Else
'        dockSettings.Left = Screen.Width / 2 - dockSettings.Width / 2
'    End If

    ' read the form's saved X/Y params from the toolSettings.ini in twips and convert to pixels
    formLeftPixels = Val(GetINISetting("Software\SteamyDockSettings", "dockSettingsXPos", toolSettingsFile)) / screenTwipsPerPixelX
    formTopPixels = Val(GetINISetting("Software\SteamyDockSettings", "dockSettingsYPos", toolSettingsFile)) / screenTwipsPerPixelY

    Call adjustFormPositionToCorrectMonitor(Me.hwnd, formLeftPixels, formTopPixels)
        
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

''---------------------------------------------------------------------------------------
'' Procedure : chkGenAlwaysAsk_Click
'' Author    : beededea
'' Date      : 13/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub chkGenAlwaysAsk_Click()
'   On Error GoTo chkGenAlwaysAsk_Click_Error
'
'    rDAlwaysAsk = chkGenAlwaysAsk.Value
'
'   On Error GoTo 0
'   Exit Sub
'
'chkGenAlwaysAsk_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenAlwaysAsk_Click of Form dockSettings"
'
'End Sub



Private Sub fraAutoHideDelay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAutoHideDuration_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub




' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False
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







Private Sub genChkShowIconSettings_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip genChkShowIconSettings.hwnd, "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here.", _
                  TTIconInfo, "Help on the automatic icon Settings Startup", , , , True
End Sub

Private Sub Label7_Click()

End Sub









' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
Private Sub lblAboutText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False
End Sub

Private Sub lblChkLabelBackgrounds_Click()
' the reason there is a separate label for certain checkboxes is due to the way that VB6 greys out checkbox labels using specific fonts causing them to be crinkled. When the label is unattached to the chkbox then it greys out correctly.
    Call chkLabelBackgrounds_Click
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
'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
'    If optGeneralReadConfig.Value = True And defaultDock = 1 And steamyDockInstalled = True And rocketDockInstalled = True Then
'        'chkGenAlwaysAsk.Enabled = True
'        'lblChkAlwaysConfirm.Enabled = True
'    End If
    
    rDGeneralReadConfig = optGeneralReadConfig.Value ' this is the nub
    
    'Call locateDockSettingsFile

   On Error GoTo 0
   Exit Sub

optGeneralReadConfig_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadConfig_Click of Form dockSettings"
End Sub

Private Sub optGeneralReadConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadConfig.hwnd, "This stores ALL SteamyDock's configuration within the user data area. This option retains future compatibility within modern versions of Windows. Not applicable for Rocketdock ", _
                  TTIconInfo, "Help on using SteamyDock's config.", , , , True
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
        'If defaultDock = 0 Then optGeneralWriteRegistry.Value = True ' if running Rocketdock the two must be kept in sync
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
        'chkGenAlwaysAsk.Enabled = False
        'lblChkAlwaysConfirm.Enabled = False
        
        rDGeneralReadConfig = optGeneralReadConfig.Value ' turns off the reading from the new location

   On Error GoTo 0
   Exit Sub

optGeneralReadRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralReadRegistry_Click of Form dockSettings"

End Sub

Private Sub optGeneralReadRegistry_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadRegistry.hwnd, "This option allows you to read Rocketdock's configuration from the Rocketdock portion of the Registry. This method is becoming increasingly incompatible with newer Windows beyond XP as it can cause some security problems on newer system as it requires admin rights to write back. Use it here in a read-only fashion to migrate from Rocketdock.", _
                  TTIconInfo, "Help on reading from the registry", , , , True
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
'        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
'            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
'            ' create a settings.ini file in the rocketdock folder
'            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
'            Close #1         ' close the file and
'             ' test it exists
'            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
'                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
'                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
'            End If
'        End If
    End If
        
'    lblGenLabel(0).Enabled = False
'    lblGenLabel(1).Enabled = False
    sliGenRunAppInterval.Enabled = False
    lblGenLabel(2).Enabled = False
    lblGenRunAppIntervalCur.Enabled = False
    'chkGenAlwaysAsk.Enabled = False
    'lblChkAlwaysConfirm.Enabled = False
    
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
Private Sub fmeSizePreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

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
        'chkGenAlwaysAsk.Value = Val(GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile))
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
        If rDGeneralReadConfig <> "" Then
            optGeneralReadConfig.Value = rDGeneralReadConfig
        Else
            optGeneralReadConfig.Value = False
        End If
    End If
    
    cmbDefaultDock.ListIndex = 1
    cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
    
    dockAppPath = sdAppPath
    txtGeneralRdLocation.Text = dockAppPath
    defaultDock = 1
    ' write the default dock to the SteamyDock settings file
    PutINISetting "Software\SteamyDockSettings", "defaultDock", defaultDock, toolSettingsFile
                
'    If steamyDockInstalled = True And rocketDockInstalled = True Then
'        If chkGenAlwaysAsk.Value = 1 Then  ' depends upon being able to read the new configuration file in the user data area
'            answer = MsgBox("Both Rocketdock and SteamyDock are installed on this system. Use SteamyDock by default? ", vbYesNo)
'            If answer = vbYes Then
'                'cmbDefaultDock.ListIndex = 1 ' steamy dock
'                dockAppPath = sdAppPath
'                txtGeneralRdLocation.Text = sdAppPath
'                defaultDock = 1
'            Else
'                'cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                txtGeneralRdLocation.Text = rdAppPath
'                defaultDock = 0
'            End If
'        Else
'            ' if the question is not being asked then use the default dock as specified in the docksettings.ini file
'            If rDDefaultDock = "steamydock" Then
'                'cmbDefaultDock.ListIndex = 1
'                dockAppPath = sdAppPath
'                txtGeneralRdLocation.Text = dockAppPath
'                defaultDock = 1
'            ElseIf rDDefaultDock = "rocketdock" Then
'                'cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                txtGeneralRdLocation.Text = rdAppPath
'                defaultDock = 0
'            Else
''                If cmbDefaultDock.ListIndex = 1 Then  ' depends upon being able to read the new configuration file in the user data area
''                    dockAppPath = sdAppPath
''                    txtGeneralRdLocation.Text = dockAppPath
''                    defaultDock = 1
''                Else
''                    cmbDefaultDock.ListIndex = 0 ' rocket dock
''                    dockAppPath = rdAppPath
''                    txtGeneralRdLocation.Text = rdAppPath
''                    defaultDock = 0
''                End If
'            End If
'        End If
'    ElseIf steamyDockInstalled = True Then ' just steamydock installed
'            cmbDefaultDock.ListIndex = 1
'            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
'
'            dockAppPath = sdAppPath
'            txtGeneralRdLocation.Text = dockAppPath
'            defaultDock = 1
'            ' write the default dock to the SteamyDock settings file
'            PutINISetting "Software\SteamyDockSettings", "defaultDock", defaultDock, toolSettingsFile
'
'    ElseIf rocketDockInstalled = True Then ' just rocketdock installed
'            cmbDefaultDock.ListIndex = 0
'            cmbDefaultDock.Enabled = False ' .11 DAEB 26/04/2021 docksettings Disable the dock select dropdown when only steamydock is present
'
'            dockAppPath = rdAppPath
'            txtGeneralRdLocation.Text = rdAppPath
'            defaultDock = 0
'    End If
    
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
                            
'        If defaultDock = 0 Then
'            rDVersion = "1.3.5"
'        Else
            rDVersion = App.Major & "." & App.Minor & "." & App.Revision
'        End If
    End If
    
    If optGeneralReadConfig.Value = False Then
        ' read the dock settings from INI or from registry
        Call readDockSettings
        Call adjustControls
    End If
    
    'if rocketdock set the automatic startup string to Rocketdock
'    If defaultDock = 0 Then ' rocketdock
'        rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "RocketDock")
'        If rdStartupRunString <> "" Then
'            rDStartupRun = "1"
'            chkGenWinStartup.Value = 1
'        End If
'    ElseIf defaultDock = 1 Then 'if rocketdock set the automatic startup string to Steamydock
        rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "SteamyDock")
        If rdStartupRunString <> "" Then
            rDStartupRun = "1"
            chkGenWinStartup.Value = 1
        End If
'    End If

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
    
'    Dim strLine As String
'    Dim strFile As String
'    Dim intFile As Integer
'
'    strLine = ""
'    strFile = App.Path & "about.txt"
'    '
'    ' If the file exists then read it and
'    ' populate the TextBox
'    '
'    If FExists(strFile) <> "" Then
'        intFile = FreeFile
'        Open strFile For Input As intFile
'        Line Input #1, strLine
'        lblAboutText.Text = strLine
'        Close intFile
'    End If


    Call LoadFileToTB(lblAboutText, App.Path & "\about.txt", False)

'    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project to allow easier configuration, bug-fixing and enhancement of Rocketdock and currently underway, a fully open source version of a Rocketdock clone."
'
'    lblAboutPara4.Caption = "The first steps are the two VB6 utilities that replicate the icons settings and dock settings screen. The subsequent step is the dock itself. I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received."
'
'    lblAboutPara1.Caption = "The original Rocketdock was developed by the Apple fanboy and fangirl team at Punklabs. They developed it as a peace offering from the Mac community to the Windows Community."
'    lblAboutPara2.Caption = "This new dock, now known as SteamyDock, was developed by a Windows/ReactOS fanboy on Windows 7 using VB6. This utility faithfully reproduces the original as created by Punklabs, originally done solely as a homage to the original as that version is no longer being supported but now it has evolved into a set of tools that has become a replacement for rocketdock itself. It must be said, the initial idea for this dock came from Punklabs and Rocketdock's OS/X dock predecessors. All HAIL to Punklabs!"
'    lblAboutPara3.Caption = "This version was developed on Windows using VisualBasic 6 as a FOSS project. It is open source to allow easier configuration, bug-fixing and enhancement of Rocketdock and community contribution towards this new dock."
'    lblAboutPara4.Caption = "The first steps were the two VB6 utilities that replicate the icons settings and dock settings screen (this utility). These are largely complete and the dock itself is now under development and 90% complete. A future step is conversion to RADBasic/TwinBasic or VB.NET for future-proofing and 64bit-ness. This next step is 1/3rd underway."
'
'    lblAboutPara5.Caption = "I do hope you enjoy using these utilities. Your software enhancements and contributions will be gratefully received if you choose to contribute."
'    lblAboutPara6.Caption = "This utility MUST run as administrator in order to access Rocketdock's " & _
'                            "registry settings (due to a Windows shadow registry feature/bug that " & _
'                            "gives incorrect shadow data). If you run it without admin rights and " & _
'                            "you want to change the values in the registry then some of the values may " & _
'                            "be incorrect and the resulting dock might look and act rather strange. " & _
'                            "You have been warned!"

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
'    Else
'        If IsUserAnAdmin() = 0 Then
'            answer = MsgBox("This program is not running as admin. Some of the settings may be strange and unwanted - be warned. Continue?", vbYesNo)
'            If answer = vbNo Then
'                Exit Sub
'            End If
'        End If
    End If
   
    ' kill the rocketdock /steamydock process first
    
'    If defaultDock = 0 Then
'        NameProcess = "RocketDock.exe"
'    Else
        NameProcess = "steamyDock.exe"
'    End If
    ans = checkAndKill(NameProcess, False, False)
            
    ' if the settings.ini has been chosen as an option then the creation of it will already have occurred,
    ' so, if the temporary settings file exists then it means that the user clicked "use settings.ini file"
    ' in which case we copy it to the main settings.ini file.
    
    debugPoint = 1
    ' Steamydock exists so we shall write to the settings file those additonal items that need to be there regardless of the location of the dock data
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralReadConfig", rDGeneralReadConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "GeneralWriteConfig", rDGeneralWriteConfig, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "RunAppInterval", rDRunAppInterval, dockSettingsFile
    'PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "DefaultDock", rDDefaultDock, dockSettingsFile
    
    'If optGeneralWriteConfig.Value = True Then ' the 3rd option
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
        
        ' the dock icon settings are empty? deanieboy
        If positionZeroFail = True And positionThreeFail = True Then
            If FExists(dockSettingsFile) Then ' does the temporary settings.ini exist?
                ' read the registry values for each of the icons and write them to the settings.ini
                'Call readIconsWriteSettings("Software\SteamyDock\IconSettings", dockSettingsFile)
            End If
        End If

'    Else
'        debugPoint = 3
'        If rocketDockInstalled = True Then
'          If optGeneralWriteSettings.Value = True Then ' use the settings file
'            debugPoint = 4
'            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
'                Call writeDockSettings("Software\RocketDock", tmpSettingsFile)
'                ' if it exists, read the registry values for each of the icons and write them to the settings.ini
'                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
'            End If
'            If FExists(tmpSettingsFile) Then ' does the tmp settings.ini exist?
'                debugPoint = 5
'                If FExists(origSettingsFile) Then ' does the tmp settings.ini exist?
'                    debugPoint = 6
'                    Kill origSettingsFile
'                End If
'                debugPoint = 7
'                Name tmpSettingsFile As origSettingsFile ' rename 'our' settings file to the one used by RD
'            End If
'          Else ' WRITE THE REGISTRY AND remove the settings file
'            Call writeRegistry
'            ' this function restarts Rocketdock so that the changes 'take'.
'            Sleep (1300) ' this is required as the o/ses' final commit of the data to the registry can be delayed
'                         ' and without the pause the restart does not picku p the committed data.
'            'see if the settings.ini exists
'            ' if it does exist, ensure it no longer does so by deleting it, RD will then use the registry.
'            If FExists(origSettingsFile) Then ' does the original settings.ini exist?
'                    Kill origSettingsFile
'            End If
'          End If
'        End If
'    End If

  
    ' From the general panel
    ' these write to registry areas available to any program not just Rocketdock
    
    picBusy.Visible = True
    busyTimer.Enabled = True
    
    If rDStartupRun = "1" Then
        'if rocketdock set the string to Rocketdock startup
'        If defaultDock = 0 Then ' rocketdock
'            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RocketDock", """" & txtGeneralRdLocation.Text & "\" & "RocketDock.exe""")
'        End If
        'if steamydock set the string to steamydock startup
        If defaultDock = 1 Then ' steamydock
            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SteamyDock", """" & txtGeneralRdLocation.Text & "\" & "SteamyDock.exe""")
        End If
    Else
'        If defaultDock = 0 Then ' rocketdock
'            Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RocketDock", "")
'        End If
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
    
'    If defaultDock = 0 Then
'        NameProcess = "RocketDock.exe"
'    Else
'        NameProcess = "steamyDock.exe"
'    End If
'    ans = checkAndKill(NameProcess, False)
    
'    If defaultDock = 0 Then
'        rDVersion = "1.3.5"
'    Else
        rDVersion = App.Major & "." & App.Minor & "." & App.Revision
'    End If
    
    rDCustomIconFolder = ""
    
'    If defaultDock = 0 Then ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'        rDHotKeyToggle = "Control+Alt+R"
'    Else
        rDHotKeyToggle = "F11"
'    End If
    cmbHidingKey.Text = rDHotKeyToggle
        
    ' removed
    'cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
'    If defaultDock = 1 Then
'        cmbHidingKey.Text = rDHotKeyToggle ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'    Else
'        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
'    End If
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
    
    ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    rDRetainIcons = "0"
    chkRetainIcons.Value = Val(rDRetainIcons)
    
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
    
    rDSoundSelection = "0"
    cmbBehaviourSoundSelection.ListIndex = Val(rDSoundSelection)
    
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
  
        lblBehaviourLabel(2).Enabled = True

        lblBehaviourLabel(8).Enabled = True
        lblAutoHideDurationMsHigh.Enabled = True
        lblAutoHideDurationMsCurrent.Enabled = True
        
        lblBehaviourLabel(4).Enabled = True
        lblBehaviourLabel(10).Enabled = True
        sliBehaviourAutoHideDelay.Enabled = True
        lblAutoHideDelayMsHigh.Enabled = True
        lblAutoHideDelayMsCurrent.Enabled = True
        
        lblBehaviourLabel(3).Enabled = True
        lblBehaviourLabel(9).Enabled = True
        lblAutoRevealDurationMsHigh.Enabled = True
        sliBehaviourPopUpDelay.Enabled = True
        
        lblBehaviourPopUpDelayMsCurrrent.Enabled = True
        
        cmbBehaviourAutoHideType.Enabled = True
        
    Else
        chkBehaviourAutoHide.Caption = "Autohide Disabled"
        sliBehaviourAutoHideDuration.Enabled = False

        lblBehaviourLabel(2).Enabled = False

        lblBehaviourLabel(8).Enabled = False
        lblAutoHideDurationMsHigh.Enabled = False
        lblAutoHideDurationMsCurrent.Enabled = False
        
        lblBehaviourLabel(4).Enabled = False
        lblBehaviourLabel(10).Enabled = False
        sliBehaviourAutoHideDelay.Enabled = False
        lblAutoHideDelayMsHigh.Enabled = False
        lblAutoHideDelayMsCurrent.Enabled = False
        
                
        lblBehaviourLabel(3).Enabled = False
        lblBehaviourLabel(9).Enabled = False
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
    Else
        chkGenDisableAnim.Enabled = True
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
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else

        If optGeneralReadConfig.Value = True Then ' steamydock
'            lblGenLabel(0).Enabled = True
'            lblGenLabel(1).Enabled = True
            sliGenRunAppInterval.Enabled = True
            lblGenLabel(2).Enabled = True
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
                
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkLabelBackgrounds.Width = 192 ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblChkLabelBackgrounds.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblChkLabelBackgrounds.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        btnStyleFont.Enabled = False
        lblStyleFontName.Enabled = False
        btnStyleShadow.Enabled = False
        lblStyleFontFontShadowColor.Enabled = False
        lblStyleFontFontShadowTest.Enabled = False
        btnStyleOutline.Enabled = False
        lblStyleOutlineColourDesc.Enabled = False
        lblStyleFontOutlineTest.Enabled = False
        lblStyleLabel(3).Enabled = False
        lblStyleLabel(4).Enabled = False
        lblStyleLabel(5).Enabled = False
        lblStyleLabel(8).Enabled = False
        lblStyleLabel(9).Enabled = False
        lblStyleLabel(10).Enabled = False
        sliStyleShadowOpacity.Enabled = False
        lblStyleShadowOpacityCurrent.Enabled = False
        sliStyleOutlineOpacity.Enabled = False
        lblStyleOutlineOpacityCurrent.Enabled = False
        
        sliStyleFontOpacity.Enabled = False
        lblStyleFontOpacityCurrent.Enabled = False
        
    Else
        chkLabelBackgrounds.Enabled = True  ' .01 docksettings DAEB added the greying out or enabling of the checkbox and label for the icon label background toggle
        'lblChkLabelBackgrounds.Enabled = True ' .01
        
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkLabelBackgrounds.Width = 2490 ' set the width to show the full check box and its intrinsic label
        lblChkLabelBackgrounds.Visible = False ' make the associated duplicate label hidden
        lblChkLabelBackgrounds.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        btnStyleFont.Enabled = True
        lblStyleFontName.Enabled = True
        btnStyleShadow.Enabled = True
        lblStyleFontFontShadowColor.Enabled = True
        lblStyleFontFontShadowTest.Enabled = True
        btnStyleOutline.Enabled = True
        lblStyleOutlineColourDesc.Enabled = True
        lblStyleFontOutlineTest.Enabled = True
        sliStyleShadowOpacity.Enabled = True
        lblStyleShadowOpacityCurrent.Enabled = True
        lblStyleLabel(3).Enabled = True
        lblStyleLabel(4).Enabled = True
        lblStyleLabel(5).Enabled = True
        lblStyleLabel(8).Enabled = True
        lblStyleLabel(9).Enabled = True
        lblStyleLabel(10).Enabled = True
        sliStyleOutlineOpacity.Enabled = True
        lblStyleOutlineOpacityCurrent.Enabled = True
                
        sliStyleFontOpacity.Enabled = True
        lblStyleFontOpacityCurrent.Enabled = True
        
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
' Procedure : cmbBehaviourSoundSelection_Click
' Author    : beededea
' Date      : 17/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbBehaviourSoundSelection_Click()
    On Error GoTo cmbBehaviourSoundSelection_Click_Error

    rDSoundSelection = cmbBehaviourSoundSelection.ListIndex

    On Error GoTo 0
    Exit Sub

cmbBehaviourSoundSelection_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbBehaviourSoundSelection_Click of Form dockSettings"
            Resume Next
          End If
    End With

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
    
    rDHoverFX = "1"  'DEAN needs to be removed later
    
    'none
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
        
        If chkToggleDialogs.Value = 0 Then
            sliIconsZoom.ToolTipText = "The maximum size after a zoom can be no smaller than 85 pixels when Zoom:Bumpy is chosen"
        Else
            sliIconsZoom.ToolTipText = ""
        End If
    Else
        
        If chkToggleDialogs.Value = 0 Then
            sliIconsZoom.ToolTipText = "The maximum size after a zoom"
        Else
            sliIconsZoom.ToolTipText = ""
        End If
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

'    If cmbDefaultDock.List(cmbDefaultDock.ListIndex) = "RocketDock" Then
'        ' check where rocketdock is installed
'        Call checkRocketdockInstallation
'        defaultDock = 0 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
'
'        ' .17 DAEB 07/09/2022 docksettings the dock folder location now changes as it is switched between Rocketdock and Steamy Dock
'        dockAppPath = rdAppPath
'        txtGeneralRdLocation.Text = rdAppPath
'
'        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
'            optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'            optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'        Else
'            optGeneralReadRegistry.Value = True
'            'optGeneralWriteRegistry.Value = True
'        End If
'
'        rDDefaultDock = "rocketdock"
'
'        ' re-enable all the controls that Rocketdock supports
'
'        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenMin.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenMin.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenMin.Visible = False ' make the associated duplicate label hidden
'        lblChkGenMin.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenDisableAnim.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenDisableAnim.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenDisableAnim.Visible = False ' make the associated duplicate label hidden
'        lblChkGenDisableAnim.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkIconsZoomOpaque.Enabled = True ' RD does not support storing the configs at the correct location
'        chkIconsZoomOpaque.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblchkIconsZoomOpaque.Visible = False ' make the associated duplicate label hidden
'        lblchkIconsZoomOpaque.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        sliIconsDuration.Enabled = True
'        lblCharacteristicsLabel(6).Enabled = True
'        lblCharacteristicsLabel(11).Enabled = True
'        lblCharacteristicsLabel(12).Enabled = True
'        lblIconsDurationMsCurrent.Enabled = True
'
'        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
''        sliIconsZoomWidth.Enabled = True
''        sliIconsDuration.Enabled = True
'
'        cmbIconsHoverFX.Enabled = True
'
''        Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
''        Call setBounceTypes
'
'        sliBehaviourAutoHideDuration.Enabled = True
'        sliAnimationInterval.Enabled = True
'
'        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        'fraZoomConfigs.Visible = True
'
'        fraAutoHideDuration.Visible = True
'        fraFontOpacity.Visible = True
'
'
'        optGeneralReadSettings.Enabled = True
'        optGeneralReadRegistry.Enabled = True
'
''        lblGenLabel(0).Enabled = False
''        lblGenLabel(1).Enabled = False
'        sliGenRunAppInterval.Enabled = False
'        lblGenLabel(2).Enabled = False
'        lblGenRunAppIntervalCur.Enabled = False
'
'        chkGenAlwaysAsk.Enabled = False
'
'        sliAnimationInterval.Enabled = False
'        lblBehaviourLabel(7).Enabled = False
'        lblAnimationIntervalMsLow.Enabled = False
'        lblAnimationIntervalMsHigh.Enabled = False
'        lblAnimationIntervalMsCurrent.Enabled = False
'        lblBehaviourLabel(12).Enabled = False
'
'        cmbBehaviourAutoHideType.Enabled = False
'
'        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
'
'        picThemeSample.Enabled = False
'        lblStyleLabel(2).Enabled = False
'
'        sliStyleThemeSize.Enabled = False
'        lblThemeSizeTextHigh.Enabled = False
'        lblStyleSizeCurrent.Enabled = False
'
'        lblBehaviourLabel(5).Enabled = False
'        lblBehaviourLabel(11).Enabled = False
'        sliContinuousHide.Enabled = False
'        lblContinuousHideMsHigh.Enabled = False
'        lblContinuousHideMsCurrent.Enabled = False
'
'        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
'
'        ' RD does not support storing the configs at the correct location
'
'        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        optGeneralReadConfig.Enabled = False ' RD does not support storing the configs at the correct location
'        optGeneralReadConfig.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lbloptGeneralReadConfig.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lbloptGeneralReadConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        optGeneralWriteConfig.Enabled = False ' RD does not support storing the configs at the correct location
'        optGeneralWriteConfig.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblOptGeneralWriteConfig.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblOptGeneralWriteConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkGenAlwaysAsk.Enabled = False ' RD does not support storing the configs at the correct location
'        chkGenAlwaysAsk.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblChkGenAlwaysAsk.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblChkGenAlwaysAsk.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkSplashStatus.Enabled = False ' RD does not support storing the configs at the correct location
'        chkSplashStatus.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblChkSplashStatus.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblChkSplashStatus.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        genChkShowIconSettings.Enabled = False ' RD does not support storing the configs at the correct location
'        genChkShowIconSettings.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblGenChkShowIconSettings.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblGenChkShowIconSettings.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'
'        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
'        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
'        chkRetainIcons.Enabled = False
'        chkRetainIcons.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
'        lblRetainIcons.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
'        lblRetainIcons.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
'        lblBehaviourLabel(14).Enabled = False
'
'    Else
        ' check where/if steamydock is installed
        Call checkSteamyDockInstallation 'defaultDock is set here
        
        rDDefaultDock = "steamydock"
        defaultDock = 1 ' .13 DAEB 29/04/2021 docksettings set the default dock for some reason not already set
        
        ' .17 DAEB 07/09/2022 docksettings the dock folder location now changes as it is switched between Rocketdock and Steamy Dock
        dockAppPath = sdAppPath
        txtGeneralRdLocation.Text = sdAppPath
        
        ' .19 DAEB 07/09/2022 docksettings when you select rocketdock it reverts to the registry but when you select steamydock it does not revert to the dock settings file.
        optGeneralReadConfig.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
        optGeneralWriteConfig.Value = True ' we just want to set this checkbox but we don't want this to trigger a click

        'disable all the controls that steamy dock does not support
                
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkGenMin.Enabled = False ' RD does not support storing the configs at the correct location
        chkGenMin.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblChkGenMin.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblChkGenMin.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        
        'lblChkMinimise.Enabled = False
        'cmbBehaviourActivationFX.Enabled = False
        'cmbStyleTheme.Enabled = False ' does not support themes yet
        'cmbPositionMonitor.Enabled = False
        'cmbIconsQuality.Enabled = False '  does not support enhanced or lower quality icons
        
        sliIconsDuration.Enabled = False ' ' does not support animations at all
        lblCharacteristicsLabel(6).Enabled = False
        lblCharacteristicsLabel(11).Enabled = False
        lblCharacteristicsLabel(12).Enabled = False
        lblIconsDurationMsCurrent.Enabled = False
        
        
        'chkGenOpen.Enabled = False ' does not support showing opening running applications, always opens new apps.
        'lblChkOpenRunning.Enabled = False
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
'        sliIconsZoomWidth.Enabled = False ' does not support zoomwidth though this is possible later
'        sliIconsDuration.Enabled = False ' does not support animations at all

        '.nn cmbIconsHoverFX.Enabled = False ' does not support hover effects other than the default
        '.nn sliBehaviourAutoHideDuration.Enabled = False ' does not support animation at all
        'sliAnimationInterval.Enabled = False ' does not support animation at all
                
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkGenDisableAnim.Enabled = False ' RD does not support storing the configs at the correct location
        chkGenDisableAnim.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblChkGenDisableAnim.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblChkGenDisableAnim.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
                
        ' allows the greying out of the checkbox label without showing a crinkly text on those fonts with serifs
        chkIconsZoomOpaque.Enabled = False ' RD does not support storing the configs at the correct location
        chkIconsZoomOpaque.Width = 192  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        lblchkIconsZoomOpaque.Visible = True ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        lblchkIconsZoomOpaque.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' Some of the controls have been bundled onto frames so that they can all be hidden entirely for Steamydock users
        
        ' 30/10/2020 docksettings .06 DAEB fraZoomConfigs containing sliIconsZoomWidth made visible by default using the IDE and the references to make them otherwise removed.
        'fraZoomConfigs.Visible = False
        
        'fraAutoHideDuration.Visible = true
        fraFontOpacity.Visible = True
        
        optGeneralReadConfig.Enabled = True


        optGeneralWriteConfig.Enabled = True

'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
        
'        If optGeneralReadConfig.Value = True And steamyDockInstalled = True And rocketDockInstalled = True Then
'            chkGenAlwaysAsk.Enabled = True
'            'lblChkAlwaysConfirm.Enabled = True
'        End If

        sliAnimationInterval.Enabled = True
        lblBehaviourLabel(7).Enabled = True
        lblAnimationIntervalMsLow.Enabled = True
        lblAnimationIntervalMsHigh.Enabled = True
        lblAnimationIntervalMsCurrent.Enabled = True
        lblBehaviourLabel(12).Enabled = True
        
        cmbBehaviourAutoHideType.Enabled = True
        
        ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        picThemeSample.Enabled = True
        lblStyleLabel(2).Enabled = True
        lblStyleLabel(3).Enabled = True
        lblStyleLabel(4).Enabled = True
        lblStyleLabel(5).Enabled = True
        lblStyleLabel(8).Enabled = True
        lblStyleLabel(9).Enabled = True
        lblStyleLabel(10).Enabled = True
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
        
        lblBehaviourLabel(5).Enabled = True
        lblBehaviourLabel(11).Enabled = True
        sliContinuousHide.Enabled = True
        lblContinuousHideMsHigh.Enabled = True
        lblContinuousHideMsCurrent.Enabled = True
        
        ' .10 ENDS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
        
        
        ' .16 DAEB 01/7/2022 docksettings DAEB added the juggling of the checkboxes and labels to allow greying out or enabling of the checkbox and labels without causing crinkly effect with serif fonts.
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        optGeneralReadConfig.Enabled = True ' RD does not support storing the configs at the correct location
        optGeneralReadConfig.Width = 5820 ' set the width to show the full check box and its intrinsic label
        lbloptGeneralReadConfig.Visible = False ' make the associated duplicate label hidden
        lbloptGeneralReadConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        optGeneralWriteConfig.Enabled = True ' RD does not support storing the configs at the correct location
        optGeneralWriteConfig.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblOptGeneralWriteConfig.Visible = False ' make the associated duplicate label hidden
        'lblOptGeneralWriteConfig.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
'        chkGenAlwaysAsk.Enabled = True ' RD does not support storing the configs at the correct location
'        chkGenAlwaysAsk.Width = 5820 ' set the width to show the full check box and its intrinsic label
'        lblChkGenAlwaysAsk.Visible = False ' make the associated duplicate label hidden
'        lblChkGenAlwaysAsk.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkSplashStatus.Enabled = True ' RD does not support storing the configs at the correct location
        chkSplashStatus.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblChkSplashStatus.Visible = False ' make the associated duplicate label hidden
        'lblChkSplashStatus.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
                
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        genChkShowIconSettings.Enabled = True ' RD does not support storing the configs at the correct location
        genChkShowIconSettings.Width = 5820 ' set the width to show the full check box and its intrinsic label
        'lblGenChkShowIconSettings.Visible = False ' make the associated duplicate label hidden
        'lblGenChkShowIconSettings.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        
        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        ' makes the checkbox label visible and able to support a balloon tooltip which the associated label cannot being a windowless control.
        chkRetainIcons.Enabled = True
        chkRetainIcons.Width = 5820  ' set the width to just show the check box itself and hide its intrinsic label that then goes 'crinkly'.
        'lblRetainIcons.Visible = False ' make the associated duplicate label visible, greyed out looks just like any other greyed out text
        'lblRetainIcons.Enabled = False ' ensure the associated label stays disabled, it should always be disabled
        lblBehaviourLabel(14).Enabled = True ' associated title label
        
    'End If
    
    Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
    Call setBounceTypes
    Call setSoundSelectionDropDown

    Call setHidingKey
        
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
Private Sub fmeMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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


    ' .23 DAEB 02/10/2022 docksettings added control logic to hide/show the scrollbar
    fraScrollbarCover.Visible = True

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
    Dim NameProcess As String: NameProcess = ""
    Dim ofrm As Form

    On Error GoTo Form_Unload_Error
    
    If debugflg = 1 Then DebugPrint "%" & "Form_Unload"
    
    ' save the current X and y position of this form to allow repositioning when restarting
    dockSettingsXPos = dockSettings.Left
    dockSettingsYPos = dockSettings.Top
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\SteamyDockSettings", "dockSettingsXPos", dockSettingsXPos, toolSettingsFile
    PutINISetting "Software\SteamyDockSettings", "dockSettingsYPos", dockSettingsYPos, toolSettingsFile
    
    NameProcess = "PersistentDebugPrint.exe"

    If debugflg = 1 Then
        checkAndKill NameProcess, False, False
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



Private Sub lblAboutPara4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub


Private Sub lblAboutPara5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblAboutPara1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    Dim answer As VbMsgBoxResult
   
    On Error GoTo lblPunklabsLink_Click_Error
   
    answer = vbNo
   
    If debugflg = 1 Then Debug.Print "%lblPunklabsLink_Click"
    
    ' .22 DAEB 02/10/2022 docksettings added a message pop up on the punklabs link
    answer = MsgBox("This link opens a browser window and connects to Punklabs Homepage. Would you like to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.punklabs.com", vbNullString, App.Path, 1)
    End If
    
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
    
    'load the gear images
    picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1.jpg")
    picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3.jpg")

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
    
    'load the gear images
    picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1Light.jpg")
    picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3Light.jpg")

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
        If useloop > 0 Then picIcon(useloop).BackColor = vbWhite
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
    
        
    lblAboutText.BackColor = RGB(redC, greenC, blueC)
    
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
Private Sub picCogs1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

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

Private Sub optGeneralReadSettings_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralReadSettings.hwnd, "This option allows you to read the configuration from Rocketdock's program files folder, this is for migrating in a read-only fashion from RocketDock to SteamyDock. Requires admin access so only select this option when migrating from Rocketdock. ", _
                  TTIconInfo, "Help on reading from the settings.ini.", , , , True
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

Private Sub optGeneralWriteConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteConfig.hwnd, "This option stores ALL configuration within the user data area retaining future compatibility in Windows. Not available to Rocketdock.", _
                  TTIconInfo, "Help on Writing SteamyDock's Config.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteRegistry_Click
' Author    : beededea
' Date      : 05/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub optGeneralWriteRegistry_Click()
'   On Error GoTo optGeneralWriteRegistry_Click_Error
'
'    If optGeneralWriteRegistry.Value = True Then
'        ' nothing to do, the checkbox value is used later to determine where to write the data
'    End If
'    If defaultDock = 0 Then optGeneralReadRegistry.Value = True ' if running Rocketdock the two must be kept in sync
'
'    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location
'
'   On Error GoTo 0
'   Exit Sub
'
'optGeneralWriteRegistry_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteRegistry_Click of Form dockSettings"
'End Sub

'Private Sub optGeneralWriteRegistry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteRegistry.hWnd, "Stores the configuration in the Rocketdock portion of the Registry, incompatible with newer version of Windows, this can cause some security problems and in all case requires admin rights to operate. Best to use option 3. ", _
'                  TTIconInfo, "Help on writing settings to the registry.", , , , True
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : optGeneralWriteSettings_Click
' Author    : beededea
' Date      : 01/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub optGeneralWriteSettings_Click()
'
'   On Error GoTo optGeneralWriteSettings_Click_Error
'
'    tmpSettingsFile = rdAppPath & "\tmpSettings.ini" ' temporary copy of Rocketdock 's settings file
'
'    If startupFlg = True Then '.NET
'        ' don't do this on the first startup run
'        Exit Sub
'    Else
'
'        If optGeneralReadSettings.Value = True Or optGeneralWriteSettings.Value = True Then
'            If defaultDock = 0 Then optGeneralWriteSettings.Value = True ' if running Rocketdock the two must be kept in sync
'            ' create a settings.ini file in the rocketdock folder
'            Open tmpSettingsFile For Output As #1 ' this wipes the file IF it exists or creates it if it doesn't.
'            Close #1         ' close the file and
'             ' test it exists
'            If FExists(tmpSettingsFile) Then ' does the temporary settings.ini exist?
'                ' if it exists, read the registry values for each of the icons and write them to the internal temporary settings.ini
'                Call readIconsWriteSettings("Software\RocketDock", tmpSettingsFile)
'            End If
'        End If
'
'        If defaultDock = 0 Then ' Rocketdock
'            If optGeneralWriteSettings.Value = True Then ' keep the two in synch.
'                If optGeneralReadSettings.Value = False Then
'                    optGeneralReadSettings.Value = True
'                End If
'            End If
'        End If
'    End If
'
'    rDGeneralWriteConfig = optGeneralWriteConfig.Value ' turns off the reading from the new location
'
'   On Error GoTo 0
'   Exit Sub
'
'optGeneralWriteSettings_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optGeneralWriteSettings_Click of Form dockSettings"
'
'End Sub

'Private Sub optGeneralWriteSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip optGeneralWriteSettings.hWnd, "Store configuration in Rocketdock's program files folder, can cause security issues on newer systems beyond XP and requires admin access. Best to move to option 3. ", _
'                  TTIconInfo, "Help on Storing within Rocketdock's program files folder.", , , , True
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : picmultipleGears3_MouseDown
' Author    : beededea
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picmultipleGears3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picmultipleGears3_MouseDown_Error
   If debugflg = 1 Then Debug.Print "%picmultipleGears3_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

picmultipleGears3_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picmultipleGears3_MouseDown of Form dockSettings"
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
    If Index < 0 Then Index = 0
    
'    If defaultDock = 0 Then
'        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
'            PutINISetting "Software\RocketDock", "OptionsTabIndex", rDOptionsTabIndex, origSettingsFile
'        Else
'            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "OptionsTabIndex", rDOptionsTabIndex)
'        End If
'    Else
        'CFG - write the current open tab to the 3rd config settings
        ' .20 DAEB 07/09/2022 docksettings tab selection fixed
        PutINISetting "Software\SteamyDock\DockSettings", "OptionsTabIndex", rDOptionsTabIndex, toolSettingsFile
        
'    End If
    
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

' reverted the picbox array to individual picboxes for the pane buttons to allow a popup help box to appear.
Private Sub picIcon_Click_event(Index As Integer)
   
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

'    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    
    ' the first is the RD settings file that only exists if RD is NOT using the registry
    ' the second is the settings file for this tool to store its own preferences
        
    ' check to see if the first settings file exists
    
    On Error GoTo readDockSettings_Error
   
    If rocketDockInstalled = True Then
        If FExists(origSettingsFile) Then ' does the original settings.ini exist?
            If optGeneralReadConfig.Value = False Then
'                optGeneralReadSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
'                optGeneralWriteSettings.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
            End If
            ' here we read from the settings file
            readDockSettingsFile "Software\RocketDock", origSettingsFile
            Call validateInputs
        Else
            If optGeneralReadConfig.Value = False Then
                optGeneralReadRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
                'optGeneralWriteRegistry.Value = True ' we just want to set this checkbox but we don't want this to trigger a click
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
'Private Sub writeRegistry()
'
'    On Error GoTo writeRegistry_Error
'
'    ' all tested and working but ONLY when run as admin
'
'    'general panel
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", rDLockIcons)
'    ' rDRetainIcons not required
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "OpenRunning", rDOpenRunning)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ShowRunning", rDShowRunning)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ManageWindows", rDManageWindows)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "DisableMinAnimation", rDDisableMinAnimation)
'
'    'icon panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconQuality", Val(rDIconQuality))
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconOpacity", rDIconOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomOpaque", rDZoomOpaque)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMin", rDIconMin)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HoverFX", rDHoverFX)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconMax", rdIconMax)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomWidth", Val(rDZoomWidth))
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ZoomTicks", rDZoomTicks)
'
'    'behaviour panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "IconActivationFX", rDIconActivationFX)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHide", rDAutoHide) '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideTicks", rDAutoHideTicks)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideDelay", rDAutoHideDelay)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "MouseActivate", rDMouseActivate)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "PopupDelay", rDPopupDelay)
'
'
'    'position panel
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Monitor", rDMonitor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", rDSide)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "zOrderMode", rDzOrderMode)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Offset", rDOffset)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "vOffset", rDvOffset)
'
'    'style panel
'    'If rDtheme = "blank" Then rDtheme = ""
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "theme", rDtheme)
'    'If rDtheme = "" Then rDtheme = "blank"
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "ThemeOpacity", rDThemeOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "HideLabels", rDHideLabels)
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontName", rDFontName) '*
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontColor", rDFontColor) '*
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontSize", rDFontSize)
'    'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontCharSet", rD)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontFlags", rDFontFlags) '*
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowColor", rDFontShadowColor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineColor", rDFontOutlineColor)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontOutlineOpacity", rDFontOutlineOpacity)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "FontShadowOpacity", rDFontShadowOpacity)
'
'   On Error GoTo 0
'   Exit Sub
'
'writeRegistry_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistry of Form dockSettings"
'End Sub




'Private Sub picIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If rDEnableBalloonTooltips = "1" Then CreateToolTip picIcon.hWnd, "This button opens the panel to configure the general options that apply to the whole dock program.", _
'                  TTIconInfo, "Help on the General Options Button", , , , True
'End Sub


Private Sub picIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim descriptiveText As String
    Dim titleText As String
    
    descriptiveText = ""
    titleText = ""
    
    If rDEnableBalloonTooltips = "1" Then
        If Index = 0 Then
            descriptiveText = "This Button will select the general pane. Use this panel to configure the general options that apply to the whole dock program. "
            titleText = "Help on the General Pane Button."
        ElseIf Index = 1 Then
            descriptiveText = "This Button will select the characteristics pane. Use this panel to configure the icon characteristics that apply only to the icons themselves. "
            titleText = "Help on the Icon Characteristics Pane Button."
        ElseIf Index = 2 Then
            descriptiveText = "This Button will select the behaviour pane. Use this panel to configure the dock settings that determine how the dock will respond to user interaction. "
            titleText = "Help on the Behaviour Pane Button."
        ElseIf Index = 3 Then
            descriptiveText = "This Button will select the style, themes and fonts pane. This is used to configure the label and font settings."
            titleText = "Help on the Style Themes and Fonts Pane Button."
        ElseIf Index = 4 Then
            descriptiveText = "This Button will select the position pane. This pane is used to control the location of the dock. "
            titleText = "Help on the Position Pane Button."
        ElseIf Index = 5 Then
            descriptiveText = "This Button will select the general pane. The About Panel provides the version number of this utility, useful information when reporting a bug. The text below this gives due credit to Punk labs for being the originator of  and gives thanks to them for coming up with such a useful tool and also to Apple who created the original idea for this whole genre of docks. This pane also gives access to some useful utilities."
            titleText = "Help on the About Pane Button."
        End If
    End If


    CreateToolTip picIcon(Index).hwnd, descriptiveText, TTIconInfo, titleText, , , , True

End Sub

Private Sub picMinSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picMinSize.hwnd, "This frame shows the icon in the small size just as it will look in the dock.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True
End Sub

Private Sub picSizePreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picSizePreview.hwnd, "This frame shows the icon in two sizes, as it looks in the dock (on the left) and is it will appear when fully enlarged during a zoom.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True
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

Private Sub picStylePreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picStylePreview.hwnd, "This panel shows a preview of the font selection - you can change the background of the preview to approximate how your font will look  on your desktop.", _
                  TTIconInfo, "Help on the Font Preview Pane.", , , , True
End Sub
Private Sub picThemeSample_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picThemeSample.hwnd, "This panel shows a portion of the dock with the current theme selected.", _
                  TTIconInfo, "Help on Theme Selection.", , , , True
End Sub

Private Sub picZoomSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picZoomSize.hwnd, "This frame shows the icon in the large size just as it looks when fully enlarged during a mouse-over zoom.", _
                  TTIconInfo, "Help on the Icon Zoom Preview.", , , , True

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

Private Sub sliAnimationInterval_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliAnimationInterval.hwnd, "The overall animation period in millisecs. 10ms is a good default but experiment with the value for your own system if the animation is not as smooth as you desire. The animation is achieved using GDI+ and is entirely CPU driven. You may see a benefit in Steamydock by changing this slider. This will have no effect on Rocketdock.", _
                  TTIconInfo, "Help on the Animation Interval.", , , , True
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

Private Sub sliBehaviourAutoHideDelay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliBehaviourAutoHideDelay.hwnd, "Determine the delay between the last usage of the dock and when it will auto-hide.", _
                  TTIconInfo, "Help on the AutoHide Delay Slider.", , , , True
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

Private Sub sliBehaviourAutoHideDuration_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliBehaviourAutoHideDuration.hwnd, "The speed at which the dock auto-hide animation will occur.", _
                  TTIconInfo, "Help on the AutoHide Duration Slider.", , , , True
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



Private Sub sliBehaviourPopUpDelay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliBehaviourPopUpDelay.hwnd, "The speed at which the dock auto-reveal animation will occur. This was previously called the Pop-up Delay in Rocketdock's settings screen.", _
                  TTIconInfo, "Help on the AutoReveal Duration Slider.", , , , True
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

Private Sub sliContinuousHide_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliContinuousHide.hwnd, "Determine the amount of time the dock will disappear when told to go away using F11 key.", _
                  TTIconInfo, "Help on the Continuous Hide Slider.", , , , True
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

Private Sub sliGenRunAppInterval_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliGenRunAppInterval.hwnd, "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency set here. The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds. ", _
                  TTIconInfo, "Help on the Running Application Timer.", , , , True
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

Private Sub sliIconsDuration_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsDuration.hwnd, "How long the effect is applied in milliseconds. ", _
                  TTIconInfo, "Help on the Icon Zoom Duration Slider", , , , True
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

Private Sub sliIconsOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsOpacity.hwnd, "The icons in the dock can be made transparent here.", _
                  TTIconInfo, "Help on the Icon Opacity Slider", , , , True

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

Private Sub sliIconsSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsSize.hwnd, "The size of all the icons in the dock prior to any zoom effect being applied. ", _
                  TTIconInfo, "Help on the Icons Size Slider", , , , True
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
            nTwips = intTwips / Screen.TwipsPerPixelX

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
            nTwips = intPixels * Screen.TwipsPerPixelX
            
            PixelsToTwips = nTwips

   On Error GoTo 0
   Exit Function

PixelsToTwips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PixelsToTwips of Form dockSettings"

End Function

Private Sub sliIconsZoom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsZoom.hwnd, "The maximum icon size after a zoom. ", _
                  TTIconInfo, "Help on the Icon Zoom Slider", , , , True

End Sub

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



Private Sub sliIconsZoomWidth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliIconsZoomWidth.hwnd, "How many icons to the left and right are also animated. Lower power machines will benefit from a lower setting. 4 is fine. ", _
                  TTIconInfo, "Help on the Icon Zoom Width Slider", , , , True


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

Private Sub sliPositionCentre_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliPositionCentre.hwnd, "You can align the dock so that it is centred or offset as you require.", _
                  TTIconInfo, "Help on the Dock Centre Position Slider ", , , , True
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

Private Sub sliPositionEdgeOffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliPositionEdgeOffset.hwnd, "Position from the bottom/top edge of the screen.", _
                  TTIconInfo, "Help on the Dock Position Edge Offset Slider ", , , , True
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

Private Sub sliStyleFontOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleFontOpacity.hwnd, "The font transparency can be changed here.", _
                  TTIconInfo, "Help on the Font Opacity Slider.", , , , True
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

Private Sub sliStyleOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleOpacity.hwnd, "This controls the transparency of the background theme.", _
                  TTIconInfo, "Help on the Opacity Slider.", , , , True
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

Private Sub sliStyleOutlineOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleOutlineOpacity.hwnd, "The label outline transparency, use the slider to change.", _
                  TTIconInfo, "Help on the Outline Opacity Slider.", , , , True
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
    'cmbDefaultDock.SelLength = 0
    cmbHidingKey.SelLength = 0
    cmbDefaultDock.SelLength = 0
    cmbBehaviourActivationFX.SelLength = 0
    cmbBehaviourSoundSelection.SelLength = 0
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
        checkAndKill NameProcess, False, False
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



Private Sub sliStyleShadowOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleShadowOpacity.hwnd, "The strength of the shadow can be altered here.", _
                  TTIconInfo, "Help on the Shadow Opacity Slider.", , , , True
End Sub

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



Private Sub sliStyleThemeSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip sliStyleThemeSize.hwnd, "This controls the size of the background theme. Only implemented on SteamyDock.", _
                  TTIconInfo, "Help on Theme Size.", , , , True
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
        
        'load the gear images
        picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1.jpg")
        picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3.jpg")

        rDSkinTheme = "dark"
    Else
        'MsgBox "Windows Alternate Theme detected"
        SysClr = GetSysColor(COLOR_BTNFACE)
        If SysClr = 13160660 Then
            Call setThemeShade(212, 208, 199)
            rDSkinTheme = "dark"
            
                    
            'load the gear images
            picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1.jpg")
            picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears1.jpg")

        Else ' 15790320
            Call setThemeShade(240, 240, 240)
            rDSkinTheme = "light"
        
            'load the gear images
            picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1Light.jpg")
            picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3Light.jpg")
            
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
        'load the gear images
        picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1.jpg")
        picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3.jpg")
        Call setThemeShade(212, 208, 199)
    Else
        'load the gear images
        picMultipleGears1.Picture = LoadPicture(App.Path & "\multipleGears1Light.jpg")
        picMultipleGears3.Picture = LoadPicture(App.Path & "\multipleGears3Light.jpg")
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
'        PutINISetting "Software\SteamyDock\DockSettings", "AlwaysAsk", rDAlwaysAsk, dockSettingsFile
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
    PutINISetting location, "RetainIcons", rDRetainIcons, settingsFile ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    
    PutINISetting location, "ManageWindows", rDManageWindows, settingsFile
    PutINISetting location, "DisableMinAnimation", rDDisableMinAnimation, settingsFile
    PutINISetting location, "ShowRunning", rDShowRunning, settingsFile
    PutINISetting location, "OpenRunning", rDOpenRunning, settingsFile
    PutINISetting location, "HoverFX", rDHoverFX, settingsFile
    PutINISetting location, "zOrderMode", rDzOrderMode, settingsFile
    PutINISetting location, "MouseActivate", rDMouseActivate, settingsFile
    PutINISetting location, "IconActivationFX", rDIconActivationFX, settingsFile
    PutINISetting location, "SoundSelection", rDSoundSelection, settingsFile
    
    PutINISetting location, "Monitor", rDMonitor, settingsFile
    PutINISetting location, "Side", rDSide, settingsFile
    PutINISetting location, "Offset", rDOffset, settingsFile
    PutINISetting location, "vOffset", rDvOffset, settingsFile
    PutINISetting location, "OptionsTabIndex", rDOptionsTabIndex, toolSettingsFile
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
'        chkGenAlwaysAsk.Value = Val(rDAlwaysAsk)
    End If

    'Rocketdock values also used by Steamydock
    
    ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    chkGenLock.Value = Val(rDLockIcons)
    chkRetainIcons.Value = Val(rDRetainIcons)

    chkGenOpen.Value = Val(rDOpenRunning)
    chkGenRun.Value = Val(rDShowRunning)
    chkGenMin.Value = Val(rDManageWindows)
    chkGenDisableAnim.Value = Val(rDDisableMinAnimation)

    If chkGenMin.Value = 0 Then
        chkGenDisableAnim.Enabled = False
    Else
        chkGenDisableAnim.Enabled = True
    End If
    
    If chkGenRun.Value = 0 Then
'        lblGenLabel(0).Enabled = False
'        lblGenLabel(1).Enabled = False
        sliGenRunAppInterval.Enabled = False
        lblGenLabel(2).Enabled = False
        lblGenRunAppIntervalCur.Enabled = False
    Else
'        lblGenLabel(0).Enabled = True
'        lblGenLabel(1).Enabled = True
        sliGenRunAppInterval.Enabled = True
        lblGenLabel(2).Enabled = True
        lblGenRunAppIntervalCur.Enabled = True
    End If
        
    ' Icons tab
    
    Call setZoomTypes ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected
    Call setBounceTypes
    Call setSoundSelectionDropDown
    
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
    
    ' if not then add the key combinations that are allowed for Steamydock
    
    Call setHidingKey
    
    ' .10 STARTS DAEB 01/02/2021 docksettings Remove some functionality not available to rocketdock
    If defaultDock = 1 Then
        picThemeSample.Enabled = True
        lblStyleLabel(2).Enabled = True
        
        sliStyleThemeSize.Enabled = True
        lblThemeSizeTextHigh.Enabled = True
        lblStyleSizeCurrent.Enabled = True
    Else
        picThemeSample.Enabled = False
        lblStyleLabel(2).Enabled = False

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

'    If defaultDock = 0 Then
'        cmbBehaviourActivationFX.AddItem "None", 0
'        cmbBehaviourActivationFX.AddItem "UberIcon Effects", 1
'        cmbBehaviourActivationFX.AddItem "Bounce", 2
'        'rDIconActivationFX = "2"
'
'    Else
        cmbBehaviourActivationFX.AddItem "None", 0
        cmbBehaviourActivationFX.AddItem "Bounce", 1
        cmbBehaviourActivationFX.AddItem "Miserable", 2
        'rDIconActivationFX = "1"
'    End If
    
    cmbBehaviourActivationFX.ListIndex = Val(rDIconActivationFX)
    

End Sub
Private Sub setSoundSelectionDropDown()

    cmbBehaviourSoundSelection.Clear

    cmbBehaviourSoundSelection.AddItem "None", 0
    cmbBehaviourSoundSelection.AddItem "Ting", 1
    cmbBehaviourSoundSelection.AddItem "Click", 2
    
    cmbBehaviourSoundSelection.ListIndex = Val(rDSoundSelection)
    

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

'    If defaultDock = 0 Then
'        cmbIconsHoverFX.AddItem "None", 0
'        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
'        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
'        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
'        rDHoverFX = "1"
'
'    Else
        cmbIconsHoverFX.AddItem "None", 0
        cmbIconsHoverFX.AddItem "Zoom: Bubble", 1
        cmbIconsHoverFX.AddItem "Zoom: Plateau", 2
        cmbIconsHoverFX.AddItem "Zoom: Flat", 3
        cmbIconsHoverFX.AddItem "Zoom: Bumpy", 4
'    End If
    
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


'---------------------------------------------------------------------------------------
' Procedure : setToolTips
' Author    : beededea
' Date      : 27/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setToolTips()
    On Error GoTo setToolTips_Error

    If chkToggleDialogs.Value = 0 Then
        Call DestroyToolTip ' destroys the current tooltip
        
        rDEnableBalloonTooltips = "0" ' this is the flag used to determine whether a new balloon tooltip is generated
        
        btnDefaults.ToolTipText = "Revert ALL settings to the defaults"
        chkToggleDialogs.ToolTipText = "When checked this toggle will display the information pop-ups and balloon tips "
        btnHelp.ToolTipText = "Click here to open tool's HTML help page in your browser"
        picBusy.ToolTipText = "The program is doing something..."
        btnClose.ToolTipText = "Exit this utility"
        btnApply.ToolTipText = "This will save your changes and restart the dock."
        lblText(0).ToolTipText = "General Configuration Options"
        picIcon(0).ToolTipText = "About the dock settings"
        picIcon(1).ToolTipText = "General Configuration Options"
        picIcon(2).ToolTipText = "Dock theme and text configuration"
        picIcon(3).ToolTipText = "Icon bounce and pop up effects"
        picIcon(4).ToolTipText = "Icon effects and quality"
        genChkShowIconSettings.ToolTipText = "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
        chkSplashStatus.ToolTipText = "Show Splash Screen on Start-up"
        optGeneralReadSettings.ToolTipText = "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
        optGeneralReadRegistry.ToolTipText = "Stores the configuration where Rocketdock stores it, in the Registry, increasingly incompatible with Windows new standards, causes some security problems and requires admin rights to operate."
        optGeneralReadConfig.ToolTipText = "This stores ALL configuration within the user data area retaining future compatibility in Windows. The trouble is, only SteamyDock can access it."
        
        sliGenRunAppInterval.ToolTipText = "The maximum time a basic VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppInterval2.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppInterval3.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenRunAppIntervalCur.ToolTipText = "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
        lblGenLabel(0).ToolTipText = "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
'        chkGenAlwaysAsk.ToolTipText = "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as above"
        btnGeneralRdFolder.ToolTipText = "Select the folder location of Rocketdock here"
        chkGenRun.ToolTipText = "After a short delay, small application indicators appear above the icon of a running program, this uses a little cpu every few seconds, frequency below"
        chkGenDisableAnim.ToolTipText = "If you dislike the minimise animation, click this"
        chkGenOpen.ToolTipText = "If you click on an icon that is already running then it can open it or fire up another instance"
        txtGeneralRdLocation.ToolTipText = "This is the extrapolated location of the currently selected dock. This is for information only."
        'cmbDefaultDock.ToolTipText = "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
        chkGenLock.ToolTipText = "This is an essential option that stops you accidentally deleting your dock icons, click it!"
        
        ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        chkRetainIcons.ToolTipText = "Dragging a program binary to the dock can take an automatically selected icon or you can retain the embedded icon."
        
        chkGenMin.ToolTipText = "This allows running applications to appear in the dock"
        chkGenWinStartup.ToolTipText = "This will cause the current dock to run when Windows starts"
        
'        optGeneralWriteSettings.ToolTipText = "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
        'optGeneralWriteRegistry.ToolTipText = "Stores the configuration where Rocketdock stores it, in the Registry, increasingly incompatible with Windows new standards, causes some security problems and requires admin rights to operate."
        optGeneralWriteConfig.ToolTipText = "This stores ALL configuration within the user data area retaining future compatibility in Windows. The trouble is, only SteamyDock can access it."

        'lblChkSplashStartup.ToolTipText = "Show Splash Screen on Start-up"
        'lblChkAlwaysConfirm.ToolTipText = "If both docks are installed then it will ask you which you would prefer to configure and operate, otherwise it will use the default dock as above"
        'lblChkOpenRunning.ToolTipText = "If you click on an icon that is already running then it can open it or fire up another instance"
        'lblRdLocation.ToolTipText = "This is the extrapolated location of the RocketDock Program, you can alter it yourself  if you have another copy of Rocketdock installed elsewhere - currently not operational, defaults to Rocketdock"
        
        lblGenLabel(2).ToolTipText = "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
        cmbHidingKey.ToolTipText = "This is the key sequence that is used to hide or restore Steamydock"
        sliContinuousHide.ToolTipText = "Determine how long Steamydock will disappear when told to hide using F11"
        
        cmbBehaviourActivationFX.ToolTipText = "Set which type of animation you want to occur on an icon mouseover. Note SteamyDock will NOT support the Ubericon effects where Rocketdock does."
        chkBehaviourAutoHide.ToolTipText = "You can determine whether the dock will auto-hide or not"
        sliBehaviourAutoHideDuration.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        sliBehaviourPopUpDelay.ToolTipText = "The dock mouse-over delay period"
        lblBehaviourPopUpDelayMsCurrrent.ToolTipText = "The dock mouse-over delay period"
        sliBehaviourAutoHideDelay.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        chkBehaviourMouseActivate.ToolTipText = "Essential functionality for the dock - pops up when  given focus"
        lblBehaviourLabel(0).ToolTipText = "which type of animation you want to occur on an icon mouseover. Note SteamyDock will NOT support the Ubericon effects but Rocketdock will."
        'lblBehaviourLabel(1).ToolTipText = "You can determine whether the dock will auto-hide or not"
        lblBehaviourLabel(2).ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblBehaviourLabel(3).ToolTipText = "The dock mouse-over delay period"
        lblBehaviourLabel(4).ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblBehaviourLabel(5).ToolTipText = "Determine how long Steamydock will disappear when told to hide for the next few minutes"
        lblBehaviourLabel(6).ToolTipText = "This is the key sequence that is used to hide or restore Steamydock"
        lblBehaviourLabel(7).ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblBehaviourLabel(8).ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblBehaviourLabel(9).ToolTipText = "The dock mouse-over delay period"
        lblBehaviourLabel(10).ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblBehaviourLabel(11).ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        lblBehaviourLabel(12).ToolTipText = "This panel is really a eulogy to Rocketdock plus a few buttons taking you to useful locations and providing additional data"
        lblBehaviourLabel(13).ToolTipText = "This is an essential option that stops you accidentally deleting your dock icons, ensure it is ticked!"
        lblBehaviourLabel(14).ToolTipText = "The original icons may be low quality."
        lblBehaviourLabel(15).ToolTipText = "Select a sound to play when an icon in the dock is clicked."
        
        cmbBehaviourSoundSelection.ToolTipText = "Select a sound to play when an icon in the dock is clicked."
        
        lblContinuousHideMsCurrent.ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        lblContinuousHideMsHigh.ToolTipText = "Determine how long Steamydock will disappear when told to go away"
        fraAutoHideType.ToolTipText = "The type of auto-hide, fade, instant or a slide like Rocketdock"
        lblAutoHideDurationMsHigh.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblAutoHideDurationMsCurrent.ToolTipText = "The speed at which the dock auto-hide animation will occur"
        lblAutoRevealDurationMsHigh.ToolTipText = "The dock mouse-over delay period"
        lblAutoHideDelayMsHigh.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        lblAutoHideDelayMsCurrent.ToolTipText = "Determine the delay between the last usage of the dock and when it will auto-hide"
        sliAnimationInterval.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second. The optimal value is probably 10ms."
        lblAnimationIntervalMsLow.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblAnimationIntervalMsHigh.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        lblAnimationIntervalMsCurrent.ToolTipText = "Certain CPUs may operate better with a different animation interval, 1ms = 1,000 animations per second"
        btnDonate.ToolTipText = "Opens a browser window and sends you to our donate page on Amazon"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs used by Rocketdock"
        btnFacebook.ToolTipText = "This will link you to the Rocket/Steamy dock users Group"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        chkLabelBackgrounds.ToolTipText = "You can toggle the icon label background on/off here"
        picThemeSample.ToolTipText = "An example preview of the chosen theme."
        sliStyleShadowOpacity.ToolTipText = "The strength of the shadow can be altered here"
        sliStyleOutlineOpacity.ToolTipText = "The label outline transparency, use the slider to change"
        sliStyleFontOpacity.ToolTipText = "The font transparency can be changed here"
        lblStyleFontOpacityCurrent.ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(3).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleOutlineOpacityCurrent.ToolTipText = "The label outline transparency, use the slider to change"
        Label35.ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleShadowOpacityCurrent.ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(9).ToolTipText = "The strength of the shadow can be altered here"
        picStylePreview.ToolTipText = "A preview of the font selection - you can change the background of the preview to approximate how your font will look  on your desktop"
        btnStyleOutline.ToolTipText = "The colour of the outline, click the button to change"
        btnStyleShadow.ToolTipText = "The colour of the shadow, click the button to change"
        btnStyleFont.ToolTipText = "The font used in the labels, click the button to change"
        chkStyleDisable.ToolTipText = "You can toggle the icon labels on/off here"
        cmbStyleTheme.ToolTipText = "The dock background theme can be selected here"
        sliStyleOpacity.ToolTipText = "The theme background opacity is here"
        sliStyleThemeSize.ToolTipText = "The theme background overall size is here"
        
        lblChkLabelBackgrounds.ToolTipText = "You can toggle the icon label background on/off here"
        lblStyleLabel(2).ToolTipText = "The theme background overall size is here"
        lblStyleFontFontShadowColor.ToolTipText = "The colour of the shadow, click the button to change"
        lblStyleFontOutlineTest.ToolTipText = "The colour of the outline, click the button to change"
        lblStyleFontFontShadowTest.ToolTipText = "The colour of the shadow, click the button to change"
        lblStyleFontName.ToolTipText = "The font used in the labels, click the button to change"
        
        lblStyleLabel(0).ToolTipText = "The dock background theme can be selected here"
        lblStyleLabel(1).ToolTipText = "The theme background opacity is set here"
        lblStyleLabel(2).ToolTipText = "The theme background overall size is set here"
        lblStyleLabel(3).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(4).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(5).ToolTipText = "The label outline transparency, use the slider to change"
        lblStyleLabel(6).ToolTipText = "The theme background opacity is set here"
        lblStyleLabel(7).ToolTipText = "The theme background overall size is set here"
        lblStyleLabel(8).ToolTipText = "The font transparency can be changed here"
        lblStyleLabel(9).ToolTipText = "The strength of the shadow can be altered here"
        lblStyleLabel(10).ToolTipText = "The label outline transparency, use the slider to change"
        

        fmeMain(0).ToolTipText = "This panel controls the positioning of the whole dock"
        cmbPositionLayering.ToolTipText = "Should the dock appear on top of other windows or underneath?"
        cmbPositionMonitor.ToolTipText = "Here you can determine upon which monitor the dock will appear"
        cmbPositionScreen.ToolTipText = "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
        sliPositionEdgeOffset.ToolTipText = "Position from the bottom/top edge of the screen"
        sliPositionCentre.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label33.ToolTipText = "Should the dock appear on top of other windows or underneath?"
        lblPositionMonitor.ToolTipText = "Here you can determine upon which monitor the dock will appear"
        Label32.ToolTipText = "Place the dock at your preferred location. Steamydock only supports top and bottom positions"
        Label31.ToolTipText = "You can align the dock so that it is centred or offas you require"
        lblPositionCentrePercCurrent.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label29.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label28.ToolTipText = "You can align the dock so that it is centred or offas you require"
        Label27.ToolTipText = "Position from the bottom/top edge of the screen"
        lblPositionEdgeOffsetPxCurrent.ToolTipText = "Position from the bottom/top edge of the screen"
        Label25.ToolTipText = "Position from the bottom/top edge of the screen"
        Label24.ToolTipText = "Position from the bottom/top edge of the screen"
        picMinSize.ToolTipText = "The icon size in the dock when static"
        picZoomSize.ToolTipText = "The maximum icon size of an animated icon"
        Label1.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label9.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label13.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        sliIconsDuration.ToolTipText = "How long the effect is applied"
        sliIconsZoomWidth.ToolTipText = "How many icons to the left and right are also animated"
        lblCharacteristicsLabel(11).ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(12).ToolTipText = "How long the effect is applied"
        lblIconsDurationMsCurrent.ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(6).ToolTipText = "How long the effect is applied"
        lblCharacteristicsLabel(10).ToolTipText = "How many icons to the left and right are also animated"
        Label14.ToolTipText = "How many icons to the left and right are also animated"
        lblIconsZoomWidth.ToolTipText = "How many icons to the left and right are also animated"
        lblCharacteristicsLabel(5).ToolTipText = "How many icons to the left and right are also animated"
        chkIconsZoomOpaque.ToolTipText = "Should the zoom be opaque too?"
        cmbIconsQuality.ToolTipText = "Lower power single/dual core machines will benefit from the lower quality setting but in reality, current machines can run with high quality enabled and suffer no degradation whatsoever."
        sliIconsZoom.ToolTipText = "The maximum icon size after a zoom"
        sliIconsSize.ToolTipText = "The size of each icon in the dock before any effect is applied"
        sliIconsOpacity.ToolTipText = "The icons in the dock can be made transparent here"
        cmbIconsHoverFX.ToolTipText = "The zoom effect to apply"
        lblCharacteristicsLabel(2).ToolTipText = "The zoom effect to apply"
        lblCharacteristicsLabel(0).ToolTipText = "Lower power machines will benefit from the lower quality setting"
        lblCharacteristicsLabel(1).ToolTipText = "The icons in the dock can be made transparent here"
        lblCharacteristicsLabel(3).ToolTipText = "The size of each icon in the dock before any effect is applied"
        lblIconsOpacity.ToolTipText = "The icons in the dock can be made transparent here"
        lblIconsSize.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        Label3.ToolTipText = "The icons in the dock can be made transparent here"
        lblCharacteristicsLabel(7).ToolTipText = "The icons in the dock can be made transparent here"
        Label5.ToolTipText = "The size of all the icons in the dock before any effect is applied"
        lblCharacteristicsLabel(8).ToolTipText = "The size of all the icons in the dock before any effect is applied"
        lblCharacteristicsLabel(4).ToolTipText = "The maximum icon size after a zoom"
        lblIconsZoom.ToolTipText = "The maximum icon size after a zoom"
        lblIconsZoomSizeMax.ToolTipText = "The maximum icon size after a zoom"
        lblCharacteristicsLabel(9).ToolTipText = "The maximum icon size after a zoom"
        picHiddenPicture.ToolTipText = "The icon size in the dock"
        Label26.ToolTipText = "Show Splash Screen on Start-up"
    
    
    Else
    
        rDEnableBalloonTooltips = "1" ' this is the flag used to determine whether a new balloon tooltip is generated

        btnDefaults.ToolTipText = ""
        chkToggleDialogs.ToolTipText = ""
        btnHelp.ToolTipText = ""
        picBusy.ToolTipText = ""
        btnClose.ToolTipText = ""
        btnApply.ToolTipText = ""
        lblText(0).ToolTipText = ""
        picIcon(0).ToolTipText = ""
        picIcon(1).ToolTipText = ""
        picIcon(2).ToolTipText = ""
        picIcon(3).ToolTipText = ""
        picIcon(4).ToolTipText = ""
        genChkShowIconSettings.ToolTipText = ""
        chkSplashStatus.ToolTipText = ""
        optGeneralReadSettings.ToolTipText = ""
        optGeneralReadRegistry.ToolTipText = ""
        optGeneralReadConfig.ToolTipText = ""
        
        sliGenRunAppInterval.ToolTipText = ""
        lblGenRunAppInterval2.ToolTipText = ""
        lblGenRunAppInterval3.ToolTipText = ""
        lblGenRunAppIntervalCur.ToolTipText = ""
        lblGenLabel(0).ToolTipText = ""
'        chkGenAlwaysAsk.ToolTipText = ""
        btnGeneralRdFolder.ToolTipText = ""
        chkGenRun.ToolTipText = ""
        chkGenDisableAnim.ToolTipText = ""
        chkGenOpen.ToolTipText = ""
        txtGeneralRdLocation.ToolTipText = ""
        'cmbDefaultDock.ToolTipText = ""
        chkGenLock.ToolTipText = ""
        chkRetainIcons.ToolTipText = ""         ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
        chkGenMin.ToolTipText = ""
        chkGenWinStartup.ToolTipText = ""
        
'        optGeneralWriteSettings.ToolTipText = ""
        'optGeneralWriteRegistry.ToolTipText = ""
        optGeneralWriteConfig.ToolTipText = ""

        'lblChkSplashStartup.ToolTipText = ""
        'lblChkAlwaysConfirm.ToolTipText = ""
        'lblChkOpenRunning.ToolTipText = ""
        'lblRdLocation.ToolTipText = ""
        
        'lblSplitter.ToolTipText = ""
        'cmbHidingKey.ToolTipText = ""
        sliContinuousHide.ToolTipText = ""
        lblBehaviourLabel(11).ToolTipText = ""
        lblBehaviourLabel(5).ToolTipText = ""
        lblContinuousHideMsCurrent.ToolTipText = ""
        lblContinuousHideMsHigh.ToolTipText = ""
        fraAutoHideType.ToolTipText = ""
        chkBehaviourAutoHide.ToolTipText = ""
        'lblBehaviourLabel(1).ToolTipText = ""
        'lblBehaviourLabel(0).ToolTipText = ""
        sliBehaviourAutoHideDuration.ToolTipText = ""
        lblBehaviourLabel(8).ToolTipText = ""
        lblAutoHideDurationMsHigh.ToolTipText = ""
        lblAutoHideDurationMsCurrent.ToolTipText = ""
        lblBehaviourLabel(2).ToolTipText = ""
        sliBehaviourPopUpDelay.ToolTipText = ""
        lblBehaviourLabel(9).ToolTipText = ""
        lblBehaviourLabel(3).ToolTipText = ""
        lblBehaviourPopUpDelayMsCurrrent.ToolTipText = ""
        lblAutoRevealDurationMsHigh.ToolTipText = ""
        sliBehaviourAutoHideDelay.ToolTipText = ""
        lblBehaviourLabel(10).ToolTipText = ""
        lblAutoHideDelayMsHigh.ToolTipText = ""
        lblAutoHideDelayMsCurrent.ToolTipText = ""
        lblBehaviourLabel(4).ToolTipText = ""
        chkBehaviourMouseActivate.ToolTipText = ""
        sliAnimationInterval.ToolTipText = ""
        lblAnimationIntervalMsLow.ToolTipText = ""
        lblAnimationIntervalMsHigh.ToolTipText = ""
        lblAnimationIntervalMsCurrent.ToolTipText = ""
        lblBehaviourLabel(7).ToolTipText = ""
        'lblBehaviourLabel(6).ToolTipText = ""
        lblBehaviourLabel(12).ToolTipText = ""
        btnDonate.ToolTipText = ""
        btnUpdate.ToolTipText = ""
        btnFacebook.ToolTipText = ""
        btnAboutDebugInfo.ToolTipText = ""
        chkLabelBackgrounds.ToolTipText = ""
        picThemeSample.ToolTipText = ""
        sliStyleShadowOpacity.ToolTipText = ""
        sliStyleOutlineOpacity.ToolTipText = ""
        sliStyleFontOpacity.ToolTipText = ""

        lblStyleFontOpacityCurrent.ToolTipText = ""
        lblStyleOutlineOpacityCurrent.ToolTipText = ""
        lblStyleShadowOpacityCurrent.ToolTipText = ""
        
        lblStyleLabel(0).ToolTipText = ""
        lblStyleLabel(1).ToolTipText = ""
        lblStyleLabel(2).ToolTipText = ""
        lblStyleLabel(3).ToolTipText = ""
        lblStyleLabel(4).ToolTipText = ""
        lblStyleLabel(5).ToolTipText = ""
        lblStyleLabel(6).ToolTipText = ""
        lblStyleLabel(7).ToolTipText = ""
        lblStyleLabel(8).ToolTipText = ""
        lblStyleLabel(9).ToolTipText = ""
        'lblStyleLabel(10).ToolTipText = ""
        lblBehaviourLabel(13).ToolTipText = ""
        lblBehaviourLabel(14).ToolTipText = ""
        lblBehaviourLabel(15).ToolTipText = ""
        
        cmbBehaviourSoundSelection.ToolTipText = ""
        
        picStylePreview.ToolTipText = ""
        btnStyleOutline.ToolTipText = ""
        btnStyleShadow.ToolTipText = ""
        btnStyleFont.ToolTipText = ""
        chkStyleDisable.ToolTipText = ""
        cmbStyleTheme.ToolTipText = ""
        sliStyleOpacity.ToolTipText = ""
        sliStyleThemeSize.ToolTipText = ""
        
        lblChkLabelBackgrounds.ToolTipText = ""
        lblStyleFontFontShadowColor.ToolTipText = ""
        lblStyleFontOutlineTest.ToolTipText = ""
        lblStyleFontFontShadowTest.ToolTipText = ""
        lblStyleFontName.ToolTipText = ""
        fmeMain(0).ToolTipText = ""
        cmbPositionLayering.ToolTipText = ""
        cmbPositionMonitor.ToolTipText = ""
        cmbPositionScreen.ToolTipText = ""
        sliPositionEdgeOffset.ToolTipText = ""
        sliPositionCentre.ToolTipText = ""
        Label33.ToolTipText = ""
        lblPositionMonitor.ToolTipText = ""
        Label32.ToolTipText = ""
        Label31.ToolTipText = ""
        lblPositionCentrePercCurrent.ToolTipText = ""
        Label29.ToolTipText = ""
        Label28.ToolTipText = ""
        Label27.ToolTipText = ""
        lblPositionEdgeOffsetPxCurrent.ToolTipText = ""
        Label25.ToolTipText = ""
        Label24.ToolTipText = ""
        picMinSize.ToolTipText = ""
        picZoomSize.ToolTipText = ""
        Label1.ToolTipText = ""
        Label9.ToolTipText = ""
        Label13.ToolTipText = ""
        sliIconsDuration.ToolTipText = ""
        sliIconsZoomWidth.ToolTipText = ""
        lblCharacteristicsLabel(11).ToolTipText = ""
        lblCharacteristicsLabel(12).ToolTipText = ""
        lblIconsDurationMsCurrent.ToolTipText = ""
        lblCharacteristicsLabel(10).ToolTipText = ""
        Label14.ToolTipText = ""
        lblIconsZoomWidth.ToolTipText = ""
        chkIconsZoomOpaque.ToolTipText = ""
        cmbIconsQuality.ToolTipText = ""
        sliIconsZoom.ToolTipText = ""
        sliIconsSize.ToolTipText = ""
        sliIconsOpacity.ToolTipText = ""
        cmbIconsHoverFX.ToolTipText = ""
        lblIconsOpacity.ToolTipText = ""
        lblIconsSize.ToolTipText = ""
        Label3.ToolTipText = ""
        Label5.ToolTipText = ""
        lblIconsZoom.ToolTipText = ""
        lblIconsZoomSizeMax.ToolTipText = ""
        lblCharacteristicsLabel(0).ToolTipText = ""
        'lblCharacteristicsLabel(1).ToolTipText = ""
        lblCharacteristicsLabel(2).ToolTipText = ""
        lblCharacteristicsLabel(3).ToolTipText = ""
        lblCharacteristicsLabel(4).ToolTipText = ""
        lblCharacteristicsLabel(5).ToolTipText = ""
        lblCharacteristicsLabel(6).ToolTipText = ""
        lblCharacteristicsLabel(7).ToolTipText = ""
        lblCharacteristicsLabel(8).ToolTipText = ""
        lblCharacteristicsLabel(9).ToolTipText = ""
        picHiddenPicture.ToolTipText = ""
        Label26.ToolTipText = ""
    
        'cmbBehaviourActivationFX.ToolTipText = ""
    
    End If

    On Error GoTo 0
    Exit Sub

setToolTips_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setToolTips of Form dockSettings"
            Resume Next
          End If
    End With
End Sub

' .21 DAEB 07/09/2022 docksettings moved hiding key definitions to own subroutine
'---------------------------------------------------------------------------------------
' Procedure : setHidingKey
' Author    : beededea
' Date      : 07/09/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setHidingKey()

    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD STARTS
    On Error GoTo setHidingKey_Error

    If defaultDock = 1 Then
        cmbHidingKey.Locked = False
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
        cmbHidingKey.Locked = True
        cmbHidingKey.Clear
        cmbHidingKey.AddItem "Control+Alt+R"
        cmbHidingKey.Text = "Control+Alt+R" ' .08 DAEB 01/02/2021 docksettings Added support for the default hiding key plus others for the two dock
    End If
    ' .15 DAEB 18/02/2021 docksettings set the default key settings for RD and SD ends
    

    On Error GoTo 0
    Exit Sub

setHidingKey_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingKey of Form dockSettings"
            Resume Next
          End If
    End With
    
End Sub



Private Sub txtGeneralRdLocation_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtGeneralRdLocation.hwnd, "This is the extrapolated location of the currently selected dock. This is for information only.", _
                  TTIconInfo, "Help on the Running Application Indicators.", , , , True
End Sub
Private Sub positionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    dockSettingsXPos = dockSettings.Left
    dockSettingsYPos = dockSettings.Top
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormXPos", dockSettingsXPos, toolSettingsFile
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormYPos", dockSettingsYPos, toolSettingsFile
End Sub

Private Sub mnuBringToCentre_click()

    dockSettings.Top = Screen.Height / 2 - dockSettings.Height / 2
    dockSettings.Left = screenWidthTwips / 2 - dockSettings.Width / 2
End Sub
