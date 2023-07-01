VERSION 5.00
Object = "{FB95F7DD-5143-4C75-88F9-A53515A946D7}#2.0#0"; "CCRTreeView.ocx"
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Object = "{FA5FEA4A-5ED5-4004-A509-2DABC30D42A7}#1.0#0"; "CCRImageList.ocx"
Begin VB.Form rDIconConfigForm 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SteamyDock Icon Settings VB6"
   ClientHeight    =   10065
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "rdIconConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   10230
   Begin VB.PictureBox picTemporaryStore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   1185
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   110
      Tag             =   "Do not delete"
      ToolTipText     =   "This is the currently selected icon scaled to 64 x 64 for the dragIcon to reside prior to conversion"
      Top             =   8835
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   450
      Top             =   5760
   End
   Begin VB.Frame Frame 
      Caption         =   "Configuration"
      Height          =   600
      Index           =   0
      Left            =   2100
      TabIndex        =   44
      Top             =   4515
      Visible         =   0   'False
      Width           =   1665
      Begin VB.CheckBox chkBiLinear 
         Caption         =   "Quality Sizing"
         Height          =   240
         Left            =   90
         TabIndex        =   45
         ToolTipText     =   "Stretch Quality Option"
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Drag && Drop, Copy && Paste too.  Unicode Compatible"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   45
         TabIndex        =   46
         ToolTipText     =   "To Paste: Click on display box and press Ctrl+V"
         Top             =   5865
         Width           =   3840
      End
   End
   Begin VB.Timer registryTimer 
      Interval        =   2500
      Left            =   450
      Top             =   6150
   End
   Begin VB.PictureBox picRdThumbFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   90
      ScaleHeight     =   705
      ScaleWidth      =   9660
      TabIndex        =   36
      Top             =   4545
      Visible         =   0   'False
      Width           =   9660
      Begin VB.HScrollBar rdMapHScroll 
         Height          =   120
         Left            =   45
         Max             =   100
         TabIndex        =   40
         Top             =   540
         Visible         =   0   'False
         Width           =   9630
      End
      Begin VB.CommandButton btnMapNext 
         Height          =   450
         Left            =   9180
         Picture         =   "rdIconConfig.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Scroll the RD map to the right"
         Top             =   45
         Width           =   450
      End
      Begin VB.CommandButton btnMapPrev 
         Height          =   450
         Left            =   45
         Picture         =   "rdIconConfig.frx":0ACF
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Scroll the RD map to the left"
         Top             =   45
         Width           =   435
      End
      Begin VB.PictureBox picCover 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   -45
         ScaleHeight     =   555
         ScaleWidth      =   570
         TabIndex        =   38
         Top             =   0
         Width           =   570
      End
      Begin VB.PictureBox back 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   9150
         ScaleHeight     =   555
         ScaleWidth      =   600
         TabIndex        =   39
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox picRdMap 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "rdIconConfig.frx":101C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   500
         Index           =   0
         Left            =   540
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   51
         Top             =   15
         Width           =   500
      End
   End
   Begin VB.PictureBox btnArrowUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9840
      Picture         =   "rdIconConfig.frx":20E6
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   35
      ToolTipText     =   "Hide the map"
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton rdMapRefresh 
      Height          =   270
      Left            =   9885
      Picture         =   "rdIconConfig.frx":2442
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Refresh the icon map"
      Top             =   4770
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame FrameFolders 
      Caption         =   "Folders"
      Height          =   4500
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "The current list of known icon folders"
      Top             =   15
      Width           =   4005
      Begin VB.PictureBox btnSettingsUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   990
         Picture         =   "rdIconConfig.frx":284B
         ScaleHeight     =   180
         ScaleWidth      =   270
         TabIndex        =   102
         ToolTipText     =   "Hide the registry form showing where details are being read from and saved to."
         Top             =   3990
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Frame fraConfigSource 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   1800
         TabIndex        =   64
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
         Begin VB.ComboBox cmbDefaultDock 
            Height          =   330
            ItemData        =   "rdIconConfig.frx":2BA7
            Left            =   570
            List            =   "rdIconConfig.frx":2BB1
            Locked          =   -1  'True
            TabIndex        =   65
            Text            =   "RocketDock"
            ToolTipText     =   "Indicates the default dock, Rocketdock or SteamyDock. Cannot be changed here but only in the dock settings utility."
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lblDefaultDock 
            Caption         =   "Dock"
            Height          =   225
            Left            =   0
            TabIndex        =   66
            Top             =   45
            Width           =   555
         End
      End
      Begin VB.PictureBox btnSettingsDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   930
         Picture         =   "rdIconConfig.frx":2BCD
         ScaleHeight     =   180
         ScaleWidth      =   375
         TabIndex        =   61
         ToolTipText     =   "Show where the details are being read from and saved to."
         Top             =   3990
         Visible         =   0   'False
         Width           =   375
      End
      Begin CCRTreeView.TreeView folderTreeView 
         Height          =   3210
         Left            =   135
         TabIndex        =   55
         ToolTipText     =   "These are the icon folders available to Rocketdock"
         Top             =   630
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5662
         VisualTheme     =   1
         LineStyle       =   1
         LabelEdit       =   1
         ShowTips        =   -1  'True
         Indentation     =   38
      End
      Begin VB.TextBox textCurrentFolder 
         Height          =   330
         Left            =   135
         TabIndex        =   21
         Text            =   "textCurrentFolder"
         ToolTipText     =   "The selected folder path"
         Top             =   225
         Width           =   3735
      End
      Begin VB.CommandButton btnRemoveFolder 
         Caption         =   "-"
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "This button can remove a custom folder from the treeview above"
         Top             =   3990
         Width           =   360
      End
      Begin VB.CommandButton btnAddFolder 
         Caption         =   "+"
         Height          =   345
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Select a target folder to add to the treeview list above"
         Top             =   3990
         Width           =   360
      End
   End
   Begin VB.PictureBox btnArrowDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9735
      Picture         =   "rdIconConfig.frx":2E77
      ScaleHeight     =   180
      ScaleWidth      =   375
      TabIndex        =   34
      ToolTipText     =   "Show the Rocketdock Map"
      Top             =   4485
      Width           =   375
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   4800
      Left            =   4230
      TabIndex        =   0
      ToolTipText     =   "The Icon Properties Window"
      Top             =   4530
      Width           =   5895
      Begin VB.Frame fraLblAppToTerminate 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         TabIndex        =   122
         Top             =   4320
         Width           =   4035
         Begin VB.TextBox txtAppToTerminate 
            Height          =   345
            Left            =   1230
            TabIndex        =   125
            ToolTipText     =   "Any program that must be terminated prior to the main program initiation will be shown here"
            Top             =   0
            Width           =   2205
         End
         Begin VB.CommandButton btnAppToTerminate 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3510
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Press to select a program to terminate prior to the main program initiation"
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblAppToTerminate 
            Caption         =   "Terminate App :"
            Height          =   255
            Left            =   0
            TabIndex        =   123
            ToolTipText     =   "If you want to run a second program after the program initiation, select it here"
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.Frame fraOptionButtons 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   4125
         TabIndex        =   117
         Top             =   3840
         Width           =   1725
         Begin VB.OptionButton optRunSecondAppAfterward 
            Enabled         =   0   'False
            Height          =   240
            Left            =   45
            TabIndex        =   119
            ToolTipText     =   "Determines whether the secondary program is run after the main program has started"
            Top             =   270
            Width           =   270
         End
         Begin VB.OptionButton optRunSecondAppBeforehand 
            Caption         =   "Run Beforehand"
            Enabled         =   0   'False
            Height          =   225
            Left            =   45
            TabIndex        =   118
            ToolTipText     =   "Determines whether the secondary program is run before the main program"
            Top             =   15
            Value           =   -1  'True
            Width           =   240
         End
         Begin VB.Label lblRunSecondAppAfterward 
            Caption         =   "Run Afterward"
            Enabled         =   0   'False
            Height          =   165
            Left            =   330
            TabIndex        =   121
            Tag             =   "this extra label compensates for the poor quality greying out on label captions"
            ToolTipText     =   "Determines whether the secondary program is run after the main program has started"
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblRunSecondAppBeforehand 
            Caption         =   "Run Beforehand"
            Enabled         =   0   'False
            Height          =   165
            Left            =   330
            TabIndex        =   120
            Tag             =   "this extra label compensates for the poor quality greying out on label captions"
            ToolTipText     =   "Determines whether the secondary program is run before the main program"
            Top             =   15
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkDisabled 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2925
         TabIndex        =   115
         ToolTipText     =   "If you want extra options to appear when you right click on an icon, enable this checkbox"
         Top             =   3045
         Width           =   240
      End
      Begin VB.PictureBox picMoreConfigUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3570
         Picture         =   "rdIconConfig.frx":3121
         ScaleHeight     =   180
         ScaleWidth      =   270
         TabIndex        =   103
         ToolTipText     =   "Hides the extra configuration section"
         Top             =   3360
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Frame fraLblRdIconNumber 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   4365
         TabIndex        =   100
         Top             =   1875
         Width           =   1320
         Begin VB.Label lblRdIconNumber 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   45
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1140
            Left            =   390
            TabIndex        =   101
            ToolTipText     =   "This is dock icon number one."
            Top             =   -75
            Width           =   480
         End
      End
      Begin VB.Frame fraLblSecondApp 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   285
         TabIndex        =   96
         Top             =   3885
         Width           =   4005
         Begin VB.CommandButton btnSecondApp 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Press to select a second program to run after the main program initiation"
            Top             =   30
            Width           =   360
         End
         Begin VB.TextBox txtSecondApp 
            Height          =   345
            Left            =   1110
            TabIndex        =   97
            ToolTipText     =   "Any second program to run after the main program initiation will be shown here"
            Top             =   30
            Width           =   2205
         End
         Begin VB.Label lblSecondApp 
            Caption         =   "Second App :"
            Height          =   225
            Left            =   15
            TabIndex        =   99
            ToolTipText     =   "If you want to run a second program after the program initiation, select it here"
            Top             =   75
            Width           =   1560
         End
      End
      Begin VB.Frame fraLblQuickLaunch 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   45
         TabIndex        =   90
         Top             =   3615
         Width           =   1590
         Begin VB.CheckBox chkQuickLaunch 
            Caption         =   "Quick Launch :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1350
            TabIndex        =   91
            ToolTipText     =   "Launch an application before the bounce has completed"
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblQuickLaunch 
            Caption         =   "Quick Launch :"
            Height          =   225
            Left            =   165
            TabIndex        =   92
            ToolTipText     =   "Launch an application before the bounce has completed"
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblConfirmDialogAfter 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1695
         TabIndex        =   87
         Top             =   3345
         Width           =   1680
         Begin VB.CheckBox chkConfirmDialogAfter 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1237
            TabIndex        =   88
            ToolTipText     =   "Shows Confirmation Dialog after the command has run."
            Top             =   -15
            Width           =   240
         End
         Begin VB.Label lblConfirmDialogAfter 
            Caption         =   "Confirm After :"
            Height          =   225
            Left            =   0
            TabIndex        =   89
            ToolTipText     =   "Shows Confirmation Dialog after the command has run."
            Top             =   -15
            Width           =   1110
         End
      End
      Begin VB.Frame fraLblConfirmDialog 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   84
         Top             =   3315
         Width           =   1575
         Begin VB.CheckBox chkConfirmDialog 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1365
            TabIndex        =   86
            ToolTipText     =   "Adds a Confirmation Dialog prior to the command running allowing you to say yes or no at runtime"
            Top             =   15
            Width           =   180
         End
         Begin VB.Label lblConfirmDialog 
            Caption         =   "Confirm Prior :"
            Height          =   225
            Left            =   240
            TabIndex        =   85
            ToolTipText     =   "Adds a Confirmation Dialog prior to the command running allowing you to say yes or no at runtime"
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblPopUp 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   45
         TabIndex        =   82
         Top             =   3045
         Width           =   1320
         Begin VB.Label lblRunElevated 
            Caption         =   "Run Elevated :"
            Height          =   225
            Left            =   210
            TabIndex        =   83
            ToolTipText     =   "If you want extra options to appear when you right click on an icon, enable this checkbox"
            Top             =   -15
            Width           =   1200
         End
      End
      Begin VB.Frame fraLblOpenRunning 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   90
         TabIndex        =   80
         Top             =   2655
         Width           =   1245
         Begin VB.Label lblOpenRunning 
            Caption         =   "New instance:"
            Height          =   225
            Left            =   135
            TabIndex        =   81
            Top             =   15
            Width           =   1515
         End
      End
      Begin VB.Frame fraLblRun 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   75
         TabIndex        =   78
         Top             =   2265
         Width           =   1260
         Begin VB.Label lblRun 
            Caption         =   "Window State:"
            Height          =   225
            Left            =   135
            TabIndex        =   79
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblArgument 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   90
         TabIndex        =   76
         Top             =   1890
         Width           =   1260
         Begin VB.Label lblArgument 
            Caption         =   "Arguments:"
            Height          =   225
            Left            =   345
            TabIndex        =   77
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblStartIn 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   74
         Top             =   1485
         Width           =   1245
         Begin VB.Label lblStartIn 
            Caption         =   "Start in:"
            Height          =   225
            Left            =   600
            TabIndex        =   75
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblTarget 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   60
         TabIndex        =   72
         Tag             =   "This frame is only here to support a mouseover for the label"
         Top             =   1080
         Width           =   1290
         Begin VB.Label lblTarget 
            Caption         =   "Target:"
            Height          =   225
            Left            =   720
            TabIndex        =   73
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.Frame fraLblCurrentIcon 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   105
         TabIndex        =   70
         Tag             =   "This frame is only here to support a mouseover for the label"
         Top             =   735
         Width           =   1260
         Begin VB.Label lblCurrentIcon 
            Caption         =   "Current Icon:"
            Height          =   225
            Left            =   255
            TabIndex        =   71
            Top             =   15
            Width           =   1215
         End
      End
      Begin VB.PictureBox picMoreConfigDown 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3510
         Picture         =   "rdIconConfig.frx":347D
         ScaleHeight     =   180
         ScaleWidth      =   375
         TabIndex        =   62
         ToolTipText     =   "Shows extra configuration items"
         Top             =   3360
         Width           =   375
      End
      Begin VB.CommandButton btnIconSelect 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Press to select an icon manually."
         Top             =   690
         Width           =   345
      End
      Begin VB.TextBox txtDbg02 
         Height          =   315
         Left            =   5370
         TabIndex        =   59
         Text            =   "txtDbg01"
         Top             =   2370
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtDbg01 
         Height          =   315
         Left            =   5370
         TabIndex        =   58
         Text            =   "txtDbg01"
         Top             =   1980
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.PictureBox picBusy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   3660
         Picture         =   "rdIconConfig.frx":3727
         ScaleHeight     =   795
         ScaleWidth      =   825
         TabIndex        =   54
         ToolTipText     =   "The program is doing something..."
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton btnSet 
         Caption         =   "&Set"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sets the icon characteristics but you will need to press the save and restart button to make it 'fix' on the running dock."
         Top             =   3195
         Width           =   1470
      End
      Begin VB.TextBox txtCurrentIcon 
         Height          =   345
         Left            =   1395
         TabIndex        =   18
         Text            =   "txtCurrentIcon"
         ToolTipText     =   "Double click on an image above to set the current icon"
         Top             =   690
         Width           =   3915
      End
      Begin VB.CommandButton btnSelectStart 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Select a start folder"
         Top             =   1470
         Width           =   345
      End
      Begin VB.CommandButton btnTarget 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Press to select a target file (or right click for a folder)"
         Top             =   1080
         Width           =   345
      End
      Begin VB.CheckBox chkRunElevated 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1395
         TabIndex        =   7
         ToolTipText     =   "If you want extra options to appear when you right click on an icon, enable this checkbox"
         Top             =   3045
         Width           =   180
      End
      Begin VB.ComboBox cmbOpenRunning 
         Height          =   330
         ItemData        =   "rdIconConfig.frx":41A2
         Left            =   1395
         List            =   "rdIconConfig.frx":41AF
         TabIndex        =   6
         Text            =   "Use Global Setting"
         ToolTipText     =   "Choose what to do if the chosen app is already running"
         Top             =   2640
         Width           =   2145
      End
      Begin VB.ComboBox cmbRunState 
         Height          =   330
         ItemData        =   "rdIconConfig.frx":41D6
         Left            =   1395
         List            =   "rdIconConfig.frx":41E3
         TabIndex        =   5
         Text            =   "Normal"
         ToolTipText     =   "Window mode for the program to operate within"
         Top             =   2250
         Width           =   2145
      End
      Begin VB.TextBox txtArguments 
         Height          =   345
         Left            =   1395
         TabIndex        =   4
         ToolTipText     =   "Add any additional arguments that the target file operation requires, eg. -s -t 00 -f "
         Top             =   1860
         Width           =   2130
      End
      Begin VB.TextBox txtStartIn 
         Height          =   345
         Left            =   1395
         TabIndex        =   3
         ToolTipText     =   "If the operation needs to be performed in a particular folder select it here"
         Top             =   1470
         Width           =   3915
      End
      Begin VB.TextBox txtTarget 
         Height          =   345
         Left            =   1395
         TabIndex        =   2
         ToolTipText     =   "The target you wish to run, a file or a folder"
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox txtLabelName 
         Height          =   345
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "The name of the icon as it appears on the dock"
         Top             =   300
         Width           =   4305
      End
      Begin VB.Frame frmLblAutoHideDock 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1680
         TabIndex        =   93
         Top             =   3615
         Width           =   1455
         Begin VB.CheckBox chkAutoHideDock 
            Caption         =   "Auto Hide Dock :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1245
            TabIndex        =   94
            ToolTipText     =   "Automatically hides the dock for the default hiding period when the program is initiated"
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblchkAutoHideDock 
            Caption         =   "Auto Hide :"
            Height          =   225
            Left            =   285
            TabIndex        =   95
            ToolTipText     =   "Automatically hides the dock for the default hiding period when the program is initiated"
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.Label lblDisabled 
         Caption         =   "Icon Disabled :"
         Height          =   225
         Left            =   1710
         TabIndex        =   116
         ToolTipText     =   "Shows Confirmation Dialog after the command has run."
         Top             =   3045
         Width           =   1110
      End
      Begin VB.Label Label2 
         Height          =   285
         Left            =   3675
         TabIndex        =   109
         Top             =   3015
         Width           =   2040
      End
      Begin VB.Label txtName 
         Caption         =   "Name:"
         Height          =   225
         Left            =   825
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame framePreview 
      Caption         =   "Preview"
      Height          =   4500
      Left            =   135
      TabIndex        =   47
      ToolTipText     =   "The Preview Pane"
      Top             =   4515
      Width           =   4000
      Begin VB.Frame fraSizeSlider 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   45
         TabIndex        =   111
         Top             =   3870
         Width           =   3810
         Begin CCRSlider.Slider sliPreviewSize 
            Height          =   300
            Left            =   0
            TabIndex        =   112
            ToolTipText     =   "Icon Size"
            Top             =   180
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   529
            Min             =   1
            Max             =   5
            Value           =   4
            LargeChange     =   1
            TickStyle       =   1
            SelStart        =   4
         End
         Begin VB.Label lblFileInfo 
            Caption         =   "File Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   90
            TabIndex        =   114
            Top             =   0
            Width           =   1950
         End
         Begin VB.Label lblWidthHeight 
            Caption         =   "width and height"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2100
            TabIndex        =   113
            Top             =   0
            Width           =   1785
         End
      End
      Begin VB.CommandButton btnPrev 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "select the previous icon"
         Top             =   240
         Width           =   195
      End
      Begin VB.CommandButton btnNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "select the next icon"
         Top             =   240
         Width           =   195
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   285
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   204
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   230
         TabIndex        =   48
         ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
         Top             =   675
         Width           =   3450
         Begin VB.Timer settingsTimer 
            Enabled         =   0   'False
            Interval        =   2500
            Left            =   2145
            Top             =   270
         End
         Begin VB.Timer positionTimer 
            Interval        =   3000
            Left            =   1695
            Top             =   180
         End
         Begin VB.Timer idleTimer 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   1260
            Top             =   180
         End
         Begin CCRImageList.ImageList imlThumbnailCache 
            Left            =   1860
            Tag             =   "Krool's image list used to store images as a cache"
            Top             =   1845
            _ExtentX        =   1005
            _ExtentY        =   1005
            InitListImages  =   "rdIconConfig.frx":4205
         End
         Begin CCRImageList.ImageList imlDragIconConverter 
            Left            =   630
            Tag             =   "Krool's imageList (with HIcon bug)"
            Top             =   1860
            _ExtentX        =   1005
            _ExtentY        =   1005
            ImageWidth      =   16
            ImageHeight     =   16
            UseMaskColor    =   0   'False
            MaskColor       =   16777215
            InitListImages  =   "rdIconConfig.frx":4225
         End
         Begin VB.Timer rdMapDragTimer 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   840
            Top             =   180
         End
         Begin VB.Timer thumbnailDragTimer 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   435
            Top             =   180
         End
         Begin VB.Timer busyTimer 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   30
            Top             =   180
         End
         Begin VB.Label lblBlankText 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Blank"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   45
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   1140
            Left            =   600
            TabIndex        =   63
            ToolTipText     =   "This is Rocketdock icon number one."
            Top             =   960
            Visible         =   0   'False
            Width           =   2220
         End
      End
   End
   Begin VB.Frame frameButtons 
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   210
      TabIndex        =   25
      Top             =   8490
      Width           =   10080
      Begin VB.PictureBox picHideConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "rdIconConfig.frx":4245
         ScaleHeight     =   180
         ScaleWidth      =   270
         TabIndex        =   104
         ToolTipText     =   "Hide the map"
         Top             =   945
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton btnGenerate 
         Caption         =   "&Auto Generate Dock"
         Height          =   360
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Makes a whole NEW dock - use with care!"
         Top             =   765
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton btnBackup 
         Caption         =   "&Backup"
         Height          =   345
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Backup or restore using a version of bkpSettings.ini"
         Top             =   405
         Width           =   1485
      End
      Begin VB.CommandButton btnSaveRestart 
         Caption         =   "&Save && Restart"
         Height          =   345
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "A save and restart of the dock is required when any icon changes have been made"
         Top             =   780
         Width           =   1485
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "&Help"
         Height          =   345
         Left            =   8430
         MousePointer    =   14  'Arrow and Question
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Help on this utility"
         Top             =   405
         Width           =   1470
      End
      Begin VB.CheckBox chkToggleDialogs 
         Caption         =   "Display Info.Dialogs"
         Height          =   225
         Left            =   4020
         TabIndex        =   27
         ToolTipText     =   "When checked this toggle will display the information pop-ups and balloon tips "
         Top             =   450
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CommandButton btnDefaultIcon 
         Caption         =   "Default Icon"
         Height          =   330
         Left            =   1830
         TabIndex        =   26
         ToolTipText     =   "Not implemented yet"
         Top             =   435
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   " &Cancel"
         Height          =   345
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel the current operation and close the window"
         Top             =   780
         Width           =   1470
      End
      Begin VB.CommandButton btnClose 
         Caption         =   " &Close"
         Height          =   345
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Cancel the current operation and close the window"
         Top             =   780
         Width           =   1470
      End
   End
   Begin VB.Frame frameIcons 
      Caption         =   "Icons"
      Height          =   4500
      Left            =   4230
      TabIndex        =   12
      ToolTipText     =   "Thumbnail or File Viewer Window"
      Top             =   15
      Width           =   5895
      Begin VB.Frame fraIconType 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   67
         Top             =   3960
         Width           =   3765
         Begin VB.ComboBox comboIconTypesFilter 
            Height          =   330
            ItemData        =   "rdIconConfig.frx":45A1
            Left            =   390
            List            =   "rdIconConfig.frx":45B7
            TabIndex        =   69
            Text            =   "All Normal Icons"
            ToolTipText     =   "Filter icon types to display"
            Top             =   0
            Width           =   2790
         End
         Begin VB.CommandButton btnKillIcon 
            Height          =   255
            Left            =   0
            Picture         =   "rdIconConfig.frx":4622
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Delete the currently selected icon file above. Use wisely!"
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.Frame frmNoFilesFound 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1635
         TabIndex        =   56
         Top             =   1860
         Visible         =   0   'False
         Width           =   2220
         Begin VB.Label lblNoFilesFound 
            Caption         =   "No files found"
            Height          =   285
            Left            =   525
            TabIndex        =   57
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "+"
         Height          =   270
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Set the current selected icon into the dock (double-click on the icon)"
         Top             =   240
         Width           =   270
      End
      Begin VB.CommandButton btnRefresh 
         Height          =   270
         Left            =   5085
         Picture         =   "rdIconConfig.frx":484F
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Refresh the Icon List"
         Top             =   225
         Width           =   210
      End
      Begin VB.TextBox textCurrIconPath 
         Height          =   330
         Left            =   1215
         TabIndex        =   16
         Text            =   "textCurrIconPath"
         ToolTipText     =   "Shows the selected icon file name"
         Top             =   210
         Width           =   3510
      End
      Begin VB.CommandButton btnGetMore 
         Caption         =   "Get More"
         Height          =   345
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Click to install more icons"
         Top             =   3975
         Width           =   1710
      End
      Begin VB.CommandButton btnFileListView 
         Height          =   270
         Left            =   5355
         Picture         =   "rdIconConfig.frx":4C58
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "View as a file listing"
         Top             =   240
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton btnThumbnailView 
         Height          =   270
         Left            =   5355
         Picture         =   "rdIconConfig.frx":5034
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "View as thumbnails"
         Top             =   240
         Width           =   285
      End
      Begin VB.PictureBox picFrameThumbs 
         BackColor       =   &H00FFFFFF&
         Height          =   3240
         Left            =   120
         ScaleHeight     =   3180
         ScaleWidth      =   5520
         TabIndex        =   22
         ToolTipText     =   "Double-click an icon to set it into the dock"
         Top             =   600
         Width           =   5580
         Begin VB.Frame fraThumbLabel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Index           =   0
            Left            =   45
            TabIndex        =   107
            Top             =   840
            Width           =   1185
            Begin VB.Label lblThumbName 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   108
               Top             =   -30
               Width           =   1000
            End
         End
         Begin VB.PictureBox picFraPicThumbIcon 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1080
            Index           =   0
            Left            =   105
            ScaleHeight     =   1080
            ScaleWidth      =   1095
            TabIndex        =   105
            Top             =   105
            Width           =   1095
            Begin VB.PictureBox picThumbIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1000
               Index           =   0
               Left            =   -15
               ScaleHeight     =   67
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   67
               TabIndex        =   106
               ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
               Top             =   -30
               Width           =   1000
            End
         End
         Begin VB.VScrollBar vScrollThumbs 
            CausesValidation=   0   'False
            Height          =   3210
            LargeChange     =   12
            Left            =   5265
            SmallChange     =   4
            TabIndex        =   23
            Top             =   -30
            Width           =   255
         End
      End
      Begin VB.FileListBox filesIconList 
         Height          =   3240
         Left            =   105
         Pattern         =   "*.jpg"
         TabIndex        =   15
         ToolTipText     =   "Select an icon, double-click to set"
         Top             =   600
         Width           =   5580
      End
      Begin VB.Label lblIconName 
         Caption         =   "Icon Name:"
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnWorking 
      Caption         =   "&Working"
      Height          =   510
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3900
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Menu mnuTrgtMenu 
      Caption         =   "Target Menu"
      Begin VB.Menu mnuTrgtSeparator 
         Caption         =   "target = Separator"
      End
      Begin VB.Menu mnuTrgtFolder 
         Caption         =   "target = Folder"
      End
      Begin VB.Menu mnuTrgtMyComputer 
         Caption         =   "target = My Computer"
      End
      Begin VB.Menu mnuTrgtMyDocuments 
         Caption         =   "target = My Documents"
      End
      Begin VB.Menu mnuTrgtMyMusic 
         Caption         =   "target = My Music"
      End
      Begin VB.Menu mnuTrgtMyPictures 
         Caption         =   "target = My Pictures"
      End
      Begin VB.Menu mnuTrgtMyVideos 
         Caption         =   "target = My Videos"
      End
      Begin VB.Menu mnuTrgtShutdown 
         Caption         =   "target = Shutdown"
      End
      Begin VB.Menu mnuTrgtRestart 
         Caption         =   "target = Restart"
      End
      Begin VB.Menu mnuTrgtSleep 
         Caption         =   "target = Sleep"
      End
      Begin VB.Menu mnuTrgtLock 
         Caption         =   "target = Lock Workstation"
      End
      Begin VB.Menu mnuTrgtLog 
         Caption         =   "target = Log Off Workstation"
      End
      Begin VB.Menu mnuTrgtNetwork 
         Caption         =   "target = Network"
      End
      Begin VB.Menu mnuTrgtWorkgroup 
         Caption         =   "target = Workgroup"
      End
      Begin VB.Menu mnuTrgtPrinters 
         Caption         =   "target = Printers"
      End
      Begin VB.Menu mnuTrgtTask 
         Caption         =   "target = Task Manager"
      End
      Begin VB.Menu mnuTrgtControl 
         Caption         =   "target = Control Panel"
      End
      Begin VB.Menu mnuTrgtProgramFiles 
         Caption         =   "target = Program Files Folder"
      End
      Begin VB.Menu mnuTrgtPrograms 
         Caption         =   "target = Programs and Features"
      End
      Begin VB.Menu mnuTrgtAdministrativeTools 
         Caption         =   "target = Administrative Tools"
         Begin VB.Menu mnuTrgtCompMgmt 
            Caption         =   "target = Computer Management"
         End
         Begin VB.Menu mnuTrgtDiscMgmt 
            Caption         =   "target = Disc Management"
         End
         Begin VB.Menu mnuTrgtDevMgmt 
            Caption         =   "target = Device Management"
         End
         Begin VB.Menu mnuTrgtEventViewer 
            Caption         =   "target = Event Viewer"
         End
         Begin VB.Menu mnuTrgtPerfMon 
            Caption         =   "target = Performance Monitor"
         End
         Begin VB.Menu mnuTrgtServices 
            Caption         =   "target = Services Management"
         End
         Begin VB.Menu mnuTrgtTaskSched 
            Caption         =   "target = Task Scheduler"
         End
      End
      Begin VB.Menu mnuTrgtRecycle 
         Caption         =   "target = Recycle Bin"
      End
      Begin VB.Menu mnuTrgtDock 
         Caption         =   "target = Dock Settings"
      End
      Begin VB.Menu mnuTrgtClearCache 
         Caption         =   "target = Clear Cache"
      End
      Begin VB.Menu mnuTrgtEnhanced 
         Caption         =   "target = Enhanced Icon Settings"
      End
      Begin VB.Menu mnuTrgtRocketdock 
         Caption         =   "target = Rocketdock Quit"
      End
      Begin VB.Menu mnuTrgtDocklet 
         Caption         =   "target = Docklet"
      End
   End
   Begin VB.Menu rdMapMenu 
      Caption         =   "The Map Menu"
      Begin VB.Menu menuRun 
         Caption         =   "Run the Item (test)"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mnuClone 
         Caption         =   "Clone Item"
      End
      Begin VB.Menu menuAddMenu 
         Caption         =   "Add Dock Item"
         Begin VB.Menu menuAddBlank 
            Caption         =   "Add Blank Item"
         End
         Begin VB.Menu mnuAddProgram 
            Caption         =   "Add a Program, DLL or EXE"
         End
         Begin VB.Menu mnuAddSeparator 
            Caption         =   "Add a Separator"
         End
         Begin VB.Menu mnuAddFolder 
            Caption         =   "Add Folder"
         End
         Begin VB.Menu mnuAddMyComputer 
            Caption         =   "Add My Computer"
         End
         Begin VB.Menu mnuAddMyDocuments 
            Caption         =   "Add My Documents"
         End
         Begin VB.Menu mnuAddMyMusic 
            Caption         =   "Add My Music"
         End
         Begin VB.Menu mnuAddMyPictures 
            Caption         =   "Add My Pictures"
         End
         Begin VB.Menu mnuAddMyVideos 
            Caption         =   "Add My Videos "
         End
         Begin VB.Menu mnuAddShutdown 
            Caption         =   "Add Shutdown"
         End
         Begin VB.Menu mnuAddRestart 
            Caption         =   "Add Restart"
         End
         Begin VB.Menu mnuAddSleep 
            Caption         =   "Add Sleep"
         End
         Begin VB.Menu mnuAddLock 
            Caption         =   "Add Lock Workstation"
         End
         Begin VB.Menu mnuAddLog 
            Caption         =   "Add Log Off Workstation"
         End
         Begin VB.Menu mnuAddWorkgroup 
            Caption         =   "Add Workgroup"
         End
         Begin VB.Menu mnuAddNetwork 
            Caption         =   "Add Network"
         End
         Begin VB.Menu mnuAddPrinters 
            Caption         =   "Add Printers"
         End
         Begin VB.Menu mnuAddTask 
            Caption         =   "Add Task Manager"
         End
         Begin VB.Menu mnuAddControl 
            Caption         =   "Add Control Panel"
         End
         Begin VB.Menu mnuAddProgramFiles 
            Caption         =   "Add Program Files Folder"
         End
         Begin VB.Menu mnuAddPrograms 
            Caption         =   "Add Programs / Features"
         End
         Begin VB.Menu mnuAddAdministrativeTools 
            Caption         =   "Add Administrative Tools"
            Begin VB.Menu mnuAddCompMgmt 
               Caption         =   "Add Computer Management"
            End
            Begin VB.Menu mnuAddDiscMgmt 
               Caption         =   "Add Disc Management"
            End
            Begin VB.Menu mnuAddDevMgmt 
               Caption         =   "Add Device Management"
            End
            Begin VB.Menu mnuAddEventViewer 
               Caption         =   "Add Event Viewer"
            End
            Begin VB.Menu mnuAddPerfMon 
               Caption         =   "Add Performance Monitor"
            End
            Begin VB.Menu mnuAddServices 
               Caption         =   "Add Services Management"
            End
            Begin VB.Menu mnuAddTaskSched 
               Caption         =   "Add Task Scheduler"
            End
         End
         Begin VB.Menu mnuAddRecycle 
            Caption         =   "Add Recycle Bin"
         End
         Begin VB.Menu mnuAddDock 
            Caption         =   "Add Dock Settings"
         End
         Begin VB.Menu mnuClearCache 
            Caption         =   "Add Clear Cache"
         End
         Begin VB.Menu mnuAddEnhanced 
            Caption         =   "Add Enhanced Icon Settings"
         End
         Begin VB.Menu mnuAddQuit 
            Caption         =   "Add Rocketdock Quit"
         End
         Begin VB.Menu mnuAddDocklet 
            Caption         =   "Add a Docklet"
         End
      End
      Begin VB.Menu menuLeft 
         Caption         =   "Move item to the left"
      End
      Begin VB.Menu menuRight 
         Caption         =   "Move item to the right"
      End
   End
   Begin VB.Menu mnupopmenu 
      Caption         =   "The main menu"
      Begin VB.Menu mnuAddPreviewIcon 
         Caption         =   "Add this icon at this position in the map"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About this utility"
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Reveal folder for this Icon Set"
         Visible         =   0   'False
      End
      Begin VB.Menu blank10 
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
         Begin VB.Menu mnuHelp 
            Caption         =   "Utility Help"
            Index           =   4
         End
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
            Caption         =   "Chat about Rocketdock functionality on Facebook"
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
      Begin VB.Menu mnuRocketDock 
         Caption         =   "Set RocketDock Location"
         Visible         =   0   'False
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
   Begin VB.Menu thumbmenu 
      Caption         =   "Thumb Menu"
      Visible         =   0   'False
      Begin VB.Menu menuAddToDock 
         Caption         =   "Add to dock"
      End
      Begin VB.Menu menuSmallerIcons 
         Caption         =   "small icons with text"
      End
      Begin VB.Menu menuLargerThumbs 
         Caption         =   "larger icons (no text)"
      End
   End
End
Attribute VB_Name = "rDIconConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Changes:

' .01 DAEB 26/10/2020 rDIconConfigForm.frm rocket1 Moved function isRunning from mdlMain (sD mdlMain.bas) to a shared common.bas module so that more than just one program can utilise it
' .02 DAEB 26/10/2020 rDIconConfigForm.frm    Created new sub readInterimAndWriteConfig to allow the save to be called more than once on a btnSaveRestart_Click
' .03 DAEB 17/11/2020 rDIconConfigForm.frm    Replaced the confirmation dialog with an automatic save when moving from one icon to another using the right/left icon buttons
' .04 DAEB 17/11/2020 rDIconConfigForm.frm    Replaced all occurrences of rocket1.exe with iconsettings.exe
' .05 DAEB 17/11/2020 rDIconConfigForm.frm Added the missing code to read/write the current theme to the tool's own settings file
' .06 DAEB 31/01/2021 rDIconConfigForm.frm Added new checkbox to determine if a post initiation dialog should appear
' .07 DAEB 01/02/2021 rDIconConfigForm.frm Modified the parameter passed to isRunning to include the full path, otherwise it does not correlate with the found processes' folder
' .08 DAEB 02/02/2021 rDIconConfigForm.frm Added menu option to clear the cache
' .09 DAEB 07/02/2021 rDIconConfigForm.frm use the fullprocess variable without adding path again - duh!
' .10 DAEB 07/02/2021 rDIconConfigForm.frm removed unused vars
' .11 DAEB 26/10/2020 rDIconConfigForm.frm No longer pops up the question if the dialog boxes are suppressed.
' .12 DAEB 07/02/2021 rDIconConfigForm.frm added as part of busy timer functionality
' .13 DAEB 27/02/2021 rdIConConfigForm.frm Moved to a subroutine for clarity
' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
' .15 DAEB 01/03/2021 rDIConConfigForm.frm added confirmation dialog prior to running the test command
' .16 DAEB 01/03/2021 rDIConConfigForm.frm added new function to allow confirmation dialogue subsequent to running the test command
' .17 DAEB 01/03/2021 rDIConConfigForm.frm moved to variable initialisation
' .18 DAEB 01/03/2021 rDIConConfigForm.frm is now a property set at design time
' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
' .20 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Documents" utility dock entry
' .21 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Music" utility dock entry
' .22 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Pictures" utility dock entry
' .23 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Videos" utility dock entry
' .24 DAEB 07/03/2021 rDIConConfigForm.frm Add a form refresh after the menu has gone away to prevent messy control image leftovers
' .25 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Documents" target utility dock entry
' .26 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Music" target utility dock entry
' .27 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Pictures" target utility dock entry
' .28 DAEB 07/03/2021 rDIConConfigForm.frm Added menu option to add a "my Videos" target utility dock entry
' .29 DAEB 14/03/2021 rDIConConfigForm.frm change to focus the icon map on the icon pre-selected
' .30 DAEB 10/04/2021 rDIConConfigForm.frm separate the initial reading of the tool's settings file from the changing of the tool's own font
' .31 DAEB 10/04/2021 rDIConConfigForm.frm initialise the value - rather important
' .32 DAEB 11/04/2021 rDIConConfigForm.frm changed all occurrences of txtTarget.Text to thisCommand to attain more compatibility with runcommand
' .33 DAEB 01/05/2021 rDIConConfigForm.frm Added double click to copy target folder above to start in field
' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
' .35 DAEB 20/04/2021 rdIconConfig.frm Added new function to identify an icon to assign to the entry
' .36 DAEB 20/04/2021 rdIconConfig.frm Add a final check that the chosen image file actually exists
' .37 DAEB 05/05/2021 rdIconConfig.frm Added the new form to the changeFont tool
' .38 DAEB 03/03/2021 rdIconConfig.frm Removed the individual references to a Windows class ID
' .39 DAEB 03/03/2021 rdIconConfig.frm check whether the prefix is present that indicates a Windows class ID is present
' .40 DAEB 09/05/2021 rdIconConfig.frm turned into a function as it returns a value
' .41 DAEB 09/05/2021 rdIconConfig.frm fix copying the dock settings file for backups
' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
' .43 DAEB 16/04/2022 rdIconConfig.frm increase the whole form height and move the bottom buttons set down
' .44 DAEB 16/04/2022 rdIconConfig.frm one menu option is not applicable in SD, adding a docklet
' .45 DAEB 16/04/2022 rdIconConfig.frm Added the word Blank to the middle of the icon display picbox, sized it
' .46 DAEB 16/04/2022 rdIconConfig.frm Made the word Blank visible
' .47 DAEB 16/04/2022 rdIconConfig.frm Added StartRecordNumber
' .48 DAEB 20/04/2022 rDIConConfig.frm All tooltips move from IDE into code to allow them to disabled at will
' .49 DAEB 24/04/2022 rDIConConfig.frm Added balloon tooltips to all controls, using frame holders for the labels
' .50 DAEB 24/04/2022 rDIConConfig.frm When a new icon collection group is selected, the first icon filename is displayed
' .51 DAEB 24/04/2022 rDIConConfig.frm Icon preview needs to be forced to refresh after a click elsewhere and a return to the same icon thumbnail.
'                                           - when the app starts, the path of the first icon is not displayed in the filename box.
' .52 DAEB 24/04/2022 rDIConConfig.frm Added up button to the two down buttons, theme them and add another at the bottom left.
' .53 DAEB 25/04/2022 rDIConConfig.frm Reformatted the code to place ico files within an outer picbox frame to stop them jumping whilst being redrawn.
' .54 DAEB 25/04/2022 rDIConConfig.frm Added rDThumbImageSize saved variable to allow the tool to open the thumbnail explorer in small or large mode.
' .55 DAEB 25/04/2022 rDIConConfig.frm Fixed bug where scrolling to the end does not quite get to the end.
' .56 DAEB 25/04/2022 rDIConConfig.frm 1st run of the thumbnail view window is done using the old method and it comes out incorrectly.
' .57 DAEB 25/04/2022 rDIConConfig.frm blank.png should be blank and not ?
' .58 DAEB 25/04/2022 rDIConConfig.frm Second click on a thumbnail should be blue.
' .59 DAEB 01/05/2022 rDIConConfig.frm Added manual drag and drop functionality.
' .60 DAEB 01/05/2022 rDIConConfig.frm Add each image from the thumbnail icon picbox array to an imageList control so that we can assign a dragIcon.
' .61 DAEB 01/05/2022 rDIConConfig.frm Added highlighting to the rdIconMap during Drag and drop.
' .62 DAEB 04/05/2022 rDIConConfig.frm Change the icon image in the map to that chosen during a manual selection.
' .63 DAEB 04/05/2022 rDIConConfig.frm Moved the thumbnail left click functionality to a separate routine.
' .64 DAEB 04/05/2022 rDIConConfig.frm Moved the fileList left click functionality to a separate routine.
' .65 DAEB 04/05/2022 rDIConConfig.frm Use the underlying control index rather than that stored in the array.
' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picDragIcon) to be used to populate the dragIcon.
' .67 DAEB 04/05/2022 rDIConConfig.frm Drag and drop from the filelist to the rdmap
' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
' .70 DAEB 16/05/2022 rDIConConfig.frm Read the chkToggleDialogs value from a file and save the value for next time
' .71 DAEB 16/05/2022 rDIConConfig.frm Move the reading of recent settings into the main read configuration procedure
' .72 DAEB 16/05/2022 rDIConConfig.frm Validate the settings read from the settings file
' .73 DAEB 16/05/2022 rDIConConfig.frm Add ability to drag and drop icons in the map to an alternate position
' .74 DAEB 22/05/2022 rDIConConfig.frm MsgboxA replacement that can be placed on top of the form instead as the middle of the screen.
' .75 DAEB 22/05/2022 rDIConConfig.frm The dropdown disclose function is calculating the positions incorrectly when the map is toggled hidden/shown.
' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS
' .77 DAEB 28/05/2022 rDIConConfig.frm Balloon tooltip on the icon name text box
' .78 DAEB 28/05/2022 rDIConConfig.frm Dragging a blank icon within the map should show a drag image a .lnk? Possibly a white box with a thin black boundary.
' .78 DAEB 28/05/2022 rDIConConfig.frm We should only fill the temporary store when this routine has been called due to a click on the map
' .79 DAEB 28/05/2022 rDIConConfig.frm new parameter to determine when to populate the dragicon
' .80 DAEB 28/05/2022 rDIConConfig.frm Change to adding the .picture to workaround the bug in Krool's imageList failing to convert to an HIcon.
' .81 DAEB 28/05/2022 rDIConConfig.frm Added check to visibility of control before running the dragIcon code.
' .82 DAEB 02/06/2022 rDIConConfig.frm Added check for moving right or left beyond the end of the RDMap.
' .83 DAEB 03/06/2022 rDIConConfig.frm Display the icon we just moved by dragging, one by one rather than the whole map
' .84 DAEB 05/06/2022 rDIConConfig.frm Additional use of an imagelist control as a cache of already-read thumbnail icons to speed access to preceding thumbnails.
' .85 DAEB 06/06/2022 rDIConConfig.frm Second app button should open the dialog box in the program files folder
' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
' .87 DAEB 06/06/2022 rDIConConfig.frm Add OLE drag and drop of applications directly to the map using code from SteamyDock
' .88 DAEB 06/06/2022 rDIConConfig.frm Added checks for the existence of .jpg files used for themeing.
' .89 DAEB 13/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public
' .90 DAEB 14/06/2022 rDIConConfig.frm Moved jpegs to \resources folder.
' .91 DAEB 25/06/2022 rDIConConfig.frm Deleting an icon from the icon thumbnail display causes a cache imageList error.
' .92 DAEB 26/06/2022 formSoftwareList.frm The auto generation of a dock pulling the start menu links and the registry (undelete) section.
' .93 DAEB 26/06/2022 formSoftwareList.frm generate dock - overwrite dock routine untested - test and repair
' .94 DAEB 26/06/2022 rDIConConfig.frm Backup and restore - fix the problem with dock entries being zeroed after a restore.
' .95 DAEB 26/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public.
' .96 DAEB 09/11/2022 rDIConConfig.frm If the target text box is just a folder and not a full file path then a click on the select button should select also select a folder and not a file
' .97 DAEB 09/11/2022 rDIConConfig.frm For all target text box swap the IME right click menu for the target selection menu.
' .98 DAEB 09/11/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.
' .99 DAEB 09/11/2022 rDIConConfig.frm With the round borders of Win 11, there is insufficient space from the frame to the border, Windows cuts it off arbitrarily. Extend.
'.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
'.101 DAEB 09/11/2022 rDIConConfig.frm Add the restart option.
'.102 DAEB 08/12/2022 rdIconConfig.frm icon settings responds to %systemroot% environment variables during testing
' when dropping down the additional button options, stretch the image box downward too?
' we need to fill the space below the image preview.
'       For win 11 bottom cut off - need to add another 100 twips
' added option buttons for running the second app beforehand/afterwards
' added new field and selection button for choosing an application to terminate
' added initialisation, reading/writing parameters to handle the beforehand/afterwards saving and setting
' added initialisation, reading/writing parameters to handle the application to terminate
' added balloon tooltip handling to new controls
' modified code to enlarge the dropdown area and pull it back as required.
' modify help to document the application to terminate and run second app beforehand/afterwards.
' take the X/Y position and store it, when restarting, set it as per FCW.
' menu option to move the utility to the centre of the main monitor
' code added to find whether the utility is off screen - works for monitor one.
' rDIconConfigForm - all routines are now byVal or byRef
' rDIconConfigForm - all routines have their local variables initialised
' adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to monitor/resolution changes
' checks the date and time of the dock settings file and reloads the map after re-reading the icon details
' fixed the prev/next button msgbox, message box now shows the question icon and title bar.
' add themeing to the dock generation form taken from the main utility.
' pulled the prev/next button code together so it shares the same code
' Messagebox msgBoxA module to save the context of the message to allow specific msgboxes to go away
' pulled the map prev/next button code together so it shares the same code
' test for the folderTreeView.DropHighlight.Key and checking that (folderTreeView.SelectedItem Is Nothing) to determine which item in the folder tree to open.

' Current Task:
' =============

' remove persistentDebug.exe and replace with logging to a file as per FCW.
' test for the height of the window using the titlebar size to properly size the main form as per the Pz Earth VB6 widget prefs that size incorrectly on Win 10/11.

' Status:
' =======
'
'   Generally the tool is complete barring some bugs to resolve and new features I would like to implement.
'
' Tasks:
'
'   test running with a blank tool settings file
'
'   test running with a blank dock settings file
'
'   Add subclass (?) scrolling to the main scrollbar using the latter method found in this thread:
'       https://www.vbforums.com/showthread.php?898786-Easy-amp-ingenious-mousewheel-scrolling
'
'   reload the icon preview image and text details when the docksettings file changes, read the current icon from a new lastIconChanged field.
'
'   add the settings timer functionality to the dockSettings tool - could be instant with no msgbox?
'
'   create an interim migration tool from rocketdock or reassess the docksettings tool's capability to read the settings file as a one-off?
'
'   flag - if you are making changes now and another to determine if you have made any changes that will be lost of the map is refreshed.
'       to be tested in the settingsTimer
'
'   add to credits
'       Procedure : adjustFormPositionToCorrectMonitor
'       Author    : Hypetia from TekTips https://www.tek-tips.com/userinfo.cfm?member=Hypetia
'
'   Krool's CCRImageList.ocx component to replace the MS imageList, we are using this already - WIP
'       OCX built successfully after OLEGUIDS.TLB problem, not available and incorrect version.
'       Dropped the ocx onto the component toolbar manually
'       Dropped the new imagelist onto the form
'       Using the new image list and populating it at run time by loading images from an existing picbox, the image comes out pure black.
'       We need to use Krool's latest code from his panoply of controls, replace the current code, rebuild OCX and re-test.
'       Then raise a problem report on the CCR thread. DONE.
'       change to adding the .picture to workaround the bug in Krool's imageList failing to convert to an HIcon.
'       Workaround providing a white drag box around the image when using .picture instead of .image.
'       Await a permanent bugfix from Krool
'       Test and implement new version of Krool's imageList control with new bugfix
'       When it arrives, apply the bugfix manually (compare the code and see if you can make the changes yourself)
'
'       Create a new resource file with the three OCXs inserted as well as the custom manifest #24
'           Test locally and on desktop with the new RES file.
'           You might have to do this each time there is a new version of CCRimageList.ocx
'
' Bugs:
' =====
'
'       when editing an icon, dragging and dropping to the map, the set button is not enabled
'       when editing an icon, dragging and dropping to the map, the close button does not change to cancel
'       cancel does not revert the icon dragged to the dock
'
'       read the docksettings and determine chkRetainIcons
'
'       the value of chkRetainIcons must be used to determine whether the icon is automatically pulled from the embedded icon or the collection
'
'       why is embeddedIcons called three times? - need to figure that out.
'
'       update the help file.
'
'       note: - Compile the other associated projects regularly as they use shared code in modules 2 & 3.
'
'       appIdent.csv - complete the list using the registry entries - test existing utilities used
'
'       appIdent.csv - add icon references to the list
'           hwinfo64 hwinfo64.png
'
'       rubberduck the code
'
'       add the build instructions
'
'       build the manifest - test the manifest for 64 bit & win 11 systems - no admin required, DPI needed.
'
'       Elroy's code to add balloon tips to comboBox
'       https://www.vbforums.com/showthread.php?893844-VB6-QUESTION-How-to-capture-the-MouseOver-Event-on-a-comboBox
'
' There are a few new functions that I'd like to add in the task list:
' ====================================================================
'
'    1. When adding a shortcut to the dock we need to determine the built-in ico and give the option to extract that.
'       The class code uses GDI+ to read the PNG files.
'       GdipLoadImageFromFile
'       GdipGetImageBounds
'       GdipCreateFromHdc
'       GdipDrawImageRectRectI
'       It is rendered into a picturebox.
'
'    2. Test with empty dock and first run with no prior install of anything.
'
'    3. Ensure the laptop OCX configuration pulls all the OCX from the local versions of the three OCX
'       Ensure the desktop OCX configuration matches that of the desktop so the build on the desktop works with the embedded RES file.
'       The OCXs should be in the same folder.
'
'    4.
'
'    5. I would like to replicate those main icons in a drawn style using the Wacom pad  Very Low Priority
'
'    6. Resize all the controls in the same way that we resize the generateDock form - I'd like to have that but all the icons would
'       need dynamic resizing and that would take time. This would need a re-engineering of the map. VERY LOW PRIORITY.
'
'    7. ANY controls loaded at runtime, MUST be Unloaded when close the form - we need to check this.
'
'>   8. Elroy's balloon tooltip adaptation for drop down controls - important to complete the balloon tips.
'       The thread on how to do this is here:
'       https://www.vbforums.com/showthread.php?893844-VB6-QUESTION-How-to-capture-the-MouseOver-Event-on-a-comboBox&p=5540805&highlight=#post5540805
'       unfortunately it requires subclassing and that is bad for development within the IDE, stop button crashes &c.
'       So, it is something to do right at the very end when you have completed all the other development.
'       Now medium priority but still to be done last. Reason for this is that the other docksettings tool makes frequent use of
'       drop downs and so it needs to be implemented to complete the balloon tooltips there.
'
'   9.  Skin the interface in a medieval manner!  Very Low Priority.
'
'   10. Use the lightweight method of reading images from SteamyDock rather than LaVolpe's method using readFromStream.
'
'
' Other Tasks:
'
'   Github
'
'   SD Messagebox msgBoxA module - ship the code to FCW to replace the native msgboxes.
'
'   SD DirectX 2D Jacob Roman's training utilities to implement 2D graphics in place of GDI+
'      in addition there is the VB6 dock version from the same author as the original GDI+ dock used as inspiration here,
'      that uses DirectX 2D.
'
'   SD Avant manager - test the animation routine for the dock, circledock might be worth looking at?
'

' BUILD:

'       Create a new resource file with the three OCXs inserted as well as the custom manifest #24
'           Test locally and on desktop with the new RES file.
'           You might have to do this each time there is a new version of CCRimageList.ocx


Option Explicit

'--------------------------------------------------------------------------------------------------------------
' Form Module : rDIconConfigFrm
' Author      : Dean Beedell
' Date        : 20/06/2019
'
' SteamyDock
'
' A VB6 GDI+ dock for XP, Vista, Win7, 8, 10 and Reactos.
' SteamyDock is a functional reproduction of the dock we all know and love - Rocketdock for Windows from Punklabs
' which aimed to replicate the Mac dock that can be seen today on the Mac osX operating system.
'
'--------------------------------------------------------------------------------------------------------------
'
' History
'           This component, one of three that comprise Steamydock was the first VB6 project that I have 'undertaken and completed' so
'           forgive the errors in coding styles and methods. Entirely self-taught and a mere 'hobbyist' in VB6.
'
'           Everyone has their reasons for coding, amateur hobbyist or professional. My primary reason for creating this SteamyDock utility,
'           I believe, is quite unique. I have a rather frightening admission to make. I am very worried about inheriting the gene that causes
'           dementia. Dementia abounds in my family, My grandmother, mother, two uncles are all suffering from it. The prospect is is a
'           little scarey and a few years ago I noticed some deterioration of my mental faculties, memory, fluency of vocabulary &c.
'           All were suffering a little. I decided to undertake a major project to try to sharpen my mental acuity and determine whether
'           I was going to suffer the same decay that my family members have been going through. Three years onward and I can say that
'           thankfully, there are no signs of it but I can say for certain, it has sharpened my mind.
'
'           The secondary reason I created it was to re-teach myself VB6, to get back into the programming 'groove'.
'
'           Back in the 90s I was programming in QB45 (and subsequently VB-DOS and the original VB from version one through to VB6) but when
'           .NET came on the scene, I left BASIC programming forever and abandoned my main project when VB6 was deprecated. My skills were
'           paltry then and had been merely picked up from the days of Sinclair ZX80s, ZX81s the ZX Spectrum and the like.
'
'           My aim now has been to resurrect such skills that I once had and improve upon them.
'           A secondary aim is to teach myself how to code in technologies that I have encountered but never fully embraced.
'
'           When I dropped BASIC I picked up Javascript and managed to hone my limited programming skills to a reasonable hobbyist level
'           but I always missed VB6 and that familiar old IDE. Javascript still has no equivalent decent IDE for what I need it to do.
'           Javascript however, is very much like BASIC in so many ways and a hobbyist BASIC programming style works in Javascript too.
'
'           Returning to VB6 after so many years, I was able to pick up the old IDE and see how surprisingly efficient it was on the hardware we
'           have today. However, having become quite used to more modern and advanced languages it was a big surprise to me to realise how
'           limited VB6 was, compared to modern languages.
'
'           VB6 had inadequate image type handling, VB6 being completely unable to handle more modern image types such as PNGs without
'           the use of a great reams of code and the use of a large number of API calls. VB6's file system dialogs seem to be unreliable under
'           modern versions of the Windows o/s. These and other problems need to be resolved...

'           However, in the process of creating this utility I learned that VB6 can 'do' anything as long as you are prepared to write the
'           code to do it. This process however, can take a lot of hard work. VB6's successor, VB.NET does a lot by default but it
'           also makes programming in general a lot more painful in general due to the syntax differences and a very different programming style
'           and methodology. Either language, programming in VB can be a hard slog.
'
'           I had a discussion with a .NET programmer of many years experience and he saw me programming in VB6 and doing hands-on work
'           with APIs and he asked why I persisted in coding in a twenty-year old language, one that had fossilised in time and could not
'           do the things current languages can do so easily. I explained that I am intending to contribute to ReactOS (a FOSS project to
'           clone Windows that is written largely in C and C++) and that it is precisely because VB6 cannot do the things that later languages
'           can do, that coding with VB6 will provide me with some of the fundamental knowledge I need to show me how Windows and its APIs
'           function and therefore how Windows and ReactOS actually do what they do.
'
'           My code and understanding are improving as I get closer to completing this dock project. I have C style syntax under my belt
'           due to my experience with javascript. To me, it appears that C++ and C are not easy languages to use in order to learn o/s
'           aspects but VB6 has been the perfect tool for what I am doing now. An accessible and strongly typed language used to implement aspects of
'           the operating system such as stopping/starting processes and maintaining lists of running applications, manual z-ordering of windows
'           and how to create graphical utilities using only the most basic graphic functions (GDI+) with minimal CPU usage.
'
'           I may yet be some way from contributing code to ReactOS yet but I feel I am making progress.
'
'           At the moment this project is largely VB6 but when this project is complete my next aim is to migrate it to VB.NET through
'           the versions to find out what problems are typically encountered in upgrading a BASIC project such as this.
'
'           I could not have made this utility without the help of code from the various projects I have listed below.
'
'           I hope you enjoy the functionality this utility provides. If you think you can improve anything then please
'           feel free to do so. If you dislike my programming style then do keep those thoughts to yourself. :)
'
'           Built on a 2.5ghz core2duo Dell Latitude E5400 running Windows 7 Ultimate 64bit using VB6 SP6.
'           Debugged on a 3.3ghz Dell Latitude E6410 running Windows 7 Ultimate 64bit using VB6 SP6.
'
' Credits : Standing on the shoulders of the following giants:

'           LA Volpe (VB Forums) for his transparent picture handling.
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           KillApp code from an unknown, untraceable source, possibly on MSN.
'           Registry reading code from ALLAPI.COM.
'           Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.
'           Active VB Germany for information on the undocumented PrivateExtractIcons API.
'           Elroy on VB forums for his Persistent debug window
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
'           Krool on the VBForums for his impressive common control replacements
'           si_the_geek for his special folder code
'           KPD-Team for the code to trawl a folder recursively KPDTeam@Allapi.net http://www.allapi.net
'           Elroy's for the balloon tooltips
' Rod Stephens vb-helper.com            Resize controls to fit when a form resizes
' KPD-Team 1999 http://www.allapi.net/  Recursive search
' IT researcher https://www.vbforums.com/showthread.php?784053-Get-installed-programs-list-both-32-and-64-bit-programs
'                                       For the idea of extracting the ununinstall keys from the registry
' CREDIT Jacques Lebrun http://www.vb-helper.com/howto_get_shortcut_info.html
'
' Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
'
'           MZ-TOOLS https://www.mztools.com/
'           CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
'           Rubberduck http://rubberduckvba.com/
'           Rocketdock https://punklabs.com/
'           Registry code ALLAPI.COM
'           La Volpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'           PrivateExtractIcons code http://www.activevb.de/rubriken/
'           Persistent debug code http://www.vbforums.com/member.php?234143-Elroy
'           Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
'           Open font dialog code without dependent OCX - unknown URL
'           Krool's replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29
'
'   Tested on :
'           ReactOS 0.4.14 32bit on virtualBox
'           Windows 7 Professional 32bit on Intel
'           Windows 7 Ultimate 64bit on Intel
'           Windows 7 Professional 64bit on Intel
'           Windows XP SP3 32bit on Intel
'           Windows 10 Home 64bit on Intel
'           Windows 10 Home 64bit on AMD
'           Windows 11 64bit on Intel
'
' Dependencies:
'           Krool's replacement for the Microsoft Windows Common Controls found in
'           mscomctl.ocx (treeview, slider) are replicated by the addition of two
'           dedicated OCX files that are shipped with this package.
'
'           CCRImageList.ocx
'           CCRSlider.ocx
'           CCRTreeView.ocx
'
'           These OCX will reside in the same folder as the utility that uses it.
'
'           OLEGuids.tlb
'
'           This is a type library that defines types, object interfaces, and more specific API definitions
'           needed for COM interop / marshalling. It is only used at design time (IDE). This is a Krool-modified
'           version of the original .tlb from the vbaccelerator website. The .tlb is compiled into the executable.
'           For the compiled .exe this is not a dependency.
'
'           From the command line, copy the tlb to a central location and register it.
'
'           COPY OLEGUIDS.TLB %SystemRoot%\System32\
'           REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
'
' Building a Manifest:
'           Using La Volpe's program
'
'
' Project References:
'           VisualBasic for Applications
'           VisualBasic Runtime Objects and Procedures
'           VisualBasic Objects and Procedures
'           OLE Automation - drag and drop
'           Microsoft Shell Controls and Automation
'
' Notes:
'           Integers are retained (rather than longs) as some of these are passed to
'           library API functions in code that is not my own so I am loathe to change.
'           A lot of the code provided (by better devs than me) seems to have code quality
'           issues (as does mine!) - I haven't gone through all the 3rd party code to fix every
'           problem but I have fixed quite a few. My own code has significant quality isses.
'
'           The icons are displayed using Lavolpe's transparent DIB image code,
'           except for the .ico files which use his earlier StdPictureEx class.
'           The original ico code caused many strange visual artifacts and complete failures to show .ico files.
'           especially when other image types were displayed on screen simultaneously.
'
' Summary:
'           The program can read a default icon folder from Rocketdock's settings.ini or registry. Use this functionality
'           to transfer the contents of Rocketdock to Steamydock. It is a one-way process, we can read Rocketdock's data
'           but we do not attempt to write any changes back to Rocketdock.
'
'           It reads the contents of the folder and sub-folders into a treeview and displays the first 12 of the
'           icons using 12 dynamically created picboxes.
'
'           VB6 does not support more modern transparent image types natively. The core of this program is Lavolpe's
'           image handling code allowing it to read and display all types of images including those that support
'           transparencies. These are then rendered into standard picture boxes.
'
'           The icons are displayed using Lavolpe's transparent DIB image code, except for the .ico files which use the earlier StdPictureEx class.
'           DLLs and EXEs with embedded icons are handled using an undocumented API named PrivateExtractIcons.
'           One selected image is extracted and displayed in larger size using the above code in the preview window.
'
'           LaVolpe's methods of image handling are not used in SteamyDock itself, only in Rocketdock.
'
'           A copy of Rocketdock's settings are transferred from the registry or settings.ini into an interim
'           settings file which provides a common method of handling the data.
'           The icon details are read from this file and the details
'           of the selected icon are displayed in the text boxes in the 'properties' frame. This data is also
'           read when the user chooses to the display the Rocketdock map.
'
'           In that 'map' each dock image is displayed in smaller form in dynamically created picboxes.
'           The RD map acts a cache of images that takes a few seconds to create but
'           doing it this way means there is no subsequent delay when viewing any other part of the map.
'           The images on the map can then be scrolled into visibility viewing fifteen icons at a time. It has
'           been tested with a map containing up to 67 icons.
'
'           The icon details are written to the registry or the settings file but only after Rocketdock
'           has been closed and just before it is restarted otherwise it will overwrite any settings
'           changes when it exits.
'
'           The utility itself has some configuration details that it stores in its own local settings.ini file.
'
'           The font selection and file/folder dialogs are generated using Win32 APIs rather than the
'           common dialog OCX which dispensed with another OCX.
'
'           I have used Krool's amazing control replacement project. The specific code for
'           just two of the controls (treeview and slider) has been incorporated rather than all 32 of
'           Krool's complete package.
'
'           In the population of the thumbnail pane we use a primitive post-fetch cache, ie. it speeds up any access
'           after each first image read. When the cache is filled (limited by a count) it adds no more and does not
'           clear up the oldest item freeing the space, it just stops populating.
'
'           The cache is populated using a VB6 imageList replacement from Krool.
'
'           For each image read from file and displayed, it is added to the imageList in its resized form.
'           Each image was given a unique name as a key relating to its position in the grid. This key's existence is
'           checked just before any image is accessed, if the key exists then the resized .picture is extracted from
'           the imageList rather than reading it from file.
'
'           Each subsequent access is much, much faster as we are retrieving from memory and only retrieving a tiny
'           image rather than the big 256x256 one on disc.
'
'           We limit the cache to certain number of image items to prevent out of memory messages, the images that
'           we are cacheing are only very small (32x32 or 16x16) so a limit of 250/500 is probably fine.
'
'           When the cache is full we do not attempt to remove the oldest image added as we would have to keep a track
'           of the insert times and this would require complexity outside the imageList.
'
'           I could attempt an array of images and see if that is faster than an imageList (I am sure it will be) but
'           having each image stored under a unique key is a very useful feature, the extra functionality it provides
'           by default I would have to build manually in code. The imageList cache is also as fast as I need it to be.
'           Remember the images are very small and are fast to load in any case, the result of very fast CPUs and SSDs.
'
'           I also created a timer that was designed to be used to populate the cache in advance, in a pre-fetch manner.
'           Used in conjunction with an idletime tester it could preload images into the cache when the app. is idle,
'           perhaps 4 icons at a time every 5 seconds or so. It would not help much when scrolling down as it would have
'           to be running in the same thread and would slow down normal operation. If VB6 supported multi threading it
'           might be sensible to implement it. Part-coded, it is disabled for now.
'
'
'    LICENCE AGREEMENTS:
'
'    Copyright © 2019 Dean Beedell
'
'    Using this program implies you have accepted the licence. The GPL licence applies to the code
'    this software Is provided 'as-is', without any express or implied warranty. In no event will the
'    author be held liable for any damages arising from the use of this software. Permission is granted to
'    anyone to use this software for any purpose, including commercial applications, and to alter it and
'    redistribute it freely, subject to the following restrictions:
'
'    1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation is required.
'    2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.
'    3. This notice may not be removed or altered from any source distribution.
'
'    This program is free software; you can redistribute it and/or modify it under the terms of the
'    GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'    even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    General Public Licence for more details.
'
'    You should have received a copy of the GNU General Public Licence along with this program; if not,
'    write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
'    USA
'
'    If you use this software in any way whatsoever then that implies acceptance of the licence. If you
'    do not wish to comply with the licence terms then please remove the download, binary and source code
'    from your systems immediately.
'
'--------------------------------------------------------------------------------------------------------------

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As OLE_COLOR, ByVal hPal As Long, ByRef lpColorRef As Long) As Long

Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean

' .13 DAEB rdIconConfig.frm 09/02/2021 Added ability to check if a window exists
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private cImage As c32bppDIB
Private cShadow As c32bppDIB

' Note: If GDI+ is available, it is more efficient for you to
' create the token then pass the token to each class.  Not required,
' but if you don't do this, then the classes will create and destroy
' a token everytime GDI+ is used to render or modify an image.
' Passing the token can result in up to 3x faster processing overall

Private m_GDItoken As Long
'Private FontDlg As CommonDlgs

Private Const COLOR_BTNFACE As Long = 15

'some variables for temporarily storing the old image name
Private previousIcon As String
Private mapImageChanged As Boolean
Private thumbPos0Pressed As Boolean
Private validIconTypes As String  ' change from VB6 to scope due to replacement of filelistbox control to simple listbox

' .10 DAEB 07/02/2021 rDIconConfigForm.frm removed unused vars STARTS
'Public rDLockIcons As String
'Public rDOpenRunning As String
'Public rDShowRunning As String
'Public rDManageWindows As String
'Public rDDisableMinAnimation As String
'Public rDDefaultDock As String
'Public rDRunAppInterval As String
'Public rDAlwaysAsk As String
'Public rdStartupRunString As String
'Public rDStartupRun As String
' .10 DAEB 07/02/2021 rDIconConfigForm.frm removed unused vars ENDS


Private totalBusyCounter As Integer ' .12 DAEB 07/02/2021 rDIconConfigForm.frm added as part of busy timer functionality

Private moreConfigVisible As Boolean
Private startRecordNumber As Integer ' .47 DAEB 16/04/2022 rdIconConfig.frm Added StartRecordNumber

Private sdMapState As String '  TBD
Private sdThumbnailCacheCount As String ' TBD the cache size can be modified in the config.

Private filesIconListClicked As Boolean
'------------------------------------------------------ STARTS
' Type defined for testing a time difference used to initiate one of the hand-coded timers
Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

' APIs defined for testing a time difference used to initiate one of the hand-coded timers
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
'------------------------------------------------------ ENDS

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private oldDockSettingsModificationTime  As Date
Private dockSettingsRunInterval As Long

Private dragToDockOperating As Boolean





'---------------------------------------------------------------------------------------
' Procedure : btnAppToTerminate_Click
' Author    : beededea
' Date      : 07/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAppToTerminate_Click()
    On Error GoTo btnAppToTerminate_Click_Error
    Dim retFileName As String: retFileName = vbNullString
    
    Call selectApplication(txtAppToTerminate.Text, retFileName)
    txtAppToTerminate.Text = retFileName
    txtAppToTerminate.ToolTipText = retFileName

    On Error GoTo 0
    Exit Sub

btnAppToTerminate_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAppToTerminate_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub btnAppToTerminate_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnAppToTerminate.hwnd, "This button will allow you to select any program that must be terminated prior to the main program initiation. The result is: When you click on the icon in the dock SteamyDock will do its very best to terminate the chosen application in advance but be aware that closing another application cannot be guaranteed - use this functionality with great care! ", _
                  TTIconInfo, "Help on Terminating an Application", , , , True
End Sub



Private Sub btnBackup_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub btnCancel_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub btnClose_Click()
    Form_Unload 0
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnClose.hwnd, "This button closes the window.", _
                  TTIconInfo, "Help on the Close Button", , , , True
End Sub

Private Sub btnGenerate_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub btnGetMore_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub btnHelp_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnIconSelect_Click
' Author    : beededea
' Date      : 16/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnIconSelect_Click()
    Dim iconPath As String: iconPath = vbNullString
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim picSize As Long: picSize = 0
    Dim suffix As String: suffix = vbNullString
  
    Dim Filename As String: Filename = vbNullString
    Dim validImageTypes As String: validImageTypes = vbNullString
    Dim savFileName As String: savFileName = vbNullString
    
    Dim extraValidImageTypes As String: extraValidImageTypes = vbNullString

    Const x_MaxBuffer = 256
    
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString


        
    ' add any remaining types that Rocketdock's code supports
    On Error GoTo btnIconSelect_Click_Error

    validImageTypes = ".jpg,.jpeg,.bmp,.ico,.png,.tif,.gif"
    
    'On Error GoTo btnTarget_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnIconSelect_Click"
    
    'On Error GoTo l_err1
 
    On Error Resume Next
    
    savFileName = txtCurrentIcon.Text
    
    ' set the default folder to the existing reference
    If Not txtCurrentIcon.Text = vbNullString Then
        If FExists(txtCurrentIcon.Text) Then
            ' extract the folder name from the string
            iconPath = getFolderNameFromPath(txtCurrentIcon.Text)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(txtCurrentIcon.Text) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = txtCurrentIcon.Text 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then
                dialogInitDir = rdAppPath
            Else
                dialogInitDir = sdAppPath
            End If
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select an icon image."
    .lpstrInitialDir = dialogInitDir
    .lpstrFilter = "PNG Files" & vbNullChar & "*.png" _
    & vbNullChar & "ICO Files" & vbNullChar & "*.ico" _
    & vbNullChar & "JPG Files" & vbNullChar & "*.jpg" _
    & vbNullChar & "JPEG Files" & vbNullChar & "*.jpeg" _
    & vbNullChar & "BMP Files" & vbNullChar & "*.bmp" _
    & vbNullChar & "GIF Files" & vbNullChar & "*.gif" _
    & vbNullChar & "TIF Files" & vbNullChar & "*.tif" _
    & vbNullChar & "TIF Files" & vbNullChar & "*.tiff" _
    & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    
    .nFilterIndex = 9
    
    .lpstrFile = String$(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With

    Call getFileNameAndTitle(retFileName, retfileTitle)
    If retFileName <> vbNullString Then
      txtCurrentIcon.Text = retFileName
      'fill in the file title and the start in automatically if they are empty and need filling
      If txtLabelName.Text = vbNullString Then txtLabelName.Text = retfileTitle
      If txtStartIn.Text = vbNullString Then txtStartIn.Text = getFolderNameFromPath(txtCurrentIcon.Text)
    End If
  
    Filename = txtCurrentIcon.Text ' this takes the filename from the field which causes the 256 buffered variable retFileName to truncate
    '                                NOT a feature that can be undone to a buffered var with a simple rtrim$.
    suffix = ExtractSuffixWithDot(Filename)
    ' DAEB TBD
    extraValidImageTypes = validImageTypes & ",.dll,.exe"
    If InStr(extraValidImageTypes, suffix) <> 0 Then
        picSize = FileLen(Filename)
        lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
            
        'refresh the preview displaying the selected image
        Call displayResizedImage(Filename, picPreview, icoSizePreset)
        
        ' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon.
        Call displayResizedImage(Filename, picTemporaryStore, 64)
                                
        ' .62 DAEB 04/05/2022 rDIConConfig.frm change the icon image in the map to that chosen
        Call displayResizedImage(Filename, picRdMap(rdIconNumber), 32)
                       
        picPreview.ToolTipText = Filename
        picPreview.Tag = Filename
        
        Call changeMapImage
    Else ' if the file type chosen is incorrect - just ignore it
        txtCurrentIcon.Text = savFileName
    
        msgBoxA "This option will only allow you to select image files", vbExclamation + vbOKOnly, "Oops - incorrect file type selected"
    End If

   On Error GoTo 0
   Exit Sub

btnIconSelect_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnIconSelect_Click of Form rDIconConfigForm"
  
End Sub




Private Sub btnSaveRestart_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub btnSet_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnSettingsDown_Click
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSettingsDown_Click()
   On Error GoTo btnSettingsDown_Click_Error

'    If frmRegistry.Visible = True Then
'        btnSettingsDown.Visible = True
'        btnSettingsUp.Visible = False
'
'        frmRegistry.Hide
'    Else
'        btnSettingsDown.Visible = False
'        btnSettingsUp.Visible = True
    
        frmRegistry.Left = rDIconConfigForm.Left + btnSettingsDown.Left
        frmRegistry.Top = rDIconConfigForm.Top + btnSettingsDown.Top + 800
        
        frmRegistry.Show
'    End If

   On Error GoTo 0
   Exit Sub

btnSettingsDown_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSettingsDown_Click of Form rDIconConfigForm"

End Sub





Private Sub btnSettingsUp_Click()
'        rDIconConfigForm.btnSettingsDown.Visible = True
'        rDIconConfigForm.btnSettingsUp.Visible = False
'
'        frmRegistry.Hide

End Sub


Private Sub btnWorking_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkConfirmDialogAfter_Click
' Author    : beededea
' Date      : 31/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkConfirmDialogAfter_Click()

   On Error GoTo chkConfirmDialogAfter_Click_Error

    btnSet.Enabled = True ' tell the program that something has changed
    btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

chkConfirmDialogAfter_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkConfirmDialogAfter_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkConfirmDialog_Click
' Author    : beededea
' Date      : 19/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkConfirmDialog_Click()

   On Error GoTo chkConfirmDialog_Click_Error

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False
    
   On Error GoTo 0
   Exit Sub

chkConfirmDialog_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkConfirmDialog_Click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkDisabled_Click
' Author    : beededea
' Date      : 30/01/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkDisabled_Click()

    On Error GoTo chkDisabled_Click_Error

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False

    On Error GoTo 0
    Exit Sub

chkDisabled_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDisabled_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub chkDisabled_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkDisabled.hwnd, "This checkbox will cause the icon to stop responding to a mouse click. It disables the icon.", _
                  TTIconInfo, "Help on Disabling", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDefaultDock_Change
' Author    : beededea
' Date      : 09/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbDefaultDock_Click()
   On Error GoTo cmbDefaultDock_Click_Error

    If cmbDefaultDock.List(cmbDefaultDock.ListIndex) = "RocketDock" Then
        ' check where rocketdock is installed
        Call checkRocketdockInstallation
        
    Else
        ' check where/if steamydock is installed
        Call checkSteamyDockInstallation

    End If

   On Error GoTo 0
   Exit Sub

cmbDefaultDock_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDefaultDock_Click of Form rDIconConfigForm"
End Sub


' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
'---------------------------------------------------------------------------------------
' Procedure : btnSecondApp_Click
' Author    : beededea
' Date      : 21/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSecondApp_Click()

    Dim retFileName As String: retFileName = vbNullString

    On Error GoTo btnSecondApp_Click_Error

    Call selectApplication(txtSecondApp.Text, retFileName)
    txtSecondApp.Text = retFileName
    txtSecondApp.ToolTipText = retFileName

    On Error GoTo 0
    Exit Sub

btnSecondApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSecondApp_Click of Form rDIconConfigForm"
    
End Sub
' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
'---------------------------------------------------------------------------------------
' Procedure : selectApplication
' Author    : beededea
' Date      : 21/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub selectApplication(ByVal inputFolderName As String, ByRef retFileName As String)
   Dim iconPath As String: iconPath = vbNullString
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString

    Const x_MaxBuffer = 256
    
    Dim retfileTitle As String: retfileTitle = vbNullString
    
    
    'On Error GoTo addTargetProgram_Error
    On Error GoTo selectApplication_Error

    If debugflg = 1 Then DebugPrint "%" & "selectApplication_Error"
    
    'On Error GoTo l_err1
    'savLblTarget = inputFoldername
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not inputFolderName = vbNullString Then
        If FExists(inputFolderName) Then
            ' extract the folder name from the string
            iconPath = getFolderNameFromPath(inputFolderName)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(inputFolderName) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = inputFolderName 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
            End If
        End If
    Else
        ' .85 DAEB 06/06/2022 rDIConConfig.frm  Second app button should open in the program files folder
        If DirExists("c:\program files") Then
            dialogInitDir = "c:\program files"
        End If
    End If
    
    If Not sDockletFile = vbNullString Then
        If FExists(sDockletFile) Then
            ' extract the folder name from the string
            dllPath = getFolderNameFromPath(sDockletFile)
            ' set the default folder to the existing reference
            dialogInitDir = dllPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(sDockletFile) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = sDockletFile 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
                dialogInitDir = rdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            End If
        End If
    End If
    
    With x_OpenFilename
    '    .hwndOwner = Me.hWnd
      .hInstance = App.hInstance
      .lpstrTitle = "Select a File Target for this icon to call"
      .lpstrInitialDir = dialogInitDir
      
      .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
      .nFilterIndex = 2
      
      .lpstrFile = String$(x_MaxBuffer, 0)
      .nMaxFile = x_MaxBuffer - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = x_MaxBuffer - 1
      .lStructSize = Len(x_OpenFilename)
    End With
    
    
    Call getFileNameAndTitle(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes
    If retFileName <> vbNullString Then
      inputFolderName = retFileName ' strips the buffered bit, leaving just the filename
      'fill in the file title and the start in automatically if they are empty and need filling
      'If txtLabelName.Text = vbNullString Then txtLabelName.Text = retfileTitle
      'If txtStartIn.Text = vbNullString Then txtStartIn.Text = getFolderNameFromPath(inputFoldername)
    End If
    

    On Error GoTo 0
    Exit Sub

selectApplication_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure selectApplication of Form rDIconConfigForm"

End Sub





Private Sub fraIconType_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    '.100 DAEB 09/11/2022 rDIConConfig.frm Add the right click menu to all the buttons and recently added frames.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub




Private Sub fraOptionButtons_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraOptionButtons.hwnd, "An additional secondary program can be made to run before or after the main program launch has completed. These controls will be disabled until a program has been selected.", _
                  TTIconInfo, "Help on Second Application", , , , True
End Sub

' This is the timer that was designed to populate the cache in advance, in a pre-fetch manner.
' Used in conjunction with an idletime tester it could preload images into the cache when the app. is idle,
' perhaps 4 icons at a time every 5 seconds or so. It would not help much when scrolling down right from the start as
' it would have to be running in the same thread and would slow down normal operation.
'
' If VB6 supported multi threading it might be sensible to implement it as the job could run in a separate thread.

Private Sub idleTimer_Timer()

    Dim lastInputVar As LASTINPUTINFO
    Dim currentIdleTime As Long: currentIdleTime = 0

    Const lngThousand As Long = 1000
    
    ' initialise vars
        
    ' check to see if the app has not been used for a while, ie it has been idle
    lastInputVar.cbSize = Len(lastInputVar)
    Call GetLastInputInfo(lastInputVar)
    currentIdleTime = GetTickCount - lastInputVar.dwTime
    
    ' only allows the function to continue if FCW has been idle for more than 30 secs
    If currentIdleTime < 30000 Then Exit Sub

    ' TBD take the current folder and add each item to the cache
    
    ' TBD cache the current image to a unique key that corresponds to the item index in the filesIconList
    ' this is a post-fetch cache, ie. it speeds up any access after each first image read.
    ' has the thumbnail grid yet been populated
    ' use a loop
    
'        For useloop = 0 To 11
'
'        'startItem = filesIconList.ListIndex ' the starting point in the file list for the thumnbnails to start
'        'when there are less than a screenful of items the.ListIndex returns -1
'
'
'        ' aside -> .NET collection can't handle going up to or beyond the count, VB6 control array copes
'        ' but the count check is here for compatibility with the .NET version.
'
'        If useloop + startItem < filesIconList.ListCount Then
'            ' take the fileame from the underlying filelist control
'            shortFilename = filesIconList.List(useloop + startItem)

        ' at each point in the loop check to see if the user is doing anything
        '  another idletime check or a flag set when the user interacts with the thumbnail
        '  read each file from the
'    If fullFilePath <> "" Then
'        If imlThumbnailCache.ListImages.Exists("cache" & useloop + startItem) Then
'            ' at the moment do nothing
'        Else
'
'            ' display the image from file within the specified picturebox
'            Call displayResizedImage(fullFilePath, picThumbIcon(useloop), imageSize)
'
'            ' add the current thumbnail to the cache with a unique key
'            Set picTemporaryStore.Picture = picThumbIcon(useloop).Image
'            imlThumbnailCache.ListImages.Add , "cache" & useloop + startItem, picTemporaryStore.Picture
'            Set picTemporaryStore.Picture = Nothing
'        End If
'    End If

End Sub




Private Sub lblDisabled_Click()
    If chkDisabled.Value = 1 Then
        chkDisabled.Value = 0
    Else
        chkDisabled.Value = 1
    End If
End Sub



Private Sub lblRdIconNumber_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

   If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub lblRunSecondAppAfterward_Click()
    optRunSecondAppAfterward.Value = True
End Sub

Private Sub lblRunSecondAppBeforehand_Click()
    optRunSecondAppBeforehand.Value = True
End Sub





'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtCompMgmt_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtCompMgmt_click()
   On Error GoTo mnuTrgtCompMgmt_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtCompMgmt_click"

    sCommand = "compmgmt.msc"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtCompMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtCompMgmt_click of Form rDIconConfigForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtDevMgmt_Click
' Author    : beededea
' Date      : 26/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtDevMgmt_Click()
    On Error GoTo mnuTrgtDevMgmt_Click_Error

    sCommand = "devmgmt.msc"
    txtTarget.Text = sCommand

    On Error GoTo 0
    Exit Sub

mnuTrgtDevMgmt_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtDevMgmt_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtDiscMgmt_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtDiscMgmt_click()
   On Error GoTo mnuTrgtDiscMgmt_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtDiscMgmt_click"

    sCommand = "diskmgmt.msc"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtDiscMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtDiscMgmt_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtEventViewer_Click
' Author    : beededea
' Date      : 26/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtEventViewer_Click()
    On Error GoTo mnuTrgtEventViewer_Click_Error

    sCommand = "eventvwr.msc"
    txtTarget.Text = sCommand

    On Error GoTo 0
    Exit Sub

mnuTrgtEventViewer_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtEventViewer_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtPerfMon_Click
' Author    : beededea
' Date      : 26/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtPerfMon_Click()
    On Error GoTo mnuTrgtPerfMon_Click_Error


    sCommand = "perfmon.msc"
    txtTarget.Text = sCommand

    On Error GoTo 0
    Exit Sub

mnuTrgtPerfMon_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtPerfMon_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtServices_Click
' Author    : beededea
' Date      : 26/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtServices_Click()
    On Error GoTo mnuTrgtServices_Click_Error

    sCommand = "services.msc"
    txtTarget.Text = sCommand


    On Error GoTo 0
    Exit Sub

mnuTrgtServices_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtServices_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtTaskSched_Click
' Author    : beededea
' Date      : 26/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtTaskSched_Click()
    On Error GoTo mnuTrgtTaskSched_Click_Error



    sCommand = "taskschd.msc"
    txtTarget.Text = sCommand

    On Error GoTo 0
    Exit Sub

mnuTrgtTaskSched_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtTaskSched_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub optRunSecondAppAfterward_Click()
    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False
End Sub

Private Sub optRunSecondAppBeforehand_Click()
    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False
End Sub

Private Sub picRdMap_OLEDragOver(ByRef Index As Integer, ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single, ByRef State As Integer)
    rdIconNumber = Index
End Sub



Private Sub positionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    rDIconConfigFormXPosTwips = rDIconConfigForm.Left
    rDIconConfigFormYPosTwips = rDIconConfigForm.Top
    
    
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormXPos", rDIconConfigFormXPosTwips, toolSettingsFile
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormYPos", rDIconConfigFormYPosTwips, toolSettingsFile
End Sub


' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
Private Sub rdMapDragTimer_Timer()
    srcDragControl = "rdMap"
    rdMapDragTimerCounter = rdMapDragTimerCounter + 1

    If rdMapDragTimerCounter >= 25 Then
        If rdMapIconMouseDown = True Then
            picRdMap(rdIconNumber).Drag vbBeginDrag
            srcRdIconNumber = rdIconNumber ' record the source icon number for reordering later
        End If
        rdMapDragTimer.Enabled = False
        rdMapDragTimerCounter = 0
    End If

End Sub



Private Sub Text1_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuTrgtMenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : settingsTimer_Timer
' Author    : beededea
' Date      : 03/03/2023
' Purpose   : Checking the date/time of the settings.ini file meaning that another tool has edited the settings, could cause an automatic restart.
'---------------------------------------------------------------------------------------
'
Private Sub settingsTimer_Timer()

    Dim timeDifferenceInSecs As Long: timeDifferenceInSecs = 0 ' max 86 years as a LONG in secs
    Dim dockSettingsModificationTime As Date: dockSettingsModificationTime = #1/1/2000 12:00:00 PM#
    Dim useloop As Integer: useloop = 0
    Dim lastChangedByWhom As String: lastChangedByWhom = vbNullString
    Dim lastIconChanged As Integer: lastIconChanged = 0
    Dim ans As VbMsgBoxResult: ans = vbNo
    Dim thisQuestionText As String: thisQuestionText = vbNullString
    
    On Error GoTo settingsTimer_Timer_Error
    
    'settingsTimer.Enabled = False
    
    dockSettingsRunInterval = dockSettingsRunInterval + 1
    If dockSettingsRunInterval < 2 Then Exit Sub
    dockSettingsRunInterval = 3

    If Not FExists(dockSettingsFile) Then
        MsgBox ("%Err-I-ErrorNumber 13 - FCW was unable to access the dock settings ini file. " & vbCrLf & dockSettingsFile)
        Exit Sub
    End If
    
    ' check the dockSettings.ini file date/time
    dockSettingsModificationTime = FileDateTime(dockSettingsFile)
    
    timeDifferenceInSecs = Int(DateDiff("s", oldDockSettingsModificationTime, dockSettingsModificationTime))

    ' if the dockSettings.ini has been modified then reload the map
    If timeDifferenceInSecs > 1 Then
    
'            arse1.Caption = oldDockSettingsModificationTime
'        arse2.Caption = dockSettingsModificationTime
        oldDockSettingsModificationTime = dockSettingsModificationTime
        
        
        ' read the lastChangedByWhom variable from the dock settings file
        ' if the lastChangedByWhom is SD then refresh
        lastChangedByWhom = GetINISetting("Software\SteamyDock\DockSettings", "lastChangedByWhom", dockSettingsFile)
        lastIconChanged = Val(GetINISetting("Software\SteamyDock\DockSettings", "lastIconChanged", dockSettingsFile))
        
        If lastChangedByWhom = "steamyDock" And lastIconChanged <> 9999 Then
                    
            thisQuestionText = " SteamyDock has modified the icon configuration, do you want to reload the map? "
            If btnSet.Enabled = True Or mapImageChanged = True Then thisQuestionText = thisQuestionText & "Bear in mind that your recent changes will be lost."
            
            ans = msgBoxA(thisQuestionText, vbQuestion + vbYesNo, "Confirm Reload.")
            If ans = 6 Then
                Call copyDockSettingsFile
                
                For useloop = 0 To rdIconMaximum
                    ' we don't bother to read the current record source here as we have already done so above.
    
                    ' read the rdsettings.ini one item up in the list
                    Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile)
    
                    ' .83 DAEB 03/06/2022 rDIConConfig.frm Display the icon we just moved by dragging, one by one rather than the whole map
                    'Call displayIconElement(useloop, picRdMap(useloop), True, 32, True, False)
    
                Next useloop
    
                Call populateRdMap(0) ' show the map from position zero
    
                Call displayIconElement(lastIconChanged, picPreview, True, icoSizePreset, True, False)
    
                btnSet.Enabled = False ' this has to be done at the end
                mapImageChanged = False
                btnCancel.Visible = False
                btnClose.Visible = True
                
            End If
        '

        End If
        
    End If
    
    'settingsTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

settingsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub textCurrentFolder_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        textCurrentFolder.Enabled = False
        textCurrentFolder.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub textCurrIconPath_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        textCurrIconPath.Enabled = False
        textCurrIconPath.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

' .77 DAEB 28/05/2022 rDIConConfig.frm Balloon tooltip on the icon name text box
Private Sub textCurrIconPath_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip textCurrIconPath.hwnd, "This displays the filename of the currently selected icon.", _
                  TTIconInfo, "Help on the Current Folder Path", , , , True

End Sub

' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
Private Sub thumbnailDragTimer_Timer()
    
    thumbnailDragTimerCounter = thumbnailDragTimerCounter + 1

    If thumbnailDragTimerCounter >= 25 Then
        ' .59 DAEB 01/05/2022 rDIConConfig.frm Added manual drag and drop functionality
        If picThumbIconMouseDown = True Then
            picThumbIcon(thumbIndexNo).Drag vbBeginDrag
        End If
        thumbnailDragTimer.Enabled = False
        thumbnailDragTimerCounter = 0
    End If

End Sub

Private Sub filesIconList_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
' .67 DAEB 04/05/2022 rDIConConfig.frm Drag and drop from the filelist to the rdmap
    filesIconList.Drag vbEndDrag
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_Initialize
' Author    : beededea
' Date      : 10/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Initialize()
    On Error GoTo Form_Initialize_Error

    rDIconConfigFormXPosTwips = ""
    rDIconConfigFormYPosTwips = ""

    sdMapState = ""
    sdChkToggleDialogs = ""
    
    ' other variable assignments
    
    moreConfigVisible = False
    iconChanged = False
    dotCount = 0 ' a variable used on the 'working...' button
    rdIconNumber = 0
    rdIconMaximum = 0  ' the final icon in the registry/settings
    thumbnailDragTimerCounter = 0
    rdMapDragTimerCounter = 0
    oldDockSettingsModificationTime = 0 ' max 86 years as a LONG in secs
    'oldDockSettingsModificationTime = Format(Now, "mm/dd/yyy hh:mm:ss")
    
    dragToDockOperating = False

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Activate
' Author    : beededea
' Date      : 21/10/2020
' Purpose   : hides the pop up data source form when anywhere on the main form is clicked upon
'---------------------------------------------------------------------------------------
'
Private Sub Form_Activate()
    On Error GoTo Form_Activate_Error
        
    ' a form activate is called after a drag and drop event, we have a flag check to show whether a drag to dock operation is underway to
    ' avoid setting focus away from the drag and drop event.
    If dragToDockOperating = False Then

        'check the map state on startup
        sdMapState = GetINISetting("Software\SteamyDockSettings", "sdMapState", toolSettingsFile)
        If sdMapState <> "hidden" Then
            Call subBtnArrowDown_Click ' .33
            If rdIconNumber > 0 Then
                 ' give the specific part of the map focus so that after startup any keypresses will operate immediately
                 picRdMap(rdIconNumber).SetFocus  ' < .net
            Else
                 ' give the map focus so that any keypresses will operate immediately
                 picRdMap(0).SetFocus  ' < .net
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : The initial subroutine for the program after the graphics code has done its stuff.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
        
    Dim answer As VbMsgBoxResult: answer = vbNo
    ReDim thumbArray(12) As Integer
    
    On Error GoTo Form_Load_Error
    If debugflg = 1 Then DebugPrint "%" & "Form_Load"
            
    ' vars set to initial values
    
    icoSizePreset = 128
    boxSpacing = 540
    storedIndex = 9999
    busyCounter = 1
    mapImageChanged = False
    rDEnableBalloonTooltips = "1"
    cacheingFlg = True ' .91 DAEB 25/06/2022 rDIConConfig.frm Deleting an icon from the icon thumbnail display causes a cache imageList error. Added cacheingFlg.
    validIconTypes = "*.jpg;*.jpeg;*.bmp;*.ico;*.png;*.tif;*.tiff;*.gif" ' add any remaining types that Rocketdock's code supports
    filesIconList.Pattern = validIconTypes ' set the filter pattern to only show the icon types supported by Rocketdock
    programStatus = "startup"

    
    ' theme variables
    
    classicTheme = False
    storeThemeColour = 13160660  '15790320 = Windows 'modern'  13160660 ' Windows classic
            
    fraProperties.Height = 3630
    fraProperties.Top = 4530
    frameButtons.Top = 7925 ' .43 DAEB 16/04/2022 rdIconConfig.frm increase the whole form height and move the bootom buttons set down
    
    lblBlankText.Visible = False
    
    ' Note: I use the obsolete call statement as it forces brackets when there is a parameter, which looks better to me!
        
    ' Clear all the message box "show again" entries in the registry
    Call clearAllMessageBoxes
        
    ' check for the existence of the rotating busy images
    Call checkBusyImageExistence
        
    ' extracts all the known drive names using Windows APIs
    Call getAllDriveNames(sAllDrives)
                      
    'if the process already exists then kill it
    Call killPreviousInstance ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity
               
    ' get the location of this tool's settings file
    Call getToolSettingsFile
    
    ' check the Windows version and where rocketdock is installed
    Call testWindowsVersion(classicThemeCapable)
    
    ' turn on the timer that tests every 10 secs whether the visual theme has changed
    Call checkClassicThemeCapable ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity
    
'    If IsUserAnAdmin() = 0 And requiresAdmin = True Then
'        msgBoxA "This tool requires to be run as administrator on Windows 8 and above in order to function. Admin access is NOT required on Win7 and below. If you aren't entirely happy with that then you'll need to remove the software now. This is a limitation imposed by Windows itself. To enable administrator access find this tool's exe and right-click properties, compatibility - run as administrator. YOU have to do this manually, I can't do it for you.", vbInformation + vbOKOnly, "Run as Administrator Warning"
'    End If
'
    ' check where rocketdock is installed
    ' Call checkRocketdockInstallation
        
    ' check where steamyDock is installed
    Call checkSteamyDockInstallation
    
    ' set the default path to the icons root
    Call setInitialPath ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity
        
    ' check the main dock settings file exists
    Call locateDockSettingsFile
    
    ' copy the dock settings file to the interim version
    Call copyDockSettingsFile
    
    'do some things for the first and only time
    Call determineFirstRun

    'read the brief config data and all the icons
    Call readIconsAndConfiguration
        
    ' if both docks are installed we need to determine which is the default
    Call checkDefaultDock

    oldDockSettingsModificationTime = FileDateTime(dockSettingsFile)
    dockSettingsRunInterval = 0
                
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties(rDIconConfigForm)

    ' various elements need to have their visibility and size modified prior to display
    Call makeVisibleFormElements
        
    ' dynamically create thumbnail picboxes and sort the captions
    Call createThumbnailLayout
    
    ' dynamically create Map thumbnail picboxes (empty)
    Call createRdMapBoxes
                
    'read this utilties own settings.ini file and set the font
    Call readAndSetUtilityFont ' .30 DAEB 10/04/2021 rDIConConfigForm.frm separate the initial reading of the tool's settings file from the changing of the tool's own font
    
    'the start record is either 0 or set by the dock calling this utility
    Call determineStartRecord
                        
    ' add to the treeview the folders that exist below the RD icons folder and the user-created entries to the folder list top right
    Call addRocketdockFolders
        
    ' add the extra steampunk icon folders to the treeview
    Call setSteampunkLocation
    
    ' add the user custom folder to the treeview
    Call readCustomLocation
            
    ' extract the previously selected default folder in the treeview
    ' open the app settings.ini and read the default folder for the tool to display
    Call readTreeviewDefaultFolder
    
    programStatus = "runtime"
    
    ' display the first icon in the preview window
    Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
    
    'programStatus = "startup"
        
    ' set the theme colour on startup
    Call setThemeSkin(Me) ' .05 17/11/2020 rDIconConfigForm.frm DAEB Added the missing code to read/write the current theme to the tool's own settings file
    
    ' select the thumbnail view rather than the file list view and populate it
    fileIconListPosition = 0
    
    ' .54 DAEB 25/04/2022 rDIConConfig.frm Added rDThumbImageSize saved variable to allow the tool to open the thumbnail explorer in small or large mode
'    rDThumbImageSize = GetINISetting("Software\SteamyDockSettings", "thumbImageSize", toolSettingsFile)
'    If rDThumbImageSize = "" Then rDThumbImageSize = "64"
'    thumbImageSize = Val(rDThumbImageSize)
    Call refreshThumbnailViewPanel
    
    ' we indicate that all changes have been lost when changes to fields are made by the program and not the user
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

        
    ' Creates an incrementally named backup of the settings.ini
    Call fbackupSettings
            
    ' check the registry for Rocketdock usage (mostly obsolete now)
    Call chkTheRegistry
    
    If FExists(interimSettingsFile) Then '
        'get the dockSettingsFile.ini for this icon alone
        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", startRecordNumber, interimSettingsFile
    End If
    
    ' .46 DAEB 16/04/2022 rdIconConfig.frm Made the word Blank visible or not during startup
    If getFileNameFromPath(sFilename) = "blank.png" Then
        lblBlankText.Visible = True
    Else
        lblBlankText.Visible = False
    End If
    
    ' .48 DAEB 20/04/2022 rDIConConfig.frm All tooltips move from IDE into code to allow them to disabled at will
    Call setToolTips

    settingsTimer.Enabled = True
    


   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form rDIconConfigForm"
                
End Sub

Private Sub clearAllMessageBoxes()
    SaveSetting App.EXEName, "Options", "Show message" & "preButtonClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "btnAdd_Click", 0
    SaveSetting App.EXEName, "Options", "Show message" & "btnSet_Click", 0
    SaveSetting App.EXEName, "Options", "Show message" & "preMapPageUpDown_Press", 0
    SaveSetting App.EXEName, "Options", "Show message" & "btnHomeRdMap", 0
    SaveSetting App.EXEName, "Options", "Show message" & "btnEndRdMap", 0
    SaveSetting App.EXEName, "Options", "Show message" & "preButtonClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "reOrderRdMap", 0
    SaveSetting App.EXEName, "Options", "Show message" & "rdMapRefresh_Click", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuHelpPdf_click", 0
    SaveSetting App.EXEName, "Options", "Show message" & "deleteRdMapPosition", 0
End Sub


Private Sub checkBusyImageExistence()
    Dim useloop As Integer: useloop = 0
    Dim busyCounter As Integer: busyCounter = 0
    Dim ans As VbMsgBoxResult: ans = vbYesNo
        
    For useloop = 1 To 6
        busyCounter = useloop
        If Not FExists(App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg") Then
            ans = msgBoxA("This file is missing - " & App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg", vbExclamation + vbOKOnly, "Checking certain files exist.")
            If ans = 6 Then
                End
            End If
        End If
    Next useloop

    For useloop = 1 To 6
        busyCounter = useloop
        If Not FExists(App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg") Then
            ans = msgBoxA("This file is missing - " & App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg", vbExclamation + vbOKOnly, "Checking certain files exist.")
            If ans = 6 Then
                End
            End If
        End If
    Next useloop
End Sub


' .48 DAEB 20/04/2022 rDIConConfig.frm All tooltips move from IDE into code to allow them to disabled at will
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

' .48 DAEB 20/04/2022 rDIConConfig.frm All tooltips move from IDE into code to allow them to disabled at will
'---------------------------------------------------------------------------------------
' Procedure : setToolTips
' Author    : beededea
' Date      : 21/04/2022
' Purpose   : Enable or disable tooltips
'---------------------------------------------------------------------------------------
'
Private Sub setToolTips()

    On Error GoTo setToolTips_Error

    If chkToggleDialogs.Value = 0 Then
        Call DestroyToolTip ' destroys the current tooltip
        rDEnableBalloonTooltips = "0" ' this is the flag used to determine wheter a new balloon toltip is generated
    
        btnMapPrev.ToolTipText = "This will scroll the icon map to the left so that you can view additional icons"
        btnMapNext.ToolTipText = "This will scroll the icon map to the right so that you can view additional icons"
        btnArrowUp.ToolTipText = "Hide the map"
        rdMapRefresh.ToolTipText = "Refresh the icon map"
        FrameFolders.ToolTipText = "The current list of known icon folders"
        cmbDefaultDock.ToolTipText = "Indicates the default dock, Rocketdock or SteamyDock. Cannot be changed here but only in the dock settings utility."
        btnSettingsDown.ToolTipText = "Show where the details are being read from and saved to."
        btnSettingsUp.ToolTipText = "Hide the registry form showing where details are being read from and saved to."
        folderTreeView.ToolTipText = "These are the icon folders available to Rocketdock"
        textCurrentFolder.ToolTipText = "The selected folder path"
        btnRemoveFolder.ToolTipText = "This button can remove a custom folder from the treeview above"
        btnAddFolder.ToolTipText = "Select a target folder to add to the treeview list above"
        btnArrowDown.ToolTipText = "Show the Dock Map"
        fraProperties.ToolTipText = "The Icon Properties Window"
        txtSecondApp.ToolTipText = "Any second program to run after the main program initiation will be shown here"
        btnSecondApp.ToolTipText = "Press to select a second program to run after the main program initiation"
        chkAutoHideDock.ToolTipText = "Automatically hides the dock for the default hiding period when the program is initiated"
        chkQuickLaunch.ToolTipText = "Launch an application before the bounce has completed"
        picMoreConfigDown.ToolTipText = "Shows extra configuration items"
        chkConfirmDialogAfter.ToolTipText = "Shows Confirmation Dialog after the command has run."
        chkConfirmDialog.ToolTipText = "Adds a Confirmation Dialog prior to the command running allowing you to say yes or no at runtime"
        btnIconSelect.ToolTipText = "Press to select an icon manually."
        picBusy.ToolTipText = "The program is doing something..."
        btnSet.ToolTipText = "Sets the icon characteristics but you will need to press the save and restart button to make it 'fix' on the running dock."
        txtCurrentIcon.ToolTipText = "Double click on an image above to set the current icon"
        btnSelectStart.ToolTipText = "Select a start folder"
        btnTarget.ToolTipText = "Press to select a target file (or right click for a folder)"
        chkRunElevated.ToolTipText = "When this checkbox is ticked, the associated app will run with elevated privileges, ie. as administrator. Some programs require this in order to operate."
        cmbOpenRunning.ToolTipText = "Choose what to do if the chosen app is already running"
        cmbRunState.ToolTipText = "Window mode for the program to operate within"
        txtArguments.ToolTipText = "Add any additional arguments that the target file operation requires, eg. -s -t 00 -f "
        txtStartIn.ToolTipText = "If the operation needs to be performed in a particular folder select it here"
        txtTarget.ToolTipText = "The target you wish to run, a file or a folder"
        txtLabelName.ToolTipText = "The name of the icon as it appears on the dock"
        lblSecondApp.ToolTipText = "If you want to run a second program after the program initiation, select it here"
        lblchkAutoHideDock.ToolTipText = "Automatically hides the dock for the default hiding period when the program is initiated"
        lblQuickLaunch.ToolTipText = "Launch an application before the bounce has completed"
        lblConfirmDialogAfter.ToolTipText = "Shows Confirmation Dialog after the command has run."
        lblConfirmDialog.ToolTipText = "Adds a Confirmation Dialog prior to the command running allowing you to say yes or no at runtime"
        lblRdIconNumber.ToolTipText = "This is the dock icon number one."
        lblRunElevated.ToolTipText = "If you want extra options to appear when you right click on an icon, enable this checkbox"
        lblRunElevated.ToolTipText = "If you want extra options to appear when you right click on an icon, enable this checkbox"
        btnPrev.ToolTipText = "Select the previous icon"
        btnNext.ToolTipText = "select the next icon"
        picPreview.ToolTipText = "This is the currently selected icon scaled to fit the preview box"
        lblBlankText.ToolTipText = "This is the dock icon number one."
        frameButtons.ToolTipText = "Makes a whole NEW dock - use with care!"
        btnBackup.ToolTipText = "Backup or restore using a version of bkpSettings.ini"
        btnSaveRestart.ToolTipText = "A save and restart of the dock is required when any icon changes have been made"
        btnCancel.ToolTipText = "Cancel the current operation"
        btnClose.ToolTipText = "Close the window"
        btnHelp.ToolTipText = "Help on this utility"
        chkToggleDialogs.ToolTipText = "This will toggle on/off most of the information pop-ups when running this utility. ie. confirmation on saves and deletes"
        btnDefaultIcon.ToolTipText = "Not implemented yet"
        frameIcons.ToolTipText = "Thumbnail or File Viewer Window"
        comboIconTypesFilter.ToolTipText = "Filter icon types to display"
        btnKillIcon.ToolTipText = "Delete the currently selected icon file above. Use wisely!"
        btnAdd.ToolTipText = "Set the current selected icon into the dock (double-click on the icon)"
        btnRefresh.ToolTipText = "Refresh the Icon List"
        textCurrIconPath.ToolTipText = "Shows the selected icon file name"
        picFrameThumbs.ToolTipText = "Double-click an icon to set it into the dock"
        'picThumbIcon.ToolTipText = "This is the currently selected icon scaled to fit the preview box"
        filesIconList.ToolTipText = "Select an icon, double-click to set"
        btnThumbnailView.ToolTipText = "View as thumbnails"
        btnFileListView.ToolTipText = "View as a file listing"
        btnGetMore.ToolTipText = "Click to install more icons"
        btnGenerate.ToolTipText = "Makes a whole NEW dock - use with care!"
        picMoreConfigUp.ToolTipText = "Hides the extra configuration section"
        picHideConfig.ToolTipText = "Hides the extra configuration section"
        lblDisabled.ToolTipText = "Disables this icon in the dock making it non-functional"
        chkDisabled.ToolTipText = "Disables this icon in the dock making it non-functional"
        txtAppToTerminate.ToolTipText = "Any program that must be terminated prior to the main program initiation will be shown here"
        btnAppToTerminate.ToolTipText = "Any program that must be terminated prior to the main program initiation will be shown here"
    Else
        rDEnableBalloonTooltips = "1"
        
        btnMapPrev.ToolTipText = ""
        btnMapNext.ToolTipText = ""
        btnPrev.ToolTipText = ""
        btnNext.ToolTipText = ""
        btnArrowUp.ToolTipText = ""
        rdMapRefresh.ToolTipText = ""
        FrameFolders.ToolTipText = ""
        cmbDefaultDock.ToolTipText = ""
        btnSettingsDown.ToolTipText = ""
        btnSettingsUp.ToolTipText = ""
        folderTreeView.ToolTipText = ""
        textCurrentFolder.ToolTipText = ""
        btnRemoveFolder.ToolTipText = ""
        btnAddFolder.ToolTipText = ""
        btnArrowDown.ToolTipText = ""
        fraProperties.ToolTipText = ""
        txtSecondApp.ToolTipText = ""
        btnSecondApp.ToolTipText = ""
        chkAutoHideDock.ToolTipText = ""
        chkQuickLaunch.ToolTipText = ""
        picMoreConfigDown.ToolTipText = ""
        chkConfirmDialogAfter.ToolTipText = ""
        chkConfirmDialog.ToolTipText = ""
        btnIconSelect.ToolTipText = ""
        picBusy.ToolTipText = ""
        btnSet.ToolTipText = ""
        txtCurrentIcon.ToolTipText = ""
        btnSelectStart.ToolTipText = ""
        btnTarget.ToolTipText = ""
        chkRunElevated.ToolTipText = ""
        cmbOpenRunning.ToolTipText = ""
        cmbRunState.ToolTipText = ""
        txtArguments.ToolTipText = ""
        txtStartIn.ToolTipText = ""
        txtTarget.ToolTipText = ""
        txtLabelName.ToolTipText = ""
        lblSecondApp.ToolTipText = ""
        lblchkAutoHideDock.ToolTipText = ""
        lblQuickLaunch.ToolTipText = ""
        lblConfirmDialogAfter.ToolTipText = ""
        lblConfirmDialog.ToolTipText = ""
        lblRdIconNumber.ToolTipText = ""
        lblRunElevated.ToolTipText = ""
        lblRunElevated.ToolTipText = ""
        framePreview.ToolTipText = ""
        btnNext.ToolTipText = ""
        picPreview.ToolTipText = ""
        lblBlankText.ToolTipText = ""
        frameButtons.ToolTipText = ""
        btnBackup.ToolTipText = ""
        btnSaveRestart.ToolTipText = ""
        btnCancel.ToolTipText = ""
        btnClose.ToolTipText = ""
        btnHelp.ToolTipText = ""
        chkToggleDialogs.ToolTipText = ""
        btnDefaultIcon.ToolTipText = ""
        frameIcons.ToolTipText = ""
        comboIconTypesFilter.ToolTipText = ""
        btnKillIcon.ToolTipText = ""
        btnAdd.ToolTipText = ""
        btnRefresh.ToolTipText = ""
        textCurrIconPath.ToolTipText = ""
        picFrameThumbs.ToolTipText = ""
        filesIconList.ToolTipText = ""
        btnThumbnailView.ToolTipText = ""
        btnFileListView.ToolTipText = ""
        btnGetMore.ToolTipText = ""
        btnGenerate.ToolTipText = ""
        picMoreConfigUp.ToolTipText = ""
        picHideConfig.ToolTipText = ""
        lblDisabled.ToolTipText = ""
        chkDisabled.ToolTipText = ""
        txtAppToTerminate.ToolTipText = ""
        btnAppToTerminate.ToolTipText = ""
        lblAppToTerminate.ToolTipText = ""
    End If

    On Error GoTo 0
    Exit Sub

setToolTips_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setToolTips of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 09/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeVisibleFormElements()

'    Dim virtualScreenWidthTwips As Long: virtualScreenWidthTwips = 0
'    Dim virtualScreenHeightTwips As Long: virtualScreenHeightTwips = 0
'    Dim MonitorCount As Integer: MonitorCount = 0

    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0
    
    ' menus are enabled in the IDE so that it is easier to maintain the menus using point and click,
    ' therefore they need to be disabled at runtime otherwise they will appear as a bar menu across the top
    ' of the screen.
    
    On Error GoTo makeVisibleFormElements_Error

    mnuTrgtMenu.Visible = False
    rdMapMenu.Visible = False
    mnupopmenu.Visible = False
    thumbmenu.Visible = False
    
    screenHeightTwips = GetDeviceCaps(rDIconConfigForm.hdc, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(rDIconConfigForm.hdc, HORZRES) * screenTwipsPerPixelX ' replaces buggy screen.width
    
    ' set the current position of the utility according to previously stored positions
    
    ' read the form X/Y params from the toolSettings.ini
'    rDIconConfigFormXPosTwips = GetINISetting("Software\SteamyDockSettings", "IconConfigFormXPos", toolSettingsFile)
'    rDIconConfigFormYPosTwips = GetINISetting("Software\SteamyDockSettings", "IconConfigFormYPos", toolSettingsFile)
'
'    ' if a current location not stored then position to the middle of the screen
'    If rDIconConfigFormXPosTwips <> "" Then
'        rDIconConfigForm.Left = Val(rDIconConfigFormXPosTwips)
'    Else
'        rDIconConfigForm.Left = screenWidthTwips / 2 - rDIconConfigForm.Width / 2
'    End If
'
'    If rDIconConfigFormYPosTwips <> "" Then
'        rDIconConfigForm.Top = Val(rDIconConfigFormYPosTwips)
'    Else
'        rDIconConfigForm.Top = Screen.Height / 2 - rDIconConfigForm.Height / 2
'    End If
'

    ' how many monitors?
    'MonitorCount = fGetMonitorCount
    
'    virtualScreenHeightTwips = fVirtualScreenHeight
'    virtualScreenWidthTwips = fVirtualScreenWidth

    ' read the form's saved X/Y params from the toolSettings.ini in twips and convert to pixels
    formLeftPixels = Val(GetINISetting("Software\SteamyDockSettings", "IconConfigFormXPos", toolSettingsFile)) / screenTwipsPerPixelX
    formTopPixels = Val(GetINISetting("Software\SteamyDockSettings", "IconConfigFormYPos", toolSettingsFile)) / screenTwipsPerPixelY

    Call adjustFormPositionToCorrectMonitor(Me.hwnd, formLeftPixels, formTopPixels)
 
    ' id the virtualScreenWidthTwips <> currentScreenWidthTwips then  we have more than one monitor
    
    ' if only one monitor - is it offscreen?
    ' if two monitors - is it on monitor one or monitor two? If neither center it on monitor one
    
    'As a result of the menus being enabled in the IDE the form height is disturbed so it needs to be corrected manually.
    ' .99 DAEB 26/06/2022 rDIConConfig.frm With the round borders of Win 11, there is insufficient space from the frame to the border, Windows cuts it off arbitrarily. Extend.
    rDIconConfigForm.Height = 9495 '9780
    
    ' if Windows 10/11 then add 250 twips to the bottom of the main form
    If Left$(LCase$(windowsVersionString), 10) = "windows 10" Then
        Me.Height = Me.Height + 100
    End If
        
    ' state and position of a few manually placed controls (easier here than in the IDE)
    picRdThumbFrame.Visible = False
    
    previewFrameGotFocus = True
    
    ' the dialog toggle unavailable to RD '.nn
'    If defaultDock = 0 Then
'        chkConfirmDialog.Visible = False
'        lblConfirmDialog.Visible = False
'        chkConfirmDialogAfter.Visible = False '.nn Added new dialog box for providing a message after a program has run
'        lblConfirmDialogAfter.Visible = False
'
'        chkQuickLaunch.Visible = False '.nn Added new check box to allow a quick launch of the chosen app
'        lblQuickLaunch.Visible = False
'
'        chkAutoHideDock.Visible = False '.nn Added new check box to allow autohide of the dock after launch of the chosen app
'        chkAutoHideDock.Visible = False
'
'        chkRunElevated.Enabled = True
'        lblRunElevated.ToolTipText = "If you want extra options to appear when you right click on an icon, enable this checkbox, Rocketdock only."
'        'lblRunElevated.Enabled = True
'        lblRunElevated.Enabled = True
'
'        cmbOpenRunning.Enabled = True
'        lblOpenRunning.Enabled = True
'
'        btnGenerate.Visible = False
'    Else
        chkConfirmDialog.Visible = True
        lblConfirmDialog.Visible = True
        chkConfirmDialogAfter.Visible = True
        lblConfirmDialogAfter.Visible = True
        
        chkQuickLaunch.Visible = True '.nn Added new check box to allow a quick launch of the chosen app
        lblQuickLaunch.Visible = True
        
        
        chkAutoHideDock.Visible = True '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        chkAutoHideDock.Visible = True
        
        chkRunElevated.Enabled = False
        lblRunElevated.ToolTipText = "Not available, SteamyDock does not provide the application context menu as the important items are already in the standard menu."
        
        chkRunElevated.Enabled = True
        lblRunElevated.Enabled = True
        
'        cmbOpenRunning.Enabled = False
'        lblOpenRunning.Enabled = False
    
        btnGenerate.Visible = True
    
'    End If
    
   On Error GoTo 0
   Exit Sub

makeVisibleFormElements_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeVisibleFormElements of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : determineStartRecord
' Author    : beededea
' Date      : 15/06/2020
' Purpose   : the start record can be extracted from a parameter passed from the dock
'---------------------------------------------------------------------------------------
'
Private Sub determineStartRecord() ' .NET
    
    Dim testingNo As Integer: testingNo = 9999 ' initialise as an invalid amount for the program

    On Error GoTo determineStartRecord_Error
   
    rdIconNumber = 0 ' normally starts at zero
    
    Randomize
    
    'parse the command line
    If Command <> "" Then
        If IsNumeric(Command) Then
            testingNo = Val(Command)
            If testingNo > 0 And testingNo <= rdIconMaximum Then
                'MsgBox testingNo
                rdIconNumber = testingNo
            Else
                rdIconNumber = rdIconMaximum * Rnd() + 1
            End If
        End If
    Else
        rdIconNumber = rdIconMaximum * Rnd() + 1
    End If
    
    'MsgBox "rdIconNumber = " & rdIconNumber
    ' set the very large icon record number displayed on the main form
    lblRdIconNumber.Caption = rdIconNumber + 1
    lblRdIconNumber.ToolTipText = "This is the dock icon number " & Str$(rdIconNumber) + 1
    
    ' .47 DAEB 16/04/2022 rdIconConfig.frm Added StartRecordNumber
    startRecordNumber = rdIconNumber
   
   On Error GoTo 0
   Exit Sub

determineStartRecord_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineStartRecord of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : setMnuPath
' Author    : beededea
' Date      : 10/04/2020
' Purpose   : set the menu path text
'---------------------------------------------------------------------------------------
'
'Private Sub setMnuPath()
'
'    Dim chkFolder As String
'
'    ' read the tool settings file
'   On Error GoTo setMnuPath_Error
'
'    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        chkFolder = GetINISetting("Software\SteamyDockSettings", "rocketDockLocation", toolSettingsFile)
'        If chkFolder <> vbNullString Then
'            If FExists(chkFolder & "\rocketDock.exe") Then
'                rdAppPath = chkFolder
'                mnuRocketDock.Caption = "Rocketdock location - " & chkFolder & " - click to change"
'            End If
'        End If
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'setMnuPath_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMnuPath of Form rDIconConfigForm"
'
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Public Sub setThemeColour(ByRef thisForm As Form)
    
    Dim SysClr As Long: SysClr = 0
    
    On Error GoTo setThemeColour_Error
    If debugflg = 1 Then DebugPrint "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(thisForm, 212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        rDSkinTheme = "dark" ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file

    Else
        'MsgBox "Windows Alternate Theme detected"
        SysClr = GetSysColor(COLOR_BTNFACE)
        If SysClr = 13160660 Then
            Call setThemeShade(thisForm, 212, 208, 199)
            rDSkinTheme = "light" ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
        Else ' 15790320
            Call setThemeShade(thisForm, 240, 240, 240)
            rDSkinTheme = "dark" ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
        End If

    End If

    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_KeyUp
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub Form_KeyUp(ByRef KeyCode As Integer, ByRef Shift As Integer)
'    ' Simple example of pasting file names on a drag drop
'   On Error GoTo Form_KeyUp_Error
'   If debugflg = 1 Then DebugPrint "%" & "Form_KeyUp"
'
'    If KeyCode = vbKeyV Then
'
'        If (Shift And vbCtrlMask) = vbCtrlMask Then
'            ' use class to load 1st file that was pasted, if any & if more than one
'            ' Unicode filenames supported
'            If cImage.LoadPicturePastedFiles(1, 256, 256) = False Then
'                ' couldn't load anything from the files, maybe image itself was pasted
'                If cImage.LoadPictureClipBoard = False Then
'                    MsgBox "Failed to load whatever was placed in the clipboard", vbInformation + vbOKOnly
'                    Exit Sub
'                End If
'            End If
'
'            If Not cShadow Is Nothing Then
'                '
'            Else
'                Call refreshPicBox(picPreview, 256)
'            End If
'            'ShowImage False, True
'
'        End If
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'Form_KeyUp_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyUp of Form rDIconConfigForm"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnAdd_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAdd_Click()
   Dim ans As VbMsgBoxResult: ans = vbNo
   Dim itemno As Integer: itemno = 0
    
   On Error GoTo btnAdd_Click_Error
   If debugflg = 1 Then DebugPrint "%" & "btnAdd_Click"
    
   mapImageChanged = False
   If storedIndex <> 9999 Then
       If chkToggleDialogs.Value = 1 Then
       
            ' .65 DAEB 04/05/2022 rDIConConfig.frm Use the underlying control index rather than that stored in the array
            itemno = filesIconList.ListIndex
            
            ans = msgBoxA(" Add this icon - " & filesIconList.List(itemno) & " - at position " & rdIconNumber + 1 & " in the dock? ", vbQuestion + vbYesNo, "Confirm dragging and dropping to the dock", True, "btnAdd_Click")
            If ans = 6 Then
                Call changeMapImage
            End If
        Else
            Call changeMapImage
        End If
    End If

   On Error GoTo 0
   Exit Sub

btnAdd_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAdd_Click of Form rDIconConfigForm"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : changeMapImage
' Author    : beededea
' Date      : 26/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub changeMapImage()
   Dim itemno As Integer: itemno = 0
   
   On Error GoTo changeMapImage_Error
   If debugflg = 1 Then DebugPrint "%changeMapImage"

    mapImageChanged = True
    
    If srcDragControl = "picThumbIcon" Then itemno = thumbArray(storedIndex)
    If srcDragControl = "filesIconList" Then itemno = filesIconList.ListIndex
    
    filesIconList.ListIndex = (itemno) ' this does a click the item in the underlying file list box
    filesIconList_DblClick

   On Error GoTo 0
   Exit Sub

changeMapImage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeMapImage of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnGenerate_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'%ProgramData%\Microsoft\Windows\Start Menu\Programs
'%AppData%\Microsoft\Windows\Start Menu\Programs


Private Sub btnGenerate_Click()
    'Dim ans As VbMsgBoxResult: ans = vbNo
    
    'Call btnArrowDown_Click ' populate the dock
    On Error GoTo btnGenerate_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnGenerate_Click"
    
    If defaultDock = 1 Then
        formSoftwareList.Show
    End If

   On Error GoTo 0
   Exit Sub

btnGenerate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGenerate_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnNext_KeyDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNext_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo btnNext_KeyDown_Error
   If debugflg = 1 Then DebugPrint "%btnNext_KeyDown"

    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

btnNext_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnNext_KeyDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPrev_KeyDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrev_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo btnPrev_KeyDown_Error
   If debugflg = 1 Then DebugPrint "%btnPrev_KeyDown"

    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

btnPrev_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrev_KeyDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRemoveFolder_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRemoveFolder_Click()
    ' remove the chosen folder to the treeview
    
    ' find the chosen node's parent
    ' if the parent is nothing then this is the top level
    ' if the parent exists then look above again
   Dim a As String: a = ""
   
   ' this next line is the MSCOMTCL.OCX usage of a Treeview node
   'Dim tNode As Node
   
   ' this is Krool's treeview replacement of the Treeview node
   Dim tNode As CCRTreeView.TvwNode
   
   On Error GoTo btnRemoveFolder_Click_Error
   If debugflg = 1 Then DebugPrint "%" & "btnRemoveFolder_Click"

    Set tNode = folderTreeView.SelectedItem
    
    If tNode Is Nothing Then Exit Sub ' in the unlikely event that no node is selected in the treeview

    If Not tNode.Parent Is Nothing Then
        Set tNode = tNode.Parent '  move up level
        Do Until tNode.Parent Is Nothing   ' if Nothing, then done
            If Not tNode.NextSibling Is Nothing Then
                Set tNode = tNode.NextSibling
                Exit Do
            End If
            Set tNode = tNode.Parent ' move up again
        Loop
    End If
    
    If tNode Is Nothing Then
        'Set treeView.selectedItem = treeView.Nodes(1).Root
        a = folderTreeView.Nodes(1).Root
    Else
        a = tNode
        'Set treeView.selectedItem = tNode
    End If
    
    If a = "my collection" Then
            msgBoxA "Cannot remove SteamyDock Enhanced Settings Utility sub-folders from the treeview.", vbExclamation + vbOKOnly
            Exit Sub
    End If
        
    If a = "icons" Then
        msgBoxA "Cannot remove SteamyDock's own sub-folders from the treeview, you have to delete the folders from Windows first then re-run this utility.", vbExclamation + vbOKOnly
        Exit Sub
    End If
        
    If folderTreeView.SelectedItem.Key = vbNullString Then
        Exit Sub
    End If
        
    If a = "custom folder" And Not folderTreeView.SelectedItem = "custom folder" Then
        msgBoxA "Cannot remove custom sub-folders from the treeview, try again at the root.", vbExclamation + vbOKOnly
        Exit Sub
    Else
        ' do the delete!
    End If
        
    folderTreeView.Nodes.Remove folderTreeView.SelectedItem.Key
    
    'write the folder to the rocketdock settings file
    'eg. rDCustomIconFolder=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\SteamyDock\DockSettings", "rDCustomIconFolder", "?", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", dockSettingsFile
    
   On Error GoTo 0
   Exit Sub

btnRemoveFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRemoveFolder_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : createThumbnailLayout
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createThumbnailLayout()
    ' code to dynamically create the labels that sit on top of the thumbnail images
    ' the problem is that labels cannot sit above the picboxes in zorder, they are windowless controlls
    ' and will always position underneath unless they are placed onto frames
    ' the code creates a frame and then puts the label onto the frame and attaches them together
    
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo createThumbnailLayout_Error
    If debugflg = 1 Then DebugPrint "%" & "createThumbnailLayout"

    storeLeft = 165
'    fraThumbLabel(0).ZOrder
'    fraThumbLabel(0).BorderStyle = 0
'    fraThumbLabel(0).Visible = True
         
    ' dynamically create the picture boxes for the thumbnails
    For useloop = 1 To 11 ' 0 is the template
        Load picFraPicThumbIcon(useloop) ' the frame for the thumb icon image to reside inside
        Load picThumbIcon(useloop) ' the thumb icon image
        Load fraThumbLabel(useloop)
        Load lblThumbName(useloop)
        
        Set picThumbIcon(useloop).Container = picFraPicThumbIcon(useloop)
        'Set fraThumbLabel(useloop).Container = picFraPicThumbIcon(useloop)
        Set lblThumbName(useloop).Container = fraThumbLabel(useloop)
    Next useloop
    
    ' .56 DAEB 25/04/2022 rDIConConfig.frm 1st run of the thumbnail view window is done using the old method and it comes out incorrectly.
    'Call placeThumbnailPicboxes(64)
    
    ' the labels for the smaller thumbnail icon view
'    For useloop = 0 To 11
'        fraThumbLabel(useloop).Visible = False
'    Next useloop
    
   On Error GoTo 0
   Exit Sub

createThumbnailLayout_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createThumbnailLayout of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createRdMapBoxes
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRdMapBoxes()
    Dim useloop As Integer: useloop = 0

   On Error GoTo createRdMapBoxes_Error
   If debugflg = 1 Then DebugPrint "%" & "createRdMapBoxes"

    storeLeft = boxSpacing
    ' dynamically create more picture boxes to the maximum number of icons
    For useloop = 1 To rdIconMaximum
        Load picRdMap(useloop)
        picRdMap(useloop).Width = 500
        picRdMap(useloop).Height = 500
        storeLeft = storeLeft + boxSpacing
        picRdMap(useloop).Left = storeLeft
        picRdMap(useloop).Top = 15
        picRdMap(useloop).Visible = True
        picRdMap(useloop).AutoRedraw = True
    Next useloop
    
   On Error GoTo 0
   Exit Sub

createRdMapBoxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRdMapBoxes of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : determineFirstRun
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read this utilties' own settings.ini file and determine first run or not
'---------------------------------------------------------------------------------------
'
Private Sub determineFirstRun()
    Dim sfirst As String: sfirst = "True" ' .31 DAEB 10/04/2021 rDIConConfigForm.frm initialise the value - rather important

    On Error GoTo determineFirstRun_Error
    If debugflg = 1 Then DebugPrint "%" & "determineFirstRun"
    
   
    If Not FExists(toolSettingsFile) Then Exit Sub ' does the tool's own settings.ini exist?
    
    'test to see if the tool has ever been run before
    sfirst = GetINISetting("Software\SteamyDockSettings", "First", toolSettingsFile)
    
    If sfirst = "True" Then
    
'        sfirst = "False"
'
'        ' insert at the final position
'        ' a link with the rocketdockSettings icon and the target
'        ' is the app.path
'
'        sFilename = "iconSettings\my collection\SteamyRocket.png" ' the default Rocketdock filename for a blank item
'
'
'        sTitle = "Rocket Settings"
'        sCommand = App.Path & "\" & "iconsettings.exe" ' 17/11/2020    .04 DAEB Replaced all occurrences of rocket1.exe with iconsettings.exe
'
'        sArguments = vbNullString
'        sWorkingDirectory = App.Path
'        sShowCmd = "1" ' .34 DAEB 05/05/2021 rDIConConfigForm.frm The value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
'
'        sOpenRunning = "0"
'        sUseContext = "0"
'        sUseDialog = "0"
'        sUseDialogAfter = "0" ' .03 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear"
'        sQuickLaunch = "0" '.nn Added new check box to allow a quick launch of the chosen app
'        sAutoHideDock = "0"   '.nn Added new check box to allow autohide of the dock after launch of the chosen app
'        sSecondApp = ""  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
'        sRunSecondAppBeforehand = ""
'
'        sAppToTerminate = ""
'        sDisabled = "0"
'
'        rdIconMaximum = rdIconMaximum + 1
        
        'write the rdsettings file
        'writeSettingsIni (rdIconMaximum)
'        Call writeIconSettingsIni("Software\SteamyDock\IconSettings" & "\Icons", rdIconMaximum, interimSettingsFile)
'
'        'amend the count in both the rdSettings.ini
'        PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", rdIconMaximum, interimSettingsFile

        'write the updated test of first run to false
        PutINISetting "Software\SteamyDockSettings", "First", sfirst, toolSettingsFile
    
'        'filecopy the rocketdockSettings png to the rocketdock icons folder
'        If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
'            If FExists(App.Path & "\" & "SteamyRocket.png") Then
'                FileCopy App.Path & "\" & "SteamyRocket.png", rdAppPath & "\icons\" & "SteamyRocket.png"
'            End If
'        End If
    End If
    

   On Error GoTo 0
   Exit Sub

determineFirstRun_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineFirstRun of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : copyDockSettingsFile
' Author    : beededea
' Date      : 05/03/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub copyDockSettingsFile()
    ' variables declared
    Dim dockSettingsDir As String: dockSettingsDir = vbNullString
        
    On Error GoTo copyDockSettingsFile_Error
    
    dockSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\steamyDock" ' just for this user alone


    interimSettingsFile = dockSettingsDir & "\interimSettings.ini" ' the third config option for steamydock alone

    If FExists(dockSettingsFile) Then
        
        ' copy the original settings file to a duplicate that we will operate upon until saved
        FileCopy dockSettingsFile, interimSettingsFile

    End If

    On Error GoTo 0
    Exit Sub

copyDockSettingsFile_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure copyDockSettingsFile of Form rDIconConfigForm"
            Resume Next
          End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : readIconsAndConfiguration
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : read the configurations, settings.ini, registry and dockSettings.ini
'---------------------------------------------------------------------------------------
'
Private Sub readIconsAndConfiguration()
    ' select the settings source STARTS
    Dim useloop As Integer: useloop = 0
    Dim location As String: location = 0
    
    'interimSettingsFile = App.Path & "\rdSettings.ini" ' a copy of the settings file that we work on
            
    On Error GoTo readIconsAndConfiguration_Error
      
    If FExists(interimSettingsFile) Then
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", interimSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", interimSettingsFile)
        'rDRunAppInterval = GetINISetting("Software\SteamyDock\DockSettings", "RunAppInterval", interimSettingsFile)
        'rDAlwaysAsk = GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", interimSettingsFile)
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", interimSettingsFile)
        
        sDShowIconSettings = GetINISetting("Software\SteamyDock\DockSettings", "ShowIconSettings", interimSettingsFile)
    End If
    
    ' .72 DAEB 16/05/2022 rDIConConfig.frm Validate the settings read from the settings file
    If rDGeneralReadConfig = "" Then rDGeneralReadConfig = "True"
    If rDGeneralWriteConfig = "" Then rDGeneralWriteConfig = "True"
    If rDDefaultDock = "" Then rDDefaultDock = "steamydock"
    
    'final check to be sure that we aren't using an incorrectly set dockSettings.ini file when in fact RD has never actually been installed
    If rocketDockInstalled = False And RDregistryPresent = False Then
        rDGeneralReadConfig = "True"
    End If
    
    If steamyDockInstalled = True And defaultDock = 1 And rDGeneralReadConfig = "True" Then ' it will always exist even if not used
        ' read the dock settings from the new configuration file into the existing interim settings file
'        chkReadRegistry.Value = 0
'        chkReadSettings.Value = 0
'        chkReadConfig.Value = 1
        
        ' read the count from the settings file and find the last icon
        theCount = Val(GetINISetting("Software\SteamyDock\IconSettings\Icons", "count", interimSettingsFile))
        ' validate
        
        rdIconMaximum = theCount - 1
        
        ' copy the original configs out of the registry and into a settings file that we will operate upon

        For useloop = 0 To rdIconMaximum
            ' get the relevant entries from the registry
            location = "Software\SteamyDock\IconSettings\Icons"
            
            sFilename = GetINISetting(location, useloop & "-FileName", interimSettingsFile)
            sFileName2 = GetINISetting(location, useloop & "-FileName2", interimSettingsFile)
            sTitle = GetINISetting(location, useloop & "-Title", interimSettingsFile)
            sCommand = GetINISetting(location, useloop & "-Command", interimSettingsFile)
            sArguments = GetINISetting(location, useloop & "-Arguments", interimSettingsFile)
            sWorkingDirectory = GetINISetting(location, useloop & "-WorkingDirectory", interimSettingsFile)
            sShowCmd = GetINISetting(location, useloop & "-ShowCmd", interimSettingsFile)
            sOpenRunning = GetINISetting(location, useloop & "-OpenRunning", interimSettingsFile)
            sRunElevated = GetINISetting(location, useloop & "-RunElevated", interimSettingsFile)
            sIsSeparator = GetINISetting(location, useloop & "-IsSeparator", interimSettingsFile)
            sUseContext = GetINISetting(location, useloop & "-UseContext", interimSettingsFile)
            sDockletFile = GetINISetting(location, useloop & "-DockletFile", interimSettingsFile)
             
            If defaultDock = 1 Then
                sUseDialog = GetINISetting(location, useloop & "-UseDialog", interimSettingsFile)
                sUseDialogAfter = GetINISetting(location, useloop & "-UseDialogAfter", interimSettingsFile) ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
                sQuickLaunch = GetINISetting(location, useloop & "-QuickLaunch", interimSettingsFile) '.nn Added new check box to allow a quick launch of the chosen app
                sAutoHideDock = GetINISetting(location, useloop & "-AutoHideDock", interimSettingsFile)  '.nn Added new check box to allow autohide of the dock after launch of the chosen app
                sSecondApp = GetINISetting(location, useloop & "-SecondApp", interimSettingsFile) ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
                
                sRunSecondAppBeforehand = GetINISetting(location, useloop & "-RunSecondAppBeforehand", interimSettingsFile)
                sAppToTerminate = GetINISetting(location, useloop & "-AppToTerminate", interimSettingsFile)
                
                sDisabled = GetINISetting(location, useloop & "-Disabled", interimSettingsFile)      ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
            End If
            
            ' write the rocketdock alternative settings.ini
            'writeSettingsIni (useloop)
            'Call writeIconSettingsIni("Software\SteamyDock\IconSettings" & "\Icons", useloop, interimSettingsFile)

        Next useloop
        
        ' make a backup of the rdSettings.ini after the intermediate file has been created
        'Call fbackupSettings("")
    End If
    
'    If rDGeneralReadConfig = "False" Then
'        ' read the Rocketdock settings from INI or from registry
'        Call readRocketDockSettings
'    End If

    ' .71 DAEB 16/05/2022 rDIConConfig.frm Move the reading of recent settings into the main read configuration procedure STARTS
    rDThumbImageSize = GetINISetting("Software\SteamyDockSettings", "thumbImageSize", toolSettingsFile)
    If rDThumbImageSize = "" Then rDThumbImageSize = "64" ' validate
    thumbImageSize = Val(rDThumbImageSize) ' set

    ' .70 DAEB 16/05/2022 rDIConConfig.frm Read the chkToggleDialogs value from a file and save the value for next time
    sdChkToggleDialogs = GetINISetting("Software\SteamyDockSettings", "sdChkToggleDialogs", toolSettingsFile)
    
    If sdChkToggleDialogs = "" Then sdChkToggleDialogs = "1" ' validate
    If sdChkToggleDialogs = "1" Then ' set
        chkToggleDialogs.Value = 1
    Else
        chkToggleDialogs.Value = 0
    End If
    ' .71 DAEB 16/05/2022 rDIConConfig.frm Move the reading of recent settings into the main read configuration procedure ENDS

    ' TBD
    sdThumbnailCacheCount = GetINISetting("Software\SteamyDockSettings", "thumbnailCacheCount", toolSettingsFile)

    If sdThumbnailCacheCount = "" Then sdThumbnailCacheCount = "250" ' default value


   On Error GoTo 0
   Exit Sub

readIconsAndConfiguration_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconsAndConfiguration of Form dockSettings on " & Erl
End Sub




'---------------------------------------------------------------------------------------
' Procedure : readRocketDockSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRocketDockSettings()
    
    ' SETTINGS: There are three (four) settings files
    ' the first is the RD settings file that only exists if RD is NOT using the registry
    ' the second is our tools copy of RD's settings file, we copy the original or create our own from RD's registry settings
    ' the third is the new config settings file that exists in the user apps folder (not obsolete as the RD locations are)
    ' the fourth is the settings file for this tool to store its own preferences
        
    ' check to see if the settings file exists
    ' (Rocketdock overwrites its own settings.ini when it closes meaning that we have to work on a copy).
   On Error GoTo readRocketDockSettings_Error
   If debugflg = 1 Then DebugPrint "%" & "readRocketDockSettings"

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
        
    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        frmRegistry.chkReadRegistry.Value = 0
        frmRegistry.chkReadSettings.Value = 1
        frmRegistry.chkReadConfig.Value = 0
        
        frmRegistry.chkWriteRegistry.Value = 0
        frmRegistry.chkWriteSettings.Value = 1
        frmRegistry.chkWriteConfig.Value = 0

        Call fbackupSettings   ' make a backup of the settings.ini file each restart
        
        ' copy the original settings file to a duplicate that we will operate upon
        FileCopy origSettingsFile, interimSettingsFile
        
        ' read the rocketdock settings.ini and find the very last icon
        theCount = Val(GetINISetting("Software\SteamyDock\IconSettings\Icons", "count", interimSettingsFile))
        rdIconMaximum = theCount - 1
    Else
        frmRegistry.chkReadRegistry.Value = 1
        frmRegistry.chkReadSettings.Value = 0
        frmRegistry.chkReadConfig.Value = 0
             
        frmRegistry.chkWriteRegistry.Value = 1
        frmRegistry.chkWriteSettings.Value = 0
        frmRegistry.chkWriteConfig.Value = 0
        
        ' read the rocketdock registry and find the last icon
        theCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count")
        rdIconMaximum = theCount - 1
                
        ' copy the original ICON configs out of the registry and into a settings file that we will operate upon
        readIconRegistryWriteSettings interimSettingsFile
        
        ' make a backup of the rdSettings.ini after the intermediate file has been created
        Call fbackupSettings
        
    End If

   On Error GoTo 0
   Exit Sub

readRocketDockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRocketDockSettings of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkSteamyDockInstallation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkSteamyDockInstallation()
        
    ' variables declared
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    SD86installed = ""
    SDinstalled = ""
    
    ' check where SteamyDock is installed
    On Error GoTo checkSteamyDockInstallation_Error

    SD86installed = driveCheck("Program Files (x86)\SteamyDock", "steamyDock.exe")
    SDinstalled = driveCheck("Program Files\SteamyDock", "steamyDock.exe")
    
    If SDinstalled = "" And SD86installed = "" Then
        steamyDockInstalled = False
        'answer = msgBoxA(" SteamyDock has not been installed in the program files (x86) folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
        Exit Sub
    Else
        steamyDockInstalled = True

        If SDinstalled <> "" Then
            'MsgBox "SteamyDock is installed in program files"
            sdAppPath = SDinstalled
        End If
        'the one in the x86 folder has precedence
        If SD86installed <> "" Then
            'MsgBox "SteamyDock is installed in program files (x86)"
            sdAppPath = SD86installed
        End If
        
        dockAppPath = sdAppPath
        defaultDock = 1
    End If
    

   On Error GoTo 0
   Exit Sub

checkSteamyDockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkSteamyDockInstallation of Form dockSettings"
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
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    Dim rDOptGeneralReadConfig As String: rDOptGeneralReadConfig = ""
    Dim chkGenAlwaysAskValue As Integer: chkGenAlwaysAskValue = 0
    
    On Error GoTo checkDefaultDock_Error
   
    'initialise the dimensioned variables
    answer = vbNo
    rDOptGeneralReadConfig = ""
    
    If steamyDockInstalled = True Then
        ' get the location of the dock's new settings file
        'Call locateDockSettingsFile
        chkGenAlwaysAskValue = Val(GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", interimSettingsFile))
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", interimSettingsFile)
        rDOptGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", interimSettingsFile)
    End If
    
'    If steamyDockInstalled = True And rocketDockInstalled = True Then
'        If chkGenAlwaysAskValue = 1 Then  ' depends upon being able to read the new configuration file in the user data area
'            answer = msgBoxA("Both Rocketdock and SteamyDock are installed on this system. Use SteamyDock by default? ", vbQuestion + vbYesNo)
'            If answer = vbYes Then
'                cmbDefaultDock.ListIndex = 1 ' steamy dock
'                dockAppPath = sdAppPath
'                'mnuRocketDock.Caption = "SteamyDock location - " & sdAppPath & " - click to change"
'                defaultDock = 1
'            Else
'                cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                'mnuRocketDock.Caption = "Rocketdock location - " & rdAppPath & " - click to change"
'                defaultDock = 0
'            End If
'        Else
'            ' if the question is not being asked then use the default dock as specified in the docksettings.ini file
'            If rDDefaultDock = "steamydock" Then
'                cmbDefaultDock.ListIndex = 1
'                dockAppPath = sdAppPath
'                'mnuRocketDock.Caption = "SteamyDock location - " & sdAppPath & " - click to change"
'                defaultDock = 1
'            ElseIf rDDefaultDock = "rocketdock" Then
'                cmbDefaultDock.ListIndex = 0 ' rocket dock
'                dockAppPath = rdAppPath
'                'mnuRocketDock.Caption = "Rocketdock location - " & rdAppPath & " - click to change"
'                defaultDock = 0
'            Else
'                If cmbDefaultDock.ListIndex = 1 Then  ' depends upon being able to read the new configuration file in the user data area
'                    dockAppPath = sdAppPath
'                    'mnuRocketDock.Caption = "SteamyDock location - " & sdAppPath & " - click to change"
'                    defaultDock = 1
'                Else
'                    cmbDefaultDock.ListIndex = 0 ' rocket dock
'                    dockAppPath = rdAppPath
'                    'mnuRocketDock.Caption = "Rocketdock location - " & rdAppPath & " - click to change"
'                    defaultDock = 0
'                End If
'            End If
'        End If
'    Else
    If steamyDockInstalled = True Then ' just steamydock installed
            cmbDefaultDock.ListIndex = 1
            dockAppPath = sdAppPath
            'mnuRocketDock.Caption = "SteamyDock location - " & sdAppPath & " - click to change"
            defaultDock = 1
            ' write the default dock to the SteamyDock settings file
            PutINISetting "Software\SteamyDockSettings", "defaultDock", defaultDock, toolSettingsFile
            
'    ElseIf rocketDockInstalled = True Then ' just rocketdock installed
'            cmbDefaultDock.ListIndex = 0
'            dockAppPath = rdAppPath
'            'mnuRocketDock.Caption = "Rocketdock location - " & rdAppPath & " - click to change"
'            defaultDock = 0
    End If
    
    ' it is possible to run this program without steamydock being installed
    If steamyDockInstalled = False And rocketDockInstalled = False Then
        answer = msgBoxA(" Neither Rocketdock nor SteamyDock has been installed on any of the drives on this system, can you please install into the correct folder and retry?", vbExclamation + vbYesNo)
         Dim ofrm As Form
         For Each ofrm In Forms
             Unload ofrm
         Next
         End
    End If
    
    ' .44 DAEB 16/04/2022 rdIconConfig.frm one menu option is not applicable in SD, ie. adding a docklet
    If defaultDock = 1 Then
        mnuAddDocklet.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

checkDefaultDock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkDefaultDock of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : locateDockSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
'Private Sub locateDockSettingsFile()
'
'    ' variables declared
'    Dim dockSettingsDir As String
'
'    'initialise the dimensioned variables
'    dockSettingsDir = ""
'
'    On Error GoTo locateDockSettingsFile_Error
'    If debugflg = 1 Then DebugPrint "%locateDockSettingsFile"
'
'    dockSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\steamyDock" ' just for this user alone
'    dockSettingsFile = dockSettingsDir & "\docksettings.ini" ' the third config option for steamydock alone
'
'    'if the folder does not exist then create the folder
'    If Not DirExists(dockSettingsDir) Then
'        MkDir dockSettingsDir
'    End If
'
'    'if the settings.ini does not exist then create the file by copying
'    If Not FExists(dockSettingsFile) Then
'        If FExists(App.Path & "\dockSettings.ini") Then
'            FileCopy App.Path & "\dockSettings.ini", dockSettingsFile
'        End If
'    End If
'
'    'confirm the settings file exists, if not use the version in the app itself
'    If Not FExists(dockSettingsFile) Then
'            dockSettingsFile = App.Path & "\settings.ini"
'    End If
'
'
'   On Error GoTo 0
'   Exit Sub
'
'locateDockSettingsFile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure locateDockSettingsFile of Form dockSettings"
'
'End Sub
'---------------------------------------------------------------------------------------
' Procedure : checkRocketdockInstallation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : we check to see if rocketdock is installed in order to know the location of the settings.ini file used by Rocketdock
'               If rocketdock Is Not installed Then test the registry read
'               If the registry settings are located then offer them as a choice.
'---------------------------------------------------------------------------------------
'
Private Sub checkRocketdockInstallation()
    ' variables declared
    Dim answer As VbMsgBoxResult:  answer = vbNo
    
    RDinstalled = ""
    RD86installed = ""
    
    ' check where rocketdock is installed
    On Error GoTo checkRocketdockInstallation_Error
    If debugflg = 1 Then DebugPrint "%" & "checkRocketdockInstallation"
        
    RD86installed = driveCheck("Program Files (x86)\Rocketdock", "RocketDock.exe")
    RDinstalled = driveCheck("Program Files\Rocketdock", "RocketDock.exe")
    
'    If RDinstalled <> "" Then mnuRocketDock.Caption = "Rocketdock location - program files - click to change"
'    If RD86installed <> "" Then mnuRocketDock.Caption = "Rocketdock location - program files (x86) - click to change"
    
    If RDinstalled = "" And RD86installed = "" Then
        rocketDockInstalled = False
        'answer = msgBoxA(" Rocketdock has not been installed in the program files (x86) folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
        
        Exit Sub
    Else
        rocketDockInstalled = True
        If RDinstalled <> "" Then
            rdAppPath = RDinstalled
        End If
        'the one in the x86 folder has precedence
        If RD86installed <> "" Then
            rdAppPath = RD86installed
        End If
        dockAppPath = rdAppPath
        'defaultDock = 0
    End If
    
    ' If rocketdock Is Not installed Then test the registry
    ' if the registry settings are not located then remove them as a source.
        
    ' rocketDockInstalled = False ' debug
    
    ' read selected random entries from the registry, if each are false then the RD registry entries do not exist.
    If rocketDockInstalled = False Then
        rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIconsd")
        rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunnings")
        rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunnings")
        rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindowsw")
        rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimations")
        If rDLockIcons = "" And rDOpenRunning = "" And rDShowRunning = "" And rDManageWindows = "" And rDDisableMinAnimation = "" Then
            ' rocketdock registry entries do not exist so RD has never been installed or it has been wiped entirely.
            RDregistryPresent = False
        Else
            RDregistryPresent = True 'rocketdock HAS been installed in the past as the registry entries are still present
        End If
    End If



   On Error GoTo 0
   Exit Sub

checkRocketdockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    Dim toolSettingsDir As String:  toolSettingsDir = ""

    On Error GoTo getToolSettingsFile_Error
    If debugflg = 1 Then DebugPrint "%getToolSettingsFile"
    
    toolSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\rocketdockEnhancedSettings"
    toolSettingsFile = toolSettingsDir & "\settings.ini"
    'toolSettingsFile = App.path & "\settings.ini"
    'sharedToolSettingsFile = SpecialFolder(SpecialFolder_CommonAppData) & "\rocketdockEnhancedSettings\settings.ini"
    
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form rDIconConfigForm"

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : checkLicenceState
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : 'check the state of the licence
''---------------------------------------------------------------------------------------
''
'Private Sub checkLicenceState()
'    Dim slicence As String
'
'    On Error GoTo checkLicenceState_Error
'    If debugflg = 1 Then DebugPrint "%" & "checkLicenceState"
'
'    ' read the tool's own settings file (
'    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        slicence = GetINISetting("Software\SteamyDockSettings", "Licence", toolSettingsFile)
'        ' if the licence state is not already accepted then display the licence form
'        If slicence = "0" Then
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkLicenceState of Form rDIconConfigForm"
'
'End Sub


' .56 DAEB 25/04/2022 rDIConConfig.frm 1st run of the thumbnail view window is done using the old method and it comes out incorrectly.
''---------------------------------------------------------------------------------------
'' Procedure : placeThumbnailPicboxes
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : ' place the thumbnails picboxes where they should go
''---------------------------------------------------------------------------------------
''
'Private Sub placeThumbnailPicboxes(ByVal imageSize As Integer)
'    Dim useloop As Integer
'    Dim storeTop As Integer
'
'
'    On Error GoTo placeThumbnailPicboxes_Error
'    If debugflg = 1 Then DebugPrint "%" & "placeThumbnailPicboxes"
'
''    picThumbIcon(0).Width = 1000
''    picThumbIcon(0).Height = 1000
''    picThumbIcon(0).Left = 165
''    picThumbIcon(0).Top = 60
'
'    For useloop = 0 To 11
'        picFraPicThumbIcon(useloop).Width = 1000
'        picFraPicThumbIcon(useloop).Height = 1000
''        picThumbIcon(useloop).Width = 1000
''        picThumbIcon(useloop).Height = 1000
'        fraThumbLabel(useloop).BorderStyle = 0
'
'        fraThumbLabel(useloop).BackColor = vbWhite
'        lblThumbName(useloop).BackColor = vbWhite
'
'        picThumbIcon(useloop).ToolTipText = filesIconList.List(useloop)
'
'
'        picFraPicThumbIcon(useloop).Left = storeLeft
'        picFraPicThumbIcon(useloop).Top = storeTop
''        picThumbIcon(useloop).Left = storeLeft
''        picThumbIcon(useloop).Top = storeTop
'
'        fraThumbLabel(useloop).Left = storeLeft - 100
'        fraThumbLabel(useloop).Top = storeTop + 800
'
'        storeLeft = storeLeft + 1200
'
'        picThumbIcon(useloop).Visible = True
'        lblThumbName(useloop).Visible = True
'        fraThumbLabel(useloop).Visible = True
'
'        picThumbIcon(useloop).ZOrder
'        fraThumbLabel(useloop).ZOrder
'
'        picThumbIcon(useloop).AutoRedraw = True
'    Next useloop
'
'
'   On Error GoTo 0
'   Exit Sub
'
'placeThumbnailPicboxes_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure placeThumbnailPicboxes of Form rDIconConfigForm"
'
'End Sub
    'carry end


'---------------------------------------------------------------------------------------
' Procedure : busyStart
' Author    : beededea
' Date      : 28/08/2019
' Purpose   : two subroutines that were previously used by a timer, now only used to set a pointer
'---------------------------------------------------------------------------------------
'
Private Sub busyStart()
   On Error GoTo busyStart_Error
   If debugflg = 1 Then DebugPrint "%busyStart"

        Me.MousePointer = 11

   On Error GoTo 0
   Exit Sub

busyStart_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure busyStart of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : busyStop
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub busyStop()
   On Error GoTo busyStop_Error
   If debugflg = 1 Then DebugPrint "%busyStop"

        Me.MousePointer = 1

   On Error GoTo 0
   Exit Sub

busyStop_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure busyStop of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnArrowDown_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub btnBackup_Click()

    Call backupDockSettings(True)

   On Error GoTo 0
   Exit Sub

btnBackup_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnBackup_Click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnGetMore_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGetMore_Click()
    ' TODO - move the link below to a right click menu as well
   On Error GoTo btnGetMore_Click_Error
   If debugflg = 1 Then DebugPrint "%btnGetMore_Click"

    Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981272/orbs-and-icons", vbNullString, App.Path, 1)

   On Error GoTo 0
   Exit Sub

btnGetMore_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGetMore_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnKillIcon_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Allows deletion of the selected icon in the file list
'---------------------------------------------------------------------------------------
'
Private Sub btnKillIcon_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    On Error GoTo btnKillIcon_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnKillIcon_Click"

        If textCurrIconPath.Text = vbNullString Then
            msgBoxA "Cannot perform a deletion as no icon has been selected. ", vbInformation + vbOKOnly, "", False
            Exit Sub
        End If

        answer = msgBoxA("This will delete the currently selected icon, " & vbCr & textCurrentFolder.Text & "\" & vbCr & textCurrIconPath.Text & "   -  are you sure?", vbExclamation + vbYesNo, "", False)
        If answer = vbNo Then
            Exit Sub
        End If
                
        'delete the icon with no confirmation
        Kill textCurrentFolder.Text & "\" & textCurrIconPath.Text
        
        'set the selection to the current list position
        filesIconList.ListIndex = fileIconListPosition
        
        'explicitly remove the one item from the cache
        'imlThumbnailCache.ListImages.Remove (fileIconListPosition + 1) ' not sensible thing to do
        
        'refresh the underlying file display
        filesIconList.Refresh
                
        ' clear the cache for re-population otherwise the cache will be screwed up after the deletion
        ' just removing one image from the cache (as above) is not enough as the whole cache is associated with each position in
        ' the thumbnail list. In effect the whole cache needs to be moved up one. Simpler to clear the cache.
        imlThumbnailCache.ListImages.Clear
        
        ' the above task appears to be asynchronous (for Krool's imageList) and as such may take a while to complete, though the
        ' program itself carries on immediately. In a compiled version it may not be complete before another request to the cache
        ' is made, so there is a flag to force the thumbnail to ignore the cache altogether.
                
        ' .91 DAEB 25/06/2022 rDIConConfig.frm Deleting an icon from the icon thumbnail display causes a cache imageList error. Added cacheingFlg.
        cacheingFlg = False ' flag to ignore the cache during the refresh
        
        ' using the current filelist as the start point on the list, repopulate the thumbs
        Call btnRefresh_Click_Event
   
        cacheingFlg = True ' switch it back on again
        
        If filesIconList.Visible = True Then
            filesIconList.SetFocus         ' return focus to the form
        Else
            picFrameThumbs.SetFocus        ' return focus to the form
        End If
        
        ' now display the current icon textual and preview details, the previous icon displayed now havin been deleted
        Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
   
   On Error GoTo 0
   Exit Sub

btnKillIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnKillIcon_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnSet_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : This just saves the current icon changes to the interim dockSettingsFile
'---------------------------------------------------------------------------------------
'
Private Sub btnSet_Click()
   
   On Error GoTo btnSet_Click_Error
   If debugflg = 1 Then DebugPrint "%" & "btnSet_Click"
   
    
    sFilename = txtCurrentIcon.Text
    
    sTitle = txtLabelName.Text
    If sDockletFile = "" Then
        sCommand = txtTarget.Text
    Else
        sDockletFile = txtTarget.Text
    End If
    sArguments = txtArguments.Text
    sWorkingDirectory = txtStartIn.Text
    sShowCmd = cmbRunState.ListIndex + 1 ' .34 DAEB 05/05/2021 rDIConConfigForm.frm The value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
    sOpenRunning = Str$(cmbOpenRunning.ListIndex)
    sRunElevated = Str$(chkRunElevated.Value)
    
    ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    If defaultDock = 1 Then
        sUseDialog = chkConfirmDialog.Value
        sUseDialogAfter = chkConfirmDialogAfter.Value
        sQuickLaunch = chkQuickLaunch.Value  '.nn Added new check box to allow a quick launch of the chosen app
        sAutoHideDock = chkAutoHideDock.Value '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        sSecondApp = txtSecondApp.Text ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
        
        sRunSecondAppBeforehand = Val(optRunSecondAppBeforehand.Value)
        sAppToTerminate = txtAppToTerminate.Text
        If sDisabled = "9" Then sDisabled = "1"
        
        sDisabled = chkDisabled.Value ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
    End If

    ' save the current fields to the settings file or registry
    If FExists(interimSettingsFile) Then '
        ' write the rocketdock settings.ini
        'writeSettingsIni (rdIconNumber) ' the settings.ini only exists when RD is set to use it
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile)
    End If
    
    ' tell the user that all has been saved
    If FExists(interimSettingsFile) Then '
        If chkToggleDialogs.Value = 1 Then
            msgBoxA "This icon change has been stored," & vbCr & "You will need to press the ""save & restart"" button " & vbCr & "to make the changes 'stick' within Rocketdock", vbInformation + vbOKOnly, "Icon Settings Saved", True, "btnSet_Click"
        End If
    End If
    
    'if the current icon has changed by a dblclick on the file list then refresh that part of the rdMap
    If iconChanged = True Then
        'only if the rdMAp has already been displayed already do we carry out the image refresh
        If Not picRdMap(0).ToolTipText = vbNullString Then ' check that the array has been populated already
            ' we just reload the sole picbox that has changed
            Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), True, 32, True, False)
        End If
        iconChanged = False
    End If
    
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

    
    If triggerRdMapRefresh = True Then
        'Call rdMapRefresh_Click
        'Call busyStart
        Call populateRdMap(0) ' show the map from position zero
        'Call busyStop

        ' we signify that there have been no changes - this is just a refresh
        btnSet.Enabled = False ' this has to be done at the end
        btnClose.Visible = True
    btnCancel.Visible = False

        triggerRdMapRefresh = False
    End If

   On Error GoTo 0
   Exit Sub

btnSet_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSet_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRefresh_Click
' Author    : beededea
' Date      : 25/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRefresh_Click()
   
    On Error GoTo btnRefresh_Click_Error
    
        If debugflg = 1 Then DebugPrint "%" & "btnRefresh_Click"
        thisRoutine = "btnRefresh_Click"
        
        ' clear the cache before the refresh
        imlThumbnailCache.ListImages.Clear
        
        Call btnRefresh_Click_Event
   
    On Error GoTo 0
    Exit Sub

btnRefresh_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRefresh_Click of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnRefresh_Click_Event
' Author    : beededea
' Date      : 27/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRefresh_Click_Event()
    On Error GoTo btnRefresh_Click_Event_Error

        Call busyStart
        
        filesIconList.Refresh ' refresh the underlying filelist that both views use
        
        ' if the thumbnail view is displaying then repopulate the thumbnail view
        
        vScrollThumbs.Value = vScrollThumbs.Min
        
        ' using the current filelist as the start point on the list, repopulate the thumbs
        Call populateThumbnails(thumbImageSize, filesIconList.ListIndex)
        
        If picFrameThumbs.Visible = True Then
            picFrameThumbs.SetFocus
        Else
            ' if it is the file viewer that is enabled just set focus to it
            filesIconList.SetFocus
        End If

        removeThumbHighlighting

        'highlight the current thumbnail
        thumbIndexNo = 0
        If thumbImageSize = 64 Then 'larger
            picFraPicThumbIcon(thumbIndexNo).BorderStyle = 1
            'picThumbIcon(thumbIndexNo).BorderStyle = 1
        ElseIf thumbImageSize = 32 Then
            lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200)
        End If
    
        Call busyStop

    On Error GoTo 0
    Exit Sub

btnRefresh_Click_Event_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRefresh_Click_Event of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : getFolderNameFromPath
'' Author    : beededea
'' Date      : 11/07/2019
'' Purpose   : get the folder or directory path as a string not including the last backslash
''---------------------------------------------------------------------------------------
''
'Private Function getFolderNameFromPath(ByRef path As String) As String
'    On Error GoTo getFolderNameFromPath_Error
'    If debugflg = 1 Then DebugPrint "%" & "getFolderNameFromPath"
'
'    If InStrRev(path, "\") = 0 Then
'        getFolderNameFromPath = ""
'        Exit Function
'    End If
'    getFolderNameFromPath = left$(path, InStrRev(path, "\") - 1)
'
'    On Error GoTo 0
'    Exit Function
'
'GetDirectory_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFolderNameFromPath of Form rDIconConfigForm"
'End Function
'---------------------------------------------------------------------------------------
' Procedure : btnTarget_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Private Sub btnTarget_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    
    ' .96 DAEB 26/06/2022 rDIConConfig.frm If the target text box is just a folder and not a full file path then a click on the select button should select also select a folder and not a file
'    If txtTarget.Text = getFolderNameFromPath(txtTarget.Text) Then
    If txtTarget.Text <> vbNullString Then
        If FExists(txtTarget.Text) Then
            Call getFileName

        ElseIf DirExists(txtTarget.Text) Then
            dialogInitDir = txtTarget.Text 'start dir, might be "C:\" or so also
            getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder
            If getFolder <> vbNullString Then txtTarget.Text = getFolder
        Else
            ' dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
             Call getFileName
        End If
    
    Else

        Call getFileName
    End If

   On Error GoTo 0
   
   Exit Sub

btnTarget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTarget_Click of Form rDIconConfigForm"
 
End Sub

Private Sub getFileName()
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString

    retFileName = addTargetProgram(txtTarget.Text)
    
    If retFileName <> vbNullString Then
        txtTarget.Text = retFileName
        'fill in the file title and the start in automatically if they are empty and need filling
        If txtLabelName.Text = vbNullString Then txtLabelName.Text = retfileTitle
        'If txtStartIn.Text = vbNullString Then txtStartIn.Text = getFolderNameFromPath(txtTarget.Text)
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : addTargetProgram
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Private Function addTargetProgram(ByVal targetText As String)
    Dim iconPath As String: iconPath = vbNullString
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    
    Const x_MaxBuffer = 256
    
    'On Error GoTo addTargetProgram_Error
    If debugflg = 1 Then DebugPrint "%" & "addTargetProgram"
    
    'On Error GoTo l_err1
    'savLblTarget = txtTarget.Text
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not targetText = vbNullString Then
        If FExists(targetText) Then
            ' extract the folder name from the string
            iconPath = getFolderNameFromPath(targetText)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(targetText) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = targetText 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
            End If
        End If
    Else
    ' .85 DAEB 06/06/2022 rDIConConfig.frm  Second app button should open in the program files folder
    If DirExists("c:\program files") Then
            dialogInitDir = "c:\program files"
        End If
    End If
    
    If Not sDockletFile = vbNullString Then
        If FExists(sDockletFile) Then
            ' extract the folder name from the string
            dllPath = getFolderNameFromPath(sDockletFile)
            ' set the default folder to the existing reference
            dialogInitDir = dllPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(sDockletFile) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = sDockletFile 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
                dialogInitDir = rdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath & "\docklets"  'start dir, might be "C:\" or so also
            End If
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target for this icon to call"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String$(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With

  Call getFileNameAndTitle(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

' this is the code left over from the use of the common dialog OCX, left here for reference
'
'    rdDialogForm.CommonDialog.DialogTitle = "Select a File" 'titlebar
'    If Not txtTarget.Text = vbNullString Then
'        If FExists(txtTarget.Text) Then
'            ' extract the folder name from the string
'            iconPath = getFolderNameFromPath(txtTarget.Text)
'            ' set the default folder to the existing reference
'            rdDialogForm.CommonDialog.InitDir = iconPath 'start dir, might be "C:\" or so also
'        ElseIf DirExists(txtTarget.Text) Then ' this caters for the entry being just a folder name
'            ' set the default folder to the existing reference
'            rdDialogForm.CommonDialog.InitDir = txtTarget.Text 'start dir, might be "C:\" or so also
'        Else
'            rdDialogForm.CommonDialog.InitDir = rdAppPath 'start dir, might be "C:\" or so also
'        End If
'    End If
'    rdDialogForm.CommonDialog.FileName = "*.*"  'Something in filenamebox
'    rdDialogForm.CommonDialog.CancelError = False 'allow escape key/cancel
'    rdDialogForm.CommonDialog.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
'    rdDialogForm.CommonDialog.ShowOpen
'
'l_err1:
'    If rdDialogForm.CommonDialog.FileName = vbNullString Then
'        txtTarget.Text = savLblTarget
'        Exit Sub
'    End If
'
'    If Err <> 32755 Then    ' User didn't chose Cancel.
'        If rdDialogForm.CommonDialog.FileName = "*.*" Then
'            txtTarget.Text = savLblTarget
'        Else
'            If txtLabelName.Text = vbNullString Then
'                txtLabelName.Text = rdDialogForm.CommonDialog.FileTitle
'            End If
'            txtTarget.Text = rdDialogForm.CommonDialog.FileName
'        End If
'    End If

    addTargetProgram = retFileName

   On Error GoTo 0
   
   Exit Function

addTargetProgram_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addTargetProgram of Form rDIconConfigForm"
 
End Function




























Private Sub lblchkAutoHideDock_Click()
    If chkAutoHideDock.Value = 1 Then
        chkAutoHideDock.Value = 0
    Else
        chkAutoHideDock.Value = 1
    End If
End Sub

Private Sub lblConfirmDialog_Click()

    If chkConfirmDialog.Value = 1 Then
        chkConfirmDialog.Value = 0
    Else
        chkConfirmDialog.Value = 1
    End If
    
End Sub

Private Sub lblConfirmDialogAfter_Click()
    If chkConfirmDialogAfter.Value = 1 Then
        chkConfirmDialogAfter.Value = 0
    Else
        chkConfirmDialogAfter.Value = 1
    End If
End Sub

Private Sub lblRunElevated_Click()
    If chkRunElevated.Value = 1 Then
        chkRunElevated.Value = 0
    Else
        chkRunElevated.Value = 1
    End If
End Sub

Private Sub lblQuickLaunch_Click()

    If chkQuickLaunch.Value = 1 Then
        chkQuickLaunch.Value = 0
    Else
        chkQuickLaunch.Value = 1
    End If
    
End Sub

Private Sub picHideConfig_Click()
    Call picMoreConfigUp_Click
    picHideConfig.Visible = False
End Sub


Private Sub picMoreConfigUp_Click()
        Dim amountToDrop As Integer: amountToDrop = 0
        

        picMoreConfigDown.Visible = True
        picMoreConfigUp.Visible = False
        picHideConfig.Visible = False
        fraProperties.Height = 3630
        moreConfigVisible = False
        amountToDrop = 1200

        ' .43 DAEB 16/04/2022 rdIconConfig.frm increase the whole form height and move the bottom buttons set down
        frameButtons.Top = frameButtons.Top - amountToDrop
        rDIconConfigForm.Height = rDIconConfigForm.Height - amountToDrop
        
        framePreview.Height = framePreview.Height - amountToDrop
        fraSizeSlider.Top = fraSizeSlider.Top - amountToDrop
        btnNext.Height = btnNext.Height - amountToDrop
        btnPrev.Height = btnPrev.Height - amountToDrop

        If chkToggleDialogs.Value = 0 Then picMoreConfigDown.ToolTipText = "Shows extra configuration items"
End Sub


Private Sub Picture3_Click()
    Call picMoreConfigUp_Click
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAddPreviewIcon_Click
' Author    : beededea
' Date      : 27/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddPreviewIcon_Click()

   On Error GoTo mnuAddPreviewIcon_Click_Error
   If debugflg = 1 Then DebugPrint "%mnuAddPreviewIcon_Click"

    'DebugPrint picPreview.Tag
    
    Call btnAdd_Click

   On Error GoTo 0
   Exit Sub

mnuAddPreviewIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPreviewIcon_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddProgram_Click
' Author    : beededea
' Date      : 04/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddProgram_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim iconImage As String: iconImage = vbNullString

   On Error GoTo mnuAddProgram_Click_Error
   If debugflg = 1 Then Debug.Print "%mnuAddProgram_Click"

    retFileName = addTargetProgram("")
       
    Refresh
    
    
    ' .35 DAEB 20/04/2021 rdIconConfig.frm Added new function to identify an icon to assign to the entry
    
    ' we should, if it is a EXE dig into it to determine the icon using privateExtractIcon
                         
    ' However, we do not extract the icon from the shortcut as it will be useless for steamydock
    ' VB6 not being able to extract and handle a transparent PNG form
    ' even if it was we have no current method of making a transparent PNG from a bitmap or ICO that
    ' I can easily transfer to the GDI collection - but I am working on it...
    ' the vast majority of default icons are far too small for steamydock in any case.
    ' the result of the above is that there is currently no icon extracted, though that may change.
    
    ' instead we have a list of apps that we can match the shortcut name against, it exists in an external comma
    ' delimited file. The list has two identification factors that are used to find a match and then we find an
    ' associated icon to use with a relative path.
       
    'MsgBox "1. retFileName " & retFileName
    iconFileName = identifyAppIcons(retFileName) ' .35 DAEB 20/04/2021 rdIconConfig.frm Added new function to identify an icon to assign to the entry
    'MsgBox "2. iconFileName " & iconFileName
                
    If FExists(iconFileName) Then
      iconImage = iconFileName
    Else
        iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-EXE.png"
        If FExists(iconFileName) Then
            iconImage = iconFileName
        Else
            iconImage = App.Path & "\Icons\help.png"
        End If
    End If
    
    'MsgBox "3. iconFileName " & iconFileName



    ' general tool to add an icon
'    iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-EXE.png"
'    If FExists(iconFileName) Then
'        iconImage = iconFileName
'    Else
'        iconImage = App.Path & "\Icons\help.png"
'    End If
        
    Call menuAddSomething(iconImage, retFileName, retFileName, vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddProgram_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddProgram_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRocketDock_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : Menu option to direct where Rocketdock may be found
'---------------------------------------------------------------------------------------
'
Private Sub mnuRocketDock_click()

    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
   
   On Error GoTo mnuRocketDock_click_Error
   If debugflg = 1 Then DebugPrint "%mnuRocketDock_click"

    dialogInitDir = "C:\" 'start dir, might be "C:\" or so also

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder
    If getFolder <> vbNullString Then
        If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
            If FExists(getFolder & "\rocketdock.exe") Then
                rdAppPath = getFolder & "\rocketdock.exe"
                'If DirExists(getFolder) Then mnuRocketDock.Caption = "RocketDock Location - " & getFolder & " - click to change."
                
                If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                    PutINISetting "Software\SteamyDockSettings", "rocketDockLocation", rdAppPath, toolSettingsFile
                End If
                
            End If
        Else
            If FExists(getFolder & "\steamydock.exe") Then
                sdAppPath = getFolder & "\steamydock.exe"
                'If DirExists(getFolder) Then mnuRocketDock.Caption = "Steamydock Location - " & getFolder & " - click to change."
                
                If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                    ' write the default dock location to the SteamyDock settings file
                    PutINISetting "Software\SteamyDockSettings", "steamyDockLocation", sdAppPath, toolSettingsFile
                End If
                
            End If
        End If
    End If
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile

   On Error GoTo 0
   Exit Sub

mnuRocketDock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRocketDock_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddSeparator_click
' Author    : beededea
' Date      : 29/09/2019
' Purpose   : Menu option to add a separator dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddSeparator_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    On Error GoTo mnuAddSeparator_click_Error
    If debugflg = 1 Then DebugPrint "mnuAddSeparator_click"
           
    iconFileName = App.Path & "\my collection" & "\separator.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If

    sIsSeparator = "1"
        
    ' general tool to add an icon
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Separator", vbNullString, vbNullString, vbNullString, vbNullString, sIsSeparator)
        
    txtLabelName.Enabled = False
    txtCurrentIcon.Enabled = False
    txtTarget.Enabled = False
    btnTarget.Enabled = False
    txtArguments.Enabled = False
    txtStartIn.Enabled = False
    cmbRunState.Enabled = False
    cmbOpenRunning.Enabled = False
    chkRunElevated.Enabled = False
    btnSelectStart.Enabled = False
        
   On Error GoTo 0
   Exit Sub

mnuAddSeparator_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddSeparator_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuaddFolder_click
' Author    : beededea
' Date      : 29/09/2019
' Purpose   : Menu option to add a folder dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuaddFolder_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
   
   On Error GoTo mnuaddFolder_click_Error
   If debugflg = 1 Then DebugPrint "%mnuaddFolder_click"

    If txtStartIn.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath
            Else
                dialogInitDir = sdAppPath
            End If

        End If
    End If

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder

    If DirExists(getFolder) Then
    
        iconFileName = App.Path & "\my collection" & "\folder-closed.png"
        If FExists(iconFileName) Then
            iconImage = iconFileName
        Else
            iconImage = "\Icons\help.png"
        End If
        
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call menuAddSomething(iconImage, "User Folder", getFolder, vbNullString, vbNullString, vbNullString, vbNullString)
    End If

   On Error GoTo 0
   Exit Sub

mnuaddFolder_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuaddFolder_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyComputer_click
' Author    : beededea
' Date      : 29/09/2019
' Purpose   : Menu option to add a "my computer" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyComputer_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
    ' check the icon exists
   On Error GoTo mnuAddMyComputer_click_Error
   If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\my collection" & "\my folder.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "My Computer", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)


   On Error GoTo 0
   Exit Sub

mnuAddMyComputer_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyComputer_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyDocuments_Click
' Author    : beededea
' Date      : 07/03/2021
' Purpose   : .20 DAEB 07/03/2021 rdIconConfig.frm Added menu option to add a "my Documents" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyDocuments_Click()
'
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
    On Error GoTo mnuAddMyDocuments_Click_Error

    If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"
    
    ' check the icon exist
    iconFileName = App.Path & "\my collection" & "\folder-closed.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\Icons\help.png"
    End If
       
    If FExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call menuAddSomething(iconImage, "My Documents", "::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}", vbNullString, vbNullString, vbNullString, vbNullString)
    Else
        MsgBox "Unable to add my Documents image as it does not exist"
    End If
        

   On Error GoTo 0
   Exit Sub

mnuAddMyDocuments_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyDocuments_Click of Form menuForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyMusic_Click
' Author    : beededea
' Date      : 07/03/2021
' Purpose   : .21 DAEB 07/03/2021 menu.frm Added menu option to add a "my Music" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyMusic_Click()
'
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim userprof As String: userprof = vbNullString

    
    ' check the icon exists
    On Error GoTo mnuAddMyMusic_Click_Error

    If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\my collection" & "\music.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\Icons\help.png"
    End If

    userprof = Environ$("USERPROFILE")
    
    If FExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        ' using the Special CLSID for the video folder this, in fact resolves to the my documents folder and not the video folder below.
        'Call menuAddSomething( iconImage, "My Music", "::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", vbNullString, vbNullString, vbNullString, vbNullString)
        Call menuAddSomething(iconImage, "My Music", userprof & "\Documents\Music", vbNullString, vbNullString, vbNullString, vbNullString)
    Else
        MsgBox "Unable to add my Music image as it does not exist"
    End If
        

   On Error GoTo 0
   Exit Sub

mnuAddMyMusic_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyMusic_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyPictures_Click
' Author    : beededea
' Date      : 07/03/2021
' Purpose   : .22 DAEB 07/03/2021 menu.frm Added menu option to add a "my Pictures" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyPictures_Click()
'
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim userprof As String: userprof = vbNullString
    
    
    ' check the icon exists
    On Error GoTo mnuAddMyPictures_Click_Error

    If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\my collection" & "\pictures.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\Icons\help.png"
    End If
       
    userprof = Environ$("USERPROFILE")

    If FExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        'Call menuAddSomething( iconImage, "My Pictures", "::{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}", vbNullString, vbNullString, vbNullString, vbNullString)
        Call menuAddSomething(iconImage, "My Pictures", userprof & "\Documents\Pictures", vbNullString, vbNullString, vbNullString, vbNullString)
    Else
        MsgBox "Unable to add my Pictures image as it does not exist"
    End If
        
   On Error GoTo 0
   Exit Sub

mnuAddMyPictures_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyPictures_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyVideos_Click
' Author    : beededea
' Date      : 07/03/2021
' Purpose   : .23 DAEB 07/03/2021 menu.frm Added menu option to add a "my Videos" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyVideos_Click()

    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim userprof As String: userprof = vbNullString
    

        
    ' check the icon exists
    On Error GoTo mnuAddMyVideos_Click_Error

    If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\my collection" & "\video-folder.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\Icons\help.png"
    End If
           
    userprof = Environ$("USERPROFILE")
       
    If FExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        'Call menuAddSomething( iconImage, "My Videos", "::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", vbNullString, vbNullString, vbNullString, vbNullString)
        Call menuAddSomething(iconImage, "My Videos", userprof & "\Documents\Videos", vbNullString, vbNullString, vbNullString, vbNullString)
    Else
        MsgBox "Unable to add my Videos image as it does not exist"
    End If
        
   On Error GoTo 0
   Exit Sub

mnuAddMyVideos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyVideos_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddEnhanced_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   : Menu option to add an enhanced settings utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddEnhanced_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
    On Error GoTo mnuAddEnhanced_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddEnhanced_click"

    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\SteamyRocket.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '[icons]
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    ' 17/11/2020    .04 DAEB Replaced all occurrences of rocket1.exe with iconsettings.exe

    Call menuAddSomething(iconImage, "Enhanced Icon Settings", App.Path & "\iconsettings.exe", vbNullString, vbNullString, vbNullString, vbNullString)
    'Call menuAddSomething( iconImage, "Enhanced Icon Settings", "[Icons]", vbNullString, vbNullString, vbNullString, vbNullString)
   
   On Error GoTo 0
   Exit Sub

mnuAddEnhanced_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddEnhanced_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddDocklet_click
' Author    : beededea
' Date      : 16/09/2019
' Purpose   : menu option to add a docklet
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDocklet_click() ' disable this if RD is not the defaultdock
   
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString


    Const x_MaxBuffer = 256
    
    On Error GoTo mnuAddDocklet_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddDocklet_click"
    
    ' set the default folder to the docklet folder under rocketdock
    If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
        dialogInitDir = rdAppPath & "\docklets"
    Else
        dialogInitDir = sdAppPath & "\docklets"
    End If
            
    With x_OpenFilename
    '    .hwndOwner = Me.hWnd
      .hInstance = App.hInstance
      .lpstrTitle = "Select a Rocketdock Docklet DLL"
      .lpstrInitialDir = dialogInitDir
      
      .lpstrFilter = "DLL Files" & vbNullChar & "*.dll" & vbNullChar & vbNullChar
      .nFilterIndex = 2
      
      .lpstrFile = String$(x_MaxBuffer, 0)
      .nMaxFile = x_MaxBuffer - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = x_MaxBuffer - 1
      .lStructSize = Len(x_OpenFilename)
    End With
          
'    Dim retFileName As String
'    Dim retfileTitle As String
    Call getFileNameAndTitle(retFileName, retfileTitle)
    txtTarget.Text = retFileName
    'txtLabelName.Text = retfileTitle
      
    If txtTarget.Text <> "" Then
        ' check the folder is valid docklet folder (beneath the docklets folder)
        ' set it to the docklet image yet to be created
        ' if it is a clock docklet use a temporary clock image just as RD does without hands?
        ' if it is a weather docklet use a temporary weather image of my own making
        ' if it is a recycling docklet use a temporary recycling image of my own making
        
        ' set the icon to that used by the docklet, it a mere guess as we cannot read the docklet DLL at this stage
        ' to determine what icon image it intends to use, later it writes to the 'other' settings.ini file in docklets
        ' but that's of no use now.
        If defaultDock = 0 Then
            If InStr(getFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
              txtCurrentIcon.Text = rdAppPath & "\icons\clock.png"
            ElseIf InStr(getFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
              txtCurrentIcon.Text = App.Path & "\my collection\recyclebin-full.png"
            Else
              txtCurrentIcon.Text = sdAppPath & "\blank.png" ' has to be an icon of some sort
            End If
        Else ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
            If InStr(getFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
              txtCurrentIcon.Text = sdAppPath & "\icons\clock.png"
            ElseIf InStr(getFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
              txtCurrentIcon.Text = App.Path & "\my collection\recyclebin-full.png"
            Else
              txtCurrentIcon.Text = sdAppPath & "\blank.png" ' has to be an icon of some sort
            End If
        End If
          
         '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call menuAddSomething(txtCurrentIcon.Text, "Docklet", vbNullString, vbNullString, vbNullString, txtTarget.Text, vbNullString)
        
        ' disable the fields, only enable the target fields and use the target field as a temporary location for the docklet data
        
        txtLabelName.Enabled = False
        txtCurrentIcon.Enabled = False
        
        sDockletFile = txtTarget.Text
        txtTarget.Enabled = True
        btnTarget.Enabled = True
        
        txtArguments.Enabled = False
        txtStartIn.Enabled = False
        cmbRunState.Enabled = False
        cmbOpenRunning.Enabled = False
        chkRunElevated.Enabled = False
        btnSelectStart.Enabled = False
    End If
    
    'triggerRdMapRefresh = True
        
   On Error GoTo 0
   Exit Sub

mnuAddDocklet_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDocklet_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFileListView_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : switch to tree view
'---------------------------------------------------------------------------------------
'
Private Sub btnFileListView_Click()
   On Error GoTo btnFileListView_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnFileListView_Click"
      
    Call busyStart
    If filesIconList.Visible = True Then
        picFrameThumbs.Visible = True
        filesIconList.Visible = False
        btnThumbnailView.Visible = False
        btnFileListView.Visible = True
    Else
        picFrameThumbs.Visible = False
        filesIconList.Visible = True
        If filesIconList.ListCount <= 0 Then
            frmNoFilesFound.Visible = True
            frmNoFilesFound.ZOrder
        End If
        btnThumbnailView.Visible = True
        btnFileListView.Visible = False
        srcDragControl = "filesIconList"

    End If
    Call busyStop
   On Error GoTo 0
   Exit Sub

btnFileListView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFileListView_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkRunElevated_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : If the checkbox for extended rocketdock menu options is selected
'---------------------------------------------------------------------------------------
'
Private Sub chkRunElevated_Click()
   On Error GoTo chkRunElevated_Click_Error
      If debugflg = 1 Then DebugPrint "%" & "chkRunElevated_Click"
           
        btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
        btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

chkRunElevated_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRunElevated_Click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : comboIconTypesFilter_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Selecting the type of icon that can be displayed from a dropbox
'---------------------------------------------------------------------------------------
'
Private Sub comboIconTypesFilter_Click()

    Dim filterType As Integer: filterType = 0
    'Dim validIconTypes As String
    
    On Error GoTo comboIconTypesFilter_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "comboIconTypesFilter_Click"

    filterType = comboIconTypesFilter.ListIndex
    
    ' read the current filter type and display the chosen images
    
    If filterType = 0 Then
        validIconTypes = "*.jpg;*.jpeg;*.bmp;*.ico;*.png;*.tif;*.gif"
    End If
    If filterType = 1 Then
        validIconTypes = "*.png"
    End If
    If filterType = 2 Then
        validIconTypes = "*.tif"
    End If
    If filterType = 3 Then
        validIconTypes = "*.bmp"
    End If
    If filterType = 4 Then
        validIconTypes = "*.jpg;*.jpeg"
    End If
    If filterType = 5 Then
        validIconTypes = "*.ico"
    End If
    
    filesIconList.Pattern = validIconTypes
    If filesIconList.ListCount > 0 Then
        filesIconList.ListIndex = (0) ' click the item in the underlying file list box
    End If
                
    ' now refresh the thumbnail display
    Call btnRefresh_Click_Event
    
    If filesIconList.ListIndex <> -1 Then ' when files found of this type
        filesIconList.ListIndex = (0)
    End If
    
    If filesIconList.Visible = True Then
        If filesIconList.ListCount <= 0 Then
            frmNoFilesFound.Visible = True
            frmNoFilesFound.ZOrder
        Else
            frmNoFilesFound.Visible = False
        End If
    Else
        frmNoFilesFound.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

comboIconTypesFilter_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comboIconTypesFilter_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbOpenRunning_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbOpenRunning_Click()
   On Error GoTo cmbOpenRunning_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbOpenRunning_Click"

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

cmbOpenRunning_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbOpenRunning_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbRunState_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbRunState_Click()
   On Error GoTo cmbRunState_Click_Error
   If debugflg = 1 Then DebugPrint "%cmbRunState_Click"

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

cmbRunState_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbRunState_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnCancel_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCancel_Click()
    Dim Filename As String: Filename = vbNullString
    
    On Error GoTo btnCancel_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnCancel_Click"
    
    settingsTimer.Enabled = False

    If mapImageChanged = True Then
        ' now change the icon image back again
        ' the target picture control and the icon size
        Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
        mapImageChanged = False
    End If
    
    If FExists(interimSettingsFile) Then '
        'get the rocketdock settings.ini for this icon alone
        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile
    'Else
        'readRegistryIconValues (rdIconNumber)
    End If
    
    'reset all the current icon displayed characteristics
    txtCurrentIcon.Text = sFilename
    txtLabelName.Text = sTitle
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory
    cmbRunState.ListIndex = Val(sShowCmd) - 1 ' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
    cmbOpenRunning.ListIndex = Val(sOpenRunning)
    chkRunElevated.Value = Val(sRunElevated)
    chkConfirmDialog.Value = Val(sUseDialog)     ' .03 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    chkConfirmDialogAfter.Value = Val(sUseDialogAfter)
    chkQuickLaunch.Value = Val(sQuickLaunch)     '.nn Added new check box to allow a quick launch of the chosen app
    chkAutoHideDock.Value = Val(sAutoHideDock)  '.nn Added new check box to allow autohide of the dock after launch of the chosen app
    txtSecondApp.Text = sSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
    optRunSecondAppBeforehand.Value = Val(sRunSecondAppBeforehand)
    txtAppToTerminate.Text = sAppToTerminate
    chkDisabled.Value = Val(sDisabled)
    
    ' display the icon from the alternative settings.ini config.
    Filename = txtCurrentIcon.Text
    
    Call displayResizedImage(Filename, picPreview, icoSizePreset)
    
    ' we signify that all changes have been lost
    iconChanged = False
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

   On Error GoTo 0
   Exit Sub

btnCancel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnCancel_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnAddFolder_Click
' Author    : beededea
' Date      : 17/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAddFolder_Click()
    Dim savTextCurrentFolder As String
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    
    ' add the custom folder to the treeview

    ' check to see if the customfolder has been previously assigned a place in the treeview
    ' read the toolSettings.ini first
        
    ' read the settings ini file
    'eg. rDCustomIconFolder=?E:\dean\steampunk theme\icons\
    On Error GoTo btnAddFolder_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnAddFolder_Click"
   
    If FExists(interimSettingsFile) Then
        rDCustomIconFolder = GetINISetting("Software\SteamyDock\DockSettings", "rDCustomIconFolder", interimSettingsFile)
    End If
    
    If rDCustomIconFolder = "?" Then
        ' currently do nothing here
    Else

        If folderTreeView.SelectedItem.Text = "my collection" Or folderTreeView.SelectedItem.Text = "icons" Then
            ' do nothing
        Else
            ' if the customfolder has been set then remove it first from the .ini
            ' and remove it from the tree
            ' remove the ?
            
            If Left$(rDCustomIconFolder, 1) = "?" Then
                folderTreeView.SelectedItem.Key = Mid$(rDCustomIconFolder, 2)
            Else
                folderTreeView.SelectedItem.Key = rDCustomIconFolder
            End If
            Call btnRemoveFolder_Click
        End If
    End If
    
    savTextCurrentFolder = textCurrentFolder.Text 'save the current default folder
    
'
'    If txtStartIn.Text <> vbNullString Then
'        If DirExists(txtStartIn.Text) Then
'            dialogInitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
'        Else
'            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
'        End If
'    End If
'
    dialogInitDir = ""

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder

    'getFolder = ChooseDir_Click ' show the dialog box to select a folder
    If getFolder = vbNullString Then
        'textCurrentFolder.Text = savTextCurrentFolder
        Exit Sub
    End If
    If getFolder <> vbNullString Then
        textCurrentFolder.Text = getFolder
    End If
    

    Call busyStart ' this was meant to cause the egg timer to appear but it no longer works from here
    
    ' add the chosen folder to the treeview
    
    folderTreeView.Nodes.Add , , textCurrentFolder.Text, textCurrentFolder.Text
    Call addtotree(textCurrentFolder.Text, folderTreeView)
    folderTreeView.Nodes.Item(textCurrentFolder.Text).Text = "custom folder"
    
    'write the folder to the rocketdock settings file
    'eg. rDCustomIconFolder=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\SteamyDock\DockSettings", "rDCustomIconFolder", "?" & textCurrentFolder.Text, interimSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    Call busyStop

   On Error GoTo 0
   Exit Sub

btnAddFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") a folder with the same name already exists in the tree view, choose another folder"
    Call busyStop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnSaveRestart_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : save the current fields to the settings file or registry, RESTART
'---------------------------------------------------------------------------------------
'
Private Sub btnSaveRestart_Click()

    picBusy.Visible = True
    busyTimer.Enabled = True

    Call btnSaveRestart_Click_event(hwnd)
    
   On Error GoTo 0
   Exit Sub

btnSaveRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSaveRestart_Click of Form rDIconConfigForm"
            
End Sub





'---------------------------------------------------------------------------------------
' Procedure : btnSelectStart_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSelectStart_Click()
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
   
    On Error GoTo btnSelectStart_Click_Error
    If debugflg = 1 Then DebugPrint "%btnSelectStart_Click"
    If txtStartIn.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
            End If
        End If
    End If

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder
    'getFolder = ChooseDir_Click ' old method to show the dialog box to select a folder
    If getFolder <> vbNullString Then txtStartIn.Text = getFolder

   On Error GoTo 0
   Exit Sub

btnSelectStart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSelectStart_Click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ChooseDir_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function ChooseDir_Click() As String
'    Dim sTempDir As String
'    On Error GoTo ChooseDir_Click_Error
'       If debugflg = 1 Then DebugPrint "%" & "ChooseDir_Click"
'
'
'
'    On Error Resume Next
'
'    sTempDir = CurDir    'Remember the current active directory
'    rdDialogForm.CommonDialog.DialogTitle = "Select a directory" 'titlebar
'    If Not txtStartIn.Text = vbNullString Then
'        If DirExists(txtStartIn.Text) Then
'            rdDialogForm.CommonDialog.InitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
'        Else
'            rdDialogForm.CommonDialog.InitDir = rdAppPath 'start dir, might be "C:\" or so also
'        End If
'    End If
'    rdDialogForm.CommonDialog.FileName = "Select a Directory"  'Something in filenamebox
'    rdDialogForm.CommonDialog.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
'    rdDialogForm.CommonDialog.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
'    rdDialogForm.CommonDialog.CancelError = True 'allow escape key/cancel ' do NOT change, causes a hang
'    rdDialogForm.CommonDialog.ShowSave   'show the dialog screen
'
'    If Err <> 32755 Then
'        ChooseDir_Click = CurDir ' User didn't chose Cancel.
'    Else
'        ChooseDir_Click = vbNullString ' User chose Cancel.
'    End If
'
'    ChDir sTempDir  'restore path to what it was at entering

'   On Error GoTo 0
'   Exit Function
'
'ChooseDir_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChooseDir_Click of Form rDIconConfigForm"
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 02/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    ' show a single help PNG with pointers as to what does what
   On Error GoTo btnHelp_Click_Error
   If debugflg = 1 Then DebugPrint "%btnHelp_Click"

    'rdHelpForm.Show
    
    Call mnuHelpPdf_click

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : rdMapPageDown_Press
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : called by getkeypress, handles a page down on the rdIconMap
'---------------------------------------------------------------------------------------
'
Private Sub rdMapPageDown_Press()
    Dim exitSubFlg As Boolean: exitSubFlg = False
    'if the modification flag is set then ask before moving to the next icon
      
    Call preMapPageUpDown(exitSubFlg)
    
    If exitSubFlg = True Then Exit Sub
    
    'increment the icon number
    rdIconNumber = rdIconNumber + 15
    'check we haven't gone too far
    If rdIconNumber > rdIconMaximum Then rdIconNumber = rdIconMaximum
    
    ' only move the map if the array has been populated,
    If Not picRdMap(0).ToolTipText = vbNullString Then
        ' I want to test to see if the picture property is populated but
        ' as the picture property is not being set by Lavolpe's method then we can't test for it
        ' testing the tooltip is one method of seeing if the map has been created
        ' as the program sets the tooltip just when the transparent image is set
        
        ' moves the RdMap on one position (one click) if it is already set at the rightmost screen position
        If rdMapHScroll.Value < rdMapHScroll.Max - 15 Then
            rdMapHScroll.Value = rdIconNumber
        Else
            If rdIconNumber < rdMapHScroll.Max Then 'ignore the following if we are already at the end
                'scroll the map 15 places before the end
                If rdMapHScroll.Max - 15 >= 0 Then ' the scrollbar doesn't like to go less than zero
                    rdMapHScroll.Value = rdMapHScroll.Max - 15
                Else
                    rdMapHScroll.Value = 0
                End If
            End If
            ' we are now at the end
            rdIconNumber = rdMapHScroll.Max
        End If
    End If
    
    
    Call postMapPageUpDown


   On Error GoTo 0
   Exit Sub

rdMapPageDown_Press_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapPageDown_Press of Form rDIconConfigForm"
End Sub

Private Sub preMapPageUpDown(ByRef exitSubFlg As Boolean)
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    If debugflg = 1 Then DebugPrint "%" & "preMapPageDown"
    
    If btnSet.Enabled = True Then
        If chkToggleDialogs.Value = 1 Then
           If btnSet.Enabled = True Or mapImageChanged = True Then
                answer = msgBoxA(" This will lose your recent changes to this icon, are you sure?", vbYesNo, "changes have been made", True, "preMapPageUpDown_Press")
                If answer = vbNo Then
                    exitSubFlg = True
                    Exit Sub
                End If
            End If
        End If
        If mapImageChanged = True Then
            ' now change the icon image
            ' the target picture control and the icon size
            Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
            mapImageChanged = False
        End If
    End If
    
    'remove and reset the highlighting on the Rocket dock map
     picRdMap(rdIconNumber).BorderStyle = 0
End Sub


Private Sub postMapPageUpDown()
        
    lblRdIconNumber.Caption = Str$(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str$(rdIconNumber) + 1
    
    Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
    
    'set the highlighting on the Rocket dock map
    picRdMap(rdIconNumber).BorderStyle = 1

    previewFrameGotFocus = True

    ' we signify that all changes have been lost
    btnSet.Enabled = False ' this has to be done at the end
    'btnCancel.Caption = "Close"
    btnClose.Visible = True
    btnCancel.Visible = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : rdMapPageUp_Press
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : called by getkeypress, handles a page up on the rdIconMap
'---------------------------------------------------------------------------------------
'
Private Sub rdMapPageUp_Press()
    
    Dim exitSubFlg As Boolean: exitSubFlg = False
    'if the modification flag is set then ask before moving to the next icon
      
    Call preMapPageUpDown(exitSubFlg)
    
    If exitSubFlg = True Then Exit Sub
    
    'increment the icon number
    rdIconNumber = rdIconNumber - 15
    'check we haven't gone too far
    If rdIconNumber < 0 Then rdIconNumber = 0
    
    ' only move the map if the array has been populated,
    If Not picRdMap(0).ToolTipText = vbNullString Then
        rdMapHScroll.Value = rdIconNumber
    End If
    
    Call postMapPageUpDown

   On Error GoTo 0
   Exit Sub

rdMapPageUp_Press_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapPageUp_Press of Form rDIconConfigForm"
End Sub
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryOnce
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : read the registry and set obtain the necessary icon data for the specific icon
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryOnce(ByVal iconNumberToRead As Integer)
'    ' read the settings from the registry
'   On Error GoTo readRegistryOnce_Error
'   If debugflg = 1 Then DebugPrint "%" & "readRegistryOnce"
'
'
'
'    sFilename = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName")
'    sFileName2 = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName2")
'    sTitle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Title")
'    sCommand = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Command")
'    sArguments = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Arguments")
'    sWorkingDirectory = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-WorkingDirectory")
'    sShowCmd = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-ShowCmd")
'    sOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-OpenRunning")
'    sIsSeparator = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-IsSeparator")
'    sUseContext = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseContext")
'    sDockletFile = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-DockletFile")
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryOnce_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryOnce of Form rDIconConfigForm"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : writeRegistryOnce
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub writeRegistryOnce(ByVal iconNumberToWrite As Integer)
'
'   On Error GoTo writeRegistryOnce_Error
'    If debugflg = 1 Then DebugPrint "%" & "writeRegistryOnce"
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", sFilename)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", sFileName2)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Title", sTitle)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Command", sCommand)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", sArguments)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", sShowCmd)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", sOpenRunning)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", sIsSeparator)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", sUseContext)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", sDockletFile)
'
'   On Error GoTo 0
'   Exit Sub
'
'writeRegistryOnce_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistryOnce of Form rDIconConfigForm"
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : ExtractSuffix
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Function ExtractSuffix(ByVal strPath As String) As String
'    Dim AY() As String ' string array
'    Dim Max As Integer
'
'    On Error GoTo ExtractSuffix_Error
'    If debugflg = 1 Then DebugPrint "%" & "ExtractSuffix"
'
'    If strPath = "" Then
'        ExtractSuffix = ""
'        Exit Function
'    End If
'
'    AY = Split(strPath, ".")
'    Max = UBound(AY)
'    ExtractSuffix = Trim$(AY(Max))
'
'   On Error GoTo 0
'   Exit Function
'
'ExtractSuffix_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffix of Form rDIconConfigForm"
'End Function


''---------------------------------------------------------------------------------------
'' Procedure : ExtractSuffixWithDot
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function ExtractSuffixWithDot(ByVal strPath As String) As String
'    Dim AY() As String ' string array
'    Dim Max As Integer
'
'    On Error GoTo ExtractSuffixWithDot_Error
'    If debugflg = 1 Then Debug.Print "%" & "ExtractSuffixWithDot"
'
'    If strPath = vbNullString Then
'        ExtractSuffixWithDot = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        AY = Split(strPath, ".")
'        Max = UBound(AY)
'        ExtractSuffixWithDot = Trim$("." & AY(Max))
'    Else
'        ExtractSuffixWithDot = Trim$(strPath)
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'ExtractSuffixWithDot_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffixWithDot of Form dock"
'End Function
'---------------------------------------------------------------------------------------
' Procedure : displayResizedImage was previously displayPreviewImage
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Displays and places resized image onto the specified picture box.
'             Just needs a filename of the file to display, a target picbox and a size.
'
'             Uses two methods to display 1. native and 2. non-native image file types
'             both methods are supplied by LaVolpe.
'---------------------------------------------------------------------------------------
'
Private Sub displayResizedImage(ByRef Filename As String, ByRef targetPicBox As PictureBox, ByRef IconSize As Integer)
    Dim suffix As String: suffix = vbNullString
    Dim picWidth As Long: picWidth = 0
    Dim picHeight As Long: picHeight = 0
    Dim picSize As Long: picSize = 0
                
    On Error GoTo displayResizedImage_Error
    If debugflg = 1 Then DebugPrint "%" & "displayResizedImage"
    ' .36 DAEB 20/04/2021 rdIconConfig.frm Add a final check that the chosen image file actually exists
    If Not FExists(Filename) Then
        If FExists(App.Path() & "\my collection\" & "red-X.png") Then
            Filename = App.Path() & "\my collection\" & "red-X.png"
        End If
        'Exit Sub    ' just a final check that the chosen image file actually exists
    End If
    
    textCurrIconPath.Text = filesIconList.Filename
    ' find and store the indexed position of the chosen file into the global variable
    fileIconListPosition = filesIconList.ListIndex
    
    ' dispose of the image prior to use
    Set targetPicBox.Picture = Nothing ' added because the two methods of drawing an image conflict leaving an image behind
    
    suffix = Trim$(ExtractSuffix(Filename))
    
    ' using Lavolpe's later method as it allows for resizing of PNGs and all other types
    If InStr("png,jpg,bmp,jpeg,tif,gif", LCase$(suffix)) <> 0 Then
        If targetPicBox.Name = "picPreview" Then
            targetPicBox.Left = 345
            targetPicBox.Top = 210
            targetPicBox.Width = 3450
            targetPicBox.Height = 3450
        End If
        
        Set cImage = New c32bppDIB
        cImage.LoadPictureFile Filename, IconSize, IconSize, False, 32
        Call refreshPicBox(targetPicBox, IconSize)
        
        ' see ref point 0001 in cPNGparser.cls for PNG size extraction
        lblWidthHeight.Caption = " width " & origWidth & " height " & origHeight & " (pixels)"

    ElseIf InStr("ico", LCase$(suffix)) <> 0 Then
        ' *.ico
        ' using Lavolpe's earlier StdPictureEx method as it allows for correct display of ICOs
        ' the later method above has a bug with some ICOs
        
        'because the earlier method draws the ico images from the top left of the
        'pictureBox we have to manually set the picbox to size and position for each icon size
        Call centrePreviewImage(targetPicBox, IconSize)
        Set targetPicBox.Picture = StdPictureEx.LoadPicture(Filename, lpsCustom, , IconSize, IconSize)
    End If


    ' display the sizes from the image types that are native to VB6
    
    ' check the size of the image and display it,
    ' unlike the .NET version, the sizing has to be done after the display of the image
    ' as it is LaVolpe's code that does the extraction of the icon count.
    'FileName = "C:\Program Files\Rocketdock\"
    If InStr("jpg,bmp,jpeg,gif", LCase$(suffix)) <> 0 Then
        Call checkImageSize(Filename, picWidth, picHeight) 'check the size of the image
        lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
    ElseIf InStr("ico", LCase$(suffix)) <> 0 Then
        ' captureIconCount is obtained elsewhere in Lavolpe's StdPictureEx code
        If captureIconCount = 1 Then

            On Error GoTo handleResizing_Error
            Call checkImageSize(Filename, picWidth, picHeight) 'check the size of the image
            GoTo displaySizes ' don't want to use a goto but error handling in VB6...
            
handleResizing_Error:
                
             ' if the ico file is damaged then display a blank icon
             ' an example of damage is an icon with an incorrect header count of thousands
             ' Note: Lavolpe's code will still display icons that are considered damaged by Windows.
             
             targetPicBox.ToolTipText = "This icon is damaged - " & Filename
             If FExists(App.Path() & "\my collection\" & "red-X.png") Then
                 Filename = App.Path() & "\my collection\" & "red-X.png"
             ElseIf FExists(sdAppPath & "\icons\" & "help.png") Then
                 Filename = sdAppPath & "\icons\" & "help.png" '' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
             ElseIf FExists(rdAppPath & "\icons\" & "help.png") Then
                 Filename = rdAppPath & "\icons\" & "help.png"
             End If
             
             'display image here after the error is handled
             Set targetPicBox.Picture = Nothing

             Set cImage = New c32bppDIB
             cImage.LoadPictureFile Filename, IconSize, IconSize, False, 32
             Call refreshPicBox(targetPicBox, IconSize)

             lblWidthHeight.Caption = " This is a damaged icon." ' < must go here.

             Exit Sub

        End If
        
displaySizes:

        If InStr("ico", LCase$(suffix)) <> 0 Then
            If captureIconCount > 1 Then
                lblWidthHeight.Caption = " multiple size (" & captureIconCount & ") ICO file"
            Else
                lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
            End If
        ElseIf InStr("TIFF", LCase$(suffix)) <> 0 Then
            lblWidthHeight.Caption = " no sizes obtained"
        Else
            ' PNG is not a native image type for VB6
            ' There is no native method of obtaining width and height for PNG, ICO, TIFF &c
            ' so instead we use a 3rd party method.
            ' see ref point 0001 in cPNGparser.cls for PNG size extraction
            'lblWidthHeight.Caption = " width " & origWidth & " height " & origHeight & " (pixels)"
            'lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
        End If
            
    End If
        
    ' DAEB TBD
    If InStr("exe,dll", LCase$(suffix)) <> 0 Then
                Call displayEmbeddedIcons(Filename, targetPicBox, IconSize)
                picSize = FileLen(Filename)
                lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (binary)"
    End If
    
    ' .46 DAEB 16/04/2022 rdIconConfig.frm Made the word Blank visible or not during an file manager icon click
'    If getFileNameFromPath(Filename) = "blank.png" Then
'        lblBlankText.Visible = True
'    Else
'        lblBlankText.Visible = False
'    End If
    
   On Error GoTo 0
   Exit Sub

displayResizedImage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayResizedImage of Form rDIconConfigForm"
                
End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkImageSize
' Author    : beededea
' Date      : 12/11/2019
' Purpose   : check the size of the image (same as the code above but without the same error checking)
'---------------------------------------------------------------------------------------
'
Private Sub checkImageSize(ByRef Filename As String, ByRef picWidth As Long, ByRef picHeight As Long)
    
    If debugflg = 1 Then DebugPrint "%checkImageSize"

    'create an original size bitmap
    Dim bmpsizingImage As StdPicture
            
    ' the on error must not be activated within this routine, if it fails it goes to the calling routine
   'On Error GoTo checkImageSize_Error
   
    ' if the ico file has a corrupt header it will fail the loadpicture
    Set bmpsizingImage = LoadPicture(Filename)
    
    ' determine the actual dimensions in pixels
    picWidth = ScaleX(bmpsizingImage.Width, vbHimetric, vbPixels)
    picHeight = ScaleY(bmpsizingImage.Height, vbHimetric, vbPixels)
    
    ' dispose of the temporarily created sizing image
    Set bmpsizingImage = Nothing

   On Error GoTo 0
   Exit Sub

checkImageSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkImageSize of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnHomeRdMap
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : on the RD Map press the HOME button to move to the beginning
'---------------------------------------------------------------------------------------
'
Private Sub btnHomeRdMap()

    Dim answer As VbMsgBoxResult: answer = vbNo
    'Dim Filename As String: Filename = vbNullString
    'Dim useloop As Integer
    'Dim ff As Long
    'if the modification flag is set then ask before moving to the next icon
    On Error GoTo btnHomeRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "btnHomeRdMap"
   
    If btnSet.Enabled = True Then
        If chkToggleDialogs.Value = 1 Then
           If btnSet.Enabled = True Or mapImageChanged = True Then

                answer = msgBoxA(" This will lose your recent changes to this icon, are you sure?", vbYesNo, "changes have been made", True, "btnHomeRdMap")
                If answer = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    rdMapHScroll.Value = rdMapHScroll.Min
    
    ' set the primary selection in the map
    'Call picRdMap_MouseDown(rdMapHScroll.Value, 1, 0, 0, 0) ' DEAN
    
    ' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
    Call picRdMap_MouseDown_event(rdMapHScroll.Value)

    
    ' we signify that all changes have been lost
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False


   On Error GoTo 0
   Exit Sub

btnHomeRdMap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHomeRdMap of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnEndRdMap
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : on the RD Map press the END button to move to the beginning
'---------------------------------------------------------------------------------------
'
Private Sub btnEndRdMap()
    Dim answer As VbMsgBoxResult: answer = vbNo
    'Dim Filename As String: Filename = vbNullString
    'Dim useloop As Integer
    'Dim ff As Long
    
    'if the modification flag is set then ask before moving to the next icon
    On Error GoTo btnEndRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "btnEndRdMap"
   
    If btnSet.Enabled = True Then
        If chkToggleDialogs.Value = 1 Then
           If btnSet.Enabled = True Or mapImageChanged = True Then
                answer = msgBoxA(" This will lose your recent changes to this icon, are you sure?", vbYesNo, "changes have been made", True, "btnEndRdMap")
                If answer = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    picRdMap(rdIconNumber).BorderStyle = 0
        
    rdMapHScroll.Value = rdMapHScroll.Max - 15
    rdIconNumber = rdMapHScroll.Value
    
        ' set the primary selection in the map
    'Call picRdMap_MouseDown(rdMapHScroll.Max, 1, 0, 0, 0) ' DEAN
    
    ' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
    Call picRdMap_MouseDown_event(rdMapHScroll.Max)
    
    'picRdMap(rdMapHScroll.Max).BorderStyle = 1

    ' we signify that all changes have been lost
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False


   On Error GoTo 0
   Exit Sub

btnEndRdMap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnEndRdMap of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnPrev_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Clicking the previous icon button next to the displayed icon
'---------------------------------------------------------------------------------------
'
Private Sub btnPrev_Click()

    Dim exitSubFlg As Boolean: exitSubFlg = False
    
    On Error GoTo btnPrev_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnPrev_Click"
    
    Call preButtonClick(exitSubFlg)
    
    If exitSubFlg = True Then Exit Sub
        
    'decrement the icon number
    rdIconNumber = rdIconNumber - 1
    If rdIconNumber < 0 Then rdIconNumber = 0     'check we haven't gone too far
    
    Call postButtonClick

   On Error GoTo 0
   Exit Sub

btnPrev_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrev_Click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnNext_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : click the next icon in the MAP
'---------------------------------------------------------------------------------------
'
Private Sub btnNext_Click()
    
    Dim exitSubFlg As Boolean: exitSubFlg = False
    
    On Error GoTo btnNext_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnNext_Click"

    Call preButtonClick(exitSubFlg)
    
    If exitSubFlg = True Then Exit Sub
        
    'increment the icon number
    rdIconNumber = rdIconNumber + 1
    'check we haven't gone too far
    If rdIconNumber > rdIconMaximum Then rdIconNumber = rdIconMaximum

    Call postButtonClick
    
   On Error GoTo 0
   Exit Sub

btnNext_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnNext_Click of Form rDIconConfigForm"
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : preButtonClick
' Author    : beededea
' Date      : 06/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub preButtonClick(ByVal exitSubFlg As Boolean)
    
    'if the modification flag is set then ask before moving to the next icon
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    On Error GoTo preButtonClick_Error

    If debugflg = 1 Then DebugPrint "%" & "btnPrev_Click"

    If btnSet.Enabled = True Then
        ' 17/11/2020    .03 DAEB Replaced the confirmation dialog with an automatic save when moving from one icon to another using the right/left icon buttons
        If chkToggleDialogs.Value = 1 Then
           If btnSet.Enabled = True Or mapImageChanged = True Then
                answer = msgBoxA(" This will lose your recent changes to this icon, are you sure?", vbQuestion + vbYesNo, "Selecting Another Icon After Changes Made", True, "preButtonClick")
                If answer = vbNo Then
                    exitSubFlg = True
                    Exit Sub
                End If
            End If
        Else
            Call btnSet_Click
        End If
        If mapImageChanged = True Then
            ' now change the icon image back again
            ' the target picture control and the icon size
            Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
            mapImageChanged = False
        End If

    End If

    On Error GoTo 0
    Exit Sub

preButtonClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure preButtonClick of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : postButtonClick
' Author    : beededea
' Date      : 06/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub postButtonClick()
    
    ' only move the map if the array has been populated,
    On Error GoTo postButtonClick_Error

    If Not picRdMap(0).ToolTipText = vbNullString Then
        ' I want to test to see if the picture property is populated but
        ' as the picture property is not being set by Lavolpe's method then we can't test for it
        ' testing the tooltip above is one method of seeing if the map has been created
        ' as the program sets the tooltip just when the transparent image is set
    
        ' moves the RdMap on one position (one click) if it is already set at the rightmost screen position
        If rdIconNumber < rdMapHScroll.Value Then
            btnMapPrev_Click
        End If
    End If

    lblRdIconNumber.Caption = Str$(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str$(rdIconNumber) + 1
    
    Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
    
    'remove and reset the highlighting on the Rocket dock map
    
    picRdMap(rdIconNumber - 1).BorderStyle = 0
    picRdMap(rdIconNumber + 1).BorderStyle = 0
    picRdMap(rdIconNumber).BorderStyle = 1
    
    previewFrameGotFocus = True

    ' we signify that all changes have been lost
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

    On Error GoTo 0
    Exit Sub

postButtonClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure postButtonClick of Form rDIconConfigForm"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : displayIconElement
' Author    : beededea
' Date      : 27/09/2019
' Purpose   : Display the icon details, text, sizing and image extracted via various methods according to source
'---------------------------------------------------------------------------------------
'
' .79 DAEB 28/05/2022 rDIConConfig.frm new parameter to determine when to populate the dragicon
Private Sub displayIconElement(ByVal iconCount As Integer, ByRef picBox As PictureBox, fillPicBox As Boolean, ByRef icoPreset As Integer, ByVal showProperties As Boolean, ByVal fillDragIcon As Boolean, Optional ByVal showBlank As Boolean)
    Dim Filename As String: Filename = vbNullString
    Dim qPos As Long: qPos = 0
    Dim filestring As String: filestring = vbNullString
    Dim suffix As String: suffix = vbNullString
    Dim picSize As Long: picSize = 0
    
    'if it is a good icon then read the data
    On Error GoTo displayIconElement_Error
    If debugflg = 1 Then DebugPrint "%" & "displayIconElement"

    If FExists(interimSettingsFile) Then '
        'get the rocketdock alternative settings.ini for this icon alone
        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", iconCount, interimSettingsFile
    End If

    ' .46 DAEB 16/04/2022 rdIconConfig.frm Made the word Blank visible or not when clicking on the icon map
    If getFileNameFromPath(sFilename) = "blank.png" Then
        If showBlank = True Then lblBlankText.Visible = True
    Else
        lblBlankText.Visible = False
    End If

    'showProperties = True
    If showProperties = True Then
        ' if the incoming text has <quote> then replace those with a " TODO ?
        txtCurrentIcon.Text = sFilename ' build the full path
        
        txtLabelName.Text = sTitle
        txtTarget.Text = sCommand
        txtArguments.Text = sArguments
        txtStartIn.Text = sWorkingDirectory
        
        'If the docklet entry in the settings.ini is populated then blank off all the target, folder and image fields
        If sDockletFile <> "" Then
              txtLabelName.Enabled = False
              txtCurrentIcon.Enabled = False
              
              'only enable the target fields and use the target field as a temporary location for the docklet data
              txtTarget.Text = sDockletFile
              txtTarget.Enabled = True
              btnTarget.Enabled = True
              
              txtArguments.Enabled = False
              txtStartIn.Enabled = False
              cmbRunState.Enabled = False
              cmbOpenRunning.Enabled = False
              chkRunElevated.Enabled = False
              btnSelectStart.Enabled = False
        Else
              txtLabelName.Enabled = True
              txtCurrentIcon.Enabled = True
              txtTarget.Enabled = True
              txtArguments.Enabled = True
              txtStartIn.Enabled = True
              cmbRunState.Enabled = True


                chkRunElevated.Enabled = True
                cmbOpenRunning.Enabled = True

              btnTarget.Enabled = True
              btnSelectStart.Enabled = True
        End If
        
        If sIsSeparator = "1" Then
              txtLabelName.Text = "Separator"
              txtLabelName.Enabled = False
              txtCurrentIcon.Enabled = False
              txtTarget.Enabled = False
              btnTarget.Enabled = False
              txtArguments.Enabled = False
              txtStartIn.Enabled = False
              cmbRunState.Enabled = False
              cmbOpenRunning.Enabled = False
              chkRunElevated.Enabled = False
              btnSelectStart.Enabled = False
        End If
        
        If sShowCmd = "0" Then ' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
            sShowCmd = "1"
        End If
        
        cmbRunState.ListIndex = Val(sShowCmd) - 1 ' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
        cmbOpenRunning.ListIndex = Val(sOpenRunning)
        chkRunElevated.Value = Val(sRunElevated)
        
        '.Value = Val(sUseContext)
        'If defaultDock = 1 Then  ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
            chkConfirmDialog.Value = Val(sUseDialog)
            chkConfirmDialogAfter.Value = Val(sUseDialogAfter)
            chkQuickLaunch.Value = Val(sQuickLaunch)     '.nn Added new check box to allow a quick launch of the chosen app
            chkAutoHideDock.Value = Val(sAutoHideDock)  '.nn Added new check box to allow autohide of the dock after launch of the chosen app
            txtSecondApp.Text = sSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
            
            optRunSecondAppBeforehand.Value = Val(sRunSecondAppBeforehand)
            If optRunSecondAppBeforehand.Value = False Then optRunSecondAppAfterward.Value = True
            txtAppToTerminate.Text = sAppToTerminate
            If sDisabled = "9" Then sDisabled = ""
            chkDisabled.Value = Val(sDisabled)
        'End If
        
        If txtSecondApp.Text = "" Then
            optRunSecondAppBeforehand.Enabled = False
            optRunSecondAppAfterward.Enabled = False
            lblRunSecondAppBeforehand.Enabled = False
            lblRunSecondAppAfterward.Enabled = False
        Else
            optRunSecondAppBeforehand.Enabled = True
            optRunSecondAppAfterward.Enabled = True
            lblRunSecondAppBeforehand.Enabled = True
            lblRunSecondAppAfterward.Enabled = True
        End If
        
    End If
    
    'If the docklet entry in the settings.ini is populated then set a helpful tooltiptext
    If sDockletFile <> "" Then
        picBox.ToolTipText = "Icon number " & iconCount + 1 & "You can modify this docklet by selecting a new target, click on the ... button next to the target field."
    Else
        picBox.ToolTipText = "Icon number " & iconCount + 1 & " = " & sFilename
    End If
    picPreview.Tag = sFilename
    
    suffix = ExtractSuffix(LCase$(sFilename))

    ' test whether it is a valid file with a path or just a relative path
    If InStr(sFilename, "?") Then
        Filename = sFilename
        lblFileInfo.Caption = ""
    ElseIf FExists(sFilename) Then
        Filename = sFilename  ' a full valid path so leave it alone
        picSize = FileLen(Filename)
        lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
    Else
        If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
            Filename = rdAppPath & "\" & sFilename ' a relative path found as per Rocketdock
        Else
            Filename = sdAppPath & "\" & sFilename ' a relative path found as per Rocketdock
        End If
        If FExists(Filename) Then
            picSize = FileLen(Filename)
            lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
            txtCurrentIcon.Text = Filename
            
            ' if the path is the relative path from the RD folder then repair it giving it a full path
            sFilename = Filename
            
        End If
    End If

    ' if the user drags an icon to the dock then RD takes a icon link of the following form:
    'FileName = "C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\vbexpress.exe?62453184"
    
    If InStr(sFilename, "?") Then ' Note: the question mark is an illegal character and test for a valid file will fail in VB.NET despite working in VB6 so we test it as a string instead
        ' does the string contain a ? if so it probably has an embedded .ICO
        qPos = InStr(1, Filename, "?")
        If qPos <> 0 Then
            ' extract the string before the ? (qPos)
            filestring = Mid$(Filename, 1, qPos - 1)
        End If
        
        ' test the resulting filestring exists
        If FExists(filestring) Then
            ' extract the suffix
            suffix = ExtractSuffix(filestring)

            'suffix = right$(filestring, Len(filestring) - InStr(1, filestring, "."))
            ' test as to whether it is an .EXE or a .DLL
            If InStr("exe,dll", LCase$(suffix)) <> 0 Then
                'FileName = txtCurrentIcon.Text ' revert to the relative path which is what is expected
                If fillPicBox = True Then Call displayEmbeddedIcons(filestring, picBox, icoPreset)
                picSize = FileLen(filestring)
                lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (binary)"

            Else
                ' the file may have a ? in the string but does not match otherwise in any useful way
                If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                    Filename = rdAppPath & "\icons\" & "help.png"
                Else
                    Filename = sdAppPath & "\icons\" & "help.png"
                End If
            End If
            
        Else ' the file doesn't exist in any form with ? or otherwise as a valid path
            If sIsSeparator = 1 Then
                Filename = App.Path & "\my collection\" & "separator.png" ' change to separator
            Else
                If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                    Filename = rdAppPath & "\icons\" & "help.png"
                Else
                    Filename = sdAppPath & "\icons\" & "help.png"
                End If
            End If
            If fillPicBox = True Then Call displayResizedImage(Filename, picBox, icoPreset)
            'dllFrame.Visible = False
        End If
    Else
        If fillPicBox = True Then
            Call displayResizedImage(Filename, picBox, icoPreset) ' fill the main picture box
            
            ' .78 DAEB 28/05/2022 rDIConConfig.frm We should only fill the temporary store when this routine has been called due to a click on the map
            
            ' now fill the temporary holder for the dragicon
            If fillDragIcon = True Then
                ' .78 DAEB 28/05/2022 rDIConConfig.frm Dragging a blank icon within the map should show a drag image a .lnk? Possibly a white box with a thin black boundary.
                If getFileNameFromPath(sFilename) = "blank.png" Then
                    Call displayResizedImage(App.Path & "\resources\mapBlank.ico", picTemporaryStore, 64) ' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon.
                Else
                    Call displayResizedImage(Filename, picTemporaryStore, 64) ' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon.
                End If
                
            End If
            
        End If
        'dllFrame.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

displayIconElement_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayIconElement of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnThumbnailView_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Click on the thumbnail view button top right to switch from file to thumb
'             also used for causing the thumbnail list to refresh
'---------------------------------------------------------------------------------------
'
Private Sub btnThumbnailView_Click()
    On Error GoTo btnThumbnailView_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnThumbnailView_Click"
    thisRoutine = "btnThumbnailView_Click"
    
    If filesIconList.ListIndex >= 0 Then
        'if the two are the same it ought not to trigger a pointless click
        If vScrollThumbs.Value <> filesIconList.ListIndex Then
            vScrollThumbs.Value = filesIconList.ListIndex
        End If
    End If
    
    ' .67 DAEB 04/05/2022 rDIConConfig.frm Drag and drop from the filelist to the rdmap
    srcDragControl = "picThumbIcon"
    
    picFrameThumbs.Visible = True
    filesIconList.Visible = False
    frmNoFilesFound.Visible = False
    btnThumbnailView.Visible = False
    btnFileListView.Visible = True
    
    refreshThumbnailView = True
    

   On Error GoTo 0
   Exit Sub

btnThumbnailView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnThumbnailView_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : refreshThumbnailViewPanel
' Author    : beededea
' Date      : 04/11/2019
' Purpose   : straight refresh
'---------------------------------------------------------------------------------------
'
Private Sub refreshThumbnailViewPanel()
 
   On Error GoTo refreshThumbnailViewPanel_Error
   If debugflg = 1 Then DebugPrint "%refreshThumbnailViewPanel"

    Call busyStart
    
    Call populateThumbnails(thumbImageSize, thumbnailStartPosition) '< this one

    picFrameThumbs.Visible = True
    filesIconList.Visible = False
    btnThumbnailView.Visible = False
    btnFileListView.Visible = True

    removeThumbHighlighting

    'highlight the current thumb
    If thumbIndexNo >= 0 Then ' -1 when there are no icons as a result of an empty filter pattern
        If thumbArray(thumbIndexNo) = 0 Or (thumbArray(thumbIndexNo) And thumbArray(thumbIndexNo) <= vScrollThumbs.Max) Then
            If thumbImageSize = 64 Then 'larger
                picFraPicThumbIcon(thumbIndexNo).BorderStyle = 1
                'picThumbIcon(thumbIndexNo).BorderStyle = 1
            ElseIf thumbImageSize = 32 Then
                lblThumbName(thumbIndexNo).BackColor = RGB(10, 36, 106) ' blue
                lblThumbName(thumbIndexNo).ForeColor = RGB(255, 255, 255) ' white
            End If
        End If
    End If

    Call busyStop

   On Error GoTo 0
   Exit Sub

refreshThumbnailViewPanel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure refreshThumbnailViewPanel of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : populateThumbnails
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Create a 12 picbox array that simulates a thumbnail icon view
'
'               The original Rocketdock opened a folder and read every single icon
'               in order to create a cache of icons that could be scrolled up and down quickly.
'               However, this could cause the read icon phase to expand interminably in time when
'               the folder contained hundreds of even thousands of icons.
'               It was not unsusual for sets such as the British Library Set to fill a folder
'               with upward of 1,500 icons. For example, on a core2duo 2.5ghz with an SSD it
'               could take up to 20 seconds to display 160 or so ordinary mixed type icons.
'               In Rocketdock this would occur each and every time the user right clicked upon
'               any icon in the dock making it unusable, not even considering the memory and
'               processing overhead implications. Creating a smaller twelve icon cache each
'               time the user selects the next twelve icons is certainly preferable.
'
'               When I created this I did not know that image lists and imageview controls existed
'               and if I had then I may have used them. However, creating something resembling an
'               imageview in code is arguably better than loading another OCx.
'
'---------------------------------------------------------------------------------------
'
Private Sub populateThumbnails(ByVal imageSize As Integer, ByRef startItem As Integer)

    Dim useloop As Integer: useloop = 0
    Dim fullFilePath As String: fullFilePath = vbNullString
    Dim shortFilename As String: shortFilename = vbNullString
    Dim tooltip As String: tooltip = vbNullString
    Dim suffix As String: suffix = vbNullString
    Dim busyFilename As String: busyFilename = vbNullString
    Dim storeTop As Integer: storeTop = 0
    Dim textLabelWidth As Integer: textLabelWidth = 0
    Dim newString As String: newString = vbNullString
    Dim leftBit As String: leftBit = vbNullString
    Dim rightBit As String: rightBit = vbNullString
    Dim thisThumbnailCacheCount As Integer: thisThumbnailCacheCount = 0
    
    On Error GoTo populateThumbnails_Error
    If debugflg = 1 Then DebugPrint "%" & "populateThumbnails"
   
    ' change the image to a tree view icon
    ' create a matrix of 12 x image objects from file

    ' starting with the startItem from filesIconList
    ' we extract the number
    
    ' There is an outer picbox that acts as the frame, we use a picbox as this border style can be set when the control has focus (64x64 icons only).
    ' Inside the frame is another picbox that stores the image. It can be moved within the frame to cope with different image types
    ' and the methods of rendering, ico files are rendered top left whilst the others are centred. There is a label that sits upon
    ' another real frame that can be styled, it sits still within the outer frame.
        
    ' changing the vScrollThumbs.Maximum causes the vScrollThumbs_changed routine to trigger in Vb.net but not in VB6 so it has the VB.NET equivalent check for a non-zero value for compatibility
    If filesIconList.ListCount - 1 > 0 Then
        vScrollThumbs.Max = filesIconList.ListCount - 1
    End If
    
    ' make the whole lot invisible first
    For useloop = 0 To 11
        picFraPicThumbIcon(useloop).Visible = False
        picThumbIcon(useloop).Visible = False
        fraThumbLabel(useloop).Visible = False
        lblThumbName(useloop).Visible = False
    Next useloop
                
    ' populate each picbox with an image
    For useloop = 0 To 11

        'startItem = filesIconList.ListIndex ' the starting point in the file list for the thumnbnails to start
        'when there are less than a screenful of items the.ListIndex returns -1
        
        If startItem = -1 Then
            startItem = 0
        End If

        ' aside -> .NET collection can't handle going up to or beyond the count, VB6 control array copes
        ' but the count check is here for compatibility with the .NET version.
        
        If useloop + startItem < filesIconList.ListCount Then
            ' take the fileame from the underlying filelist control
            shortFilename = filesIconList.List(useloop + startItem)
            fullFilePath = textCurrentFolder.Text & "\" & shortFilename ' changed from filelistbox.path to different path source for Vb.NET compatibility
            ' if any file does not exist
            If FExists(fullFilePath) = False Then
                If FExists(App.Path & "\resources\" & "blank.jpg") Then fullFilePath = App.Path & "\resources\" & "blank.jpg"
                picThumbIcon(useloop).ToolTipText = vbNullString
                lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
                picThumbIcon(useloop).Picture = LoadPicture(fullFilePath)
                
                ' display the image within the specified picturebox
                Call displayResizedImage(fullFilePath, picThumbIcon(useloop), imageSize)
            End If
            
            If filesIconList.List(useloop + startItem) <> vbNullString Then ' check it has a filename in the list
            
                picThumbIcon(useloop).Width = 1000  ' these set here specifically so I can play
                picThumbIcon(useloop).Height = 1000
                picFraPicThumbIcon(useloop).Width = 1000  ' these set here specifically so I can play
                picFraPicThumbIcon(useloop).Height = 1000
                fraThumbLabel(useloop).BorderStyle = 0
                fraThumbLabel(useloop).BackColor = vbWhite
                lblThumbName(useloop).BackColor = vbWhite
                lblThumbName(useloop).Alignment = 0 ' centre align the label
                lblThumbName(useloop).WordWrap = True
                lblThumbName(useloop).BorderStyle = 0
                lblThumbName(useloop).AutoSize = False
                
                ' .53 DAEB 25/04/2022 rDIConConfig.frm Reformatted the code to place ico files within an outer picbox frame to stop them jumping whilst being redrawn. STARTS
                If shortFilename <> vbNullString Then
                    suffix = ExtractSuffix(LCase$(shortFilename))
                    picThumbIcon(useloop).Left = 0
                    
                    ' Code to deal with ICO positioning.
                    ' If the thumbnail is an ico, nudge the picbox over a bit as it displays ICOs from the top left.
                    ' The picbox sits within another picbox which acts as a frame that itself can be styled with an outline border.
                    
                    ' This is all dealt with in the VB.NET version by padding the image it just after displaying it, so easy.
                    
                    ' first the top left rendering ICOs.
                    If suffix = "ico" Then
                        If thumbImageSize = 32 Then
                            If useloop = 0 Or useloop = 4 Or useloop = 8 Then storeLeft = 165 ' reset left on each row
                            If useloop >= 0 And useloop <= 3 Then storeTop = 100
                            If useloop >= 4 And useloop <= 7 Then storeTop = 1150
                            If useloop >= 8 And useloop <= 11 Then storeTop = 2270
                            
                            fraThumbLabel(useloop).Left = storeLeft - 100
                            fraThumbLabel(useloop).Top = storeTop + 550
                            
                            fraThumbLabel(useloop).Visible = True
                            lblThumbName(useloop).Visible = True
                            picThumbIcon(useloop).Left = 300
                            picThumbIcon(useloop).Top = 0
                        Else
                            fraThumbLabel(useloop).Visible = False
                            lblThumbName(useloop).Visible = False
                            If useloop = 0 Or useloop = 4 Or useloop = 8 Then storeLeft = 295 ' reset left
                            If useloop >= 0 And useloop <= 3 Then storeTop = 20
                            If useloop >= 4 And useloop <= 7 Then storeTop = 1080
                            If useloop >= 8 And useloop <= 11 Then storeTop = 2100
                            picThumbIcon(useloop).Left = 0
                            picThumbIcon(useloop).Top = 0
                        End If
                        picFraPicThumbIcon(useloop).Left = storeLeft - 60
                        picFraPicThumbIcon(useloop).Top = storeTop
                    Else
                        If thumbImageSize = 32 Then
                            If useloop = 0 Or useloop = 4 Or useloop = 8 Then storeLeft = 165 ' reset left on each row
                            If useloop >= 0 And useloop <= 3 Then storeTop = -200
                            If useloop >= 4 And useloop <= 7 Then storeTop = 880
                            If useloop >= 8 And useloop <= 11 Then storeTop = 1970
                            
                            fraThumbLabel(useloop).Left = storeLeft - 100
                            fraThumbLabel(useloop).Top = storeTop + 830
                            
                            fraThumbLabel(useloop).Visible = True
                            lblThumbName(useloop).Visible = True
                        Else
                            fraThumbLabel(useloop).Visible = False
                            lblThumbName(useloop).Visible = False
                            
                            If useloop = 0 Or useloop = 4 Or useloop = 8 Then storeLeft = 295 ' reset left
                            If useloop >= 0 And useloop <= 3 Then storeTop = 20
                            If useloop >= 4 And useloop <= 7 Then storeTop = 1080
                            If useloop >= 8 And useloop <= 11 Then storeTop = 2100
                        End If
                        picFraPicThumbIcon(useloop).Left = storeLeft - 60
                        picFraPicThumbIcon(useloop).Top = storeTop
                    End If
                    
                    picThumbIcon(useloop).ZOrder ' give it topmost
                    fraThumbLabel(useloop).ZOrder
                    
                    picThumbIcon(useloop).AutoRedraw = True
                End If
                
                ' set the tooltip to the short filename
                picThumbIcon(useloop).ToolTipText = filesIconList.List(useloop + startItem)
        
                ' synch. the label to the tooltiptext
                lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
                                
                ' lower case the filename
                lblThumbName(useloop).Caption = LCase$(lblThumbName(useloop).Caption)
                
                ' swap spaces in the filename with underscores. This prevents the weird VB6 text boxes from wrapping on a space unexpectedly.
                lblThumbName(useloop).Caption = Replace(lblThumbName(useloop).Caption, " ", "_")
                                               
                textLabelWidth = TextWidth(lblThumbName(useloop).Caption)
                If textLabelWidth > 1000 Then
                    lblThumbName(useloop).Caption = Left$(lblThumbName(useloop).Caption, 10) & "..."
                End If
                
                ' .84 DAEB 05/06/2022 rDIConConfig.frm Additional use of an imagelist control as a cache of already-read thumbnail icons to speed access to preceding thumbnails STARTS
                
                ' Cache the current image to a unique key that corresponds to the item index in the filesIconList
                ' this is a primitive post-fetch cache, ie. it speeds up any access after each first image read.
                ' when the cache is filled it adds no more and does not clear up the old items, it just stops populating.
                
                If fullFilePath <> "" Then
                    ' .91 DAEB 25/06/2022 rDIConConfig.frm Deleting an icon from the icon thumbnail display causes a cache imageList error. Added cacheingFlg.
                    If imlThumbnailCache.ListImages.Exists("cache" & textCurrentFolder.Text & useloop + startItem) And cacheingFlg = True Then
                        
                        ' display the image within the specified picturebox but extract from the cache
                        picThumbIcon(useloop).Picture = imlThumbnailCache.ListImages("cache" & textCurrentFolder.Text & useloop + startItem).ExtractIcon
                    Else
                        ' display the image (from file) within the specified picturebox
                        Call displayResizedImage(fullFilePath, picThumbIcon(useloop), imageSize)
                                                                        
                        If cacheingFlg = True Then
                            thisThumbnailCacheCount = Val(sdThumbnailCacheCount)
                            If thisThumbnailCacheCount = 0 Then thisThumbnailCacheCount = 250 ' default value
                            
                            ' limit the cache to certain number of image items to prevent out of memory messages
                            If imlThumbnailCache.ListImages.count <= thisThumbnailCacheCount Then
                                
                                ' add the current thumbnail to the cache with a unique key
                                Set picTemporaryStore.Picture = picThumbIcon(useloop).Image
                                imlThumbnailCache.ListImages.Add , "cache" & textCurrentFolder.Text & useloop + startItem, picTemporaryStore.Picture
                                Set picTemporaryStore.Picture = Nothing
                            End If
                        End If
                    End If
                
                End If
                
                ' .84 DAEB 05/06/2022 rDIConConfig.frm Additional use of an imagelist control as a cache of already-read thumbnail icons to speed access to preceding thumbnails ENDS

                              
                'now the picture and captions are populated, make them visible
                picFraPicThumbIcon(useloop).Visible = True
                picThumbIcon(useloop).Visible = True
                
            End If
            thumbArray(useloop) = useloop + startItem
            lblThumbName(useloop).ZOrder
            
            storeLeft = storeLeft + 1200 ' move the next icon to the right a set amount
            
            ' .53 DAEB 25/04/2022 rDIConConfig.frm Reformatted the code to place ico files within an outer picbox frame to stop them jumping whilst being redrawn. ENDS
        Else
            
            picThumbIcon(useloop).ToolTipText = vbNullString
            lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
            If FExists(App.Path & "\resources\" & "blank.jpg") Then picThumbIcon(useloop).Picture = LoadPicture(App.Path & "\resources\" & "blank.jpg")
            
        End If
  
        ' do the hourglass timer display, moved from the non-functional timer to here where it works
        If displayHourglass = True Then

            picBusy.Visible = True
            busyCounter = busyCounter + 1
            If busyCounter >= 7 Then busyCounter = 1
            If classicTheme = True Then
                busyFilename = App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg"
            Else
                busyFilename = App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg"
            End If
            picBusy.Picture = LoadPicture(busyFilename)
            
            ' attempted to load using LaVolpe's method of cImage.LoadPictureFile but to no avail
        End If
    Next useloop
    
    ' .60 DAEB 01/05/2022 rDIConConfig.frm Add each image from the thumbnail icon picbox array to an imageList control so that we can assign a dragIcon
'    For useloop = 0 To 11
'        imlDragIconConverter.ListImages.Clear
'
'        imlDragIconConverter.ListImages.Add , , picTemporaryStore.Image ' adds the icon to key position 1
'
'        imlDragIconConverter.ListImages.Add , , picThumbIcon(useloop).Image ' adds the icon to key position 1
'        picThumbIcon(useloop).DragIcon = imlDragIconConverter.ListImages(1).ExtractIcon
'    Next useloop
    
    picBusy.Visible = False
    
    On Error GoTo 0
   Exit Sub

populateThumbnails_Error:

    MsgBox "fullFilePath = " & fullFilePath & " cache Key = " & "cache" & textCurrentFolder.Text & useloop + startItem

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateThumbnails of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : populateRdMap
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : redraws the whole rdMap
'---------------------------------------------------------------------------------------
'
Private Sub populateRdMap(ByVal xDeviation As Integer)

    Dim useloop As Integer: useloop = 0
    Dim busyFilename As String: busyFilename = vbNullString
    Dim dotString As String: dotString = vbNullString
   
    On Error GoTo populateRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "populateRdMap"

    dotString = vbNullString
    dotCount = 0

    'Refresh ' display the results prior to the for loop
    ' the above command allows the working button to show as it should
                
    rDIconConfigForm.Refresh
    
    ' populate each with an image
    For useloop = 0 To rdIconMaximum
        
        picRdMap(useloop).BorderStyle = 1 ' put a border around the picboxes to show an update
        
        ' using the deviation from the extracted start
        ' visit the filelist at that point and extract the filename
        '  and extract the file path
        
        ' the target picture control and the icon size
        Call displayIconElement(useloop + xDeviation, picRdMap(useloop), True, 32, False, False, False)
        picRdMap(useloop).BorderStyle = 0
        
        'do the 'working...' text on the button
        dotCount = dotCount + 1
        If dotCount = 5 Then
            'Refresh
            dotCount = 0
            dotString = dotString & "."
            btnWorking.Caption = "Working " & dotString
            btnWorking.Refresh
            If dotString = "..." Then dotString = vbNullString
        End If
    
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        picBusy.Visible = True
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        If classicTheme = True Then
            busyFilename = App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        picBusy.Visible = False
 
    Next useloop
    
    rDIconConfigForm.Refresh

   On Error GoTo 0
   Exit Sub

populateRdMap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateRdMap of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : deleteRdMap
' Author    : beededea
' Date      : 14/12/2019
' Purpose   : unused
'---------------------------------------------------------------------------------------
'
Public Sub deleteRdMap(Optional ByVal backupFirst As Boolean = False, Optional ByVal refreshDisplay As Boolean)

    Dim useloop As Integer: useloop = 0
    Dim busyFilename As String: busyFilename = vbNullString
    Dim dotString As String: dotString = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    
   On Error GoTo deleteRdMap_Error
   If debugflg = 1 Then DebugPrint "%deleteRdMap"

    'If picRdMapGotFocus <> True Then Exit Sub
    answer = msgBoxA(" This will delete all the icons in your dock , are you sure?", vbQuestion + vbYesNo)
    If answer = vbNo Then
        Exit Sub
    End If

    If backupFirst = True Then Call fbackupSettings

    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    For useloop = 1 To rdIconMaximum
            
        removeSettingsIni (useloop)
        'clear the icon
        picRdMap(useloop).BackColor = &H8000000F
        Set picRdMap(useloop).Picture = LoadPicture(vbNullString)
        Unload picRdMap(useloop)
        
                'do the 'working...' text on the button
        dotCount = dotCount + 1
        If dotCount = 5 Then
            'Refresh
            dotCount = 0
            dotString = dotString & "."
            btnWorking.Caption = "Working " & dotString
            If dotString = "..." Then dotString = vbNullString
        End If
    
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        picBusy.Visible = True
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        If classicTheme = True Then
            busyFilename = App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        picBusy.Visible = False
            
    Next useloop
        
    'decrement the icon count and the maximum icon
    theCount = 1
    rdIconMaximum = 0
    
    'amend the count
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, interimSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    'set the slider bar
    rdMapHScroll.Max = 1

    rdIconNumber = 0
    
    If refreshDisplay = True Then
        ' load the new icon as an image in the picturebox
        Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), True, 32, True, False)
        
        Call populateRdMap(0) ' regenerate the map from position zero
    End If

   On Error GoTo 0
   Exit Sub

deleteRdMap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteRdMap of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnMapNext_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   : Scroll the RD map to the left
'---------------------------------------------------------------------------------------
'
Private Sub btnMapNext_Click()
    
   On Error GoTo btnMapNext_Click_Error
   If debugflg = 1 Then DebugPrint "%btnMapNext_Click"

    Call picRdMapSetFocus
    
    If rdMapHScroll.Value < rdMapHScroll.Max Then
        rdMapHScroll.Value = rdMapHScroll.Value + 1
    End If

   On Error GoTo 0
   Exit Sub

btnMapNext_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnMapNext_Click of Form rDIconConfigForm"
     
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnMapPrev_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Clicking on the rd Map button to scroll the map
'---------------------------------------------------------------------------------------
'
Private Sub btnMapPrev_Click()

    On Error GoTo btnMapPrev_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnMapPrev_Click"

    Call picRdMapSetFocus
    
    If rdMapHScroll.Value >= 1 Then
        rdMapHScroll.Value = rdMapHScroll.Value - 1
    End If

   On Error GoTo 0
   Exit Sub

btnMapPrev_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnMapPrev_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Clicking on the window that holds the icon files listing
'---------------------------------------------------------------------------------------
' the click event must be retained even though it appears to be useless, it is called automatically when the filesListBox control listindex is changed
Private Sub filesIconList_Click()
    
    '.64 DAEB 04/05/2022 rDIConConfig.frm Moved the fileList left click functionality to a separate routine
    Call filesIconListLeftMouseDown_event

   On Error GoTo 0
   Exit Sub

filesIconList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_Click of Form rDIconConfigForm"

End Sub


'.64 DAEB 04/05/2022 rDIConConfig.frm Moved the fileList left click functionality to a separate routine
'---------------------------------------------------------------------------------------
' Procedure : filesIconListLeftMouseDown_event
' DateTime  : 04/05/2022 16:03
' Author    : beededea
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub filesIconListLeftMouseDown_event()

    Dim Filename As String: Filename = vbNullString
    Dim picSize As Long: picSize = 0
    Dim suffix As String: suffix = vbNullString
    
    On Error GoTo filesIconListLeftMouseDown_event_Error
    thisRoutine = "filesIconListLeftMouseDown_event"

    picPreview.AutoRedraw = True
    picPreview.AutoSize = False
    
    Filename = textCurrentFolder.Text ' textCurrentFolder.Text ' changed from .path to use alternate path source to be compatible with VB.NET
    If Right$(Filename, 1) <> "\" Then
        Filename = Filename & "\"
    End If
    Filename = Filename & filesIconList.Filename
    
    If filesIconList.Filename = "" Then
        Exit Sub
    End If
    
    suffix = ExtractSuffix(Filename)
    picSize = FileLen(Filename)
    lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
    
    'If picFrameThumbsGotFocus = True Then
         
    'refresh the preview displaying the selected image
    Call displayResizedImage(Filename, picPreview, icoSizePreset)
    Call displayResizedImage(Filename, picTemporaryStore, 64) ' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon.
    
    filesIconListGotFocus = True
    
    picPreview.ToolTipText = Filename
    picPreview.Tag = Filename
    
    storedIndex = thumbIndexNo
    'mapImageChanged = True
    
    ' .picture is the graphic itself
    ' .image property is a bitmap handle to the actual rendered "canvas" of the (resized) container
    
    ' .80 DAEB 28/05/2022 rDIConConfig.frm Change to adding the .picture to workaround the bug in Krool's imageList failing to convert to an HIcon.
    ' .81 DAEB 28/05/2022 rDIConConfig.frm Added check to visibility of control before running the dragIcon code.
    If filesIconList.Visible = True Then
        imlDragIconConverter.ListImages.Clear ' clear the CCR imageList
        Set picTemporaryStore.Picture = picTemporaryStore.Image 'convert the original bitmap handle into a graphic that the CCR imageList can handle
        imlDragIconConverter.ListImages.Add , "arse", picTemporaryStore.Picture ' add the picture to CCR imageList, adds the icon to key position 1
        Set picTemporaryStore.Picture = Nothing ' clear the temporary picBox
        filesIconList.DragIcon = imlDragIconConverter.ListImages("arse").ExtractIcon ' extract the icon and assign to the dragICon on the simple file listing
    End If
        
    ' we signify that no changes have been made
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False


    On Error GoTo 0
    Exit Sub

filesIconListLeftMouseDown_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconListLeftMouseDown_event of Form rDIconConfigForm"
End Sub






' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font STARTS
'---------------------------------------------------------------------------------------
' Procedure : mnuFont_Click
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFont_Click()
    Dim storedFont As String: storedFont = vbNullString
    
    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False

    'storedFont = txtTextFont.Text 'TBD
    
    fntFont = SDSuppliedFont
    fntSize = SDSuppliedFontSize
    fntItalics = CBool(SDSuppliedFontItalics)
    fntColour = CLng(SDSuppliedFontColour)
    
    Call changeFont(Me, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    If fntFont = vbNullString Then
        fntFont = storedFont
        fntSize = "8"
    End If
    
    If fntSize = "0" Then
        fntSize = "8"
    End If
    
    SDSuppliedFont = CStr(fntFont)
    SDSuppliedFontSize = CStr(fntSize)
    SDSuppliedFontItalics = CStr(fntItalics)
    SDSuppliedFontColour = CStr(fntColour)

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\SteamyDockSettings", "defaultFont", SDSuppliedFont, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultSize", SDSuppliedFontSize, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultItalics", SDSuppliedFontItalics, toolSettingsFile
        PutINISetting "Software\SteamyDockSettings", "defaultColour", SDSuppliedFontColour, toolSettingsFile
    End If
    
    
' - o -

'    Dim suppliedFont As String
'    Dim suppliedSize As Integer
'    Dim suppliedStrength As Boolean
'    Dim suppliedStyle As Boolean
'
'    On Error GoTo mnuFont_Click_Error
'    If debugflg = 1 Then DebugPrint "%" & "mnuFont_Click"
'
'    Set FontDlg = New CommonDlgs
'
'    With FontDlg
'            .DialogTitle = "Select a Font"
'            .flags = cdlCFScreenFonts _
'                  Or cdlCFBoth _
'                  Or cdlCFEffects _
'                  Or cdlCFApply _
'                  Or cdlCFForceFontExist
'    End With
'
'    On Error Resume Next
'    If FontDlg.ShowFont(hwnd, hdc) Then
'        SDSuppliedFont = FontDlg.FontName
'        SDSuppliedFontSize = FontDlg.FontSize
'        SDSuppliedFontStrength = FontDlg.FontBold
'        SDSuppliedFontStyle = FontDlg.FontItalic
'    End If
'
'l_err1:
''    If dlgFontForm.dlgFont.FontName = vbNullString Then
''        Exit Sub
''    End If
'
'    If Err <> 32755 Then    ' User didn't chose Cancel.
'        'suppliedFont = dlgFontForm.dlgFont.FontName
'    End If
    
'    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
'        PutINISetting "Software\SteamyDockSettings", "defaultFont", SDSuppliedFont, toolSettingsFile
'        PutINISetting "Software\SteamyDockSettings", "defaultSize", SDSuppliedFontSize, toolSettingsFile
'SDSuppliedFontItalics
'        PutINISetting "Software\SteamyDockSettings", "defaultStrength", SDSuppliedFontStrength, toolSettingsFile
'        PutINISetting "Software\SteamyDockSettings", "defaultStyle", SDSuppliedFontStyle, toolSettingsFile
'    End If

''    If suppliedFont <> vbNullString Then
''        Call changeFont(FontDlg.FontName, FontDlg.FontSize, FontDlg.FontBold, FontDlg.FontItalic)
''    End If

   On Error GoTo 0
   Exit Sub

mnuFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFont_Click of Form rDIconConfigForm"
    
End Sub
' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_DblClick
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Double-Clicking on the window that holds the icon files listing to select a specific icon
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_DblClick()
    
    ' takes the result from the treeview
    On Error GoTo filesIconList_DblClick_Error
    If debugflg = 1 Then DebugPrint "%" & "filesIconList_DblClick"
    
    ' note the old icon image just in case the user decides to not save
    previousIcon = txtCurrentIcon.Text
    
    ' change the text in the icon field
    txtCurrentIcon.Text = relativePath & "\" & filesIconList.Filename
    
    ' now change the icon image
    ' the target picture control and the icon size
    Call displayResizedImage(txtCurrentIcon.Text, picRdMap(rdIconNumber), 32)
        
    ' we signify that no changes have been made
    btnSet.Enabled = True ' this has to be done at the end
    '
    btnCancel.Visible = True
    btnClose.Visible = False
    
    iconChanged = True

   On Error GoTo 0
   Exit Sub

filesIconList_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_DblClick of Form rDIconConfigForm"
  
End Sub

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_GotFocus
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Sets some variables to determine what currently has focus
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_GotFocus()
   On Error GoTo filesIconList_GotFocus_Error
    If debugflg = 1 Then DebugPrint "%" & "filesIconList_GotFocus"

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False
   On Error GoTo 0
   Exit Sub

filesIconList_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each control that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo filesIconList_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "filesIconList_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
        Exit Sub
    End If
    
    ' the click event must be retained even though it appears to be useless, it is called automatically when the control listindex is changed
    ' this reference is left here as a reminder.
    'Call filesIconListLeftMouseDown_event

    ' .67 DAEB 04/05/2022 rDIConConfig.frm Drag and drop from the filelist to the rdmap
    filesIconList.Drag vbBeginDrag


    'imlDragIconConverter.ListImages.Remove ("arse")
    
   On Error GoTo 0
   Exit Sub

filesIconList_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_MouseDown of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : readTreeviewDefaultFolder
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the default folder, the folder the user previously selected
'---------------------------------------------------------------------------------------
'
Private Sub readTreeviewDefaultFolder()
    ' extract the default folder in the treeview, this is an enhancement over the original tool
    ' it stores the last used default folder as shown in the tree view top left
    
    Dim iX As Integer: iX = 0
    Dim iFound As Boolean: iFound = False
    Dim defaultFolderNodeKey As String: defaultFolderNodeKey = vbNullString
    
    ' this is Krool's treeview replacement of the Treeview node
    'Dim Node As CCRTreeView.TvwNode
        
    ' read the tool settings file
    'eg. defaultFolderNodeKey=?E:\dean\steampunk theme\icons\
    On Error GoTo readTreeviewDefaultFolder_Error
    If debugflg = 1 Then DebugPrint "%" & "readTreeviewDefaultFolder"

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        defaultFolderNodeKey = GetINISetting("Software\SteamyDockSettings", "defaultFolderNodeKey", toolSettingsFile)
    End If

    folderTreeView.HideSelection = False ' Ensures found item highlighted

    If defaultFolderNodeKey <> vbNullString Then
        For iX = 1 To folderTreeView.Nodes.count
            If Trim$(folderTreeView.Nodes(iX).Key) = Trim$(defaultFolderNodeKey) Then
                iFound = True
                Exit For
            End If
        Next
        If iFound Then
            ' highlight the treeview item
            folderTreeView.Nodes(iX).EnsureVisible
            folderTreeView.SelectedItem = folderTreeView.Nodes(iX)
            folderTreeView.Nodes(iX).Selected = True
            ' expand the current node
            folderTreeView.SelectedItem.Expanded = True
            
            ' the above line makes the following line obsolete
            ' click on the selected item to expand as it does not automatically trigger a thumbnail refresh on startup
            'folderTreeView_Click
        Else
            'MsgBox ("String not found")
        End If
    End If

   On Error GoTo 0
   Exit Sub

readTreeviewDefaultFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readTreeviewDefaultFolder of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : readRegistryWriteSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the registry one line at a time and create a temporary settings file
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryWriteSettings()
    Dim useloop As Integer: useloop = 0
    
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
   On Error GoTo readRegistryWriteSettings_Error
      If debugflg = 1 Then DebugPrint "%" & "readRegistryWriteSettings"
   
    For useloop = 0 To rdIconMaximum
         ' get the relevant entries from the registry
         readRegistryIconValues (useloop)
         Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile)
         
     Next useloop

   On Error GoTo 0
   Exit Sub

readRegistryWriteSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryWriteSettings of Form rDIconConfigForm"
End Sub


''---------------------------------------------------------------------------------------
'' Procedure : driveCheck
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : check for the existence of the rocketdock binary
''---------------------------------------------------------------------------------------
''
'Public Function driveCheck(folder As String, filename As String)
'
'
'   Dim sAllDrives As String
'   Dim sDrv As String
'   Dim sDrives() As String
'   Dim cnt As Long
'   Dim folderString As String
'   Dim testAppPath As String
'
'  'get the list of all drives
'   On Error GoTo driveCheck_Error
'
'   sAllDrives = GetDriveString()
'
'  'Change nulls to spaces, then trim.
'  'This is required as using Split()
'  'with Chr$(0) alone adds two additional
'  'entries to the array drives at the end
'  'representing the terminating characters.
'   sAllDrives = Replace$(sAllDrives, Chr$(0), Chr$(32))
'   sDrives() = Split(Trim$(sAllDrives), Chr$(32))
'
'    For cnt = LBound(sDrives) To UBound(sDrives)
'        sDrv = sDrives(cnt)
'        ' on 32bit windows the folder is "Program Files\Rocketdock"
'        folderString = sDrv & folder
'        If DirExists(folderString) = True Then
'           'test for the yahoo widgets binary
'            testAppPath = folderString
'            If FExists(testAppPath & "\" & filename) Then
'                'MsgBox "YWE folder exists"
'                driveCheck = testAppPath
'                Exit Function
'            End If
'        End If
'    Next
'
'   On Error GoTo 0
'   Exit Function
'
'driveCheck_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure driveCheck of Form rDIconConfigForm"
'
'End Function


''---------------------------------------------------------------------------------------
'' Procedure : GetDriveString
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Determine the number and name of drives using VB alone
''---------------------------------------------------------------------------------------
''
'Private Function GetDriveString() As String
'
'    'Used by both demos
'
'    ' returns string of available
'    ' drives each separated by a null
'    ' Dim sBuff As String
'    '
'    ' possible 26 drives, three characters
'    ' each plus a trailing null for each
'    ' drive letter and a terminating null
'    ' for the string
'
'    Dim I As Long
'    Dim builtString As String
'
'    '===========================
'    'pure VB approach, no controls required
'    'drive letters are found in positions 1-UBound(Letters)
'    '"C:\ D:\ E:\ &C:"
'
'    On Error GoTo GetDriveString_Error
'       If debugflg = 1 Then DebugPrint "%" & "GetDriveString"
'
'
'
'    For I = 1 To 26
'        If ValidDrive(Chr(96 + I)) = True Then
'            builtString = builtString + uCase$(Chr(96 + I)) & ":\    "
'        End If
'    Next I
'
'    GetDriveString = builtString
'
'   On Error GoTo 0
'   Exit Function
'
'GetDriveString_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriveString of Form rDIconConfigForm"
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : ValidDrive
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Check if the drive found is a valid one
''---------------------------------------------------------------------------------------
''
'Public Function ValidDrive(ByVal d As String) As Boolean
'   On Error GoTo ValidDrive_Error
'      If debugflg = 1 Then DebugPrint "%" & "ValidDrive"
'
'
'
'  On Error GoTo driveerror
'  Dim Temp As String
'
'    Temp = CurDir
'    ChDrive d
'
'    ChDir Temp
'    ValidDrive = True
'
'  Exit Function
'driveerror:
'
'   On Error GoTo 0
'   Exit Function
'
'ValidDrive_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidDrive of Form rDIconConfigForm"
'End Function

'---------------------------------------------------------------------------------------
' Procedure : addRocketdockFolders
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Add the rocketdock icon folders that exist below the RD icons folder
'---------------------------------------------------------------------------------------
'
Private Sub addRocketdockFolders()
    Dim pathCheck As String: pathCheck = vbNullString
    
    On Error GoTo addRocketdockFolders_Error
    If debugflg = 1 Then DebugPrint "%" & "addRocketdockFolders"

    If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
        pathCheck = rdAppPath & "\icons"
    Else
        pathCheck = sdAppPath & "\icons"
    End If
        
    If Not pathCheck = vbNullString Then
        ' add the chosen folder to the treeview
        folderTreeView.Nodes.Add , , pathCheck, pathCheck

        Call addtotree(pathCheck, folderTreeView)
        folderTreeView.Nodes(pathCheck).Text = "icons"
    End If

   On Error GoTo 0
   Exit Sub

addRocketdockFolders_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addRocketdockFolders of Form rDIconConfigForm"
    
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : setSteampunkLocation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : add the extra steampunk icon folders to the treeview
'---------------------------------------------------------------------------------------
'
Private Sub setSteampunkLocation()

    ' add the custom folder to the treeview
    Dim SteampunkIconFolder As String: SteampunkIconFolder = vbNullString
    
   On Error GoTo setSteampunkLocation_Error
   If debugflg = 1 Then DebugPrint "%" & "setSteampunkLocation"


    SteampunkIconFolder = App.Path & "\my collection"
    
    If DirExists(SteampunkIconFolder) Then
        ' add the chosen folder to the treeview
        folderTreeView.Nodes.Add , , SteampunkIconFolder, SteampunkIconFolder
        Call addtotree(SteampunkIconFolder, folderTreeView)
        folderTreeView.Nodes(SteampunkIconFolder).Text = "my collection"
    End If

   On Error GoTo 0
   Exit Sub

setSteampunkLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setSteampunkLocation of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : readCustomLocation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the location of the user's custom folder
'---------------------------------------------------------------------------------------
'
Private Sub readCustomLocation()

    ' add the custom folder to the treeview
    'Dim rDCustomIconFolder As String
    
    ' read the settings ini file
    'eg. rDCustomIconFolder=?E:\dean\steampunk theme\icons\
    On Error GoTo readCustomLocation_Error
    If debugflg = 1 Then DebugPrint "%" & "rDCustomIconFolder"

    If FExists(interimSettingsFile) Then
        rDCustomIconFolder = GetINISetting("Software\RocketDock", "rDCustomIconFolder", interimSettingsFile)
    End If
    
    If Not rDCustomIconFolder = vbNullString Then
        rDCustomIconFolder = Mid$(rDCustomIconFolder, 2) ' remove the question mark
        If DirExists(rDCustomIconFolder) Then
            ' add the chosen folder to the treeview
            folderTreeView.Nodes.Add , , rDCustomIconFolder, rDCustomIconFolder
            Call addtotree(rDCustomIconFolder, folderTreeView)
            folderTreeView.Nodes(rDCustomIconFolder).Text = "custom folder"
        Else ' .NET
            'MsgBox ("Error reading the previously specified custom location " & rDCustomIconFolder)
        End If
    End If

   On Error GoTo 0
   Exit Sub

readCustomLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readCustomLocation of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo Form_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%" & "Form_MouseDown"
   
    If moreConfigVisible = True Then Call picMoreConfigDown_Click ' .nn cause the new expanding section to close
   
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : frameButtons_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub frameButtons_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo frameButtons_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%" & "frameButtons_MouseDown"
   
   'If moreConfigVisible = True Then Call picMoreConfigDown_Click ' .nn cause the new expanding section to close
   If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

frameButtons_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure frameButtons_MouseDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FrameFolders_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub FrameFolders_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo FrameFolders_MouseDown_Error
      If debugflg = 1 Then DebugPrint "%" & "FrameFolders_MouseDown"

    If Button = 2 Then
    ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
    
   On Error GoTo 0
   Exit Sub

FrameFolders_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FrameFolders_MouseDown of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : frameIcons_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub frameIcons_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo frameIcons_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "frameIcons_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

frameIcons_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure frameIcons_MouseDown of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : framePreview_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub framePreview_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo framePreview_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "framePreview_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

framePreview_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure framePreview_MouseDown of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : fraProperties_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub fraProperties_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo fraProperties_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "fraProperties_MouseDown"
   
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If


   On Error GoTo 0
   Exit Sub

fraProperties_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fraProperties_MouseDown of Form rDIconConfigForm"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : picFrameThumbs_LostFocus
' Author    : beededea
' Date      : 16/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFrameThumbs_LostFocus()
   On Error GoTo picFrameThumbs_LostFocus_Error
   'If debugflg = 1 Then DebugPrint "%picFrameThumbs_LostFocus"
   
   'MsgBox Me.ActiveControl.Name
   

   Call thumbsLostFocus

   On Error GoTo 0
   Exit Sub

picFrameThumbs_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbs_LostFocus of Form rDIconConfigForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : picMoreConfigDown_Click
' Author    : beededea
' Date      : 20/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picMoreConfigDown_Click()
    Dim amountToDrop As Integer: amountToDrop = 0
    On Error GoTo picMoreConfigDown_Click_Error

        picMoreConfigDown.Visible = False
        picMoreConfigUp.Visible = True
        picHideConfig.Visible = True
        amountToDrop = 1200
        fraProperties.Height = 3630 + amountToDrop
        
        moreConfigVisible = True
        
        ' .43 DAEB 16/04/2022 rdIconConfig.frm increase the whole form height and move the bootom buttons set down
        rDIconConfigForm.Height = rDIconConfigForm.Height + amountToDrop
        frameButtons.Top = frameButtons.Top + amountToDrop
        
        framePreview.Height = framePreview.Height + amountToDrop
        fraSizeSlider.Top = fraSizeSlider.Top + amountToDrop
        btnPrev.Height = btnPrev.Height + amountToDrop
        btnNext.Height = btnNext.Height + amountToDrop
        
        
        If chkToggleDialogs.Value = 0 Then picMoreConfigDown.ToolTipText = "Hides the extra configuration items"
        
'    Else
'        btnSettingsDown.Visible = False
'        btnSettingsUp.Visible = True
'        fraProperties.Height = 3630
'        moreConfigVisible = False
'
'        ' .43 DAEB 16/04/2022 rdIconConfig.frm increase the whole form height and move the bootom buttons set down
'        frameButtons.Top = frameButtons.Top - 645
'        rDIconConfigForm.Height = rDIconConfigForm.Height - 645
'
'        If chkToggleDialogs.Value = 0 Then picMoreConfigDown.ToolTipText = "Shows extra configuration items"
'    End If

    On Error GoTo 0
    Exit Sub

picMoreConfigDown_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picMoreConfigDown_Click of Form rDIconConfigForm"
End Sub


   'this is the preliminary code to allow sliding of the icons by mouse and cursor
   
'    Dim picX As Integer
'    picRdMap(Index).ZOrder 'the drag pic
'    'so it appears over the top of other controls
'    picX = X - 500
    
    'MsgBox "X = " & picRdMap(Index).Left

'---------------------------------------------------------------------------------------
' Procedure : picRdMap_DragDrop
' Author    : beededea
' Date      : 03/02/2021
' Purpose   :
'---------------------------------------------------------------------------------------
' .59 DAEB 01/05/2022 rDIConConfig.frm Added Drag and drop functionality
Private Sub picRdMap_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

   On Error GoTo picRdMap_DragDrop_Error
       
    Dim useloop As Integer: useloop = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim Filename As String: Filename = vbNullString
    
    If debugflg = 1 Then DebugPrint "%" & "picRdMap_DragDrop"
    
    'if dragging from one part of the map to another but landing on the same icon as the one we started from
    If srcDragControl = "rdMap" And rdMapIconSrcIndex = Index Then
        rdMapIconMouseDown = False
        picRdMap(rdIconNumber).Drag vbEndDrag

        Exit Sub
    End If
    
    'remove the current highlighting on the Rocket dock map
    picRdMap(rdIconNumber).BorderStyle = 0
   
    ' The icon is dragged to a position in the dock
    ' Now read and display the characteristics of the icon that is being replaced
    ' but this time do NOT show the icon in preview form, bottom left as we just want the text.
    
    rdIconNumber = Index
    lblRdIconNumber.Caption = Str$(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str$(rdIconNumber) + 1
    
    ' show the target icon details only
    Call displayIconElement(rdIconNumber, picPreview, False, icoSizePreset, True, False) ' < False to showing the image but show the target details

    If Index <= rdIconMaximum Then
        picRdMap(Index).BorderStyle = 1 ' highlight the new icon position
    End If
    
    lastHighlightedRdMapIndex = Index
    
    ' .73 DAEB 16/05/2022 rDIConConfig.frm Add ability to drag and drop icons in the map to an alternate position
    If srcDragControl = "rdMap" Then
        trgtRdIconNumber = rdIconNumber
        Call reOrderRdMap(srcRdIconNumber, trgtRdIconNumber)
        '
    Else ' add an icon from the thumbnail list
        Call btnAdd_Click
    End If
    
    If mapImageChanged = True Then
        ' now determine the new filename
        If textCurrentFolder.Text <> vbNullString Then
            Filename = textCurrentFolder.Text
            If Right$(Filename, 1) <> "\" Then Filename = Filename & "\"
            Filename = Filename & filesIconList.Filename
            ' refresh the image display on the map
            Call displayResizedImage(Filename, picRdMap(rdIconNumber), 32)
        End If
        
        Call btnSet_Click ' automatically press the set button
        mapImageChanged = False
    End If

    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False
    
    dragToDockOperating = False ' flag to show whether a drag to dock operation is underway

   On Error GoTo 0
   Exit Sub

picRdMap_DragDrop_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_DragDrop of Form rDIconConfigForm"
End Sub

' .73 DAEB 16/05/2022 rDIConConfig.frm Add ability to drag and drop icons in the map to an alternate position STARTS
'---------------------------------------------------------------------------------------
' Procedure : reOrderRdMap
' Author    : beededea
' Date      : 22/05/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub reOrderRdMap(ByVal srcRdIconNumber As Integer, ByVal trgtRdIconNumber As Integer)

    ' take the source icon details and store those in the stored vars
    ' read the target icon details and store those in the standards vars
        
    Dim srcFilename  As String: srcFilename = vbNullString
    Dim srcFileName2  As String: srcFileName2 = vbNullString
    Dim srcTitle  As String: srcTitle = vbNullString
    Dim srcCommand  As String: srcCommand = vbNullString
    Dim srcArguments  As String: srcArguments = vbNullString
    Dim srcWorkingDirectory  As String: srcWorkingDirectory = vbNullString
    Dim srcShowCmd  As String: srcShowCmd = vbNullString
    Dim srcOpenRunning  As String: srcOpenRunning = vbNullString
    Dim srcIsSeparator  As String: srcIsSeparator = vbNullString
    Dim srcUseContext  As String: srcUseContext = vbNullString
    Dim srcDockletFile  As String: srcDockletFile = vbNullString
    Dim srcUseDialog  As String: srcUseDialog = vbNullString
    Dim srcUseDialogAfter  As String: srcUseDialogAfter = vbNullString
    Dim srcQuickLaunch  As String: srcQuickLaunch = vbNullString
    Dim srcAutoHideDock  As String: srcAutoHideDock = vbNullString
    Dim srcSecondApp   As String: srcSecondApp = vbNullString
    Dim srcRunSecondAppBeforehand As String: srcRunSecondAppBeforehand = vbNullString
    Dim srcAppToTerminate   As String: srcAppToTerminate = vbNullString
    Dim srcDisabled   As String: srcDisabled = vbNullString
    Dim srcRunElevated  As String: srcRunElevated = vbNullString
    
    
    Dim trgtFilename  As String: trgtFilename = vbNullString
    Dim trgtFileName2  As String: trgtFileName2 = vbNullString
    Dim trgtTitle  As String: trgtTitle = vbNullString
    Dim trgtCommand  As String: trgtCommand = vbNullString
    Dim trgtArguments  As String: trgtArguments = vbNullString
    Dim trgtWorkingDirectory  As String: trgtWorkingDirectory = vbNullString
    Dim trgtShowCmd  As String: trgtShowCmd = vbNullString
    Dim trgtOpenRunning  As String: trgtOpenRunning = vbNullString
    Dim trgtIsSeparator  As String: trgtIsSeparator = vbNullString
    Dim trgtUseContext  As String: trgtUseContext = vbNullString
    Dim trgtDockletFile  As String: trgtDockletFile = vbNullString
    Dim trgtUseDialog  As String: trgtUseDialog = vbNullString
    Dim trgtUseDialogAfter  As String: trgtUseDialogAfter = vbNullString
    Dim trgtQuickLaunch  As String: trgtQuickLaunch = vbNullString
    Dim trgtAutoHideDock  As String: trgtAutoHideDock = vbNullString
    Dim trgtSecondApp   As String: trgtSecondApp = vbNullString
    Dim trgtRunSecondAppBeforehand As String: trgtRunSecondAppBeforehand = vbNullString
    Dim trgtAppToTerminate  As String: trgtAppToTerminate = vbNullString
    Dim trgtDisabled   As String: trgtDisabled = vbNullString
    Dim trgtRunElevated  As String: trgtRunElevated = vbNullString
    
    Dim firstRecord As Integer: firstRecord = 0
    Dim secondRecord As Integer: secondRecord = 0
    
    Dim useloop As Integer: useloop = 0
    Dim thisIcon As Integer: thisIcon = 0
    Dim notQuiteTheTop As Integer: notQuiteTheTop = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    ' ask if you wish to move the icon
    On Error GoTo reOrderRdMap_Error

    If chkToggleDialogs.Value = 1 Then
        'Prompt, Buttons, Title, HelpFile, Optional ByVal Context As Long
        
        answer = msgBoxA(" Move the icon from position " & srcRdIconNumber + 1 & " to position " & trgtRdIconNumber + 1 & " in the dock? ", vbQuestion + vbYesNo, "Confirm dragging and dropping to the dock", True, "reOrderRdMap")
        'answer = MsgBox(" Move the icon from position " & srcRdIconNumber + 1 & " to position " & trgtRdIconNumber + 1 & " in the dock? ", vbExclamation + vbYesNo, "Confirm dragging and dropping to the dock")
        If answer = vbNo Then
            mapImageChanged = False
            Exit Sub
        End If
        Refresh
    End If
        
    picThumbIconMouseDown = False
    picRdMap(rdIconNumber).Drag vbEndDrag

    ' TBD
    Me.MousePointer = 0 ' this changes the drag icon back to a normal pointer much quicker than waiting for the vbEndDrag to do its stuff.
    
    mapImageChanged = True

    notQuiteTheTop = rdIconMaximum - 1
    
    

    ' take the source icon details and store those in the stored vars
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", srcRdIconNumber, interimSettingsFile
        
    srcFilename = sFilename
    srcFileName2 = sFileName2
    srcTitle = sTitle
    srcCommand = sCommand
    srcArguments = sArguments
    srcWorkingDirectory = sWorkingDirectory
    srcShowCmd = sShowCmd
    srcOpenRunning = sOpenRunning
    srcRunElevated = sRunElevated
    srcIsSeparator = sIsSeparator
    srcUseContext = sUseContext
    srcDockletFile = sDockletFile
    
    If defaultDock = 1 Then
        srcUseDialog = sUseDialog
        srcUseDialogAfter = sUseDialogAfter
        srcQuickLaunch = sQuickLaunch
        srcAutoHideDock = sAutoHideDock
        srcSecondApp = sSecondApp
        srcRunSecondAppBeforehand = sRunSecondAppBeforehand
            
        srcAppToTerminate = sAppToTerminate
        srcDisabled = sDisabled
    End If
    
    ' read the target icon details and store those in the trgt vars
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", trgtRdIconNumber, interimSettingsFile

    trgtFilename = sFilename
    trgtFileName2 = sFileName2
    trgtTitle = sTitle
    trgtCommand = sCommand
    trgtArguments = sArguments
    trgtWorkingDirectory = sWorkingDirectory
    trgtShowCmd = sShowCmd
    trgtOpenRunning = sOpenRunning
    trgtRunElevated = sRunElevated
    trgtIsSeparator = sIsSeparator
    trgtUseContext = sUseContext
    trgtDockletFile = sDockletFile

    If defaultDock = 1 Then
        trgtUseDialog = sUseDialog
        trgtUseDialogAfter = sUseDialogAfter
        trgtQuickLaunch = sQuickLaunch
        trgtAutoHideDock = sAutoHideDock
        trgtSecondApp = sSecondApp
        trgtRunSecondAppBeforehand = sRunSecondAppBeforehand
        
        trgtAppToTerminate = sAppToTerminate
        trgtDisabled = sDisabled
    End If
            
    ' we determine which record will be encountered first, the source or the target
    If srcRdIconNumber < trgtRdIconNumber Then
        firstRecord = srcRdIconNumber
        secondRecord = trgtRdIconNumber
    Else
        secondRecord = srcRdIconNumber
        firstRecord = trgtRdIconNumber
    End If
          
    ' if the first record is the source then the copying is from left to right,
    ' the source record location must be filled, overwritten with the next one up...
    If firstRecord = srcRdIconNumber Then
        
'        For useloop = 0 To srcRdIconNumber
'            ' do nothing for the records to the left
'        Next useloop
        
        For useloop = firstRecord To secondRecord
            ' we don't bother to read the current record source here as we have already done so above.
            
            ' read the rdsettings.ini one item up in the list
            Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop + 1, interimSettingsFile)
            
            'write the the next item up at the current source location effectively overwriting the current record
            Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile)
            
            ' .83 DAEB 03/06/2022 rDIConConfig.frm Display the icon we just moved by dragging, one by one rather than the whole map
            Call displayIconElement(useloop, picRdMap(useloop), True, 32, True, False)
        
        Next useloop
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        
        ' write the the source item at the target location effectively overwriting it
        ' reassign the source details to the standard vars
        sFilename = srcFilename
        sFileName2 = srcFileName2
        sTitle = srcTitle
        sCommand = srcCommand
        sArguments = srcArguments
        sWorkingDirectory = srcWorkingDirectory
        sShowCmd = srcShowCmd
        sOpenRunning = srcOpenRunning
        sRunElevated = srcRunElevated
        sIsSeparator = srcIsSeparator
        sUseContext = srcUseContext
        sDockletFile = srcDockletFile
        
        If defaultDock = 1 Then
            sUseDialog = srcUseDialog
            sUseDialogAfter = srcUseDialogAfter
            sQuickLaunch = srcQuickLaunch
            sAutoHideDock = srcAutoHideDock
            sSecondApp = srcSecondApp
            sRunSecondAppBeforehand = srcRunSecondAppBeforehand
            
            sAppToTerminate = srcAppToTerminate
            sDisabled = srcDisabled
            
        End If
        
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", trgtRdIconNumber, interimSettingsFile)
            
        Call displayIconElement(trgtRdIconNumber, picRdMap(trgtRdIconNumber), True, 32, True, False)

'        For useloop = secondRecord + 1 To rdIconMaximum
'            ' do nothing for the records to the right
'        Next useloop
    End If
    
    ' if the first record is the target then copying is occurring from right towards the left
    If firstRecord = trgtRdIconNumber Then
    
'        For useloop = 0 To trgtRdIconNumber - 1
'            ' do nothing
'        Next useloop
                        
        ' move all the records up one space, leaving space for the target record to be inserted
        For useloop = srcRdIconNumber To trgtRdIconNumber Step -1
        
            ' read the rdsettings.ini one item up in the list
            Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop - 1, interimSettingsFile)
            
            
            'write the the next item up at the current source location effectively overwriting it
            Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile)
        
            Call displayIconElement(useloop, picRdMap(useloop), True, 32, True, False)

        Next useloop
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        
        ' write the the source item at the target location effectively overwriting it
        ' reassign the source details to the standard vars
        sFilename = srcFilename
        sFileName2 = srcFileName2
        sTitle = srcTitle
        sCommand = srcCommand
        sArguments = srcArguments
        sWorkingDirectory = srcWorkingDirectory
        sShowCmd = srcShowCmd
        sOpenRunning = srcOpenRunning
        sRunElevated = srcRunElevated
        sIsSeparator = srcIsSeparator
        sUseContext = srcUseContext
        sDockletFile = srcDockletFile
        
        If defaultDock = 1 Then
            sUseDialog = srcUseDialog
            sUseDialogAfter = srcUseDialogAfter
            sQuickLaunch = srcQuickLaunch
            sAutoHideDock = srcAutoHideDock
            sSecondApp = srcSecondApp
            
            sRunSecondAppBeforehand = srcRunSecondAppBeforehand
            
            sAppToTerminate = srcAppToTerminate
            sDisabled = srcDisabled
            
        End If
        
        Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", trgtRdIconNumber, interimSettingsFile)
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        
        ' .83 DAEB 03/06/2022 rDIConConfig.frm Display the icon we just moved by dragging, one by one rather than the whole map
        Call displayIconElement(trgtRdIconNumber, picRdMap(trgtRdIconNumber), True, 32, True, False)

    End If
                    
    thisIcon = trgtRdIconNumber
    
    Call btnSet_Click
    ' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
    Call picRdMap_MouseDown_event(thisIcon)
    
    'set the fields for this icon to the correct value as supplied
    txtLabelName.Text = sTitle
    
    If sDockletFile <> vbNullString Then
        txtTarget.Text = sDockletFile
    Else
        txtTarget.Text = sCommand
    End If
    
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory
    
    cmbRunState.ListIndex = 1 ' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
    cmbOpenRunning.ListIndex = 0 ' "Use Global Setting"
    chkRunElevated.Value = 0
    
     
    btnSet.Enabled = False ' tell the program that nothing has changed
    btnClose.Visible = True
    btnCancel.Visible = False


    On Error GoTo 0
    Exit Sub

reOrderRdMap_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reOrderRdMap of Form rDIconConfigForm"
            Resume Next
          End If
    End With

End Sub
' .73 DAEB 16/05/2022 rDIConConfig.frm Add ability to drag and drop icons in the map to an alternate position ENDS

'---------------------------------------------------------------------------------------
' Procedure : picRdMap_DragOver
' Author    : beededea
' Date      : 03/02/2021
' Purpose   :
'---------------------------------------------------------------------------------------
' .61 DAEB 01/05/2022 rDIConConfig.frm Added highlighting to the rdIconMap during Drag and drop.
Private Sub picRdMap_DragOver(ByRef Index As Integer, ByRef Source As Control, ByRef X As Single, ByRef Y As Single, ByRef State As Integer)

   On Error GoTo picRdMap_DragOver_Error
   
    ' if the rdMap already has a highlighted icon, then clear the highlight
    If lastHighlightedRdMapIndex >= 0 Then picRdMap(lastHighlightedRdMapIndex).BorderStyle = 0

    'highlight the map entry hovered over
    picRdMap(Index).BorderStyle = 1
    
    lastHighlightedRdMapIndex = Index ' store the last highlighted
    
   On Error GoTo 0
   Exit Sub

picRdMap_DragOver_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_DragOver of Form rDIconConfigForm"
End Sub
'
'Private Sub picRdMap_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
'    MsgBox "OLE Dragged over"
'End Sub
'

' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
Private Sub picRdMap_MouseUp(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    rdMapIconMouseDown = False
    
    ' add a vbEndDrag here
    picRdMap(Index).Drag vbEndDrag
End Sub

' .61 DAEB 01/05/2022 rDIConConfig.frm Added highlighting to the rdIconMap during Drag and drop.
Private Sub picRdThumbFrame_DragOver(ByRef Source As Control, ByRef X As Single, ByRef Y As Single, ByRef State As Integer)
    ' as you leave the map the frame surrounds and gaps are interspersed between the map elements
    ' if the rdMap already has a highlighted icon, then clear the highlight
    picRdMap(lastHighlightedRdMapIndex).BorderStyle = 0
End Sub

Private Sub picThumbIcon_Click(ByRef Index As Integer)
    ' the click event for this control was handled by the mouseDown to allow the right click to be handled too
End Sub




' .63 DAEB 04/05/2022 rDIConConfig.frm Moved the left click functionality to a separate routine
'---------------------------------------------------------------------------------------
' Procedure : picThumbIconMouseDown_event
' DateTime  : 01/05/2022 23:54
' Author    : beededea
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picThumbIconMouseDown_event(ByVal Index As Integer)
    Dim thumbItemNo As Integer: thumbItemNo = 0
    
    On Error GoTo picThumbIconMouseDown_event_Error
    thisRoutine = "picThumbIconMouseDown_event"
    
    If Index = 0 Then ' a temporary kludge to fix a bug
        thumbPos0Pressed = True
    Else
        thumbPos0Pressed = False
    End If
    
    'Do not refresh the whole thumbnail view array
    refreshThumbnailView = False
    keyPressOccurred = True

    thumbIndexNo = Index ' allow other functions access to the chosen index number

    ' extract the filename from the associated array
    If Not picThumbIcon(Index).ToolTipText = vbNullString Then ' we use the tooltip because the .picture property is not populated
        thumbItemNo = thumbArray(Index)
        'this next line change is meant to trigger a re-click but it does not when the index is unchanged from previous click
        vScrollThumbs.Value = thumbItemNo

         ' this next 'if then' checks to see if the stored click is the same as the current, if so it triggers a click on the item in the underlying file list box
        If storedIndex = Index Or storedIndex = 9999 Then  ' if the storedindex = 9999 it is the first time the icon has been pressed so it triggers
           '
           'Call vScrollThumbs_Change ' TBD1
        End If
        storedIndex = Index
    End If
    
    lblBlankText.Visible = False ' TBD
   
    On Error GoTo 0
    Exit Sub

picThumbIconMouseDown_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIconMouseDown_event of Form rDIconConfigForm"
End Sub

' .59 DAEB 01/05/2022 rDIConConfig.frm Added Drag and drop functionality, moved mouseDown code to dragDrop event STARTS
Private Sub picThumbIcon_MouseUp(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    picThumbIconMouseDown = False
    picThumbIcon(Index).Drag vbEndDrag
    'MsgBox "Dropped"
End Sub



'Private Sub picRdThumbFrame_OLEDragDrop(Data As DataObject, Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
'    MsgBox "OLE Dragged over"
'End Sub

'
'Private Sub picRdMap_MouseUp(Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
'    MsgBox "picRdMap_MouseUp"
'End Sub

'Private Sub picThumbIcon_MouseUp(Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
'    picThumbIcon(Index).DragMode = 0 'disabled
'
'End Sub

'Private Sub fraThumbLabel_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
'    fraThumbLabel(Index).ZOrder
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : rdMapHScroll_Change
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Scrolls the whole RdMap array of picboxes using .move
'             The rdMap sits within a frame and the picboxes to the left and right are obscured.
'             It was easier to just slide the boxes right or left rather than maintain a
'             view on the array. VB allows .left and .right values greater that 32768
'             all you have to do is be careful with manipulating those values in variables
'---------------------------------------------------------------------------------------
'
Private Sub rdMapHScroll_Change()
    Dim useloop As Long: useloop = 0
    Dim startPos As Long: startPos = 0
    Dim maxPos As Long: maxPos = 0
    Dim rdIconMaxLong As Long: rdIconMaxLong = 0
    Dim spacing As Integer: spacing = 0
    
    On Error GoTo rdMapHScroll_Change_Error
    If debugflg = 1 Then DebugPrint "%" & "rdMapHScroll_Change"
   
    spacing = 540

    rdIconMaxLong = rdIconMaximum
    rdMapHScroll.Min = 0
    rdMapHScroll.Max = theCount - 1
    
    startPos = rdMapHScroll.Value - 1
    
    'xlabel.Caption = startPos
    'nLabel.Caption = (startPos * spacing)
    
    maxPos = rdIconMaxLong * spacing
    
    For useloop = 0 To rdIconMaximum
            picRdMap(useloop).Move ((useloop * spacing) - (startPos * spacing)), 30, 500, 500
    Next useloop

   On Error GoTo 0
   Exit Sub

rdMapHScroll_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapHScroll_Change of Form rDIconConfigForm"
    
End Sub



Private Sub sliPreviewSize_Click()
    ' do NOT remove
End Sub





Private Sub txtAppToTerminate_Change()
        btnSet.Enabled = True ' tell the program that something has changed
            btnCancel.Visible = True
    btnClose.Visible = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtAppToTerminate_MouseDown
' Author    : beededea
' Date      : 08/02/2023
' Purpose   : strange code to enable a menu right click on a text area
'---------------------------------------------------------------------------------------
'
Private Sub txtAppToTerminate_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo txtAppToTerminate_MouseDown_Error

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtAppToTerminate.Enabled = False
        txtAppToTerminate.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtAppToTerminate_MouseDown_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAppToTerminate_MouseDown of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub txtAppToTerminate_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtAppToTerminate.hwnd, "Any program that must be terminated prior to the main program initiation will be shown here. The text placed here must be the correct and full path/filename of the application to kill. The program name is selected using the program selection button on the right. The result is: when you click on the icon in the dock SteamyDock will do its very best to terminate the chosen application in advance but be aware that closing another application cannot be guaranteed - use this functionality with great care! ", _
                  TTIconInfo, "Help on Terminating an Application", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtArguments_MouseDown
' Author    : beededea
' Date      : 08/02/2023
' Purpose   : strange code to enable a menu right click on a text area
'---------------------------------------------------------------------------------------
'
Private Sub txtArguments_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .98 DAEB 26/06/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.

    On Error GoTo txtArguments_MouseDown_Error

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtArguments.Enabled = False
        txtArguments.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtArguments_MouseDown_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtArguments_MouseDown of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

Private Sub txtCurrentIcon_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .98 DAEB 26/06/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtCurrentIcon.Enabled = False
        txtCurrentIcon.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub txtLabelName_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .98 DAEB 26/06/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtLabelName.Enabled = False
        txtLabelName.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
'---------------------------------------------------------------------------------------
' Procedure : txtSecondApp_Change
' Author    : beededea
' Date      : 08/02/2023
' Purpose   : strange code to enable a menu right click on a text area
'---------------------------------------------------------------------------------------
'
Private Sub txtSecondApp_Change()
    On Error GoTo txtSecondApp_Change_Error

        btnSet.Enabled = True ' tell the program that something has changed
    btnCancel.Visible = True
    btnClose.Visible = False
    optRunSecondAppBeforehand.Enabled = True
        optRunSecondAppAfterward.Enabled = True
        lblRunSecondAppBeforehand.Enabled = True
        lblRunSecondAppAfterward.Enabled = True

    On Error GoTo 0
    Exit Sub

txtSecondApp_Change_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtSecondApp_Change of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 20/09/2019
' Purpose   : Test to see if the system colour settings have changed due to a theme changing
'---------------------------------------------------------------------------------------
'
Public Sub themeTimer_Timer()
    Dim SysClr As Long: SysClr = 0

' This should only be required on a machine that can give the Windows classic theme to the UI
' that excludes windows 8 and 10 so this timer can be switched off on these o/s.

   On Error GoTo themeTimer_Timer_Error
    If debugflg = 1 Then DebugPrint "%themeTimer_Timer"

    SysClr = GetSysColor(COLOR_BTNFACE)
    If debugflg = 1 Then DebugPrint "COLOR_BTNFACE = " & SysClr ' generates too many debug statements in the log
    If SysClr <> storeThemeColour Then
    
        Call setThemeColour(Me)

    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : folderTreeView_GotFocus
' Author    : beededea
' Date      : 16/09/2019
' Purpose   : remove the focus from the small thumbnails
'---------------------------------------------------------------------------------------
'
Private Sub folderTreeView_GotFocus()

    
   On Error GoTo folderTreeView_GotFocus_Error
   If debugflg = 1 Then DebugPrint "%folderTreeView_GotFocus"

   Call thumbsLostFocus

   On Error GoTo 0
   Exit Sub

folderTreeView_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure folderTreeView_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : thumbsLostFocus
' Author    : beededea
' Date      : 16/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub thumbsLostFocus()

   On Error GoTo thumbsLostFocus_Error
   If debugflg = 1 Then DebugPrint "%thumbsLostFocus"

    If picFrameThumbsLostFocus = False Then
        lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200) ' grey
        lblThumbName(thumbIndexNo).ForeColor = RGB(0, 0, 0) ' black
    End If

   On Error GoTo 0
   Exit Sub

thumbsLostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thumbsLostFocus of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : txtArguments_Change
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtArguments_Change()
   On Error GoTo txtArguments_Change_Error
   If debugflg = 1 Then DebugPrint "%txtArguments_Change"

    btnSet.Enabled = True ' tell the program that something has changed
    btnCancel.Visible = True
    btnClose.Visible = False
    
   On Error GoTo 0
   Exit Sub

txtArguments_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtArguments_Change of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtLabelName_Change
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtLabelName_Change()
   On Error GoTo txtLabelName_Change_Error
   If debugflg = 1 Then DebugPrint "%txtLabelName_Change"

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

txtLabelName_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtLabelName_Change of Form rDIconConfigForm"
End Sub


Private Sub txtSecondApp_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .98 DAEB 26/06/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtSecondApp.Enabled = False
        txtSecondApp.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtStartIn_Change
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtStartIn_Change()
   On Error GoTo txtStartIn_Change_Error
   If debugflg = 1 Then DebugPrint "%txtStartIn_Change"

    btnSet.Enabled = True ' tell the program that something has changed
    btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

txtStartIn_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtStartIn_Change of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtStartIn_DblClick
' Author    : beededea
' Date      : 01/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtStartIn_DblClick()
    On Error GoTo txtStartIn_DblClick_Error

    txtStartIn.Text = getFolderNameFromPath(txtTarget.Text)

    On Error GoTo 0
    Exit Sub

txtStartIn_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtStartIn_DblClick of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtStartIn_MouseDown
' Author    : beededea
' Date      : 08/02/2023
' Purpose   : strange code to enable a menu right click on a text area
'---------------------------------------------------------------------------------------
'
Private Sub txtStartIn_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .98 DAEB 26/06/2022 rDIConConfig.frm For all the text boxes swap the IME right click menu for a useful one, in context.

    On Error GoTo txtStartIn_MouseDown_Error

    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtStartIn.Enabled = False
        txtStartIn.Enabled = True
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtStartIn_MouseDown_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtStartIn_MouseDown of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtTarget_Change
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtTarget_Change()
   On Error GoTo txtTarget_Change_Error
   If debugflg = 1 Then DebugPrint "%txtTarget_Change"

    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False

   On Error GoTo 0
   Exit Sub

txtTarget_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtTarget_Change of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
'---------------------------------------------------------------------------------------
'
Private Sub getKeyPress(ByVal KeyCode As Integer)

    '36 home
    '40 is down
    '38 is up
    '37 is left
    '39 is right
    ' 33 page up
    ' 34 page down
    ' 35 end
    
    On Error GoTo getkeypress_Error
    thisRoutine = "getKeyPress"

    
    If debugflg = 1 Then DebugPrint "%" & "getkeypress"
    If debugflg = 1 Then DebugPrint "%" & "keycode= " & KeyCode
    
    keyPressOccurred = True
    displayHourglass = False
        
    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True
        Case vbKeyD
            CTRL_2 = True
    End Select
    If CTRL_1 And CTRL_2 Then
            CTRL_1 = False
            CTRL_2 = False
            Call deleteRdMap(True, True)
    End If
    
    'f5 refresh button as per all browsers
    If KeyCode = 116 Then

        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then refresh on f5
            rdMapRefresh_Click
        End If
    End If

    ' home key
    If KeyCode = 36 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the first icon
            btnHomeRdMap
        End If
    End If
    
    ' end key pressed
    If KeyCode = 35 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the end of the rdMap
            btnEndRdMap
        End If
    End If
    

    If debugflg = 1 Then DebugPrint "%" & "picFrameThumbsGotFocus= " & picFrameThumbsGotFocus

    '38 is key press up
    If KeyCode = 38 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll up one line
        End If
    End If
    
    '40 is down
    If KeyCode = 40 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll down one line 'TODO
        End If
    End If

    '37 is left
    If KeyCode = 37 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnPrev_Click

        End If
    End If
    
    '39 is right
    If KeyCode = 39 Then
        
        'DebugPrint "########################### getkeypress right STARTS"
    
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnNext_Click
        End If
    End If
    
    '33 is the page up key
    If KeyCode = 33 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            ' scroll toward the end of the map by 15 items allowing some visual overlap
            Call rdMapPageUp_Press
        End If
    End If

    '34 is PAGE down
    If KeyCode = 34 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            ' scroll toward the start of the map by 15 items allowing some visual overlap
            Call rdMapPageDown_Press
        End If
    End If
    
    'xlabel.Caption = thumbIndexNo
    'nLabel.Caption = filesIconList.ListIndex
    
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : thumbnailGetKeyPress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
'---------------------------------------------------------------------------------------
'
Private Sub thumbnailGetKeyPress(ByVal KeyCode As Integer)

    '36 home
    '40 is down
    '38 is up
    '37 is left
    '39 is right
    ' 33 page up
    ' 34 page down
    ' 35 end
    
    On Error GoTo thumbnailGetKeyPress_Error
    thisRoutine = "thumbnailGetKeyPress"

    
    If debugflg = 1 Then DebugPrint "%" & "thumbnailGetKeyPress"
    If debugflg = 1 Then DebugPrint "%" & "keycode= " & KeyCode
    
    keyPressOccurred = True
    displayHourglass = False
        
    Select Case KeyCode
        Case vbKeyControl
            CTRL_1 = True
        Case vbKeyD
            CTRL_2 = True
    End Select
    If CTRL_1 And CTRL_2 Then
            CTRL_1 = False
            CTRL_2 = False
            Call deleteRdMap(True, True)
    End If
    
    'f5 refresh button as per all browsers
    If KeyCode = 116 Then
        If picFrameThumbsGotFocus = True Or filesIconListGotFocus = True Then
            ' if the thumbframe has focus then refresh on f5
            Call btnRefresh_Click_Event
        End If
    End If

    ' home key
    If KeyCode = 36 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = True
            triggerStartCalc = True
            thumbIndexNo = 0
            vScrollThumbs.Value = vScrollThumbs.Min
        End If
    End If
    
    ' end key pressed
    If KeyCode = 35 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = True
            triggerStartCalc = True
            vScrollThumbs.Value = vScrollThumbs.Max
        End If
    End If
    

    If debugflg = 1 Then DebugPrint "%" & "picFrameThumbsGotFocus= " & picFrameThumbsGotFocus

    '38 is key press up
    If KeyCode = 38 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = False
            If thumbIndexNo >= 4 Then
                refreshThumbnailView = True
                thumbIndexNo = thumbIndexNo - 4
                If thumbIndexNo < 0 Then thumbIndexNo = 0
                vScrollThumbs.Value = thumbArray(thumbIndexNo)
            Else
                 ' if there are no more icons preceding this then move on
                 If filesIconList.ListIndex - thumbIndexNo > 3 Then
                   ' there are any icons preceding this do we will scroll to them
                   refreshThumbnailView = True
                    thumbnailStartPosition = thumbnailStartPosition - 4
                    If vScrollThumbs.Value - 4 < vScrollThumbs.Min Then
                        vScrollThumbs.Value = vScrollThumbs.Min
                        'the above line does not trigger a change to vscroll
                        vScrollThumbs_Change ' do it manually
                    Else
                        vScrollThumbs.Value = vScrollThumbs.Value - 4
                    End If
                    
                End If
            End If
        End If
    End If
    
    '40 is down
    If KeyCode = 40 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = False
            If thumbIndexNo <= 7 Then
                If thumbArray(thumbIndexNo + 4) <= vScrollThumbs.Max Then
                    thumbIndexNo = thumbIndexNo + 4
                End If
                
                If thumbArray(thumbIndexNo) Then
                    If thumbArray(thumbIndexNo) <= vScrollThumbs.Max Then
                        vScrollThumbs.Value = thumbArray(thumbIndexNo)
                    Else
                        vScrollThumbs.Value = vScrollThumbs.Max
                    End If
                End If
            Else
                refreshThumbnailView = True
                If vScrollThumbs.Value + 4 > vScrollThumbs.Max Then
                    vScrollThumbs.Value = vScrollThumbs.Max
                Else
                    vScrollThumbs.Value = vScrollThumbs.Value + 4
                End If
            End If
        End If
    End If

    '37 is left
    If KeyCode = 37 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = False
            ' if this is the first icon in the array 0-11
            If thumbIndexNo <= 0 Then
                 ' if there are any icons preceding this then scroll to it
                 If filesIconList.ListIndex > 0 Then
                    refreshThumbnailView = True
                    thumbnailStartPosition = thumbnailStartPosition - 1
                    If vScrollThumbs.Value > vScrollThumbs.Min Then
                        vScrollThumbs.Value = vScrollThumbs.Value - 1
                    Else
                        thumbnailStartPosition = thumbnailStartPosition - 1
                        vScrollThumbs.Value = vScrollThumbs.Value
                    End If

                    vScrollThumbsGotFocus = False
                    picFrameThumbsGotFocus = True
                    thumbIndexNo = 0
                 Else
                    thumbIndexNo = 0
                    vScrollThumbs.Value = thumbArray(thumbIndexNo)
                 End If
            Else
                thumbIndexNo = thumbIndexNo - 1
                vScrollThumbs.Value = thumbArray(thumbIndexNo)
            End If
            
'            DebugPrint "thumbnailGetKeyPress left "
'            DebugPrint thumbIndexNo
'            Sleep (1000)
        End If
    End If
    
    '39 is right
    If KeyCode = 39 Then
        
        'DebugPrint "########################### thumbnailGetKeyPress right STARTS"
    
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = False
            If thumbArray(thumbIndexNo + 1) <= vScrollThumbs.Max Then
                thumbIndexNo = thumbIndexNo + 1
            End If

            If thumbIndexNo > 11 Then
                thumbIndexNo = 11
                ' check if there are any icons subsequent to this
                ' if so then scroll down one line using the vertical scroll bar
                ' and select the next icon, the first icon on that line
                
                If filesIconList.ListIndex = filesIconList.ListCount - 1 Then Exit Sub ' .55 DAEB 25/04/2022 rDIConConfig.frm Fixed bug where scrolling to the end does not quite get to the end.
                refreshThumbnailView = True
                
                vScrollThumbs.Value = vScrollThumbs.Value + 1
                vScrollThumbsGotFocus = False
                picFrameThumbsGotFocus = True
        
            Else
            
                If thumbArray(thumbIndexNo) Then
                    If thumbArray(thumbIndexNo) <= vScrollThumbs.Max Then
                        vScrollThumbs.Value = thumbArray(thumbIndexNo)
                    Else
                        vScrollThumbs.Value = vScrollThumbs.Max
                    End If
                End If
            End If
'            DebugPrint thumbIndexNo
'            DebugPrint "########################### thumbnailGetKeyPress right ENDS"
'            Sleep (1000)

        End If
    End If
    
    '33 is the page up key
    If KeyCode = 33 Then
        If picFrameThumbsGotFocus = True Then
            refreshThumbnailView = True
            triggerStartCalc = True
            If vScrollThumbs.Value - 12 < vScrollThumbs.Min Then
                vScrollThumbs.Value = vScrollThumbs.Min
            Else
                vScrollThumbs.Value = vScrollThumbs.Value - 12
            End If
        End If
    End If

    '34 is PAGE down
    If KeyCode = 34 Then
        If picFrameThumbsGotFocus = True Then
                refreshThumbnailView = True
                triggerStartCalc = True
                If vScrollThumbs.Value + 12 > vScrollThumbs.Max Then
                    vScrollThumbs.Value = vScrollThumbs.Max
                Else
                    vScrollThumbs.Value = vScrollThumbs.Value + 12
                End If
        End If
    End If
    
    'xlabel.Caption = thumbIndexNo
    'nLabel.Caption = filesIconList.ListIndex
    
    On Error GoTo 0
   Exit Sub

thumbnailGetKeyPress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thumbnailGetKeyPress of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : lblThumbName_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblThumbName_Click(ByRef Index As Integer)
    ' clicking on the text below the icon triggers a click on the icon itself
   On Error GoTo lblThumbName_Click_Error
   If debugflg = 1 Then DebugPrint "%lblThumbName_Click"

    Call picThumbIcon_MouseDown(Index, 1, 0, 0, 0)

   On Error GoTo 0
   Exit Sub

lblThumbName_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblThumbName_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picFrameThumbs_GotFocus
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFrameThumbs_GotFocus()
   On Error GoTo picFrameThumbs_GotFocus_Error
   If debugflg = 1 Then DebugPrint "%picFrameThumbs_GotFocus"

    picFrameThumbsGotFocus = True

   On Error GoTo 0
   Exit Sub

picFrameThumbs_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbs_GotFocus of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : picFrameThumbs_KeyDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picFrameThumbs_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo picFrameThumbs_KeyDown_Error
   If debugflg = 1 Then DebugPrint "%picFrameThumbs_KeyDown"

    Call thumbnailGetKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

picFrameThumbs_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbs_KeyDown of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : picFrameThumbs_MouseDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   : Each frame in VB6 has a mousedown to catch a right click and select which menu to display.
'---------------------------------------------------------------------------------------
'
Private Sub picFrameThumbs_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo picFrameThumbs_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%picFrameThumbs_MouseDown"

   If Button = 2 Then
        'storedIndex = Index ' get the icon number from the array's index
        menuAddToDock.Caption = "Add icon at position " & rdIconNumber + 1 & " in the map"

        Me.PopupMenu thumbmenu, vbPopupMenuRightButton
        
    End If

   On Error GoTo 0
   Exit Sub

picFrameThumbs_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbs_MouseDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picThumbIcon_MouseDown (was picThumbIcon_Click)
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : There is a filelist control underneath the thumbnail view
'             A click on any icon thumbnail causes a click on the underlying filelist
'             control which in turn shows a preview
'---------------------------------------------------------------------------------------
' Initially, changed from a click to a mousedown as it allows it to catch the right button press and retain the index
Private Sub picThumbIcon_MouseDown(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    
    thisRoutine = "picThumbIcon_MouseDown"
    picThumbIconMouseDown = True

    If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_MouseDown"
    On Error GoTo picThumbIcon_MouseDown_Error
        
    If Button = 2 Then
        menuAddToDock.Caption = "Add icon at position " & rdIconNumber + 1 & " in the map"
        storedIndex = Index ' get the icon number from the array's index
        Me.PopupMenu thumbmenu, vbPopupMenuRightButton
        Exit Sub
    End If
    
    If picRdThumbFrame.Visible = True Then
    
        srcDragControl = "picThumbIcon"
        dragToDockOperating = True
        
        ' .63 DAEB 04/05/2022 rDIConConfig.frm Moved the left click functionality to a separate routine
        Call picThumbIconMouseDown_event(Index) ' I thought I might use this elsewhere but in the end I did not need to
        
         ' .66 DAEB 04/05/2022 rDIConConfig.frm use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon STARTS
        imlDragIconConverter.ListImages.Clear ' first clear it down
        
        ' .picture is the graphic itself
        ' .image property is a bitmap handle to the actual rendered "canvas" of the (resized) container
        
        ' .80 DAEB 28/05/2022 rDIConConfig.frm Change to adding the .picture to workaround the bug in Krool's imageList failing to convert to an HIcon.
        If thumbImageSize = 32 Then
            Set picTemporaryStore.Picture = picTemporaryStore.Image
            imlDragIconConverter.ListImages.Add , "arse", picTemporaryStore.Picture
        Else
            Set picThumbIcon(Index).Picture = picThumbIcon(Index).Image
            imlDragIconConverter.ListImages.Add , "arse", picThumbIcon(Index).Picture
        End If
        Set picTemporaryStore.Picture = Nothing
        
        ' the use of the imgList should convert the picture to an icon with a transparent background for dragging.
        picThumbIcon(Index).DragIcon = imlDragIconConverter.ListImages("arse").ExtractIcon
            
        ' .66 DAEB 04/05/2022 rDIConConfig.frm use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon ENDS
        ' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
        thumbnailDragTimer.Enabled = True ' initiates the vbBeginDrag after n millisecs
    End If
    
   On Error GoTo 0
   Exit Sub

picThumbIcon_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIcon_MouseDown of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picThumbIcon_DblClick
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the filename from the associated array
'---------------------------------------------------------------------------------------

Private Sub picThumbIcon_DblClick(ByRef Index As Integer)
    Dim itemno As Integer: itemno = 0

    On Error GoTo picThumbIcon_DblClick_Error
    If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_DblClick"

    Call btnAdd_Click

   On Error GoTo 0
   Exit Sub

picThumbIcon_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIcon_DblClick of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picThumbIcon_GotFocus
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picThumbIcon_GotFocus(ByRef Index As Integer)
   On Error GoTo picThumbIcon_GotFocus_Error
   If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_GotFocus"

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False
    
    On Error GoTo 0
    Exit Sub

picThumbIcon_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIcon_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picThumbIcon_KeyDown
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : A key down when the individual icons are in focus and a key is pressed
'---------------------------------------------------------------------------------------
'
Private Sub picThumbIcon_KeyDown(ByRef Index As Integer, ByRef KeyCode As Integer, ByRef Shift As Integer)
    If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_KeyDown"
    On Error GoTo picThumbIcon_KeyDown_Error

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False
    
    'Label2.Caption = "picThumbIcon_KeyDown"

    Call thumbnailGetKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

picThumbIcon_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIcon_KeyDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picThumbIcon_MouseMove
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : As the mouse is moved over the icons bring the label to the fore
'---------------------------------------------------------------------------------------
'
Private Sub picThumbIcon_MouseMove(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo picThumbIcon_MouseMove_Error
   'If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_MouseMove"

    fraThumbLabel(Index).ZOrder

   On Error GoTo 0
   Exit Sub

picThumbIcon_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picThumbIcon_MouseMove of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picPreview_DblClick
' Author    : beededea
' Date      : 19/06/2019
' Purpose   : on a double click search the file list and select the image that is being shown by filename match
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_DblClick()

    Dim useloop As Integer: useloop = 0

    On Error GoTo picPreview_DblClick_Error
    If debugflg = 1 Then DebugPrint "%" & "picPreview_DblClick"
   
    For useloop = 0 To filesIconList.ListCount - 1
    ' TODO - extract just the filename from txtCurrentIcon.Text
        If filesIconList.List(useloop) = getFileNameFromPath(txtCurrentIcon.Text) Then
            filesIconList.ListIndex = useloop
            GoTo l_found_file ' if the file is found no need to process the whole list
        End If
    Next useloop
    MsgBox ("The icon " & getFileNameFromPath((txtCurrentIcon.Text)) & " is not found in the currently selected folder, please select the " & getFolderNameFromPath(txtCurrentIcon.Text) & " folder")
l_found_file:

    ' using the current preview image as the start point on the list, repopulate the thumbs
    If picFrameThumbs.Visible = True Then
        Call populateThumbnails(thumbImageSize, filesIconList.ListIndex)
    
        removeThumbHighlighting

        'highlight the current thumbnail
        thumbIndexNo = 0
        If thumbImageSize = 64 Then 'larger
            picFraPicThumbIcon(thumbIndexNo).BorderStyle = 1
            'picThumbIcon(thumbIndexNo).BorderStyle = 1
        ElseIf thumbImageSize = 32 Then

            lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200)
        End If
    End If

   On Error GoTo 0
   Exit Sub

picPreview_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_DblClick of Form rDIconConfigForm"

End Sub

''---------------------------------------------------------------------------------------
'' Procedure : getFileNameFromPath
'' Author    : beededea
'' Date      : 01/06/2019
'' Purpose   : A function to getFileNameFromPath
''---------------------------------------------------------------------------------------
''
'Public Function getFileNameFromPath(ByRef strFullPath As String) As String
'   On Error GoTo getFileNameFromPath_Error
'   If debugflg = 1 Then DebugPrint "%" & "getFileNameFromPath"
'
'   getFileNameFromPath = right$(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
'
'   On Error GoTo 0
'   Exit Function
'
'getFileNameFromPath_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFileNameFromPath of Form rDIconConfigForm"
'End Function
    '

'---------------------------------------------------------------------------------------
' Procedure : picPreview_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - but add one new option for adding this icon to the map -
'             this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo picPreview_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "picPreview_MouseDown"
   
    If Button = 2 Then
        If picPreview.Tag <> txtCurrentIcon.Text Then
            mnuAddPreviewIcon.Caption = "Add this icon at position " & rdIconNumber + 1 & " in the map"
            mnuAddPreviewIcon.Visible = True
        Else
            mnuAddPreviewIcon.Visible = False
        End If
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
        'Me.PopupMenu previewMenu, vbPopupMenuRightButton  ' not using the preview menu but adding to the main menu instead
    End If

   On Error GoTo 0
   Exit Sub

picPreview_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_MouseDown of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : picRdMap_GotFocus
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : if the rdMap has focus then a keypress affects the scrolling of that left/right, otherwise it is the thumbnail view
'---------------------------------------------------------------------------------------
'
Private Sub picRdMap_GotFocus(ByRef Index As Integer)
    On Error GoTo picRdMap_GotFocus_Error
    If debugflg = 1 Then DebugPrint "%" & "picRdMap_GotFocus"

    picRdMapGotFocus = True
    picFrameThumbsGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False
    
    On Error GoTo 0
    Exit Sub

picRdMap_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRdMap_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get left/right keypresse to scroll the RD map
'---------------------------------------------------------------------------------------
'
Private Sub picRdMap_KeyDown(ByRef Index As Integer, ByRef KeyCode As Integer, ByRef Shift As Integer)
   On Error GoTo picRdMap_KeyDown_Error
    If debugflg = 1 Then DebugPrint "%" & "picRdMap_KeyDown"

    Call getKeyPress(KeyCode)

   On Error GoTo 0
   Exit Sub

picRdMap_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_KeyDown of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRdMap_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : Mousedown rather than click as it allows us to respond to left click as well as right click events#
'             if the right click is selected it offers the choice to add or delete an icon
'---------------------------------------------------------------------------------------
'
Private Sub picRdMap_MouseDown(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    Dim useloop As Integer: useloop = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    On Error GoTo picRdMap_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "picRdMap_MouseDown"

    If Button = 2 Then
        rdIconNumber = Index ' get the icon number from the array's index
        Me.PopupMenu rdMapMenu, vbPopupMenuRightButton
        Exit Sub
    End If
    
    ' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
    Call picRdMap_MouseDown_event(Index)
    
    ' .66 DAEB 04/05/2022 rDIConConfig.frm Use a hidden picbox (picTemporaryStore) to be used to populate the dragIcon.
    imlDragIconConverter.ListImages.Clear
    'imlDragIconConverter.ListImages.Add , "arse", picTemporaryStore.Image  ' adds the icon to key position
    
    ' .picture is the graphic itself
    ' .image property is a bitmap handle to the actual rendered "canvas" of the (resized) container
    
    ' .80 DAEB 28/05/2022 rDIConConfig.frm Change to adding the .picture to workaround the bug in Krool's imageList failing to convert to an HIcon.
    Set picTemporaryStore.Picture = picTemporaryStore.Image
    imlDragIconConverter.ListImages.Add , "arse", picTemporaryStore.Picture
    Set picTemporaryStore.Picture = Nothing
    
    picRdMap(Index).DragIcon = imlDragIconConverter.ListImages("arse").ExtractIcon
    
    ' .68 DAEB 04/05/2022 rDIConConfig.frm Added a timer to activate Drag and drop from the thumbnails to the rdmap only after 25ms
    rdMapIconMouseDown = True
    rdMapIconSrcIndex = Index
    rdMapDragTimer.Enabled = True ' initiates the vbBeginDrag after n millisecs
    
   On Error GoTo 0
   Exit Sub

picRdMap_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_MouseDown of Form rDIconConfigForm"
End Sub


' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
'---------------------------------------------------------------------------------------
' Procedure : picRdMap_MouseDown_event
' Author    : beededea
' Date      : 16/05/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub picRdMap_MouseDown_event(Index)
    Dim answer As VbMsgBoxResult: answer = vbNo
    On Error GoTo picRdMap_MouseDown_event_Error

    If btnSet.Enabled = True Then
        ' 17/11/2020    .03 DAEB Replaced the confirmation dialog with an automatic save when moving from one icon to another using the right/left icon buttons
'        If chkToggleDialogs.Value = 1 Then
'            answer = msgBoxA(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
'            If answer = vbNo Then
'                Exit Sub
'            End If
'
'        Else
'            Call btnSet_Click
'        End If

        

        If mapImageChanged = True Then
            ' now change the icon image back again
            ' the target picture control and the icon size
            'Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
            mapImageChanged = False
        End If
    End If
    
    'remove the highlighting on the Rocket dock map
    picRdMap(rdIconNumber).BorderStyle = 0
   
    rdIconNumber = Index
    
    lblRdIconNumber.Caption = Str$(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str$(rdIconNumber) + 1
    Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, True, False)

    'set the highlighting on the Rocket dock map
    If Index <= rdIconMaximum Then
        picRdMap(Index).BorderStyle = 1
    End If
    
    ' TBD
    lastHighlightedRdMapIndex = Index
    
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

        

    On Error GoTo 0
    Exit Sub

picRdMap_MouseDown_event_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_MouseDown_event of Form rDIconConfigForm"
            Resume Next
          End If
    End With

End Sub

Private Sub picRdMap_MouseMove(ByRef Index As Integer, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
' code retained in case I want to do a graphical drag and drop of one item in the map to another

' Dim picX As Integer
'
' With picRdMap(Index)
' If Button Then
'  .Move .Left + (X) - picX
' End If
' End With

   If rDEnableBalloonTooltips = "1" Then CreateToolTip picRdMap(Index).hwnd, "This is the icon map. It maps your dock exactly, showing you the same icons that appear in your dock. You can add or delete icons to/from the map. Press save and restart and they will appear in your dock.", _
                  TTIconInfo, "Help on the Icon Map", , , , True

End Sub




'---------------------------------------------------------------------------------------
' Procedure : picPreview_OLEDragDrop
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : drag and drop of an image to the icon preview
'---------------------------------------------------------------------------------------
'
'Private Sub picPreview_OLEDragDrop(data As DataObject, Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
'
'    ' use class to load 1st file that was dropped, if more than one. Unicode compatible
'   On Error GoTo picPreview_OLEDragDrop_Error
'      If debugflg = 1 Then DebugPrint "%" & "picPreview_OLEDragDrop"
'
'
'
'    If cImage.LoadPictureDroppedFiles(data, 1, 256, 256) Then
'
'        Call refreshPicBox(picPreview, 256)
'
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'picPreview_OLEDragDrop_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_OLEDragDrop of Form rDIconConfigForm"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : rdMapRefresh_Click
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : the refresh button for the rocketdock map
'             it redraws the whole rdMap
'---------------------------------------------------------------------------------------
'
Private Sub rdMapRefresh_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
   
    On Error GoTo rdMapRefresh_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "rdMapRefresh_Click"
    
    If btnSet.Enabled = True Or mapImageChanged = True Then
        If chkToggleDialogs.Value = 1 Then
            answer = msgBoxA("This will lose your recent changes to the map. Proceed?", vbQuestion + vbYesNo, "Lose your changes?", True, "rdMapRefresh_Click")
            If answer = vbNo Then
                Exit Sub
            Else
                Me.Refresh ' just to clear the dialog box remnants
            End If
        End If
    End If
    
    mapImageChanged = False
    
    Call busyStart
    
    
    Call recreateTheMap(rdIconMaximum)
    
    Call busyStop
    
    ' we signify that there have been no changes - this is just a refresh
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False


   On Error GoTo 0
   Exit Sub

rdMapRefresh_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapRefresh_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : recreateTheMap
' Author    : beededea
' Date      : 17/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub recreateTheMap(ByVal oldRdIconMax As Integer)
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo recreateTheMap_Error

    useloop = 0
    
    Call readIconsAndConfiguration
    
    If rdIconMaximum > oldRdIconMax Then
        ' if you do a refresh and the old rdIconMaximum is less than the recently read
        ' then items have been added to the Rocketdock via RD itself
        ' in which case you need to create the extra slots in the RD map
        
        'loop from the old rdIconMaximum to the new rdIconMaximum and create a new slot in the map
        For useloop = oldRdIconMax To rdIconMaximum
        '       test to see if the picturebox has already been created
            If CheckControlExists(picRdMap(useloop)) Then
                'do nothing
            Else
                Load picRdMap(useloop) ' dynamically extend the number of picture boxes by one
                picRdMap(useloop).Width = 500
                picRdMap(useloop).Height = 500
                picRdMap(useloop).Left = picRdMap(useloop - 1).Left + boxSpacing
                picRdMap(useloop).Top = 30
                picRdMap(useloop).Visible = True
            End If
        Next useloop
    End If
    
    Call populateRdMap(0) ' show the map from position zero

    On Error GoTo 0
    Exit Sub

recreateTheMap_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure recreateTheMap of Form rDIconConfigForm"
            Resume Next
          End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : registryTimer_Timer
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : check the existence of a settings.ini file in the RD folder and...
'---------------------------------------------------------------------------------------
'
Private Sub registryTimer_Timer()
   On Error GoTo registryTimer_Timer_Error
    'If debugflg = 1 Then DebugPrint "%" & "registryTimer_Timer" ' no messages thankyou

    chkTheRegistry

   On Error GoTo 0
   Exit Sub

registryTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure registryTimer_Timer of Form rDIconConfigForm on line " & Erl
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliPreviewSize_Change
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : The size slider that determines the icon size to display
'---------------------------------------------------------------------------------------
'
Private Sub sliPreviewSize_Change()
    Dim Filename As String: Filename = vbNullString
           
    On Error GoTo sliPreviewSize_Change_Error
    If debugflg = 1 Then DebugPrint "%" & "sliPreviewSize_Change"
    
    If programStatus = "startup" Then Exit Sub ' prevent the control from registering a change just from modifying its theme at startup

    ' change the parameters that govern the display of the icons at specified sizes

    If sliPreviewSize.Value = 1 Then
        icoSizePreset = 16
        sliPreviewSize.ToolTipText = "16x16"
    End If
    If sliPreviewSize.Value = 2 Then
        icoSizePreset = 32
        sliPreviewSize.ToolTipText = "32x32"
    End If
    If sliPreviewSize.Value = 3 Then
        icoSizePreset = 64
        sliPreviewSize.ToolTipText = "64x64"
    End If
    If sliPreviewSize.Value = 4 Then
        icoSizePreset = 128
        sliPreviewSize.ToolTipText = "128x128"
    End If
    If sliPreviewSize.Value = 5 Then
        icoSizePreset = 256
        sliPreviewSize.ToolTipText = "256x256"
    End If
    
    ' display the icon from the settings config.
    
    'if the thumbview or fileiconlist have the focus
    If picRdMapGotFocus = True Or previewFrameGotFocus = True Then
        ' if the map or the preview have focus
        Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
    Else
        If textCurrentFolder.Text <> vbNullString Then ' changed from filesIconList.path to textCurrentFolder.Text for compatibility with VB.net
            Filename = textCurrentFolder.Text ' changed from filesIconList.path to textCurrentFolder.Text for compatibility with VB.net
            If Right$(Filename, 1) <> "\" Then Filename = Filename & "\"
            Filename = Filename & filesIconList.Filename
            ' refresh the image display
            Call displayResizedImage(Filename, picPreview, icoSizePreset)
        End If
    End If

   On Error GoTo 0
   Exit Sub

sliPreviewSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPreviewSize_Change of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : folderTreeView_MouseMove
' Author    : beededea
' Date      : 23/06/2019
' Purpose   : As the user moves the cursor across the various treeview elements
'               change the displayed tooltip
'---------------------------------------------------------------------------------------
'
Private Sub folderTreeView_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   ' this next line is the MSCOMTCL.OCX usage of a Treeview node
   'Dim n As Node
   
   ' this is Krool's treeview replacement of the Treeview node
   Dim N As CCRTreeView.TvwNode

  On Error GoTo folderTreeView_MouseMove_Error
   'If debugflg = 1 Then DebugPrint "%" & "folderTreeView_MouseMove" ' we don't want too many notifications in the debug log

  Set N = folderTreeView.HitTest(X, Y)
   If N Is Nothing Then
    folderTreeView.ToolTipText = "Click a folder to show the icons contained within"
    ElseIf N.Text = "icons" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are Rocketdock/SteamyDock's own in-built icons"
    ElseIf N.Text = "custom folder" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are the custom folders that the user can add using the + button below."
    ElseIf N.Text = "my collection" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are the default folders that come with this enhanced settings utility."
    Else
     folderTreeView.ToolTipText = N.Text
   End If
   
    If rDEnableBalloonTooltips = "1" Then CreateToolTip folderTreeView.hwnd, "These are the icon folders available to SteamyDock. Select each to view the icon sets contained within. You can also add any other existing folders here so that RocketDock or SteamyDock can use the icons within..", _
                TTIconInfo, "Help on the Folder Treeview", , , , True

   On Error GoTo 0
   Exit Sub

folderTreeView_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure folderTreeView_MouseMove of Form rDIconConfigForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : folderTreeView_Click
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : populate the file icon list or thumb view from the chosen folder
'---------------------------------------------------------------------------------------
'
'Private Sub folderTreeView_Click()
Private Sub folderTreeView_NodeSelect(ByVal Node As CCRTreeView.TvwNode)

    On Error GoTo folderTreeView_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "folderTreeView_Click"
   
    Dim Path As String: Path = vbNullString
    Dim defaultFolderNodeKey As String: defaultFolderNodeKey = vbNullString
   
    displayHourglass = True
   
    Call busyStart

    On Error GoTo l_bypass_parent
    
    If Not folderTreeView.SelectedItem Is Nothing Then

         Path = folderTreeView.SelectedItem.Key
         relativePath = Path
    End If
     
l_bypass_parent:
   On Error GoTo folderTreeView_Click_Error
    
    If Not folderTreeView.SelectedItem Is Nothing Then
        textCurrentFolder.Text = Path
        If DirExists(textCurrentFolder.Text) Then
            filesIconList.Path = textCurrentFolder.Text
        End If
        
        defaultFolderNodeKey = folderTreeView.SelectedItem.Key
        'eg. defaultFolderNodeKey=?E:\dean\steampunk theme\icons\
        If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                PutINISetting "Software\SteamyDockSettings", "defaultFolderNodeKey", defaultFolderNodeKey, toolSettingsFile
        End If
            
        If picFrameThumbs.Visible = True Then
            
            Call btnRefresh_Click_Event
        End If
            
        '  .50 when a new icon collection group is selected, the first icon filename is displayed
        filesIconList.ListIndex = 0         'refresh the preview displaying the selected image
        'Call filesIconList_Click
    End If
    
    Call busyStop
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
   On Error GoTo 0
   Exit Sub

folderTreeView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure folderTreeView_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : addtotree
' Author    : beededea
' Date      : 17/06/2019
' Purpose   : add a chosen folder to the treeview
'---------------------------------------------------------------------------------------
'
' this next line is the MSCOMTCL.OCX usage of the Treeview object
'Private Sub addtotree(path As String, tv As TreeView)
   
' this is Krool's treeview replacement passing the Treeview object itself
Private Sub addtotree(ByVal Path As String, ByRef tv As CCRTreeView.TreeView)
    Dim folder1 As Object
    Dim FS As Object
    Dim busyFilename As String: busyFilename = vbNullString

    On Error GoTo addtotree_Error
    If debugflg = 1 Then DebugPrint "%" & "addtotree"

    Set FS = CreateObject("Scripting.FileSystemObject")
    If DirExists(Path) Then
        For Each folder1 In FS.getFolder(Path).SubFolders
        
            ' this next line is the MSCOMTCL.OCX usage of the Treeview TvwChild property
            'tv.Nodes.Add path, CCRTreeView.TvwChild, path & "\" & folder1.Name, folder1.Name
            
            ' this next line is the Krool replacement usage of the Treeview TvwChild property
            tv.Nodes.Add Path, CCRTreeView.TvwNodeRelationshipChild, Path & "\" & folder1.Name, folder1.Name
            
            Call addtotree(Path & "\" & folder1.Name, tv)
            
            ' do the hourglass timer
            If displayHourglass = True Then
                picBusy.Visible = True
                busyCounter = busyCounter + 1
                If busyCounter >= 7 Then busyCounter = 1
                If classicTheme = True Then
                    busyFilename = App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg"
                Else
                    busyFilename = App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg"
                End If
                picBusy.Picture = LoadPicture(busyFilename) ' imageList candidate
            End If
            
        Next
    End If
    picBusy.Visible = False

   On Error GoTo 0
   Exit Sub

addtotree_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") a folder with the same name already exists in the tree view, choose another folder"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : folderTreeView_DblClick
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : open the selected folder in an explorer window
'---------------------------------------------------------------------------------------
'
Private Sub folderTreeView_DblClick()
    
   On Error GoTo folderTreeView_DblClick_Error
      If debugflg = 1 Then DebugPrint "%" & "folderTreeView_DblClick"
   
  
'   Dim a As String
'   Dim fromNode As String
'
'    If DirExists(folderTreeView.SelectedItem.Key) Then
'        ShellExecute 0, vbNullString, folderTreeView.SelectedItem.Key, vbNullString, vbNullString, 1
'    End If


   On Error GoTo 0
   Exit Sub

folderTreeView_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure folderTreeView_DblClick of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : folderTreeView_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub folderTreeView_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo folderTreeView_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "folderTreeView_MouseDown"
    
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        
        mnuOpenFolder.Visible = True
        blank10.Visible = True
        
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
        
        mnuOpenFolder.Visible = False
        blank10.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

folderTreeView_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure folderTreeView_MouseDown of Form rDIconConfigForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : txtCurrentIcon_Change
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : When the user modifies this field the save button should activate
'---------------------------------------------------------------------------------------
'
Private Sub txtCurrentIcon_Change()
   On Error GoTo txtCurrentIcon_Change_Error
      If debugflg = 1 Then DebugPrint "%" & "txtCurrentIcon_Change"
   
   
    'Dim savIt As String

    'savIt = txtCurrentIcon.Text
    
    btnSet.Enabled = True ' tell the program that something has changed
        btnCancel.Visible = True
    btnClose.Visible = False
    
    'txtCurrentIcon.Text = savIt
    
   On Error GoTo 0
   Exit Sub

txtCurrentIcon_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtCurrentIcon_Change of Form rDIconConfigForm"
End Sub



Private Sub txtTarget_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    ' .97 DAEB 26/06/2022 rDIConConfig.frm For the target text box swap the IME right click menu for the target selection menu.
    If Button = 2 Then
        mnuAddPreviewIcon.Visible = False
        txtTarget.Enabled = False
        txtTarget.Enabled = True
        Me.PopupMenu mnuTrgtMenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vScrollThumbs_Change
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Vertical scrollbar for the simulated thumbnail image box
'             a single (small) change of the vscrollbar is a single line of the thumbnail view
'             a large change is the same as a page down
'             IMPORTANT : A change to this scrollbar is passed to the underlying file list box
'---------------------------------------------------------------------------------------
'
Private Sub vScrollThumbs_Change()
    Dim storedIndexCheck As Integer: storedIndexCheck = 0
        
    On Error GoTo vScrollThumbs_Change_Error
    'thisRoutine = "vScrollThumbs_Change"

    If debugflg = 1 Then DebugPrint "%" & "vScrollThumbs_Change"
   
    ' .51 DAEB 24/04/2022 rDIConConfig.frm Icon preview needs to be forced to refresh after a click elsewhere and a return to the same icon thumbnail.
    'Dim picFrameThumbsLostFocus As Boolean
    
    If picFrameThumbsGotFocus = False Then
        picFrameThumbsLostFocus = True
        picFrameThumbsGotFocus = True
    End If
    
    triggerStartCalc = True
    
    'keyPressOccurred = False ' TBD
    If keyPressOccurred = True Then
        picFrameThumbsGotFocus = True
        If picFrameThumbs.Visible Then picFrameThumbs.SetFocus ' important
    Else
        Me.SetFocus
        triggerStartCalc = True
    End If
    
    ' update the underlying file list control that determines which icon has been selected
    ' causes the preview to be refreshed
    If filesIconList.ListCount > 0 Then
        If vScrollThumbs.Value <= vScrollThumbs.Max Then
            
'            Dim a As Integer
'            If thumbIndexNo = 11 Then
'                a = 1
'            End If
            
            storedIndexCheck = filesIconList.ListIndex
            'if the .ListIndex and the vScrollThumbs.Value are the same it does not trigger a click...
            filesIconList.ListIndex = (vScrollThumbs.Value) 'Causes a click on the window that holds the icon files listing in text mode
            
            ' if focus from the thumbnail list has been lost (by a click on the map) then returning and clicking on the first icon in the
            ' thumbnail list will result in the above line: filesIconList.ListIndex = (vScrollThumbs.Value) to be ignored as the index is
            ' the same. Instead we watch for the loss of focus and cause a hard click on filesIconList. This causes the preview image to be
            ' populated with the correct icon in the collection.
            
            ' .51 DAEB 24/04/2022 rDIConConfig.frm Icon preview needs to be forced to refresh after a click elsewhere and a return to the same icon thumbnail.
            ' the second condition prevents the routine being called a second time
            If picFrameThumbsLostFocus = True And storedIndexCheck = filesIconList.ListIndex Then Call filesIconListLeftMouseDown_event     '.64 DAEB 04/05/2022 rDIConConfig.frm Moved the fileList left click functionality to a separate routine
        Else
            filesIconList.ListIndex = (vScrollThumbs.Max) 'Causes a click on the window that holds the icon files listing in text mode
        End If
    End If
    
    ' if a specific thumbnail is selected or if the vscrollbar is used on its own without an icon being selected
    If thumbIndexNo > 0 Or triggerStartCalc = True Then
        thumbnailStartPosition = (filesIconList.ListIndex - thumbIndexNo)
        triggerStartCalc = False
    Else
        
    End If
    
    ' if the thumbnail start position is negative then remove that portion to obtain the correct thumbnail position in the array
    If thumbnailStartPosition < 0 Then
        thumbIndexNo = thumbIndexNo - Abs(thumbnailStartPosition)
    End If
    
    If thumbIndexNo = 0 And thumbPos0Pressed = False Then
        refreshThumbnailView = True ' this causes the first press of the scrollbar to select the correct thumb
        thumbPos0Pressed = False
    End If

    If thumbnailStartPosition <= 0 Then
        thumbnailStartPosition = 0
    End If
            
    'sometimes we want to refresh and other times we do not
    If refreshThumbnailView = True Then
        Call refreshThumbnailViewPanel ' click the button switching to thumbnail view causing a thumbnail list refresh
    Else
        refreshThumbnailView = True ' the refresh flag is set back to true immediately TBD1
    End If
    
    'remove the highlighting
    Call removeThumbHighlighting
    
    'highlight the current thumb
    If thumbIndexNo >= 0 Then ' -1 when there are no icons as a result of an empty filter pattern
        If thumbArray(thumbIndexNo) = 0 Or (thumbArray(thumbIndexNo) And thumbArray(thumbIndexNo) <= vScrollThumbs.Max) Then
            If thumbImageSize = 64 Then 'larger
                picFraPicThumbIcon(thumbIndexNo).BorderStyle = 1
                'picThumbIcon(thumbIndexNo).BorderStyle = 1
            ElseIf thumbImageSize = 32 Then
                ' .58 DAEB 25/04/2022 rDIConConfig.frm second click on a thumbnail should be blue.
                lblThumbName(thumbIndexNo).BackColor = RGB(10, 36, 106) ' blue
                lblThumbName(thumbIndexNo).ForeColor = RGB(255, 255, 255) ' white
            End If
        End If
    End If
    
    txtDbg01.Text = vScrollThumbs.Value
    txtDbg02.Text = vScrollThumbs.Max
    
    lblIconName.Caption = "Icon " & filesIconList.ListIndex + 1 & " Name:"
            
    'Label2.Caption = thisRoutine
   On Error GoTo 0
   Exit Sub

vScrollThumbs_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vScrollThumbs_Change of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : removeThumbHighlighting
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : remove any prior thumbnail highlighting
'---------------------------------------------------------------------------------------
'
Private Sub removeThumbHighlighting()
    Dim useloop As Integer: useloop = 0
    
    'remove the highlighting
   On Error GoTo removeThumbHighlighting_Error
      If debugflg = 1 Then DebugPrint "%" & "removeThumbHighlighting"
    
    'remove the highlighting
    For useloop = 0 To 11
        picFraPicThumbIcon(useloop).BorderStyle = 0
        'picThumbIcon(useloop).BorderStyle = 0
        lblThumbName(useloop).BackColor = &HFFFFFF
        lblThumbName(useloop).ForeColor = &H80000012
    Next useloop

   On Error GoTo 0
   Exit Sub

removeThumbHighlighting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeThumbHighlighting of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelpPdf_click
' Author    : beededea
' Date      : 30/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpPdf_click()
   Dim answer As VbMsgBoxResult: answer = vbNo

   On Error GoTo mnuHelpPdf_click_Error
   If debugflg = 1 Then DebugPrint "%mnuHelpPdf_click"

    answer = msgBoxA("This option opens a browser window and displays this tool's help. Proceed?", vbQuestion + vbYesNo, "Display Help for this tool? ", True, "mnuHelpPdf_click")
    If answer = vbYes Then
        If FExists(App.Path & "\help\Rocketdock Enhanced Settings.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\Rocketdock Enhanced Settings.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("The help file -Rocketdock Enhanced Settings.html- is missing from the help folder.")
        End If
    End If

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpPdf_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFacebook_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFacebook_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo

    On Error GoTo mnuFacebook_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuFacebook_Click"

    answer = msgBoxA("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbQuestion + vbYesNo, "Connect to FaceBook? ", False)
    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFacebook_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelp_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelp_Click(ByRef Index As Integer)

    On Error GoTo mnuHelp_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuHelp_Click"

    rdHelpForm.Visible = True
    
    On Error GoTo 0
    Exit Sub

mnuHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelp_Click of Form quartermaster"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLatest_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLatest_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo

    On Error GoTo mnuLatest_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuLatest_Click"

    answer = msgBoxA("Download latest version of the program - this button opens a browser window and connects to the widget download page where you can check and download the latest zipped file). Proceed?", vbQuestion + vbYesNo)

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
    Dim answer As VbMsgBoxResult: answer = vbNo
    'Dim hWnd As Long

    On Error GoTo mnuSupport_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuSupport_Click"

    answer = msgBoxA("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbQuestion + vbYesNo)

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
    Dim answer As VbMsgBoxResult: answer = vbNo
    'Dim hWnd As Long

    On Error GoTo mnuSweets_Click_Error
       If debugflg = 1 Then DebugPrint "%" & "mnuSweets_Click"
    
    
    answer = msgBoxA(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbQuestion + vbYesNo)

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
    Dim answer As VbMsgBoxResult: answer = vbNo
    'Dim hWnd As Long

    On Error GoTo mnuWidgets_Click_Error
       If debugflg = 1 Then DebugPrint "%" & "mnuWidgets_Click"
    
    

    answer = msgBoxA(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbQuestion + vbYesNo)

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
    If debugflg = 1 Then DebugPrint "%mnuDebug_Click"
    
    Call btnCancel_Click

   On Error GoTo 0
   Exit Sub

mnuClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClose_Clickg_Click of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuDebug_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Run the runtime debugging window exectuable
'---------------------------------------------------------------------------------------
'
Private Sub mnuDebug_Click()
'    Dim NameProcess As String: NameProcess = ""
'    Dim debugPath As String: debugPath = vbNullString
    
    On Error GoTo mnuDebug_Click_Error
    If debugflg = 1 Then DebugPrint "%mnuDebug_Click"

'    NameProcess = "PersistentDebugPrint.exe"
'    debugPath = App.Path() & "\" & NameProcess
    
    If debugflg = 0 Then
        debugflg = 1
'        mnuDebug.Caption = "Turn Debugging OFF"
'        If FExists(debugPath) Then
'            Call ShellExecute(hwnd, "Open", debugPath, vbNullString, App.Path, 1)
'        End If
    Else
        debugflg = 0
'        mnuDebug.Caption = "Turn Debugging ON"
'        checkAndKill NameProcess, False, False
    End If

   On Error GoTo 0
   Exit Sub

mnuDebug_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDebug_Click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click(ByRef Index As Integer)
    
    On Error GoTo mnuAbout_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAbout_Click"
     
     'MsgBox App.Major & ":" & App.Minor & ":" & App.Revision
     
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
' Procedure : mnuMoreIcons_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuMoreIcons_Click()
   On Error GoTo mnuMoreIcons_Click_Error
   If debugflg = 1 Then DebugPrint "%mnuMoreIcons_Click"

    Call btnGetMore_Click

   On Error GoTo 0
   Exit Sub

mnuMoreIcons_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuMoreIcons_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuLeft_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Click event from the menu option that slides a icon in the map one step to the left
'---------------------------------------------------------------------------------------
'
Private Sub menuLeft_Click()
    Dim storedFilename  As String: storedFilename = vbNullString
    Dim storedFileName2  As String: storedFileName2 = vbNullString
    Dim storedTitle  As String: storedTitle = vbNullString
    Dim storedCommand  As String: storedCommand = vbNullString
    Dim storedArguments  As String: storedArguments = vbNullString
    Dim storedWorkingDirectory  As String: storedWorkingDirectory = vbNullString
    Dim storedShowCmd  As String: storedShowCmd = vbNullString
    Dim storedOpenRunning  As String: storedOpenRunning = vbNullString
    Dim storedIsSeparator  As String: storedIsSeparator = vbNullString
    Dim storedUseContext  As String: storedUseContext = vbNullString
    Dim storedDockletFile  As String: storedDockletFile = vbNullString
    Dim storedUseDialog  As String: storedUseDialog = vbNullString
    Dim storedUseDialogAfter  As String: storedUseDialogAfter = vbNullString
    Dim storedQuickLaunch  As String: storedQuickLaunch = vbNullString
    Dim storedAutoHideDock  As String: storedAutoHideDock = vbNullString '.nn Added new check box to allow autohide of the dock after launch of the chosen app
    Dim storedSecondApp   As String: storedSecondApp = vbNullString  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
    Dim storedRunSecondAppBeforehand As String: storedRunSecondAppBeforehand = vbNullString
    Dim storedAppToTerminate As String: storedAppToTerminate = vbNullString
    Dim storedDisabled   As String: storedDisabled = vbNullString
    Dim storedRunElevated   As String: storedRunElevated = vbNullString

    ' take the current icon -1 and read its details and store it
    On Error GoTo menuLeft_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuLeft_Click"
    
    ' .82 DAEB 02/06/2022 rDIConConfig.frm Added check for moving right or left beyond the end of the RDMap.
    If rdIconNumber - 1 < 0 Then Exit Sub
    
    Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber - 1, interimSettingsFile)
        
    storedFilename = sFilename
    storedFileName2 = sFileName2
    storedTitle = sTitle
    storedCommand = sCommand
    storedArguments = sArguments
    storedWorkingDirectory = sWorkingDirectory
    storedShowCmd = sShowCmd
    storedOpenRunning = sOpenRunning
    storedRunElevated = sRunElevated
    storedIsSeparator = sIsSeparator
    storedUseContext = sUseContext
    storedDockletFile = sDockletFile
    
    ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    If defaultDock = 1 Then
        storedUseDialog = sUseDialog
        storedUseDialogAfter = sUseDialogAfter
        storedQuickLaunch = sQuickLaunch '.nn Added new check box to allow a quick launch of the chosen app
        storedAutoHideDock = sAutoHideDock   '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        storedSecondApp = sSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
        storedRunSecondAppBeforehand = sRunSecondAppBeforehand
        
        storedAppToTerminate = sAppToTerminate
        storedDisabled = sDisabled
    End If
    
    ' take the current icon details and write it into the place of the one to the left (-1)
    Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile)
    
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber - 1, interimSettingsFile)
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    sFilename = storedFilename
        
    sFileName2 = storedFileName2
    sTitle = storedTitle
    sCommand = storedCommand
    sArguments = storedArguments
    sWorkingDirectory = storedWorkingDirectory
    sShowCmd = storedShowCmd
    sOpenRunning = storedOpenRunning
    sRunElevated = storedRunElevated
    sIsSeparator = storedIsSeparator
    sUseContext = storedUseContext
    sDockletFile = storedDockletFile
    
    ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    If defaultDock = 1 Then
        sUseDialog = storedUseDialog
        sUseDialogAfter = storedUseDialogAfter
        sQuickLaunch = storedQuickLaunch '.nn Added new check box to allow a quick launch of the chosen app
        sAutoHideDock = storedAutoHideDock     '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        sSecondApp = storedSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
        
        sRunSecondAppBeforehand = storedRunSecondAppBeforehand
        
        sAppToTerminate = storedAppToTerminate
        sDisabled = storedDisabled
    End If
    
    ' take the stored icon details and write it into the current location
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile)
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), True, 32, True, False)
    Call displayIconElement(rdIconNumber - 1, picRdMap(rdIconNumber - 1), True, 32, True, False)
    
    btnSet.Enabled = False ' tell the program that nothing has changed
    btnClose.Visible = True
    btnCancel.Visible = False


   On Error GoTo 0
   Exit Sub

menuLeft_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuLeft_Click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : menuright_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Click event from the menu option that slides a icon in the map one step to the right
'---------------------------------------------------------------------------------------
'
Private Sub menuright_Click()
    Dim storedFilename  As String: storedFilename = vbNullString
    Dim storedFileName2  As String: storedFileName2 = vbNullString
    Dim storedTitle  As String: storedTitle = vbNullString
    Dim storedCommand  As String: storedCommand = vbNullString
    Dim storedArguments  As String: storedArguments = vbNullString
    Dim storedWorkingDirectory  As String: storedWorkingDirectory = vbNullString
    Dim storedShowCmd  As String: storedShowCmd = vbNullString
    Dim storedOpenRunning  As String: storedOpenRunning = vbNullString
    Dim storedIsSeparator  As String: storedIsSeparator = vbNullString
    Dim storedUseContext  As String: storedUseContext = vbNullString
    Dim storedDockletFile  As String: storedDockletFile = vbNullString
    Dim storedUseDialog  As String: storedUseDialog = vbNullString
    Dim storedUseDialogAfter  As String: storedUseDialogAfter = vbNullString
    Dim storedQuickLaunch  As String: storedQuickLaunch = vbNullString
    Dim storedAutoHideDock  As String: storedAutoHideDock = vbNullString '.nn Added new check box to allow autohide of the dock after launch of the chosen app
    Dim storedSecondApp   As String: storedSecondApp = vbNullString  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
    Dim storedRunSecondAppBeforehand As String: storedRunSecondAppBeforehand = vbNullString
    Dim storedAppToTerminate As String: storedAppToTerminate = vbNullString
    Dim storedDisabled   As String: storedDisabled = vbNullString
    Dim storedRunElevated   As String: storedRunElevated = vbNullString
     
    On Error GoTo menuright_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuright_Click"

    ' .82 DAEB 02/06/2022 rDIConConfig.frm Added check for moving right or left beyond the end of the RDMap.
    If rdIconNumber + 1 > rdIconMaximum Then Exit Sub
    
    ' take the current icon plus one and read its details and store it
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", rdIconNumber + 1, interimSettingsFile
            
    storedFilename = sFilename
    storedFileName2 = sFileName2
    storedTitle = sTitle
    storedCommand = sCommand
    storedArguments = sArguments
    storedWorkingDirectory = sWorkingDirectory
    storedShowCmd = sShowCmd
    storedOpenRunning = sOpenRunning
    storedRunElevated = sRunElevated
    storedIsSeparator = sIsSeparator
    storedUseContext = sUseContext
    storedDockletFile = sDockletFile
    ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    If defaultDock = 1 Then
        storedUseDialog = sUseDialog
        storedUseDialogAfter = sUseDialogAfter
        storedQuickLaunch = sQuickLaunch '.nn Added new check box to allow a quick launch of the chosen app
        storedAutoHideDock = sAutoHideDock       '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        storedSecondApp = sSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
        storedRunSecondAppBeforehand = sRunSecondAppBeforehand
        
        storedAppToTerminate = sAppToTerminate
        storedDisabled = sDisabled
    End If
    ' take the current icon details and write it into the place of the one to the right
    'readSettingsIni (rdIconNumber)
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile
    
    'writeSettingsIni (rdIconNumber + 1)
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber + 1, interimSettingsFile)
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile

    sFilename = storedFilename
    sFileName2 = storedFileName2
    sTitle = storedTitle
    sCommand = storedCommand
    sArguments = storedArguments
    sWorkingDirectory = storedWorkingDirectory
    sShowCmd = storedShowCmd
    sOpenRunning = storedOpenRunning
    sRunElevated = storedRunElevated
    sIsSeparator = storedIsSeparator
    sUseContext = storedUseContext
    sDockletFile = storedDockletFile
    ' .06 DAEB 31/01/2021 rdIconConfig.frm Added new checkbox to determine if a post initiation dialog should appear
    If defaultDock = 1 Then
        sUseDialog = storedUseDialog
        sUseDialogAfter = storedUseDialogAfter
        sQuickLaunch = storedQuickLaunch '.nn Added new check box to allow a quick launch of the chosen app
        sAutoHideDock = storedAutoHideDock      '.nn Added new check box to allow autohide of the dock after launch of the chosen app
        sSecondApp = storedSecondApp  ' .42 DAEB 21/05/2021 rdIconConfig.frm Added new field for second program to be run
        sRunSecondAppBeforehand = storedRunSecondAppBeforehand
        sAppToTerminate = storedAppToTerminate
        sDisabled = storedDisabled
    End If
    ' take the stored icon details and write it into the current location
    'writeSettingsIni (rdIconNumber)
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconNumber, interimSettingsFile)
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), True, 32, True, False)
    Call displayIconElement(rdIconNumber + 1, picRdMap(rdIconNumber + 1), True, 32, True, False)

    btnSet.Enabled = False ' tell the program that nothing has changed
    btnClose.Visible = True
    btnCancel.Visible = False


   On Error GoTo 0
   Exit Sub

menuright_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuright_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuAddSomething
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add something to the RD icon map, called by all the menuAdd functions that follow
'---------------------------------------------------------------------------------------
'
Private Sub menuAddSomething(ByVal thisFilename As String, ByVal thisTitle As String, _
    ByVal thisCommand As String, _
    ByVal thisArguments As String, ByVal thisWorkingDirectory As String, ByVal thisDocklet As String, _
    ByVal thIsSeparator As String, _
    Optional ByVal thisShowCmd As String, Optional ByVal thisOpenRunning As String, _
    Optional ByVal thisUseDialog As String, Optional ByVal thisUseDialogAfter As String, _
    Optional ByVal thisQuickLaunch As String, _
    Optional ByVal thisAutoHideDock As String, Optional ByVal thisSecondApp As String, Optional ByVal thisDisabled As String, Optional ByVal sRunElevated As String)

    Dim useloop As Integer: useloop = 0
    Dim thisIcon As Integer

    On Error GoTo menuAddSomething_Error
    If debugflg = 1 Then DebugPrint "%" & "menuAddSomething"
  
    Call busyStart


    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    ' starting at the end of the rocketdock map, step backward and decrement the number
    ' until we reach the current position.
    
    For useloop = rdIconMaximum To rdIconNumber Step -1
         
         Call zeroAllIconCharacteristics
         
         readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile
        
        ' and increment the identifier by one
         Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop + 1, interimSettingsFile)
         
    Next useloop
    
    'increment the new icon count
    theCount = theCount + 1
    
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, interimSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    rdIconMaximum = theCount - 1 '

    'set the slider bar to the new maximum
    rdMapHScroll.Max = theCount - 1

    ' .24 Add a form refresh after the menu has gone away to prevent messy control image leftovers
    rDIconConfigForm.Refresh

    ' test to see if the picturebox has already been created
    If CheckControlExists(picRdMap(rdIconMaximum)) Then
        'do nothing
    Else
        Load picRdMap(rdIconMaximum) ' dynamically extend the number of picture boxes by one
        picRdMap(rdIconMaximum).Width = 500
        picRdMap(rdIconMaximum).Height = 500
        picRdMap(rdIconMaximum).Left = picRdMap(rdIconMaximum - 1).Left + boxSpacing
        picRdMap(rdIconMaximum).Top = 30
        picRdMap(rdIconMaximum).Visible = True
    End If
    
    thisIcon = useloop + 1
    
    Call zeroAllIconCharacteristics

    'when we arrive at the original position then add a blank item
    ' with the following blank characteristics
    sFilename = thisFilename ' the default Rocketdock filename for a blank item
    
    sTitle = thisTitle
    sCommand = thisCommand
    sArguments = thisArguments
    sWorkingDirectory = thisWorkingDirectory
    sDockletFile = thisDocklet
    sIsSeparator = thIsSeparator
    
    sShowCmd = "1" ' .34 DAEB 05/05/2021 rDIConConfigForm.frm The value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
    
    'set the fields for this icon to the correct value as supplied
    txtLabelName.Text = sTitle
    
    If sDockletFile <> vbNullString Then
        txtTarget.Text = sDockletFile
    Else
        txtTarget.Text = sCommand
    End If
    
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory
    
    cmbRunState.ListIndex = 1 ' .34 DAEB 05/05/2021 rDIConConfigForm.frm sShowCmd value must be at least 1 to open a normal window and needs to be calculated from the dropdown value +1
    cmbOpenRunning.ListIndex = 0 ' "Use Global Setting"
    chkRunElevated.Value = 0
    
    'writeSettingsIni (thisIcon)
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", thisIcon, interimSettingsFile)
    
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    Call displayIconElement(thisIcon, picRdMap(thisIcon), True, 32, True, False)
    
    Call populateRdMap(0) ' regenerate the map from position zero
      
    btnSet.Enabled = False ' tell the program that nothing has changed
    btnClose.Visible = True
    btnCancel.Visible = False


    'Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1) ' click on the picture box
    ' .69 DAEB 16/05/2022 rDIConConfig.frm Moved the core left click code to a separate routine to avoid the clicks-via-code from activating a start drag
    Call btnSet_Click
    
    Call picRdMap_MouseDown_event(thisIcon)
    
    
    Call busyStop

   On Error GoTo 0
   Exit Sub

menuAddSomething_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuAddSomething of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuClone_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Menu option for the RD Map to clone the current item.
'---------------------------------------------------------------------------------------
'
Private Sub mnuClone_Click()
        
    On Error GoTo mnuClone_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuClone_Click"
      
    ' when we arrive at the original position then add a blank item
    ' with the following blank characteristics
    ' "\Icons\help.png" ' the default Rocketdock filename for a blank item
    
    ' general tool to clone an icon
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory
    Call menuAddSomething(sFilename, sTitle, sCommand, sArguments, sWorkingDirectory, sDockletFile, sIsSeparator)
    
   On Error GoTo 0
   Exit Sub

mnuClone_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClone_Click of Form rDIconConfigForm"
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : menuClearCache_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Menu option for the RD Map to add a blank picture item.
'---------------------------------------------------------------------------------------
'
Private Sub mnuClearCache_Click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
    On Error GoTo menuAdd_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuClearCache_Click"
      
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\recyclebin-full.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    ' thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Clear Cache", Environ$("windir") & "\System32\RUNDLL32.exe", "advapi32.dll , ProcessIdleTasks", vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

menuAdd_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClearCache_Click of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : menuAddBlank_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Menu option for the RD Map to add a blank picture item.
'---------------------------------------------------------------------------------------
'
Private Sub menuAddBlank_Click()
        
    On Error GoTo menuAdd_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuAddBlank_Click"
      
    ' when we arrive at the original position then add a blank item
    ' with the following blank characteristics
    ' "\Icons\help.png" ' the default Rocketdock filename for a blank item
    
    ' general tool to add an icon
    ' .57 DAEB 25/04/2022 rDIConConfig.frm blank.png should be blank and not ?
    Call menuAddSomething("blank.png", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    
   On Error GoTo 0
   Exit Sub

menuAdd_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuAdd_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddShutdown_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a shutdown icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddShutdown_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
   On Error GoTo mnuAddShutdown_click_Error
      If debugflg = 1 Then DebugPrint "%" & "mnuAddShutdown_click"
   
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\shutdown.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Shutdown", Environ$("windir") & "\System32\shutdown.exe", "/s /t 00 /f /i", vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddShutdown_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddShutdown_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddRestart_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a Restart icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'.101 DAEB 09/11/2022 rDIConConfig.frm Add the restart option.
Private Sub mnuAddRestart_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    
   On Error GoTo mnuAddRestart_click_Error
      If debugflg = 1 Then DebugPrint "%" & "mnuAddRestart_click"
   
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\Reboot.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Shutdown", Environ$("windir") & "\System32\shutdown.exe", "/r", vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddRestart_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddRestart_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddSleep_click
' Author    : beededea
' Date      : 17/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddSleep_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
   
    ' check the icon exists
    On Error GoTo mnuAddSleep_click_Error

    iconFileName = App.Path & "\my collection" & "\sleep.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\Icons\help.png"
    End If
           
    If FExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call menuAddSomething(iconImage, "Sleep", Environ$("windir") & "\System32\RUNDLL32.exe", "powrprof.dll,SetSuspendState 0,1,0", vbNullString, vbNullString, vbNullString)
    Else
        MsgBox "Unable to add sleep image as it does not exist"
    End If

    On Error GoTo 0
    Exit Sub

mnuAddSleep_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddSleep_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddLog_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddLog_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    On Error GoTo mnuAddLog_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddLog_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\console-green-screen-logout.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Log Out", Environ$("windir") & "\system32\shutdown.exe", "/l", "%windir%", vbNullString, vbNullString)

    On Error GoTo 0
    Exit Sub

mnuAddLog_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddLog_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddLock_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddLock_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    On Error GoTo mnuAddLock_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddLock_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\padlockLockWorkstation.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Lock Workstation", Environ$("windir") & "\system32\rundll32.exe", "user32.dll, LockWorkStation", "%windir%", vbNullString, vbNullString)

    On Error GoTo 0
    Exit Sub

mnuAddLock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddLock_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddNetwork_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a network icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddNetwork_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    On Error GoTo mnuAddNetwork_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddNetwork_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\big-globe(network).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    ' thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Network", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddNetwork_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddNetwork_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddWorkgroup_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a network icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddWorkgroup_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    On Error GoTo mnuAddWorkgroup_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddWorkgroup_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\big-globe(network).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Network", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddWorkgroup_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddWorkgroup_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddPrinters_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a PRINTERS icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddPrinters_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddPrinters_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddPrinters_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\printer.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Printers", "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddPrinters_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPrinters_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddTask_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a task manager icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddTask_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    ' check the icon exists
    On Error GoTo mnuAddTask_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddTask_click"
    
    iconFileName = App.Path & "\my collection" & "\task-manager(tskmgr).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    'Call menuAddSomething( iconImage, "Task Manager", "taskmgr", vbNullString, vbNullString, vbNullString, vbNullString)

    If Is64bit() Then
        ' if a 32 bit application on a 64bit o/s, regardless of the command, the o/s calls C:\Windows\SysWOW64\taskmgr.exe
        If FExists(iconImage) Then
            '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
            Call menuAddSomething(iconImage, "Task Manager", Environ$("windir") & "\SysWOW64\" & "taskmgr.exe", vbNullString, vbNullString, vbNullString, vbNullString)
        Else
            MsgBox "Unable to add Task Manager image as it has been deleted"
        End If
    Else
        ' if a 32 bit application on a 32bit o/s, regardless of the o/s calls C:\Windows\System32\taskmgr.exe
        If FExists(iconImage) Then
            '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
            Call menuAddSomething(iconImage, "Task Manager", Environ$("windir") & "\System32\" & "taskmgr.exe", vbNullString, vbNullString, vbNullString, vbNullString)
        Else
            MsgBox "Unable to add Task Manager image as it has been deleted"
        End If
    End If



   On Error GoTo 0
   Exit Sub

mnuAddTask_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddTask_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddControl_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a control panel icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddControl_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    ' check the icon exists
    On Error GoTo mnuAddControl_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddControl_click"

    iconFileName = App.Path & "\my collection" & "\control-panel(control).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Control panel", "control", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddControl_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddControl_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddPrograms_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddPrograms_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddPrograms_click_Error
       If debugflg = 1 Then DebugPrint "%" & "mnuAddPrograms_click"
    
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\programs and features.ico"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Programs and Features", "appwiz.cpl", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddPrograms_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPrograms_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddDiscMgmt_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDiscMgmt_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddDiscMgmt_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddDiscMgmt_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\discMgmt.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Disc Management", "diskmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddDiscMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDiscMgmt_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddDevMgmt_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDevMgmt_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddDevMgmt_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddDevMgmt_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Device Management", "devmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddDevMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDevMgmt_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddEventViewer_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddEventViewer_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddEventViewer_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddEventViewer_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\event-viewer(CEventVwr.msc).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Event Viewer", "eventvwr.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddEventViewer_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddEventViewer_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddPerfMon_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddPerfMon_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddPerfMon_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddPerfMon_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\perfmon.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Performance Monitor", "perfmon.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddPerfMon_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPerfMon_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddServices_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddServices_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddServices_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddServices_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Services Management", "services.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddServices_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddServices_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddTaskSched_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a programs and features icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddTaskSched_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddTaskSched_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddTaskSched_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\glass-clipboard.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Task Scheduler", "taskschd.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddTaskSched_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddTaskSched_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddDock_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a dock settings icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDock_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    ' check the icon exists
    On Error GoTo mnuAddDock_click_Error
      If debugflg = 1 Then DebugPrint "%" & "mnuAddDock_click"

    iconFileName = App.Path & "\my collection" & "\dock settings.ico"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Dock Settings", "[Settings]", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddDock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDock_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddAdministrative_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddAdministrative_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    ' check the icon exists
    On Error GoTo mnuAddAdministrative_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddAdministrative_click"

    iconFileName = App.Path & "\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Administration Tools", "compmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddAdministrative_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddAdministrative_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddRecycle_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddRecycle_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddRecycle_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddRecycle_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\recyclebin-full.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Recycle Bin", "::{645ff040-5081-101b-9f08-00aa002f954e}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddRecycle_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddRecycle_click of Form rDIconConfigForm"

End Sub




' .08 DAEB 02/02/2021 rDIconConfigForm.frm Added menu option to clear the cache
'---------------------------------------------------------------------------------------
' Procedure : mnuAddClearCache_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add the menu option of clearing the Windows cache
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddClearCache_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    On Error GoTo mnuAddClearCache_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddClearCache_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\my collection" & "\recyclebin-full.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Clear Cache", Environ$("windir") & "\system32\rundll32.exe", "advapi32.dll , ProcessIdleTasks", "%windir%", vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddClearCache_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddClearCache_click of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAddQuit_click
' Author    : beededea
' Date      : 19/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddQuit_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    ' check the icon exists
    On Error GoTo mnuAddQuit_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddQuit_click"
   
    iconFileName = App.Path & "\my collection" & "\quit.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Quit", "[Quit]", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddQuit_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddQuit_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddProgramFiles_click
' Author    : beededea
' Date      : 19/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddProgramFiles_click()
    Dim iconImage As String: iconImage = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString

    ' check the icon exists
    On Error GoTo mnuAddProgramFiles_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddProgramFiles_click"
   
    iconFileName = App.Path & "\my collection" & "\hard-drive-indicator-D.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSomething(iconImage, "Program Files", "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddProgramFiles_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddProgramFiles_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtSeparator_click
' Author    : beededea
' Date      : 27/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtSeparator_click()
       
   On Error GoTo mnuTrgtSeparator_click_Error
   If debugflg = 1 Then DebugPrint "mnuTrgtSeparator_click"
      
    sIsSeparator = "1"
        
    txtLabelName.Text = "Separator"
        
    ' set fields to blank
    txtCurrentIcon.Text = ""
    txtTarget.Text = ""
    txtArguments.Text = ""
    txtStartIn.Text = ""
    
    txtLabelName.Enabled = False
    txtCurrentIcon.Enabled = False
    txtTarget.Enabled = False
    btnTarget.Enabled = False
    txtArguments.Enabled = False
    txtStartIn.Enabled = False
    cmbRunState.Enabled = False
    cmbOpenRunning.Enabled = False
    chkRunElevated.Enabled = False
    btnSelectStart.Enabled = False

   On Error GoTo 0
   Exit Sub

mnuTrgtSeparator_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtFolder_click of Form rDIconConfigForm"
    
End Sub

' .08 DAEB 02/02/2021 rDIconConfigForm.frm Added menu option to clear the cache
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtClearCache_click
' Author    : beededea
' Date      : 27/09/2019
' Purpose   : clear the cache
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtClearCache_click()
    
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
   
    On Error GoTo mnuTrgtClearCache_click_Error
    If debugflg = 1 Then DebugPrint "%mnuTrgtClearCache_click"

    sCommand = Environ$("windir") & "\System32\RUNDLL32.exe"
    sArguments = "advapi32.dll , ProcessIdleTasks"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments


   On Error GoTo 0
   Exit Sub

mnuTrgtClearCache_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtClearCache_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtFolder_click
' Author    : beededea
' Date      : 27/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtFolder_click()
    
    Dim getFolder As String: getFolder = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
   
   On Error GoTo mnuTrgtFolder_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtFolder_click"

    If txtTarget.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtTarget.Text 'start dir, might be "C:\" or so also
        Else
            If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
                dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
            Else
                dialogInitDir = sdAppPath 'start dir, might be "C:\" or so also
            End If
        End If
    End If

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder
    'getFolder = ChooseDir_Click ' old method to show the dialog box to select a folder
    If getFolder <> vbNullString Then txtTarget.Text = getFolder

    sCommand = getFolder

   On Error GoTo 0
   Exit Sub

mnuTrgtFolder_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtFolder_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtShutdown_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtShutdown_click()

   On Error GoTo mnuTrgtShutdown_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtShutdown_click"

    sCommand = Environ$("windir") & "\System32\shutdown.exe"
    sArguments = "/s /t 00 /f /i"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments

   On Error GoTo 0
   Exit Sub

mnuTrgtShutdown_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtShutdown_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtRestart_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'.101 DAEB 09/11/2022 rDIConConfig.frm Add the restart option.
Private Sub mnuTrgtRestart_click()

   On Error GoTo mnuTrgtRestart_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtRestart_click"

    sCommand = Environ$("windir") & "\System32\shutdown.exe"
    sArguments = "/r"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments

   On Error GoTo 0
   Exit Sub

mnuTrgtRestart_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtRestart_click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtSleep_click
' Author    : beededea
' Date      : 17/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtSleep_click()

    On Error GoTo mnuTrgtSleep_click_Error

    sCommand = Environ$("windir") & "\System32\RUNDLL32.exe"
    sArguments = "powrprof.dll,SetSuspendState 0,1,0"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments

    On Error GoTo 0
    Exit Sub

mnuTrgtSleep_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtSleep_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtEnhanced_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtEnhanced_click()
   On Error GoTo mnuTrgtEnhanced_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtEnhanced_click"

    sCommand = App.Path & "\iconsettings.exe" ' 17/11/2020    .04 DAEB Replaced all occurrences of rocket1.exe with iconsettings.exe

    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtEnhanced_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtEnhanced_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtLog_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtLog_click()

   On Error GoTo mnuTrgtLog_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtLog_click"

    sCommand = Environ$("windir") & "\System32\shutdown.exe"
    sArguments = "/l"
    sWorkingDirectory = "%windir%"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory

   On Error GoTo 0
   Exit Sub

mnuTrgtLog_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtLog_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtLock_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtLock_click()

   On Error GoTo mnuTrgtLock_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtLock_click"

    sCommand = Environ$("windir") & "\System32\rundll32.exe"
    sArguments = "user32.dll, LockWorkStation"
    sWorkingDirectory = "%windir%"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory

   On Error GoTo 0
   Exit Sub

mnuTrgtLock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtLock_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtWorkgroup_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtWorkgroup_click()

   On Error GoTo mnuTrgtWorkgroup_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtWorkgroup_click"

    sCommand = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtWorkgroup_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtWorkgroup_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtNetwork_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtNetwork_click()

   On Error GoTo mnuTrgtNetwork_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtNetwork_click"

    sCommand = "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtNetwork_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtNetwork_click of Form rDIconConfigForm"
    
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtMyComputer_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtMyComputer_click()
 
   On Error GoTo mnuTrgtMyComputer_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtMyComputer_click"

    sCommand = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyComputer_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyComputer_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtMyDocuments_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : .25 DAEB 07/03/2021 rdIconConfig.frm Added menu option to add a "my Documents" target utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtMyDocuments_click()
 
   On Error GoTo mnuTrgtMyDocuments_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtMyDocuments_click"

    sCommand = "::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyDocuments_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyDocuments_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtMyMusic_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : .26 DAEB 07/03/2021 rdIconConfig.frm Added menu option to add a "my Music" target utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtMyMusic_click()
    Dim userprof As String: userprof = vbNullString
    
    ' initialise the vars above
    
    userprof = ""
    
   On Error GoTo mnuTrgtMyMusic_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtMyMusic_click"

    userprof = Environ$("USERPROFILE")

    sCommand = userprof & "\Documents\Music"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyMusic_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyMusic_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtMyPictures_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : .27 DAEB 07/03/2021 rdIconConfig.frm Added menu option to add a "my Pictures" target utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtMyPictures_click()
    Dim userprof As String: userprof = vbNullString
    
    ' initialise the vars above
    
    userprof = ""
    
   On Error GoTo mnuTrgtMyPictures_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtMyPictures_click"

    userprof = Environ$("USERPROFILE")

    sCommand = userprof & "\Documents\Pictures"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyPictures_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyPictures_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtMyVideos_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : .28 DAEB 07/03/2021 rdIconConfig.frm Added menu option to add a "my Videos" target utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtMyVideos_click()
    Dim userprof As String: userprof = vbNullString
    
    ' initialise the vars above
    
    userprof = ""
    
   On Error GoTo mnuTrgtMyVideos_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtMyVideos_click"

    userprof = Environ$("USERPROFILE")

    sCommand = userprof & "\Documents\Videos"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyVideos_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyVideos_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtPrinters_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtPrinters_click()
   On Error GoTo mnuTrgtPrinters_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtPrinters_click"

    sCommand = "::{2227a280-3aea-1069-a2de-08002b30309d}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtPrinters_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtPrinters_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtTask_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtTask_click()
   On Error GoTo mnuTrgtTask_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtTask_click"

    sCommand = "taskmgr"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtTask_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtTask_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtControl_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtControl_click()
   On Error GoTo mnuTrgtControl_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtControl_click"

    sCommand = "control"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtControl_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtControl_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtProgramFiles_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtProgramFiles_click()
   On Error GoTo mnuTrgtProgramFiles_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtProgramFiles_click"

    sCommand = "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtProgramFiles_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtProgramFiles_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtPrograms_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtPrograms_click()
   On Error GoTo mnuTrgtPrograms_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtPrograms_click"

    sCommand = "appwiz.cpl"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtPrograms_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtPrograms_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtRecycle_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtRecycle_click()
   On Error GoTo mnuTrgtRecycle_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtRecycle_click"

    sCommand = "::{645ff040-5081-101b-9f08-00aa002f954e}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtRecycle_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtRecycle_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtDocklet_click
' Author    : beededea
' Date      : 27/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtDocklet_click()
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retFileName As String: retFileName = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    
    Const x_MaxBuffer = 256
    ' set the default folder to the docklet folder under rocketdock
   On Error GoTo mnuTrgtDocklet_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtDocklet_click"
   
    If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
        dialogInitDir = rdAppPath & "\docklets"
    Else
        dialogInitDir = sdAppPath & "\docklets"
    End If
 
    With x_OpenFilename
    '    .hwndOwner = Me.hWnd
      .hInstance = App.hInstance
      .lpstrTitle = "Select a Rocketdock Docklet DLL"
      .lpstrInitialDir = dialogInitDir
      
      .lpstrFilter = "DLL Files" & vbNullChar & "*.dll" & vbNullChar & vbNullChar
      .nFilterIndex = 2
      
      .lpstrFile = String$(x_MaxBuffer, 0)
      .nMaxFile = x_MaxBuffer - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = x_MaxBuffer - 1
      .lStructSize = Len(x_OpenFilename)
    End With
      

    Call getFileNameAndTitle(retFileName, retfileTitle)
    txtTarget.Text = retFileName
    'txtLabelName.Text = retfileTitle
    
    sDockletFile = txtTarget.Text
    
    'it chooses the icon here as with a docklet no alternative icon is allowed, the docklet determines that
    If InStr(getFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
        If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
            txtCurrentIcon.Text = rdAppPath & "\icons\clock.png"
        Else
            txtCurrentIcon.Text = sdAppPath & "\icons\clock.png"
        End If
    ElseIf InStr(getFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
      txtCurrentIcon.Text = App.Path & "\my collection\recyclebin-full.png"
    Else
        If defaultDock = 0 Then ' ' .19 DAEB 01/03/2021 rDIConConfigForm.frm Separated the Rocketdock/Steamydock specific actions
            txtCurrentIcon.Text = rdAppPath & "\icons\blank.png" ' has to be an icon of some sort
        Else
            txtCurrentIcon.Text = sdAppPath & "\blank.png" ' has to be an icon of some sort
        End If
    End If
    
    triggerRdMapRefresh = True

   On Error GoTo 0
   Exit Sub

mnuTrgtDocklet_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtDocklet_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtDock_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtDock_click()
   On Error GoTo mnuTrgtDock_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtDock_click"

    sCommand = "[Settings]"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtDock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtDock_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtRocketdock_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtRocketdock_click()
   On Error GoTo mnuTrgtRocketdock_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtRocketdock_click"

    sCommand = "[Quit]"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtRocketdock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtRocketdock_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckControlExists
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Function to see whether a control exists - no .NET equivalent
'---------------------------------------------------------------------------------------
'
Public Function CheckControlExists(ByRef ctl As Object) As Boolean
   On Error GoTo CheckControlExists_Error
    If debugflg = 1 Then DebugPrint "%" & "CheckControlExists"
   
    CheckControlExists = (VarType(ctl) <> vbObject)

   On Error GoTo 0
   Exit Function

CheckControlExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckControlExists of Form rDIconConfigForm"
End Function

'---------------------------------------------------------------------------------------
' Procedure : mnuDelete_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : a menu item for the RD Map allowing deletion of an item.
'               in the rdSettings.ini file read the mnu items from the next item
'               and decrement the identifier by one
'---------------------------------------------------------------------------------------
'
Private Sub mnuDelete_Click()
    On Error GoTo mnuDelete_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuDelete_Click"
    
    Call deleteRdMapPosition(rdIconNumber)
    
   On Error GoTo 0
   Exit Sub

mnuDelete_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDelete_Click of Form rDIconConfigForm"
    
End Sub

Private Sub deleteRdMapPosition(ByVal thisIconNumber As Integer, Optional confirmDialog As Boolean = True)
    Dim useloop As Integer: useloop = 0
    Dim thisIcon As Integer: thisIcon = 0
    Dim notQuiteTheTop As Integer: notQuiteTheTop = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
        
    If confirmDialog = True Then
        If chkToggleDialogs.Value = 1 Then
            answer = msgBoxA("This will delete the currently selected entry in the Rocketdock map, " & vbCr & txtCurrentIcon & "   -  are you sure?", vbQuestion + vbYesNo, "Deleting from the Map", True, "deleteRdMapPosition")
            If answer = vbNo Then
                Exit Sub
            End If
            Refresh
        End If
        
        If thisIconNumber = 0 And rdIconMaximum = 1 Then
            MsgBox "Cannot currently delete the last icon, one icon must always be present for the dock to operate - apologies.", vbInformation + vbYesNo
            Exit Sub
        End If
    End If
    
    Call busyStart
    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    If thisIconNumber < rdIconMaximum Then 'if not the top icon loop through them all and reassign the values
        notQuiteTheTop = rdIconMaximum - 1
        For useloop = thisIconNumber To notQuiteTheTop
            
            ' read the rocketdock alternative rdsettings.ini one item up in the list
            'readSettingsIni (useloop + 1) ' the alternative rdsettings.ini only exists when RD is set to use it
            
            readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop + 1, interimSettingsFile
            
            'write the the new item at the current location effectively overwriting it
            'writeSettingsIni (useloop)
            Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile)
        Next useloop
    End If
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    ' to tidy up we need to overwrite the final data from the rdsettings.ini, we will write sweet nothings to it
    removeSettingsIni (rdIconMaximum)
    
    'clear the icon
    picRdMap(rdIconMaximum).BackColor = &H8000000F
'    Set picRdMap(rdIconMaximum).Picture = LoadPicture(vbNullString)
    Set picRdMap(rdIconMaximum).Picture = Nothing
    Unload picRdMap(rdIconMaximum)
    
    ' the picbox positioning
    storeLeft = storeLeft - boxSpacing
        
    'decrement the icon count and the maximum icon
    theCount = theCount - 1
    rdIconMaximum = theCount - 1
    
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, interimSettingsFile
    
    'set the slider bar to the new maximum
    rdMapHScroll.Max = theCount - 1

    If thisIconNumber > rdIconMaximum Then rdIconNumber = rdIconMaximum
    thisIcon = rdIconNumber
    
    ' load the new icon as an image in the picturebox
    Call displayIconElement(thisIcon, picRdMap(thisIcon), True, 32, True, False)
    
    Call populateRdMap(0) ' regenerate the map from position zero
    
    btnSet.Enabled = False ' tell the program that nothing has changed
    btnClose.Visible = True
    btnCancel.Visible = False


    ' emulate a click on the appropriate icon in the map so that the image and details are shown
    'Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1)
   
    Call busyStop
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(ByRef Index As Integer)
    Dim answer As VbMsgBoxResult: answer = vbNo

    On Error GoTo mnuCoffee_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuCoffee_Click"
    
    answer = msgBoxA(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbQuestion + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form quartermaster"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : menuAddToDock
' Author    : beededea
' Date      : 25/10/2019
' Purpose   : right click and add to dock
'---------------------------------------------------------------------------------------
'
Private Sub menuAddToDock_click()

   On Error GoTo menuAddToDock_Error
   If debugflg = 1 Then DebugPrint "%menuAddToDock"

    Call btnAdd_Click

   On Error GoTo 0
   Exit Sub

menuAddToDock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuAddToDock of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuSmallerIcons_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : right click on the thumbnails moves the picboxes to a more appropriate position
'             for displaying the smaller thumbnail icons
'---------------------------------------------------------------------------------------
'
Private Sub menuSmallerIcons_Click()
    
    ' set the icon size
    
    ' the labels for the smaller thumbnail icon view
    On Error GoTo menuSmallerIcons_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuSmallerIcons_Click"
   
    If thumbImageSize = 64 Then ' change to 32
        thumbImageSize = 32
    End If
    
    ' .54 DAEB 25/04/2022 rDIConConfig.frm Added rDThumbImageSize saved variable to allow the tool to open the thumbnail explorer in small or large mode
    rDThumbImageSize = Str$(thumbImageSize)
    PutINISetting "Software\SteamyDockSettings", "thumbImageSize", rDThumbImageSize, toolSettingsFile
    
    removeThumbHighlighting
    
    imlThumbnailCache.ListImages.Clear
    
    'then populate them and refresh
    Call btnRefresh_Click_Event
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
   On Error GoTo 0
   Exit Sub

menuSmallerIcons_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuSmallerIcons_Click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : menuLargerThumbs_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : right click on the thumbnails moves the picboxes back to their original position
'---------------------------------------------------------------------------------------
'
Private Sub menuLargerThumbs_Click()
    'Dim useloop As Integer
    'Dim tooltip As String
    'Dim suffix As String: suffix = vbnullstring
    
    On Error GoTo menuLargerThumbs_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuLargerThumbs_Click"
    
    If thumbImageSize = 32 Then
        thumbImageSize = 64
    End If
      
    ' .54 DAEB 25/04/2022 rDIConConfig.frm Added rDThumbImageSize saved variable to allow the tool to open the thumbnail explorer in small or large mode
    rDThumbImageSize = Str$(thumbImageSize)
    PutINISetting "Software\SteamyDockSettings", "thumbImageSize", rDThumbImageSize, toolSettingsFile

    imlThumbnailCache.ListImages.Clear

    'then populate them, perhaps refresh?
    Call btnRefresh_Click_Event
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
   On Error GoTo 0
   Exit Sub

menuLargerThumbs_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuLargerThumbs_Click of Form rDIconConfigForm"

End Sub

'Private Sub chkBiLinear_Click()
'
'    If chkBiLinear.Value = 1 Then
'        If chkBiLinear.Tag = vbNullString Then
'            If cImage.isGDIplusEnabled = False Then
'                If Not 0 Then
'                    chkBiLinear.Tag = "noMsg"
'                    On Error Resume Next
'                    DebugPrint 1 / 0
'                    If Err Then ' uncompiled
'                        Err.Clear
'                        MsgBox "Non-GDI+ rotation with bilinear interpolation is painfully slow in IDE." & vbCrLf & _
'                            "But is acceptable when the routines are compiled", vbInformation + vbOKOnly
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    cImage.HighQualityInterpolation = chkBiLinear.Value
'    Call refreshPicBox(picPreview, 256)
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : What to do when unloading the main form
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(ByRef Cancel As Integer)
    Dim ofrm As Form
    Dim NameProcess As String: NameProcess = ""
    Dim fcount As Integer: fcount = 0
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo Form_Unload_Error
    
    ' save the current X and y position of this form to allow repositioning when restarting
    rDIconConfigFormXPosTwips = rDIconConfigForm.Left
    rDIconConfigFormYPosTwips = rDIconConfigForm.Top
    
    ' now write those params to the toolSettings.ini
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormXPos", rDIconConfigFormXPosTwips, toolSettingsFile
    PutINISetting "Software\SteamyDockSettings", "IconConfigFormYPos", rDIconConfigFormYPosTwips, toolSettingsFile
    
    Call DestroyToolTip ' destroys any current tooltip
    
    ' ANY controls loaded at runtime, MUST be Unloaded when close the form.
    For useloop = 1 To rdIconMaximum
        If CheckControlExists(picRdMap(useloop)) Then
            Unload picRdMap(useloop)
        End If
    Next useloop
    
'    For useloop = 1 To 11
'        If CheckControlExists(picFraPicThumbIcon(useloop)) Then
'            Unload picFraPicThumbIcon(useloop)
'        End If
'    Next useloop
'    For useloop = 1 To 11
'        If CheckControlExists(picThumbIcon(useloop)) Then
'            Unload picThumbIcon(useloop)
'        End If
'    Next useloop
'    For useloop = 1 To 11
'        If CheckControlExists(fraThumbLabel(useloop)) Then
'            Unload fraThumbLabel(useloop)
'        End If
'    Next useloop
'    For useloop = 1 To 11
'        If CheckControlExists(lblThumbName(useloop)) Then
'            Unload lblThumbName(useloop)
'        End If
'    Next useloop

    
    ' when you create a token to be shared, you must
    ' destroy it in the Unload or Terminate event
    ' and also reset gdiToken property for each existing class

    'If debugflg = 1 Then DebugPrint "%" & "Form_Unload"
    
    NameProcess = "PersistentDebugPrint.exe"
    
    If debugflg = 1 Then
        checkAndKill NameProcess, False, False
    End If
        
    If m_GDItoken Then
        If Not cShadow Is Nothing Then cShadow.gdiToken = 0&
        If Not cImage Is Nothing Then
            cImage.gdiToken = 0&
            cImage.DestroyGDIplusToken m_GDItoken
        End If
    End If
    
    'this was initially commented out as it caused a crash on exit in Win 7 (only) subsequent to the two Krool's
    'controls being added or perhaps it was the failure to close GDI properly
    'then I added it back in as an END is the wrong thing to do supposedly - but I do like a good END.
    
    For Each ofrm In Forms
        'fcount = fcount + 1
        'Sleep 250
        'MsgBox ("Unloading form " & fcount)
        Unload ofrm
        
    Next
    
'    Sleep 5000
'    MsgBox ("END " & fcount)
    
    'End ' on 32bit Windows this causes a crash and untidy exit so removed.

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuSubOpts_Click
' Author    : beededea
' Date      : 11/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub mnuSubOpts_Click(ByRef Index As Integer)
'
'    ' The 1st two options will be disabled if you do not have GDI+ installed
'
'    On Error GoTo mnuSubOpts_Click_Error
'    If debugflg = 1 Then DebugPrint "%" & "mnuSubOpts_Click"
'
'    Select Case Index
'    Case 0: ' do not use GDI+
'        If mnuSubOpts(Index).Checked = True Then Exit Sub
'        cImage.isGDIplusEnabled = False
'        mnuSubOpts(0).Checked = Not mnuSubOpts(0).Checked
'        mnuSubOpts(1).Checked = False
'
'        If m_GDItoken Then  ' when using token, we'll clean up here
'            cImage.DestroyGDIplusToken m_GDItoken
'            m_GDItoken = 0&
'            cImage.gdiToken = m_GDItoken ' reset the token
'            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
'        End If
'
'        Call refreshPicBox(picPreview, 256)
'
'    Case 1: ' always usge GDI+.
'        If mnuSubOpts(Index).Checked = True Then Exit Sub
'        mnuSubOpts(0).Checked = False ' remove checkmark on "Don't Use GDI+"
'        mnuSubOpts(1).Checked = True  ' show using GDI+
'        cImage.isGDIplusEnabled = True
'        ' verify it enabled correct and get a token to share
'        If cImage.isGDIplusEnabled Then
'            m_GDItoken = cImage.CreateGDIplusToken()
'            cImage.gdiToken = m_GDItoken
'            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
'        End If
'        ' tell GDI+ that we want high quality interpolation
'        If chkBiLinear.Value = 0 Then chkBiLinear.Value = 1 Else Call refreshPicBox(picPreview, 256)
'
'    Case 7: ' save as
'
'    End Select
''ExitRoutine:
'
'   On Error GoTo 0
'   Exit Sub
'
'mnuSubOpts_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSubOpts_Click of Form rDIconConfigForm"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : refreshPicBox
' Author    : beededea
' Date      : 14/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub refreshPicBox(ByRef picBox As PictureBox, ByVal iconSizing As Integer)

    Dim newWidth As Long: newWidth = 0
    Dim newHeight As Long: newHeight = 0
    
    Dim mirrorOffsetX As Long: mirrorOffsetX = 0
    Dim mirrorOffsetY As Long: mirrorOffsetY = 0
    
    Dim X As Long: X = 0
    Dim Y As Long: Y = 0
    
    Dim ShadowOffset As Long: ShadowOffset = 0
    Dim LightAdjustment As Single
        
    On Error GoTo refreshPicBox_Error
    If debugflg = 1 Then DebugPrint "%" & "refreshPicBox"

    mirrorOffsetX = 1
    mirrorOffsetY = 1

    newWidth = iconSizing: newHeight = iconSizing
    
    X = (picBox.ScaleWidth - newWidth) \ 2
    Y = (picBox.ScaleHeight - newHeight) \ 2
    
    picBox.Cls
    If Not cShadow Is Nothing Then
        picBox.CurrentX = 20
        picBox.CurrentY = 5
        picBox.Print "See c32bppDIB.CreateDropShadow for more ": picBox.CurrentX = 20
        picBox.Print "Color, Opacity, Blur Effect, ": picBox.CurrentX = 20
        picBox.Print "  and X,Y Position are adjustable"
    End If
    
    
    ' Generally, when rotating and/or resizing, it is easier to calculate the center of where you want the image rotated vs
    '   calculating the top/left coordinate of the resized, rotated image.  The last parameter of the Render call (CenterOnDestXY)
    '   will render around that center point if that paremeter is set.  So, what about when an image is not rotated? The Render
    '   function will still draw around that center point if the parameter is true. Or render, starting at the passed
    '   DestX,DestY coordinates if that parameter is false.
    
    ' The Render call only has one required parameter.  All others are optional and defaulted as follows
        ' srcX, srcY, destX, destY defaults are zero
        ' srcWidth, destWidth defaults are the image's width
        ' srcHeight, destHeight defaults are the image's height
        ' Opacity (Global Alpha) default is 100% opaque, pixel LigthAdjustmnet default is zero (no additional adjustment)
        ' GrayScale default is not grayscaled
        ' Rotation angle is at zero degrees
        ' Rendering image around a center point is false
    
    ' the cboAngle entries are at 15 degree intervals, so we simply multiply ListIndex by 15
    
    If Not cShadow Is Nothing Then
        ' the 55 below is the shadow's opacity; hardcoded here but can be modified to your heart's delight
        cShadow.Render picBox.hdc, X + newWidth \ 2 + ShadowOffset, Y + newHeight \ 2 + ShadowOffset, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
            55, , , , , LightAdjustment, 0, True
    End If
    
    Dim ttemp As Integer
    ttemp = -1
    
    cImage.Render picBox.hdc, X + newWidth \ 2, Y + newHeight \ 2, newWidth * 1, newHeight * 1, , , , , _
        100, , , , -1, 0, 0, True
    
    picBox.Refresh

   On Error GoTo 0
   Exit Sub

refreshPicBox_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure refreshPicBox of Form rDIconConfigForm"

End Sub
''---------------------------------------------------------------------------------------
'' Procedure : vScrollThumbs_KeyDown
'' Author    : beededea
'' Date      : 14/07/2019
'' Purpose   : key press whilst the vertical scroll bar is in focus
''---------------------------------------------------------------------------------------
''
'Private Sub vScrollThumbs_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
'   On Error GoTo vScrollThumbs_KeyDown_Error
'      If debugflg = 1 Then DebugPrint "%" & "vScrollThumbs_KeyDown"
'
'
'
'    Call getKeyPress(KeyCode)
'
'   On Error GoTo 0
'   Exit Sub
'
'vScrollThumbs_KeyDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vScrollThumbs_KeyDown of Form rDIconConfigForm"
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : LoadFileToTB
'' Author    : beededea
'' Date      : 26/08/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function LoadFileToTB(TxtBox As TextBox, FilePath As _
'   String, Optional Append As Boolean = False) As Boolean
'
'    'PURPOSE: Loads file specified by FilePath into textcontrol
'    '(e.g., Text Box, Rich Text Box) specified by TxtBox
'
'    'If Append = true, then loaded text is appended to existing
'    ' contents else existing contents are overwritten
'
'    'Returns: True if Successful, false otherwise
'
'    Dim iFile As Integer
'    Dim s As String
'
'   On Error GoTo LoadFileToTB_Error
'   If debugflg = 1 Then DebugPrint "%" & "LoadFileToTB"
'
'    If Dir(FilePath) = "" Then Exit Function
'
'    On Error GoTo ErrorHandler:
'    s = TxtBox.Text
'
'    iFile = FreeFile
'    Open FilePath For Input As #iFile
'    s = Input(LOF(iFile), #iFile)
'    If Append Then
'        TxtBox.Text = TxtBox.Text & s
'    Else
'        TxtBox.Text = s
'    End If
'
'    LoadFileToTB = True
'
'ErrorHandler:
'    If iFile > 0 Then Close #iFile
'
'   On Error GoTo 0
'   Exit Function
'
'LoadFileToTB_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function LoadFileToTB of Form rDIconConfigForm"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : btnTarget_MouseDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTarget_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo btnTarget_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%btnTarget_MouseDown"

    If Button = 2 Then
        'Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
        Me.PopupMenu mnuTrgtMenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

btnTarget_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTarget_MouseDown of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuLight_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_click()
   On Error GoTo mnuLight_click_Error
   If debugflg = 1 Then DebugPrint "%mnuLight_click"
    
    'MsgBox "Auto Theme Selection Manually Disabled"
    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    rDSkinTheme = "light" ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
    
    Call setThemeShade(Me, 240, 240, 240)
    
   On Error GoTo 0
   Exit Sub

mnuLight_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_click of Form rDIconConfigForm"
End Sub
''---------------------------------------------------------------------------------------
'' Procedure : setThemeLight
'' Author    : beededea
'' Date      : 26/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub setThemeLight()
'
'    Dim a As Long
'    Dim Ctrl As Control
'
'    On Error GoTo setThemeLight_Error
'    If debugflg = 1 Then DebugPrint "%setThemeLight"
'
'    classicTheme = False
'    mnuLight.Checked = True
'    mnuDark.Checked = False
'
'    ' custom button pictures that need to be skinned according to the theme
'    btnArrowDown.Picture = LoadPicture(App.path & "\arrowDown10.gif")
'    btnMapPrev.Picture = LoadPicture(App.path & "\leftArrow10.jpg")
'    btnMapNext.Picture = LoadPicture(App.path & "\rightArrow10.jpg")
'    btnArrowUp.Picture = LoadPicture(App.path & "\arrowUp10.jpg")
'
'    ' RGB(240, 240, 240) is the background colour used by the lighter themes
'
'    Me.BackColor = RGB(240, 240, 240)
'    ' a method of looping through all the controls that require reversion of any background colouring
'    For Each Ctrl In rDIconConfigForm.Controls
'        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
'          Ctrl.BackColor = RGB(240, 240, 240)
'        End If
'    Next
'
'    ' these elements are normal elements that should have their styling reverted,
'    ' the loop above changes the background colour and we don't want that for all items
'
'    'all other buttons go here
'
'    picPreview.BackColor = RGB(240, 240, 240)
'    picRdThumbFrame.BackColor = RGB(240, 240, 240)
'    btnRemoveFolder.BackColor = RGB(240, 240, 240)
'    picCover.BackColor = RGB(240, 240, 240)
'    back.BackColor = RGB(240, 240, 240)
'    sliPreviewSize.BackColor = RGB(240, 240, 240)
'
'    ' on NT6 plus using the MSCOMCTL slider with the lighter default theme, the slider
'    ' fails to pick up the new theme colour fully
'    ' the following lines triggers a partial colour change on the treeview that has no backcolor property
'    ' this also causes a refresh of the preview pane - so don't remove it.
'    ' I will have to create a new slider to overcome this - not yet tested the VB.NET version
'    ' do not remove - essential
'
'    'a = sliPreviewSize.Value
'    'sliPreviewSize.Value = 1
'    'sliPreviewSize.Value = a
'
'    ' the slider has a redrawing problem after changing the theme
'
'    ' the above no longer required with Krool's replacement controls
'
'   On Error GoTo 0
'   Exit Sub
'
'setThemeLight_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeLight of Form rDIconConfigForm"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_click()
   On Error GoTo mnuDark_click_Error
   If debugflg = 1 Then DebugPrint "%mnuDark_click"
    
    'MsgBox "Auto Theme Selection Manually Disabled"
    mnuAuto.Caption = "Auto Theme Enable"
    themeTimer.Enabled = False
    
    rDSkinTheme = "dark" ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
        
    Call setThemeShade(Me, 212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_click of Form rDIconConfigForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAuto
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuAuto_click()
   On Error GoTo mnuAuto_Error
   If debugflg = 1 Then DebugPrint "%mnuAuto"

    ' set the menu checks
    
    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Enable"
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Disable"
            themeTimer.Enabled = True
            Call setThemeColour(Me)
    End If
    
   On Error GoTo 0
   Exit Sub

mnuAuto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuRun_click
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : equivalent to SteamyDock's runcommand function
'---------------------------------------------------------------------------------------
' .32 DAEB 11/04/2021 rDIConConfigForm.frm changed all occurrences of txtTarget.Text to thisCommand to attain more compatibility with runcommand
'
Private Sub menuRun_click()
    Dim testURL As String: testURL = vbNullString
    Dim validURL As Boolean: validURL = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim folderPath As String: folderPath = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    Dim rmessage As String: rmessage = vbNullString ' .15 DAEB 01/03/2021 rDIConConfigForm.frm added confirmation dialog prior to running the test command
    Dim userLevel As String: userLevel = vbNullString
    Dim userprof As String: userprof = vbNullString
    Dim intShowCmd As Integer: intShowCmd = 0
    Dim arrStr() As String 'cannot initialise arrays in VB6
    Dim strCnt As Integer: strCnt = 0
    Dim suffix As String: suffix = vbNullString
    Dim listOfTypes As String: listOfTypes = vbNullString
    Dim useloop As Integer: useloop = 0
    
    userLevel = "open"
    
    On Error GoTo menuRun_click_Error
    If debugflg = 1 Then DebugPrint "%menuRun_click"
    
    intShowCmd = Val(sShowCmd)
    
    lblRdIconNumber.Caption = Str$(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str$(rdIconNumber) + 1
    Call displayIconElement(rdIconNumber, picPreview, True, icoSizePreset, True, False)
    
    ' we signify that all changes have been lost so the "save this icon" will not appear when switching icons
    btnSet.Enabled = False ' this has to be done at the end
    btnClose.Visible = True
    btnCancel.Visible = False

    
    thisCommand = txtTarget.Text
    
    If chkConfirmDialog.Value = 1 Then
        ' .15 DAEB 01/03/2021 rDIConConfigForm.frm added confirmation dialog prior to running the test command
        rmessage = "Are you sure you wish to run the following command - " & txtLabelName.Text & "?" & vbCr & thisCommand
        If txtArguments.Text <> "" Then rmessage = rmessage & " " & txtArguments.Text
        answer = msgBoxA(rmessage, vbExclamation + vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    If sIsSeparator = "1" Then
        MsgBox "That is a separator, you can't test it!"
        Exit Sub
    End If
    
    'If userLevel = vbNullString Then userLevel = "open"
    
    'now deal with the special extras
    ' contains "shutdown"
    If InStr(thisCommand, "shutdown.exe") <> 0 Then
        MsgBox "I am sure you don't really want me to shutdown... test cancelled."
        Exit Sub
    End If
        
    ' is the target a URL?
    testURL = Left$(thisCommand, 3)
    If testURL = "htt" Or testURL = "www" Then
        validURL = True
        Call executeCommand("Open", thisCommand, vbNullString, vbNullString, intShowCmd) 'change to call new function as part of .16
    End If

    ' control panel
    If thisCommand = "control" Then
        Call shellCommand("rundll32.exe shell32.dll,Control_RunDLL", intShowCmd) 'change to call new function as part of .16
        Exit Sub
    End If
    ' RD quit
    If thisCommand = "[Quit]" Then
        MsgBox "I am sure you don't really want me to quit SteamyDock... test cancelled."
        Exit Sub
    End If
    
    ' RD settings
    If thisCommand = "[Settings]" Then
        'thisCommand = App.Path & "\resources\dockSettings.exe"
        thisCommand = "C:\Program Files (x86)\SteamyDock\dockSettings\dockSettings.exe"
        If FExists(thisCommand) Then
            If debugflg = 1 Then Debug.Print "ShellExecute " & thisCommand
            Call executeCommand("runas", thisCommand, vbNullString, vbNullString, intShowCmd) 'change to call new function as part of .16
        Else
            MsgBox "Cannot find " & thisCommand
        End If
        
        Exit Sub
    End If
'    ' RD icon settings
'    If thisCommand = "[Icons]" Then
'        thisCommand = "C:\Program Files (x86)\SteamyDock\iconSettings\rocket1.exe"
'
'
'        If FExists(thisCommand) Then
'            If debugflg = 1 Then Debug.Print "ShellExecute " & thisCommand
'            Call executeCommand("runas", thisCommand, vbNullString, vbNullString, intShowCmd)
'        Else
'            MsgBox "Cannot find " & thisCommand
'        End If
'
'        Exit Sub
'    End If
    
    ' program files folder ' .38 DAEB 03/03/2021 rdIconConfig.frm Removed the individual references to a Windows class ID
'    If thisCommand = "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}" Then
'        Call shellCommand("explorer.exe /e,::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", intShowCmd) 'change to call new function as part of .16
'        Exit Sub
'    End If


    ' .39 DAEB 03/03/2021 rdIconConfig.frm check whether the prefix is present that indicates a Windows class ID is present
    ' this allows SD to act like Rocketdock which only needs the CLSID to operate eg. ::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}
    If InStr(thisCommand, "::{") Then
        Call shellCommand("explorer.exe /e," & thisCommand, intShowCmd)
        Exit Sub
    End If
    
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If

    '.102 DAEB 08/12/2022 rdIconConfig.frm icon settings responds to %systemroot% environment variables during testing
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
     ' applications And features
'    If thisCommand = "appwiz.cpl" Then
'        If debugflg = 1 Then DebugPrint "Shell " & "rundll32.exe shell32.dll,Control_RunDLL " & thisCommand
'        Call shellCommand("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl", intShowCmd) 'change to call new function as part of .16
'        Exit Sub
'    End If
    ' recycle bin ' .38 DAEB 03/03/2021 rdIconConfig.frm Removed the individual references to a Windows class ID
    If thisCommand = "[RecycleBin]" Then
        Call shellCommand("explorer.exe /e,::{645ff040-5081-101b-9f08-00aa002f954e}", intShowCmd) 'change to call new function as part of .16
        Exit Sub
    End If
        
    ' cpanel files with cpl suffix can be called from the command line
    If InStr(thisCommand, ".cpl") <> 0 Then
        Call shellCommand("rundll32.exe shell32.dll,Control_RunDLL " & thisCommand, intShowCmd)
        Exit Sub
    End If
    
    ' admin tools with msc suffix (management console controls) can be called from the command line
    If InStr(thisCommand, ".msc") <> 0 Then
        If FExists(thisCommand) Then ' if the file exists and is valid - run it
            Call executeCommand(userLevel, thisCommand, sArguments, vbNullString, intShowCmd)
            Exit Sub ' .89 DAEB 08/12/2022 frmMain.frm Fixed duplicate run of .msc files.
        Else
            folderPath = getFolderNameFromPath(thisCommand)  ' extract the default folder from the full path

            ' .45 DAEB 01/04/2021 frmMain.frm Changed the logic to remove the code around a folder path existing...
            If Not DirExists(folderPath) Then
                 ' if there is no folder path provided then attempt it on its own hoping that the windows PATH will find it
                On Error GoTo tryMSCFullPAth ' apologies for this GOTO - testing to see if it is in the path, then it will run.
                Call executeCommand(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
                Exit Sub
tryMSCFullPAth:
                On Error GoTo menuRun_click_Error
                ' now run it in the system32 folder
                Call executeCommand(userLevel, Environ$("windir") & "\SYSTEM32\" & getFileNameFromPath(thisCommand), sArguments, sWorkingDirectory, intShowCmd)
                Exit Sub
            End If

        End If
    End If
    
    ' task manager
    If thisCommand = "taskmgr" Then
        Call executeCommand("Open", Environ$("windir") & "\SYSTEM32\taskmgr", 0&, 0&, intShowCmd)
        Exit Sub
    End If
    ' RocketdockEnhancedSettings.exe (the .NET version of this program)
    If getFileNameFromPath(thisCommand) = "RocketdockEnhancedSettings.exe" Then
        answer = msgBoxA("It might not be a good idea to run the .NET and VB6 versions of the Rocketdock Utility at the same time. The two might conflict, and the results might not be positive. Are you sure you want me to?", vbExclamation + vbYesNo)
        If answer = 6 Then
            Call executeCommand("Open", thisCommand, txtArguments.Text, vbNullString, intShowCmd) 'change to call new function as part of .16
        Else
            Exit Sub
        End If
    End If
    ' rocket1.exe (this program)
    If getFileNameFromPath(thisCommand) = "iconsettings.exe" Then ' 17/11/2020    .04 DAEB Replaced all occurrences of rocket1.exe with iconsettings.exe

        MsgBox "If you run the Icon Settings Utility, the first thing it does is to kill any existing instance, ie. this program you are running now - and I'm sure you don't really want me to do that... test cancelled."
        Exit Sub
    End If
    
    ' bat files
    If ExtractSuffixWithDot(UCase$(thisCommand)) = ".BAT" Then
        If debugflg = 1 Then Debug.Print "ShellExecute " & thisCommand
        thisCommand = """" & sCommand & """" ' put the command in quotes so it handles spaces in the path
        folderPath = getFolderNameFromPath(thisCommand)  ' extract the default folder from the batch full path
        If FExists(thisCommand) Then
            Call executeCommand("Open", thisCommand, vbNullString, folderPath, intShowCmd) 'change to call new function as part of .16
        Else
            MsgBox (thisCommand & " - this batch file does not exist")
        End If
        Exit Sub
    End If
    
    'anything else remaining
    If FExists(thisCommand) Then ' checks the current folder for the named target
        'If debugflg = 1 Then debugLog "ShellExecute " & thisCommand
        If sWorkingDirectory <> vbNullString Then
            Call executeCommand(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
            Exit Sub
        Else
            Call executeCommand(userLevel, thisCommand, sArguments, vbNullString, intShowCmd)
            Exit Sub
        End If
    ElseIf DirExists(thisCommand) Then ' checks if a folder of the same name exists in the current folder
        Call executeCommand("open", thisCommand, sArguments, sWorkingDirectory, intShowCmd)
        Exit Sub
    End If
    
    ' items with no suffix not found in default folder - look in system32
    suffix = LCase(ExtractSuffixWithDot(thisCommand))
    If suffix = "" Then
        listOfTypes = ".exe|.msc|.cpl|.lnk|.bat"
        arrStr = Split(listOfTypes, "|")
        strCnt = UBound(arrStr) + 1
        
        For useloop = 0 To strCnt - 1
            userprof = Environ$("SYSTEMROOT") & "\system32\" & thisCommand & arrStr(useloop)
            If FExists(userprof) Then ' ' checks the windows system 32 folder for the named target
                Call executeCommand(userLevel, userprof, sArguments, sWorkingDirectory, intShowCmd)
                Exit Sub
            ElseIf validURL = False Then
                ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
                MessageBox Me.hwnd, thisCommand & " - That isn't valid as a target file or a folder, or it doesn't exist - so SteamyDock can't run it.", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
            End If
        Next useloop
    End If

   On Error GoTo 0
   Exit Sub

menuRun_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuRun_click of Form rDIconConfigForm"
End Sub

' .16 DAEB 01/03/2021 rDIConConfigForm.frm added new function to allow confirmation dialogue subsequent to running the test command
'---------------------------------------------------------------------------------------
' Procedure : executeCommand
' Author    : beededea
' Date      : 31/01/2021
' Purpose   : runs the shellexecute API function and puts up a dialogue box if required
'---------------------------------------------------------------------------------------
'
Private Sub executeCommand(ByVal userLevel As String, ByVal sCommand As String, ByVal sArguments As String, ByVal sWorkingDirectory As String, ByVal lastVal As Integer)

   On Error GoTo executeCommand_Error
   
    ' run the selected program
    Call ShellExecute(hwnd, userLevel, sCommand, sArguments, sWorkingDirectory, lastVal)
            
    userLevel = "open" ' return to default
    
    ' call up a dialog box if required
    If chkConfirmDialog.Value = 1 Then
        MsgBox sTitle & " Command Issued - " & sCommand, vbSystemModal + vbInformation, "SteamyDock Confirmation Message"
    End If
    
   On Error GoTo 0
   Exit Sub

executeCommand_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure executeCommand of Form rDIconConfigForm"
End Sub



' .16 DAEB 01/03/2021 rDIConConfigForm.frm added new function to allow confirmation dialogue subsequent to running the test command
'---------------------------------------------------------------------------------------
' Procedure : shellCommand
' Author    : beededea
' Date      : 31/01/2021
' Purpose   : runs the shell function and puts up a dialogue box if required
'---------------------------------------------------------------------------------------
'
Private Sub shellCommand(ByVal shellparam1 As String, ByVal vbNormalFocus)

   On Error GoTo shellCommand_Error

    ' run the selected program
    Call Shell(shellparam1, vbNormalFocus)
    'userLevel = "open" ' return to default

    ' call up a dialog box if required
    If chkConfirmDialogAfter.Value = 1 Then
        MsgBox sTitle & " Command Issued - " & sCommand, vbSystemModal + vbInformation, "SteamyDock Confirmation Message"
    End If

   On Error GoTo 0
   Exit Sub

shellCommand_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure shellCommand of Form rDIconConfigForm"
End Sub







Private Sub btnArrowDown_Click()
    Call subBtnArrowDown_Click
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnArrowDown_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub subBtnArrowDown_Click()
    Dim growBit As Integer: growBit = 0
    Dim amountToDrop As Integer: amountToDrop = 0
    
    On Error GoTo btnArrowDown_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnArrowDown_Click"
    
    growBit = 670
    amountToDrop = 1200
    
    If btnArrowDown.Visible = True Then btnArrowDown.Visible = False
            
    If picRdThumbFrame.Visible = False Then
        'has to do this first as redrawing errors occur otherwise
        'rDIconConfigForm.Refresh
        
        btnWorking.Visible = True
        Call busyStart
        'If picRdMap(0).Picture = 0 Then ' only recreate the map if the array is empty
        ' we used to check the .picture property but using lavolpe's 2nd method this proprty is not set.
        ' now we check for the tooltiptext which is only set when the image is populated.
        If picRdMap(0).ToolTipText = vbNullString Then ' only recreate the map if the array is empty
            Call populateRdMap(0) 'show the map from position zero
            ' set the primary selection in the map
            picRdMap(0).BorderStyle = 1
            
            'vb6 won't let the rdiconmap receive focus here using setfocus
            
        End If
        'rDIconConfigForm.Refresh
        
        Call setRdIconConfigFormHeight
        
        rDIconConfigForm.Height = rDIconConfigForm.Height + growBit

        framePreview.Top = 4545 + growBit
        fraProperties.Top = 4545 + growBit
        frameButtons.Top = 7925 + growBit
                
        ' .75 DAEB 22/05/2022 rDIConConfig.frm The dropdown disclose function is calculating the positions incorrectly when the map is toggled hidden/shown.
        If moreConfigVisible = True Then
            rDIconConfigForm.Height = rDIconConfigForm.Height + amountToDrop
            frameButtons.Top = frameButtons.Top + amountToDrop
        End If
                
        btnArrowUp.Visible = True
        picRdThumbFrame.Visible = True
        
        rdMapRefresh.Visible = True
        If rdIconMaximum > 16 Then
            'rdMapHScroll.Visible = True
        End If
        rdMapHScroll.Visible = True
        
        rdMapHScroll.Max = theCount - 1
                
        ' we signify that all changes have been lost
        btnSet.Enabled = False ' this has to be done at the end
        btnClose.Visible = True
        btnCancel.Visible = False

        
        btnWorking.Visible = False
        rDIconConfigForm.Refresh
        
        Call busyStop

        'write the visible state
        PutINISetting "Software\SteamyDockSettings", "sdMapState", "visible", toolSettingsFile

    
    End If
    
    ' .29 DAEB 14/03/2021 rDIConConfigForm.frm change to focus the icon map on the icon pre-selected
    If rdIconNumber > 0 Then
        Call picRdMapSetFocus ' set focus to the rocketdock icon map
        rdMapHScroll.Value = rdIconNumber ' set the map scroll value to the pre-selected number
        picRdMap(rdIconNumber).BorderStyle = 1 ' put a border around the selected map image
        ' give the specifc part of the map focus so that any keypresses will operate immediately
        'picRdMap(rdIconNumber).SetFocus  ' < .net
    Else

        ' give the map focus so that any keypresses will operate immediately
        'picRdMap(0).SetFocus  ' < .net
    End If
    
   On Error GoTo 0
   Exit Sub

btnArrowDown_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowDown_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnArrowDown_MouseUp
' Author    : beededea
' Date      : 15/11/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnArrowDown_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo btnArrowDown_MouseUp_Error
   If debugflg = 1 Then DebugPrint "%btnArrowDown_MouseUp"

        btnWorking.Visible = True

   On Error GoTo 0
   Exit Sub

btnArrowDown_MouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowDown_MouseUp of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnArrowUp_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : moves the lower frames back up to accommodate the RD Map becoming invisible
'---------------------------------------------------------------------------------------
'
Private Sub btnArrowUp_Click()
   On Error GoTo btnArrowUp_Click_Error
   If debugflg = 1 Then DebugPrint "%" & "btnArrowUp_Click"

   If picRdThumbFrame.Visible = True Then
   
'        rDIconConfigForm.Height = rDIconConfigForm.Height + 645
'        frameButtons.Top = frameButtons.Top + 645

        framePreview.Top = 4545
        fraProperties.Top = 4545
        frameButtons.Top = 7910
        
        Call setRdIconConfigFormHeight
    
        'rDIconConfigForm.dllFrame.Top = 7530
        
        ' .75 DAEB 22/05/2022 rDIConConfig.frm The dropdown disclose function is calculating the positions incorrectly when the map is toggled hidden/shown.
        If moreConfigVisible = True Then
            rDIconConfigForm.Height = rDIconConfigForm.Height - 750
            frameButtons.Top = frameButtons.Top - 750
        End If
        
        
        btnArrowDown.Visible = True
        btnArrowUp.Visible = False
        picRdThumbFrame.Visible = False
        
        If chkToggleDialogs.Value = 0 Then btnArrowDown.ToolTipText = "Show the Dock Map"
        
        rdMapRefresh.Visible = False
        rdMapHScroll.Visible = False
        
        'write the hidden state
        PutINISetting "Software\SteamyDockSettings", "sdMapState", "hidden", toolSettingsFile
                
   End If

   On Error GoTo 0
   Exit Sub

btnArrowUp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowUp_Click of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setRdIconConfigFormHeight
' Author    : beededea
' Date      : 17/12/2022
' Purpose   : Windows 10/11+ cut off the bottom of the traditional windows, add another 100 twips to compensate.
'---------------------------------------------------------------------------------------
'
Private Sub setRdIconConfigFormHeight()

    On Error GoTo setRdIconConfigFormHeight_Error

        rDIconConfigForm.Height = 9525
        
        ' if Windows 10/11 then add 250 twips to the bottom of the main form
        If Left$(LCase$(windowsVersionString), 10) = "windows 10" Then
            Me.Height = Me.Height + 100
        End If

    On Error GoTo 0
    Exit Sub

setRdIconConfigFormHeight_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setRdIconConfigFormHeight of Form rDIconConfigForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picRdMapSetFocus
' Author    : beededea
' Date      : 17/11/2019
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub picRdMapSetFocus()
    
    On Error GoTo picRdMapSetFocus_Error
    If debugflg = 1 Then DebugPrint "%picRdMapSetFocus"

    picRdMapGotFocus = True
    picFrameThumbsGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False

    On Error GoTo 0
    Exit Sub

picRdMapSetFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMapSetFocus of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : picFrameThumbsSetFocus
' Author    : beededea
' Date      : 17/11/2019
' Purpose   : as panels cannot getFocus in VB.NET we have to kludge it by setting a variable
'             that indicates which control we want to have focus.
'---------------------------------------------------------------------------------------
'
    Private Sub picFrameThumbsSetFocus()

   On Error GoTo picFrameThumbsSetFocus_Error
   If debugflg = 1 Then DebugPrint "%picFrameThumbsSetFocus"

            picFrameThumbsGotFocus = True
            picRdMapGotFocus = False
            previewFrameGotFocus = False
            filesIconListGotFocus = False
            vScrollThumbsGotFocus = False

   On Error GoTo 0
   Exit Sub

picFrameThumbsSetFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbsSetFocus of Form rDIconConfigForm"

    End Sub

' .12 DAEB 07/02/2021 rDIconConfigForm.frm added as part of busy timer functionality
'---------------------------------------------------------------------------------------
' Procedure : busyTimer_Timer
' Author    : beededea
' Date      : 25/05/2020
' Purpose   : rotates the hourglass timer
'---------------------------------------------------------------------------------------
'
Private Sub busyTimer_Timer()
        Dim thisWindow As Long: thisWindow = 0
        Dim busyFilename As String: busyFilename = vbNullString
        
        On Error GoTo busyTimer_Timer_Error

        thisWindow = fFindWindowHandle("SteamyDock")
        busyFilename = ""
        
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        totalBusyCounter = totalBusyCounter + 1
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        If classicTheme = True Then
            busyFilename = App.Path & "\resources\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.Path & "\resources\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename) ' imageList candidate
        
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


' .13 DAEB rdIconConfig.frm 09/02/2021 Added ability to check if a window exists
'---------------------------------------------------------------------------------------
' Procedure : fFindWindowHandle
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function fFindWindowHandle(ByVal Caption As String) As Long
   On Error GoTo fFindWindowHandle_Error

  fFindWindowHandle = FindWindow(vbNullString, Caption)

   On Error GoTo 0
   Exit Function

fFindWindowHandle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFindWindowHandle of Module mdlMain"
End Function



'---------------------------------------------------------------------------------------
' Procedure : checkClassicThemeCapable
' Author    : beededea
' Date      : 26/02/2021
' Purpose   :
'               turn on the timer that tests every 10 secs whether the visual theme has changed
'               only on those o/s versions that need it
'---------------------------------------------------------------------------------------
'
Private Sub checkClassicThemeCapable() ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity
    
   On Error GoTo checkClassicThemeCapable_Error

    If classicThemeCapable = True Then
        rDIconConfigForm.mnuAuto.Caption = "Auto Theme Disable"
        rDIconConfigForm.themeTimer.Enabled = True
    Else
        rDIconConfigForm.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"
        rDIconConfigForm.themeTimer.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

checkClassicThemeCapable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkClassicThemeCapable of Form rDIconConfigForm"
End Sub
    




'---------------------------------------------------------------------------------------
' Procedure : setInitialPath
' Author    : beededea
' Date      : 26/02/2021
' Purpose   : set the default path to the icons root, this will be superceded later if the user has chosen a default folder
'---------------------------------------------------------------------------------------
'
Private Sub setInitialPath() ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity

   On Error GoTo setInitialPath_Error

    If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
        filesIconList.Path = rdAppPath & "\Icons" ' rdAppPath is defined in driveCheck above
        textCurrentFolder.Text = rdAppPath & "\Icons"
        relativePath = "\Icons"
    Else
        filesIconList.Path = sdAppPath & "\Icons" ' rdAppPath is defined in driveCheck above
        textCurrentFolder.Text = sdAppPath & "\Icons"
        relativePath = "\Icons"
    End If

   On Error GoTo 0
   Exit Sub

setInitialPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setInitialPath of Form rDIconConfigForm"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : killPreviousInstance
' Author    : beededea
' Date      : 26/02/2021
' Purpose   : if the process already exists then kill it
'---------------------------------------------------------------------------------------
'
Private Sub killPreviousInstance() ' .13 DAEB 27/02/2021 rdIConConfigFrm moved to a subroutine for clarity
    Dim NameProcess As String: NameProcess = ""
    Dim AppExists As Boolean: AppExists = False

   On Error GoTo killPreviousInstance_Error

    ReDim thumbArray(12) As Integer
    
    ' initial values assigned
    
    NameProcess = ""
    AppExists = False

    '
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "iconsettings.exe" ' 17/11/2020    .04 DAEB Replaced all occurrences of rocket1.exe with iconsettings.exe

        'MsgBox "You now have two instances of this utility running, they will conflict..."
        checkAndKill NameProcess, False, False
    End If

   On Error GoTo 0
   Exit Sub

killPreviousInstance_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure killPreviousInstance of Form rDIconConfigForm"

End Sub




' .30 DAEB 10/04/2021 rDIConConfigForm.frm separate the initial reading of the tool's settings file from the changing of the tool's own font STARTS
'---------------------------------------------------------------------------------------
' Procedure : readAndSetUtilityFont
' Author    : beededea
' Date      : 10/04/2021
' Purpose   : reads the tool's font settings from the local tool file
'---------------------------------------------------------------------------------------
'
Private Sub readAndSetUtilityFont()
    
    ' variables declared
    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    SDSuppliedFontColour = vbBlack

    ' set the tool's default font
   On Error GoTo readAndSetUtilityFont_Error

    SDSuppliedFont = GetINISetting("Software\SteamyDockSettings", "defaultFont", toolSettingsFile)
    SDSuppliedFontSize = Val(GetINISetting("Software\SteamyDockSettings", "defaultSize", toolSettingsFile))
    SDSuppliedFontItalics = Val(GetINISetting("Software\SteamyDockSettings", "defaultItalics", toolSettingsFile))
    SDSuppliedFontColour = Val(GetINISetting("Software\SteamyDockSettings", "defaultColour", toolSettingsFile))
'    SDSuppliedFontStrength = GetINISetting("Software\SteamyDockSettings", "defaultStrength", toolSettingsFile)
'    SDSuppliedFontStyle = GetINISetting("Software\SteamyDockSettings", "defaultStyle", toolSettingsFile)
    rDSkinTheme = GetINISetting("Software\SteamyDockSettings", "SkinTheme", toolSettingsFile) ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
    
    
    'storedFont = txtTextFont.Text 'TBD
    
    fntFont = SDSuppliedFont
    fntSize = SDSuppliedFontSize
    fntItalics = CBool(SDSuppliedFontItalics)
    fntColour = CLng(SDSuppliedFontColour)

    If Not SDSuppliedFont = "" Then
        ' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font
        Call changeFont(Me, False, SDSuppliedFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
        'Call changeFont(SDSuppliedFont, SDSuppliedFontSize, SDSuppliedFontStrength, SDSuppliedFontStyle)
    End If


   On Error GoTo 0
   Exit Sub

readAndSetUtilityFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readAndSetUtilityFont of Form rDIconConfigForm"
End Sub
' .30 DAEB 10/04/2021 rDIConConfigForm.frm separate the initial reading of the tool's settings file from the changing of the tool's own font ENDS

Private Sub chkQuickLaunch_Click()
        btnSet.Enabled = True ' tell the program that something has changed
            btnCancel.Visible = True
    btnClose.Visible = False
End Sub

Private Sub chkAutoHideDock_Click()
        btnSet.Enabled = True ' tell the program that something has changed
            btnCancel.Visible = True
    btnClose.Visible = False
End Sub


' .49 DAEB 20/04/2022 rDIConConfig.frm Added balloon tooltips STARTS

Private Sub txtTarget_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip txtTarget.hwnd, "This field should contain the full path and filename of the target application.", _
                  TTIconInfo, "Help on the Target Path Box", , , , True
End Sub

Private Sub btnSet_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSet.hwnd, "This button sets and stores the icon characteristics that you have entered. However, you will need to press the save and restart button below to make it 'fix' onto the running dock. ", _
                  TTIconInfo, "Help on Additional Arguments", , , , True
End Sub

Private Sub btnAdd_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnAdd.hwnd, "This button takes the currently selected icon and places it onto the Dock Map, the same as double-clicking on an icon.", _
                  TTIconInfo, "Help on the Add an Icon Button", , , , True
End Sub

Private Sub btnAddFolder_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnAddFolder.hwnd, "This button works with the folder treelist above. It allows you to add an existing folder location to SteamyDock so that you can also select your own icons.", _
                  TTIconInfo, "Help on Adding a Folder", , , , True
End Sub

Private Sub btnArrowDown_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnArrowDown.hwnd, "This small button will show the icon map.", _
                  TTIconInfo, "Help on the Show Icon Map Button", , , , True
End Sub

Private Sub btnArrowUp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnArrowUp.hwnd, "This small button will hide the icon map.", _
                  TTIconInfo, "Help on the Hide Icon Map Button", , , , True
End Sub

Private Sub btnBackup_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnBackup.hwnd, "This button takes an immediate backup and optionally opens the backup folder so that you can review the backup files.", _
                  TTIconInfo, "Help on the Backup Button", , , , True
End Sub

Private Sub btnCancel_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnCancel.hwnd, "This button cancels the current operation.", _
                  TTIconInfo, "Help on the Cancel Button", , , , True
End Sub

Private Sub btnFileListView_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnFileListView.hwnd, "This button switches from image display mode to file detail mode in icon file display window.", _
                  TTIconInfo, "Help on the File Detail Mode Button", , , , True
End Sub

Private Sub btnGenerate_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGenerate.hwnd, "Pressing this button causes a utility to appear that will wipe the dock and make a whole NEW dock -  use with care! ", _
                  TTIconInfo, "Help on Auto-Generating a Dock", , , , True
End Sub

Private Sub btnGetMore_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGetMore.hwnd, "This button will open the browser at a page where you can download more icons.", _
                  TTIconInfo, "Help on the More Icons Button", , , , True
End Sub

Private Sub btnHelp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnHelp.hwnd, "This button opens the help page in your default browser.", _
                  TTIconInfo, "Help on the Help Button", , , , True
End Sub

Private Sub btnIconSelect_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnIconSelect.hwnd, "Press this button to select an icon manually using a file browser. Select a PNG, ICO, JPG or BMP file. Ensure the file is square and is an icon.", _
                  TTIconInfo, "Help on the Manual Icon Select Button", , , , True
End Sub

Private Sub btnKillIcon_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnKillIcon.hwnd, "This button allows you to delete the currently selected icon in the icon file window above. Use wisely! Once it has gone, it has gone forever!", _
                  TTIconInfo, "Help on the Delete Icon Button", , , , True
End Sub

Private Sub btnMapNext_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnMapNext.hwnd, "This will scroll the icon map to the right so that you can view additional icons.", _
                  TTIconInfo, "Help on the Scroll Map Right Button", , , , True
End Sub

Private Sub btnMapPrev_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnMapPrev.hwnd, "This will scroll the icon map to the left so that you can view additional icons.", _
                  TTIconInfo, "Help on the Scroll Map Left Button", , , , True
End Sub

Private Sub btnNext_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnNext.hwnd, "This will select the next icon to the right within the icon map.", _
                  TTIconInfo, "Help on the Next Icon Button", , , , True
End Sub

Private Sub btnPrev_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnPrev.hwnd, "This will select the next icon to the left within the icon map.", _
                  TTIconInfo, "Help on the Refresh Icon Map Button", , , , True
End Sub

Private Sub btnRefresh_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnRefresh.hwnd, "This button refreshes the icon file display.", _
                  TTIconInfo, "Help on the Refresh Button", , , , True
End Sub

Private Sub btnRemoveFolder_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnRemoveFolder.hwnd, "This button works with the folder treelist above. It will allow you to remove the selected folder from the folder treelist. Note that the default application folders cannot be removed, only those that you add manually.", _
                  TTIconInfo, "Help on Removing a Folder", , , , True

End Sub

Private Sub btnSaveRestart_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSaveRestart.hwnd, "A press of the save and restart button is required when any icon changes have been made. This causes the dock to restart and in so doing, it picks up the latest changes and displays them.", _
                  TTIconInfo, "Help on Saving and Restarting", , , , True
End Sub

Private Sub btnSecondApp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSecondApp.hwnd, "This button will open a file explorer window allowing you to specify any additional secondary program to run after the main program launch has completed. ", _
                  TTIconInfo, "Help on Second Application Selection Button", , , , True
End Sub

Private Sub btnSelectStart_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSelectStart.hwnd, "Press this button to select a target folder for this icon, using a folder browser from which you can select a specific folder. Some apps require a default folder from which to operate. If you double click on the empty text box to the left then it will automatically fill in the folder using the folder of the target application. ", _
                  TTIconInfo, "Help on the Start Folder Select Button", , , , True
End Sub


Private Sub btnSettingsDown_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSettingsDown.hwnd, "This button displays the location of the current settings. This tells you where the configuration details are being stored and where they are being read from and saved to. The help has more information.", _
                  TTIconInfo, "Help on the cConfiguration Settings Location", , , , True
End Sub
Private Sub btnSettingsUp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnSettingsUp.hwnd, "Hide the registry form showing where details are being read from and saved to.", _
                  TTIconInfo, "Help on hiding the Configuration Settings", , , , True
'Hide the registry form showing where details are being read from and saved to.
End Sub

Private Sub btnTarget_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnTarget.hwnd, "This button will open a file browser from which you can select an application for this icon. Typically, you would select a binary or a .EXE file to run when the selected icon is clicked upon. If you RIGHT CLICK ON THIS BUTTON, a menu will become visible where you can select a target and all the fields will be filled out automatically.", _
                  TTIconInfo, "Help on the Target Application Button", , , , True
End Sub

Private Sub btnThumbnailView_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnThumbnailView.hwnd, "This button switches from file detail mode to image display mode in icon file display window.", _
                  TTIconInfo, "Help on the Image Mode Button", , , , True
End Sub

Private Sub btnWorking_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnWorking.hwnd, "This is an informational button that simply tells you that this utility is doing something...", _
                  TTIconInfo, "Help on the Working Button", , , , True
End Sub

Private Sub chkAutoHideDock_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkAutoHideDock.hwnd, "This causes the dock to hide immediately before the application launches. This allows full screen apps to run uninterrupted by the dock. The dock will re-appear 1.5 seconds after the application is closed. ", _
                  TTIconInfo, "Help on Auto-Hiding the Dock", , , , True
End Sub

Private Sub chkConfirmDialog_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkConfirmDialog.hwnd, "This causes a Confirmation Dialog to pop up prior to the specified command running, allowing you a chance to say yes or no at runtime. ", _
                  TTIconInfo, "Help on Confirming Beforehand", , , , True
End Sub

Private Sub chkConfirmDialogAfter_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkConfirmDialogAfter.hwnd, "Some programs run without producing any output. Checking this causes a Confirmation Dialog to pop up after the specified command has run. Please note it does not confirm the application was successful in its task, it just gives you confirmation that the command was successfully issued.", _
                  TTIconInfo, "Help on Confirming Afterward", , , , True

End Sub

Private Sub chkRunElevated_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkRunElevated.hwnd, "When this checkbox is ticked, the associated app will run with elevated privileges, ie. as administrator. Some programs require this in order to operate.", _
                  TTIconInfo, "Help on Running Elevated", , , , True
End Sub

Private Sub chkToggleDialogs_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkToggleDialogs.hwnd, "When this checkbox is ticked this will display the information pop-ups (the confirmation on saves and deletes) and balloon tooltips. When it is unchecked only the standard single-line tooltips will appear and there will be no warning dialogs. ", _
                  TTIconInfo, "Help on the Dialog Checkbox", , , , True
End Sub

Private Sub chkQuickLaunch_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip chkQuickLaunch.hwnd, "This causes the application to launch before any dock animation has occurred speeding up launch times. This setting can also be controlled globally via the Dock Settings Utility in the Icon Behaviour Pane via the setting named Icon Attention Effect", _
                  TTIconInfo, "Help on Quick Launch", , , , True
End Sub

Private Sub fraConfigSource_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraConfigSource.hwnd, "This Dropdown displays the currently selected dock. The current dock is configured in the dock settings utility, this is just for informational purposes only.", _
                  TTIconInfo, "Help on the Dock Selection Dropdown", , , , True
End Sub

Private Sub fraIconType_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraIconType.hwnd, "This dropdown allows yout to filter icon types to display in the icon file window above.", _
                  TTIconInfo, "Help on the Drop Down Icon Filter", , , , True
End Sub

Private Sub fraLblArgument_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblArgument.hwnd, "An optional field, add any additional arguments that the target file operation requires, eg. -s -t 00 -f . ", _
                  TTIconInfo, "Help on Additional Arguments", , , , True
End Sub

Private Sub fraLblConfirmDialogAfter_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblConfirmDialogAfter.hwnd, "Some programs run without producing any output. Checking this causes a Confirmation Dialog to pop up after the specified command has run. Please note it does not confirm the application was successful in its task, it just gives you confirmation that the command was successfully issued.", _
                  TTIconInfo, "Help on Confirming Afterward", , , , True
End Sub

Private Sub fraLblCurrentIcon_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblCurrentIcon.hwnd, "This displays the full path of the currently selected icon. Just double-click on an icon in the icon window above and it will automatically populate this field, replacing the current icon.", _
                  TTIconInfo, "Help on the Icon Path Text Box", , , , True
End Sub

Private Sub fraLblOpenRunning_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblOpenRunning.hwnd, "Choose whether to open a new instance if the chosen app is already running. The global setting normally determines whether you open new or existing instances of all apps but here you can set a specific action for particular programs.", _
                  TTIconInfo, "Help on Open Running Behaviour.", , , , True
End Sub

Private Sub fraLblPopUp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblPopUp.hwnd, "When this checkbox is ticked, the associated app will run with elevated privileges, ie. as administrator. Some programs require this in order to operate.", _
                  TTIconInfo, "Help on Running Elevated", , , , True
End Sub

Private Sub fraLblQuickLaunch_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblQuickLaunch.hwnd, "This causes the application to launch before any dock animation has occurred speeding up launch times. This setting can also be controlled globally via the Dock Settings Utility in the Icon Behaviour Pane via the setting named Icon Attention Effect", _
                  TTIconInfo, "Help on Quick Launch", , , , True
End Sub

Private Sub fraLblRdIconNumber_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblRdIconNumber.hwnd, "This is number of the current icon that is being displayed in the preview or in the map above.", _
                  TTIconInfo, "Help on Icon Numbering", , , , True
End Sub

Private Sub fraLblRun_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblRun.hwnd, "This dropdown selects the Window mode for the program to operate within. If you want to force an application to run in a full screen size window then select Maximised (note this is not a requirement for most full screen-type apps such as games). You might also want to start an app fully minimised on the taskbar. In other cases choose normal.", _
                  TTIconInfo, "Help on Window Mode Selection", , , , True
End Sub

Private Sub fraLblStartIn_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblStartIn.hwnd, "An optional field, that only needs to contain a value if the starting app requires a start folder. Press the square button on the right to select a start folder for this icon. If you double click here then it will automatically fill in the folder using the target file path immediately above. ", _
                  TTIconInfo, "Help on Start Folder Selection", , , , True
End Sub

Private Sub fraLblTarget_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblTarget.hwnd, "This field should contain the full path and filename of the target application.", _
                  TTIconInfo, "Help on the Target Path Box", , , , True
End Sub

Private Sub fraLblConfirmDialog_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblConfirmDialog.hwnd, "Adds a Confirmation Dialog prior to the command running allowing you to say yes or no at runtime.", _
                  TTIconInfo, "Help on the Confirming Dialog", , , , True
End Sub

Private Sub fraLblSecondApp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblSecondApp.hwnd, "Specify any additional secondary program to run after the main program launch has completed. ", _
                  TTIconInfo, "Help on Second Application", , , , True
End Sub

Private Sub fraLblAppToTerminate_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLblAppToTerminate.hwnd, "Specify any program that must be terminated prior to the main program initiation will be shown here. ", _
                  TTIconInfo, "Help on Terminating an Application", , , , True
End Sub
Private Sub frmLblAutoHideDock_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip frmLblAutoHideDock.hwnd, "This causes the dock to hide immediately before the application launches. This allows full screen apps to run uninterrupted by the dock. The dock will re-appear 1.5 seconds after the application is closed. ", _
                  TTIconInfo, "Help on Auto-Hiding the Dock", , , , True
End Sub


Private Sub picHideConfig_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picHideConfig.hwnd, "Hides the extra configuration section.", _
                  TTIconInfo, "Help on Hiding Configuration", , , , True
End Sub
Private Sub picMoreConfigUp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picMoreConfigUp.hwnd, "Hides the extra configuration section.", _
                  TTIconInfo, "Help on Hiding Configuration", , , , True
End Sub

Private Sub txtLabelName_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtLabelName.hwnd, "This field should contain the label of the icon as it appears on the dock.", _
                  TTIconInfo, "Help on the Icon Label", , , , True
End Sub

Private Sub picMoreConfigDown_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picMoreConfigDown.hwnd, "Press this button to display extra configuration items in the dropdown area at the base of this utility.", _
                  TTIconInfo, "Help on the More Configuration Button", , , , True
End Sub

Private Sub picPreview_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip picPreview.hwnd, "This is the currently selected icon scaled to fit the preview box, the size is controlled using the slider below.", _
                  TTIconInfo, "Help on the Icon Preview", , , , True
End Sub

Private Sub picRdThumbFrame_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip picRdThumbFrame.hwnd, "This is the icon map. It maps your dock exactly, showing you the same icons that appear in your dock. You can add or delete icons to/from the map. Press save and restart and they will appear in your dock.", _
                  TTIconInfo, "Help on the Icon Map", , , , True
End Sub

Private Sub rdMapRefresh_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip rdMapRefresh.hwnd, "This button refreshes the Icon Map. If you ever worry about mistakes in about your recent changes, just refresh.", _
                  TTIconInfo, "Help on the Refresh Icon Map Button", , , , True
End Sub

Private Sub sliPreviewSize_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip sliPreviewSize.hwnd, "This will size the chosen icon so you can see how it looks when it is shown at different sizes in the dock.", _
                  TTIconInfo, "Help on the Icon Size Slider", , , , True
End Sub

Private Sub textCurrentFolder_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip textCurrentFolder.hwnd, "This displays the full path of the currently selected folder in the treelist.", _
                  TTIconInfo, "Help on the Current Folder Path", , , , True
End Sub

Private Sub txtArguments_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtArguments.hwnd, "An optional field, add any additional arguments that the target file operation requires, eg. -s -t 00 -f . ", _
                  TTIconInfo, "Help on Additional Arguments", , , , True
End Sub

Private Sub txtCurrentIcon_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtCurrentIcon.hwnd, "This displays the full path of the currently selected icon. Just double-click on an icon in the icon window above and it will automatically populate this field, replacing the current icon.", _
                  TTIconInfo, "Help on the Icon Path Text Box", , , , True
End Sub

Private Sub txtSecondApp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtSecondApp.hwnd, "Specify any additional secondary program to run after the main program launch has completed. ", _
                  TTIconInfo, "Help on Second Application", , , , True
End Sub
Private Sub txtStartIn_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip txtStartIn.hwnd, "An optional field, that only needs to contain a value if the starting app requires a start folder. Press the square button on the right to select a start folder for this icon. If you double click here then it will automatically fill in the folder using the target file path immediately above. ", _
                  TTIconInfo, "Help on Start Folder Selection", , , , True
End Sub

' .49 DAEB 20/04/2022 rDIConConfig.frm Added balloon tooltips ENDS










' .87 DAEB 06/06/2022 rDIConConfig.frm Add OLE drag and drop of applications directly to the map using code from SteamyDock - STARTS
'---------------------------------------------------------------------------------------
' Procedure : Form_OLEDragDrop
' Author    : beededea
' Date      : 15/04/2020
' Purpose   : Handles drag and drop to the dock, only file types accepted. If an image, drops it straight onto the dock.
'             If it is a binary then we use code to try to extract the embededded icons using privateExtractIcons API
'             especially when the icon is a bigger one, if it is only a low resolution icon then we give it an icon based upon its suffix.
'             direct from the icon collection.
'             If it is a special binary, msc, cpl then it is given an icon from the collection
'             If it is a shortcut we have some code to investigate the shortcut for the link details
'
'             I have made the decision not to use the embedded icons by default as for the majority of
'             Win o/ses before 10 the embedded icons are low resolution and puny. Instead we use document types from the collection.
'             This IS STEAMYDOCK!
'---------------------------------------------------------------------------------------
'
Private Sub picRdMap_OLEDragDrop(ByRef Index As Integer, ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

    'The Format numbers used in the OLE DragDrop data structure, are:
    'Text = 1 (vbCFText)
    'Bitmap = 2 (vbCFBitmap)
    'Metafile = 3
    'Emetafile = 14
    'DIB = 8
    'Palette = 9
    'Files = 15 (vbCFFiles)
    'RTF = -16639
    
    Dim suffix As String: suffix = vbNullString
    Dim Filename As String: Filename = vbNullString
    Dim iconImage As String: iconImage = vbNullString
    Dim iconTitle As String: iconTitle = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim iconCommand As String: iconCommand = vbNullString
    Dim iconArguments As String: iconArguments = vbNullString
    Dim iconWorkingDirectory As String: iconWorkingDirectory = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim nname As String: nname = vbNullString
    Dim npath As String: npath = vbNullString
    Dim ndesc As String: ndesc = vbNullString
    Dim nwork As String: nwork = vbNullString
    Dim nargs As String: nargs = vbNullString
    'Dim shortCutMethod As Integer
    Dim thisShortcut As Link

    
    On Error GoTo picRdMap_OLEDragDrop
    
'    ' if the dock is not the bottom layer then pop up a message box
'    ' ie. don't pop it up if layered underneath everything as no-one will see the msgbox
'    If rDLockIcons = 1 Then
'        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
'        MessageBox Me.hwnd, "Sorry, the dock is currently locked, so drag and drop is disabled!", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
'        '        MsgBox "Sorry, the dock is currently locked, so drag and drop is disabled!"
'        Exit Sub
'    End If
    
    iconImage = vbNullString
    iconTitle = vbNullString
    iconArguments = vbNullString
    iconWorkingDirectory = vbNullString
    
    ' if there is more than one file dropped reject the drop
    ' if the dock is not the bottom layer then pop up a message box
    ' ie. don't pop it up if layered underneath everything as no-one will see the msgbox
    If Data.Files.count > 1 Then
       ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Sorry, can only accept one icon drop at a time, you have dropped " & Data.Files.count, "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Sorry, can only accept one icon drop at a time, you have dropped " & Data.Files.count
        Exit Sub
    End If
    
    If Data.GetFormat(vbCFFiles) Then
        ' if it is of type 'file' then determine what to do, I have specific catch-alls and
        ' sections for each, just in case any specific tasks are required for each type.
        ' this could all be removed later as the actions seem more or less the same for each.
        
        ' Data.Files.Item
    
        iconTitle = Data.Files(1) ' set the title for all types
        iconCommand = Data.Files(1) ' set the command for all types
        
        ' is it a folder, does the folder exist
        If DirExists(iconTitle) Then
            iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-dir.png"
            If FExists(iconFileName) Then
                iconImage = iconFileName
            End If
        Else ' otherwise it is a file
    
            suffix = LCase$(ExtractSuffixWithDot(Data.Files(1)))
            If InStr(".exe,.bat,.msc,.cpl,.lnk", suffix) <> 0 Then
                  
                  Effect = vbDropEffectCopy
                 
                  'if an exe is dragged and dropped onto RD it is given an id, that it appends to the binary name after an additional "?"
                  ' that ? signifies what? Well, possibly it is the handle of the embedded icon only added the one time, so that when the binary is read in the future the handle is already there
                  ' and that can be used to populate image array? Untested.
                  ' in this case we just need to note the ? and then query the binary for an embedded icon handle and compare it to the id that RD has given it.
                  ' if it is the same then we can perhaps simulate the same.
                  
                  If suffix = ".exe" Then
                    ' we should, if it is a EXE dig into it to determine the icon using privateExtractIcon
                                         
                    ' However, we do not extract the icon from the shortcut as it will be useless for steamydock
                    ' VB6 not being able to extract and handle a transparent PNG form
                    ' even if it was we have no current method of making a transparent PNG from a bitmap or ICO that
                    ' I can easily transfer to the GDI collection - but I am working on it...
                    ' the vast majority of default icons are far too small for steamydock in any case.
                    ' the result of the above is that there is currently no icon extracted, though that may change.
                    
                    ' instead we have a list of apps that we can match the shortcut name against, it exists in an external comma
                    ' delimited file. The list has two identification factors that are used to find a match and then we find an
                    ' associated icon to use with a relative path.
                       
                    iconFileName = identifyAppIcons(iconCommand) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                       
                    If FExists(iconFileName) Then
                      iconImage = iconFileName
                    Else
                      iconImage = App.Path & "\my collection\steampunk icons MKVI" & "\document-EXE.png"
                    End If
                    
                  End If
                  
                  If suffix = ".msc" Then
                      ' if it is a MSC then  give it a SYSTEM type icon (EVENT)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-msc.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
                  If suffix = ".bat" Then
                      ' if it is a BAT then give it a BATCH type icon (NOTEPAD)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-bat.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
                  If suffix = ".cpl" Then
                      ' if it is a CPL then give it a SYSTEM type icon (CONSOLE)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\my collection\steampunk icons MKVI" & "\document-cpl.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
            '       If it is a shortcut we have some code to investigate the shortcut for the link details
                  If suffix = ".lnk" Then
                        ' if it is a short cut then you can use two methods, the first is currently limited to only
                        ' producing the path alone but it does avoid using the shell method that causes FPs to occur in av tools

                        Call GetShortcutInfo(iconCommand, thisShortcut) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                                       
                        iconTitle = getFileNameFromPath(thisShortcut.Filename)
                        
                        If Not thisShortcut.Filename = "" Then
                            iconCommand = LCase$(thisShortcut.Filename)
                        End If
                        iconArguments = thisShortcut.Arguments
                        iconWorkingDirectory = thisShortcut.RelPath
                        
                        ' .55 DAEB 19/04/2021 frmMain.frm Added call to the older function to identify an icon using the shell object
                        'if the icontitle and command are blank then this is user-created link that only provides the relative path
                        If iconTitle = "" And thisShortcut.Filename = "" And Not iconWorkingDirectory = "" Then
                            Call GetShellShortcutInfo(iconCommand, nname, npath, ndesc, nwork, nargs)
                    
                            iconTitle = nname
                            iconCommand = npath
                            iconArguments = nargs
                            iconWorkingDirectory = nwork
                        End If
                       
                      ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                      
                      ' we do not extract the icon from the shortcut as it will be useless for steamydock
                      ' VB6 not being able to extract and handle a transparent PNG form
                      ' even if it was we have no current method of making a transparent PNG from a bitmap or ICO that
                      ' I can easily transfer to the GDI collection - but I am working on it...
                      ' the vast majority of default icons are far too small for steamydock in any case.
                      ' the result of the above is that there is currently no icon extracted, though that may change.
                      
                      ' instead we have a list of apps that we can match the shortcut name against, it exists in an external comma
                      ' delimited file. The list has two identification factors that are used to find a match and then we find an
                      ' associated icon to use with a relative path.
                      
                      iconFileName = identifyAppIcons(iconCommand)
                       
                      If FExists(iconFileName) Then
                        iconImage = iconFileName
                      Else
                        iconImage = App.Path & "\my collection\steampunk icons MKVI" & "\document-lnk.png"
                      End If
                  End If
            
              ElseIf InStr(".png, .bmp, .gif, .jpg, .jpeg, .ico, .tif, .tiff", suffix) <> 0 Then
                  ' See if this is a file name ending in bmp, gif, jpg, or jpeg or tiff.
                  ' if so use the filename and drop it onto the dock
                  
                  Effect = vbDropEffectCopy
                  
                  iconImage = iconCommand
                  If Not FExists(iconImage) Then
                      Exit Sub
                  End If
              
              ElseIf InStr(".zip, .7z, .arj, .deb, .pkg, .rar, .rpm, .gz, .z, .bck", suffix) <> 0 Then
                  
                '    .7z - 7-Zip compressed file
                '    .arj - ARJ compressed file
                '    .deb - Debian software package file
                '    .pkg - Package file
                '    .rar - RAR file
                '    .rpm - Red Hat Package Manager
                '    .tar.gz - Tarball compressed file
                '    .z - Z compressed file
                '    .zip - Zip compressed file
                
                ' See if this is a file name ending in the above
                ' if so use the filename and drop it onto the dock
                  
                Effect = vbDropEffectCopy
                                  
                suffix = LCase$(ExtractSuffix(Data.Files(1)))
                iconImage = App.Path & "\my collection\steampunk icons MKVI\document-" & suffix & ".png"
                iconCommand = Data.Files(1)
                If Not FExists(iconImage) Then
                    iconImage = App.Path & "\my collection\steampunk icons MKVI" & "\document-zip.png"
                End If
                
                If Not FExists(iconImage) Then
                    Exit Sub
                End If
            
                      
              Else ' does not match any given type so see if we have an icon in the collection ready for it.
              
                  ' take the suffix and find a file in the collection that matches
                  ' if the file exists then add it to the menu
                  ' otherwise just do an empty default icon
                  
                  Effect = vbDropEffectCopy
                  
                  suffix = LCase$(ExtractSuffix(Data.Files(1)))
                  iconImage = App.Path & "\my collection\steampunk icons MKVI\document-" & suffix & ".png"
                  iconCommand = Data.Files(1)
                  If Not FExists(iconImage) Then
                      iconImage = App.Path & "\nixietubelargeQ.png"
                  End If
                      
              End If
    
        End If
        
        ' if no specific image found
        If iconImage = vbNullString Then
            iconImage = App.Path & "\nixietubelargeQ.png"
        End If
        
        If FExists(iconImage) Then ' last check that the default ? image has not been deleted.
            Call menuAddSomething(iconImage, iconTitle, iconCommand, iconArguments, iconWorkingDirectory, vbNullString, vbNullString)
        Else
            ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
             'MessageBox Me.hwnd, iconImage & " missing default image, " & App.Path & "\nixietubelargeQ.png" & " drop unsuccessful. ", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
             '            MsgBox iconImage & " missing default image, " & App.Path & "\nixietubelargeQ.png" & " drop unsuccessful. ", vbSystemModal
        End If
        
        
        'Call menuForm.mnuIconSettings_Click
        
    Else
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, " unknown file Object OLE dropped onto SteamyDock.", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        'MsgBox " unknown file Object OLE dropped onto SteamyDock."
    End If

    On Error GoTo 0
    Exit Sub

picRdMap_OLEDragDrop:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_OLEDragDrop of Form dock"

End Sub
' .87 DAEB 06/06/2022 rDIConConfig.frm Add OLE drag and drop of applications directly to the map using code from SteamyDock - END


'---------------------------------------------------------------------------------------
' Procedure : mnuBringToCentre_click
' Author    : beededea
' Date      : 19/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuBringToCentre_click()

   On Error GoTo mnuBringToCentre_click_Error

    rDIconConfigForm.Top = Screen.Height / 2 - rDIconConfigForm.Height / 2
    rDIconConfigForm.Left = screenWidthTwips / 2 - rDIconConfigForm.Width / 2

   On Error GoTo 0
   Exit Sub

mnuBringToCentre_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuBringToCentre_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuOpenFolder_click
' Author    : beededea
' Date      : 09/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuOpenFolder_click()
    Dim selectedKey As String
    Dim rightClickHighlightedKey As String
    Dim normallySelectedKey As String
    
    On Error GoTo mnuOpenFolder_click_Error
    'On Error Resume Next
    
    If Not (folderTreeView.DropHighlight Is Nothing) Then
        selectedKey = folderTreeView.DropHighlight.Key
    ElseIf Not (folderTreeView.SelectedItem Is Nothing) Then
        selectedKey = folderTreeView.SelectedItem.Key
    Else
        Exit Sub
    End If
    
    If DirExists(selectedKey) Then
        ShellExecute 0, vbNullString, selectedKey, vbNullString, vbNullString, 1
    End If
    mnuOpenFolder.Visible = False
    blank10.Visible = False
   On Error GoTo 0
   Exit Sub

mnuOpenFolder_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuOpenFolder_click of Form rDIconConfigForm"
End Sub



