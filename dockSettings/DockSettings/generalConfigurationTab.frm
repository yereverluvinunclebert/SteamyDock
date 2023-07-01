VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form generalConfigurationTab 
   BorderStyle     =   0  'None
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
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
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "These are the main settings for the dock"
      Top             =   0
      Width           =   6930
      Begin VB.Frame fraWriteOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   510
         TabIndex        =   24
         Top             =   2340
         Width           =   6165
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
            TabIndex        =   27
            ToolTipText     =   $"generalConfigurationTab.frx":0000
            Top             =   615
            Width           =   225
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
            TabIndex        =   26
            ToolTipText     =   $"generalConfigurationTab.frx":0095
            Top             =   300
            Width           =   5500
         End
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
            TabIndex        =   25
            ToolTipText     =   "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
            Top             =   0
            Width           =   5500
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
            TabIndex        =   28
            ToolTipText     =   "Writes the configuration data to a new location that is compatible with the methods used by current Windows"
            Top             =   615
            Width           =   5445
         End
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
         TabIndex        =   23
         ToolTipText     =   "This will cause the current dock to run when Windows starts"
         Top             =   360
         Width           =   1440
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
         TabIndex        =   22
         ToolTipText     =   "This allows running applications to appear in the dock"
         Top             =   3585
         Width           =   255
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
         TabIndex        =   21
         ToolTipText     =   "This is an essential option that stops you accidentally deleting your dock icons, click it!"
         Top             =   5865
         Width           =   4620
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
         ItemData        =   "generalConfigurationTab.frx":0158
         Left            =   2085
         List            =   "generalConfigurationTab.frx":0162
         TabIndex        =   20
         Text            =   "Rocketdock"
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock, these utilities are compatible with both"
         Top             =   6255
         Width           =   2310
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
         TabIndex        =   19
         Text            =   "C:\programs"
         ToolTipText     =   $"generalConfigurationTab.frx":017E
         Top             =   6960
         Width           =   4710
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
         TabIndex        =   18
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   5520
         Width           =   240
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
         TabIndex        =   17
         ToolTipText     =   "If you dislike the minimise animation, click this"
         Top             =   3915
         Value           =   1  'Checked
         Width           =   255
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
         TabIndex        =   16
         ToolTipText     =   $"generalConfigurationTab.frx":0214
         Top             =   4350
         Width           =   2985
      End
      Begin VB.CommandButton btnGeneralRdFolder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   5745
         TabIndex        =   15
         ToolTipText     =   "Select the folder location of Rocketdock here"
         Top             =   6975
         Visible         =   0   'False
         Width           =   300
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
         TabIndex        =   14
         ToolTipText     =   $"generalConfigurationTab.frx":02B3
         Top             =   7395
         Width           =   210
      End
      Begin VB.Frame fraRunAppIndicators 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   450
         TabIndex        =   8
         Top             =   4635
         Width           =   5955
         Begin CCRSlider.Slider sliGenRunAppInterval 
            Height          =   315
            Left            =   1020
            TabIndex        =   9
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
            TabIndex        =   13
            ToolTipText     =   "This function consumes cpu on  low power computers so keep it above 15 secs, preferably 30."
            Top             =   120
            Width           =   3210
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
            TabIndex        =   12
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   1215
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
            TabIndex        =   11
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   585
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
            TabIndex        =   10
            ToolTipText     =   "The maximum time a VB6 timer can extend to is 65,536 ms or 65 seconds"
            Top             =   495
            Width           =   630
         End
      End
      Begin VB.Frame fraReadOptionButtons 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   540
         TabIndex        =   3
         Top             =   900
         Width           =   6315
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
            TabIndex        =   6
            ToolTipText     =   $"generalConfigurationTab.frx":034C
            Top             =   780
            Width           =   225
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
            TabIndex        =   5
            ToolTipText     =   $"generalConfigurationTab.frx":03E1
            Top             =   465
            Width           =   5500
         End
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
            TabIndex        =   4
            ToolTipText     =   "Store configuration in Rocketdock's program files folder, causes security issues and requires admin access,"
            Top             =   165
            Width           =   5500
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
            TabIndex        =   7
            ToolTipText     =   "Reads the configuration data from a new location that is compatible with the methods used by current Windows"
            Top             =   780
            Width           =   6135
         End
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
         TabIndex        =   2
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   7815
         Width           =   255
      End
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
         TabIndex        =   1
         ToolTipText     =   "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
         Top             =   8145
         Width           =   225
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
         TabIndex        =   38
         ToolTipText     =   "Choose which dock you are using Rocketdock or SteamyDock - currently not operational, defaults to Rocketdock"
         Top             =   6300
         Width           =   1530
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
         TabIndex        =   37
         ToolTipText     =   $"generalConfigurationTab.frx":04A4
         Top             =   6690
         Width           =   1695
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
         TabIndex        =   36
         Top             =   3660
         Width           =   3510
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
         TabIndex        =   35
         Top             =   3975
         Width           =   2505
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
         TabIndex        =   34
         ToolTipText     =   "If you click on an icon that is already running then it can open it or fire up another instance"
         Top             =   5595
         Width           =   3465
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
         TabIndex        =   33
         ToolTipText     =   $"generalConfigurationTab.frx":056E
         Top             =   7455
         Width           =   5310
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
         TabIndex        =   32
         Top             =   720
         Width           =   1800
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1800
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
         TabIndex        =   30
         ToolTipText     =   "Show Splash Screen on Start-up"
         Top             =   7815
         Width           =   3870
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
         TabIndex        =   29
         ToolTipText     =   "When you drag or add an item to the dock it will always show the icon settings utility unless you disable it here"
         Top             =   8145
         Width           =   4995
      End
   End
End
Attribute VB_Name = "generalConfigurationTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
