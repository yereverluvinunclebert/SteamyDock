VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form rDIconConfigForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rocketdock Icon Settings"
   ClientHeight    =   8880
   ClientLeft      =   150
   ClientTop       =   -135
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rocketdockIconConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame dllFrame 
      Height          =   705
      Left            =   660
      TabIndex        =   69
      Top             =   7530
      Width           =   2940
      Begin VB.Label Arses 
         Caption         =   "Embedded icons within DLLs and EXEs currently only show as 32bpp"
         Height          =   450
         Left            =   135
         TabIndex        =   70
         Top             =   180
         Width           =   2745
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Configuration"
      Height          =   600
      Index           =   0
      Left            =   2100
      TabIndex        =   57
      Top             =   4515
      Visible         =   0   'False
      Width           =   1665
      Begin VB.CheckBox chkBiLinear 
         Caption         =   "Quality Sizing"
         Height          =   240
         Left            =   90
         TabIndex        =   58
         ToolTipText     =   "Stretch Quality Option"
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Drag && Drop, Copy && Paste too.  Unicode Compatible"
         Height          =   255
         Index           =   3
         Left            =   45
         TabIndex        =   59
         ToolTipText     =   "To Paste: Click on display box and press Ctrl+V"
         Top             =   5865
         Width           =   3840
      End
   End
   Begin VB.CommandButton btnWorking 
      Caption         =   "Working"
      Height          =   510
      Left            =   8100
      TabIndex        =   56
      Top             =   3915
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer registryTimer 
      Interval        =   2500
      Left            =   3180
      Top             =   7095
   End
   Begin VB.PictureBox picRdThumbFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   75
      ScaleHeight     =   705
      ScaleWidth      =   9660
      TabIndex        =   48
      Top             =   4545
      Visible         =   0   'False
      Width           =   9660
      Begin VB.HScrollBar rdMapHScroll 
         Height          =   120
         Left            =   15
         Max             =   100
         TabIndex        =   53
         Top             =   570
         Visible         =   0   'False
         Width           =   9630
      End
      Begin VB.CommandButton btnMapNext 
         Caption         =   ">"
         Height          =   450
         Left            =   9210
         TabIndex        =   50
         ToolTipText     =   "Scroll the RD map to the left (or press HOME)"
         Top             =   60
         Width           =   450
      End
      Begin VB.CommandButton btnMapPrev 
         Caption         =   "<"
         Height          =   450
         Left            =   45
         TabIndex        =   49
         ToolTipText     =   "Scroll the RD Map to the left (or press END)"
         Top             =   45
         Width           =   435
      End
      Begin VB.PictureBox picCover 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   -45
         ScaleHeight     =   555
         ScaleWidth      =   570
         TabIndex        =   51
         Top             =   0
         Width           =   570
      End
      Begin VB.PictureBox back 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   0
         Left            =   9150
         ScaleHeight     =   555
         ScaleWidth      =   600
         TabIndex        =   52
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox picRdMap 
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
         Height          =   500
         Index           =   0
         Left            =   540
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   65
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
      Picture         =   "rocketdockIconConfig.frx":1856A
      ScaleHeight     =   180
      ScaleWidth      =   270
      TabIndex        =   47
      ToolTipText     =   "Hide the map"
      Top             =   4485
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton rdMapRefresh 
      Height          =   270
      Left            =   9885
      Picture         =   "rocketdockIconConfig.frx":188C6
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Refresh the icon map"
      Top             =   4785
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame FrameFolders 
      Caption         =   "Folders"
      Height          =   4500
      Left            =   105
      TabIndex        =   16
      ToolTipText     =   "The current list of known icon folders"
      Top             =   15
      Width           =   4005
      Begin VB.PictureBox picThingCover 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1005
         ScaleHeight     =   375
         ScaleWidth      =   1200
         TabIndex        =   79
         Top             =   3975
         Width           =   1200
      End
      Begin VB.Frame frmRegistry 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1020
         TabIndex        =   74
         Top             =   3945
         Width           =   1230
         Begin VB.CheckBox chkRegistry 
            Height          =   255
            Left            =   45
            TabIndex        =   76
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   90
            Width           =   165
         End
         Begin VB.CheckBox chkSettings 
            Height          =   225
            Left            =   645
            TabIndex        =   75
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   105
            Width           =   210
         End
         Begin VB.Label Label10 
            Caption         =   "Reg."
            Height          =   240
            Left            =   270
            TabIndex        =   78
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   105
            Width           =   450
         End
         Begin VB.Label Label11 
            Caption         =   "Set."
            Height          =   240
            Left            =   870
            TabIndex        =   77
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   105
            Width           =   240
         End
      End
      Begin VB.TextBox textCurrentFolder 
         Height          =   330
         Left            =   105
         TabIndex        =   33
         Text            =   "textCurrentFolder"
         ToolTipText     =   "The selected folder path"
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton btnRemoveFolder 
         Caption         =   "-"
         Height          =   345
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "This button can remove a custom folder from the treeview above"
         Top             =   3990
         Width           =   390
      End
      Begin ComctlLib.TreeView treeView 
         Height          =   3210
         Left            =   105
         TabIndex        =   24
         ToolTipText     =   "These are the icon folders available to Rocketdock"
         Top             =   630
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   5662
         _Version        =   327682
         HideSelection   =   0   'False
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton btnAddFolder 
         Caption         =   "+"
         Height          =   345
         Left            =   555
         TabIndex        =   17
         ToolTipText     =   "Select a target folder to add to the treeview list above"
         Top             =   3990
         Width           =   390
      End
      Begin VB.Label nLabel 
         Caption         =   "0"
         Height          =   270
         Left            =   2640
         TabIndex        =   73
         Top             =   4155
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label xlabel 
         Caption         =   "0"
         Height          =   210
         Left            =   2640
         TabIndex        =   72
         Top             =   3915
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.PictureBox btnArrowDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9735
      Picture         =   "rocketdockIconConfig.frx":18D09
      ScaleHeight     =   180
      ScaleWidth      =   375
      TabIndex        =   46
      ToolTipText     =   "Show the Rocketdock Map"
      Top             =   4485
      Width           =   375
   End
   Begin VB.Frame frameIcons 
      Caption         =   "Icons"
      Height          =   4500
      Left            =   4230
      TabIndex        =   19
      Top             =   15
      Width           =   5895
      Begin VB.CommandButton btnAdd 
         Caption         =   "+"
         Height          =   270
         Left            =   4755
         TabIndex        =   71
         ToolTipText     =   "Set the current selected icon into the dock (double-click on the icon)"
         Top             =   240
         Width           =   270
      End
      Begin VB.CommandButton btnRefresh 
         Height          =   270
         Left            =   5085
         Picture         =   "rocketdockIconConfig.frx":190D0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Refresh the Icon List"
         Top             =   240
         Width           =   210
      End
      Begin VB.CommandButton btnKillIcon 
         Height          =   255
         Left            =   165
         Picture         =   "rocketdockIconConfig.frx":19513
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Delete the currently selected icon file"
         Top             =   4020
         Width           =   240
      End
      Begin VB.TextBox textCurrIconPath 
         Height          =   330
         Left            =   1035
         TabIndex        =   25
         Text            =   "textCurrIconPath"
         ToolTipText     =   "Shows the selected icon file name"
         Top             =   210
         Width           =   3660
      End
      Begin VB.ComboBox comboIconTypesFilter 
         Height          =   345
         ItemData        =   "rocketdockIconConfig.frx":19740
         Left            =   510
         List            =   "rocketdockIconConfig.frx":19756
         TabIndex        =   21
         Text            =   "All Normal Icons"
         ToolTipText     =   "Filter icon types to display"
         Top             =   3975
         Width           =   2790
      End
      Begin VB.CommandButton btnGetMore 
         BackColor       =   &H8000000A&
         Caption         =   "Get More"
         Height          =   345
         Left            =   3960
         TabIndex        =   20
         ToolTipText     =   "Click to install more icons"
         Top             =   3975
         Width           =   1710
      End
      Begin VB.CommandButton btnThumbnailView 
         Height          =   270
         Left            =   5355
         Picture         =   "rocketdockIconConfig.frx":197C1
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "View as thumbnails"
         Top             =   240
         Width           =   270
      End
      Begin VB.CommandButton btnTreeView 
         Height          =   255
         Left            =   5355
         Picture         =   "rocketdockIconConfig.frx":199CF
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "View as a file listing"
         Top             =   255
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picFrameThumbs 
         BackColor       =   &H00FFFFFF&
         Height          =   3210
         Left            =   120
         ScaleHeight     =   3150
         ScaleWidth      =   5475
         TabIndex        =   34
         ToolTipText     =   "Double-click an icon to set it into the dock"
         Top             =   615
         Width           =   5535
         Begin VB.Frame frmThumbLabel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   350
            Index           =   0
            Left            =   60
            TabIndex        =   67
            Top             =   840
            Width           =   1185
            Begin VB.Label lblThumbName 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "01234567890123"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   68
               Top             =   -15
               Width           =   1095
               WordWrap        =   -1  'True
            End
         End
         Begin VB.VScrollBar vScrollThumbs 
            Height          =   3180
            LargeChange     =   12
            Left            =   5250
            SmallChange     =   4
            TabIndex        =   35
            Top             =   -15
            Width           =   240
         End
         Begin VB.PictureBox picIconList 
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
            Left            =   165
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   66
            ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
            Top             =   60
            Width           =   1000
         End
      End
      Begin VB.FileListBox filesIconList 
         Height          =   3240
         Left            =   105
         Pattern         =   "*.jpg"
         TabIndex        =   23
         ToolTipText     =   "Select an icon, double-click to set"
         Top             =   600
         Width           =   5565
      End
      Begin VB.Label Label7 
         Caption         =   "Icon Name:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.Frame frameProperties 
      Caption         =   "Properties"
      Height          =   3495
      Left            =   4230
      TabIndex        =   0
      Top             =   4530
      Width           =   5895
      Begin VB.PictureBox picBusy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   3660
         Picture         =   "rocketdockIconConfig.frx":19DAB
         ScaleHeight     =   795
         ScaleWidth      =   825
         TabIndex        =   81
         ToolTipText     =   "The program is doing something..."
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4335
         TabIndex        =   45
         ToolTipText     =   "Sets the icon characteristics but you will need to restart to make it 'fix'"
         Top             =   3030
         Width           =   1470
      End
      Begin VB.TextBox txtCurrentIcon 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1395
         TabIndex        =   29
         Text            =   "txtCurrentIcon"
         ToolTipText     =   "The name of the icon as it appears on the dock"
         Top             =   690
         Width           =   4305
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   22
         ToolTipText     =   "Select a target folder"
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
         TabIndex        =   18
         ToolTipText     =   "Select a target file"
         Top             =   1080
         Width           =   345
      End
      Begin VB.CheckBox checkPopupMenu 
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
         Top             =   3030
         Width           =   180
      End
      Begin VB.ComboBox comboOpenRunning 
         Height          =   345
         ItemData        =   "rocketdockIconConfig.frx":1A801
         Left            =   1395
         List            =   "rocketdockIconConfig.frx":1A80E
         TabIndex        =   6
         Text            =   "Use Global Setting"
         ToolTipText     =   "Choose what to do if the chosen app is already running"
         Top             =   2640
         Width           =   2145
      End
      Begin VB.ComboBox comboRun 
         Height          =   345
         ItemData        =   "rocketdockIconConfig.frx":1A835
         Left            =   1395
         List            =   "rocketdockIconConfig.frx":1A842
         TabIndex        =   5
         Text            =   "Normal"
         ToolTipText     =   "Window mode for the program to operate within"
         Top             =   2250
         Width           =   2145
      End
      Begin VB.TextBox lblArguments 
         Height          =   345
         Left            =   1395
         TabIndex        =   4
         ToolTipText     =   "Add any additional arguments that the target file operation requires eg. -s -t 00 -f "
         Top             =   1860
         Width           =   2130
      End
      Begin VB.TextBox lblStartIn 
         Height          =   345
         Left            =   1395
         TabIndex        =   3
         ToolTipText     =   "If the operation needs to be performed in a particular folder select it here"
         Top             =   1470
         Width           =   3915
      End
      Begin VB.TextBox lblTarget 
         Height          =   345
         Left            =   1395
         TabIndex        =   2
         ToolTipText     =   "The target you wish to run, a file or a folder"
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox lblName 
         Height          =   345
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "The name of the icon as it appears on the dock"
         Top             =   300
         Width           =   4305
      End
      Begin VB.Label Label8 
         Caption         =   "Current Icon:"
         Height          =   225
         Left            =   345
         TabIndex        =   30
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label lblRdIconNumber 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1035
         Left            =   4440
         TabIndex        =   27
         ToolTipText     =   "This is Rocketdock icon number one."
         Top             =   1830
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Popup Menu:"
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Display Special Actions"
         Height          =   225
         Left            =   1665
         TabIndex        =   14
         ToolTipText     =   "If you want extra options to appear when you right click on an icon, enable this checkbox"
         Top             =   3030
         Width           =   1965
      End
      Begin VB.Label Label5 
         Caption         =   "Arguments:"
         Height          =   225
         Left            =   450
         TabIndex        =   13
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Start in:"
         Height          =   225
         Left            =   735
         TabIndex        =   12
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Run:"
         Height          =   225
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Open Running:"
         Height          =   225
         Left            =   225
         TabIndex        =   10
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Target:"
         Height          =   225
         Index           =   0
         Left            =   780
         TabIndex        =   9
         Top             =   1110
         Width           =   1215
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
      Height          =   4275
      Left            =   105
      TabIndex        =   60
      Top             =   4530
      Width           =   4000
      Begin ComctlLib.Slider sliPreviewSize 
         Height          =   300
         Left            =   60
         TabIndex        =   64
         ToolTipText     =   "Icon Size"
         Top             =   3795
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   529
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         Max             =   5
         SelStart        =   4
         TickStyle       =   1
         Value           =   4
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
         TabIndex        =   63
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
         Left            =   3735
         TabIndex        =   62
         ToolTipText     =   "select the next icon"
         Top             =   255
         Width           =   195
      End
      Begin VB.PictureBox picPreview 
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
         Height          =   3450
         Left            =   270
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   230
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   230
         TabIndex        =   61
         ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
         Top             =   330
         Width           =   3450
      End
   End
   Begin VB.Frame frameButtons 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   210
      TabIndex        =   37
      Top             =   7680
      Width           =   10080
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate Dock"
         Height          =   360
         Left            =   3960
         TabIndex        =   80
         ToolTipText     =   "Makes a whole NEW rocketdock - use with care!"
         Top             =   750
         Width           =   1755
      End
      Begin VB.CommandButton btnBackup 
         Caption         =   "Backup"
         Height          =   345
         Left            =   6885
         TabIndex        =   55
         ToolTipText     =   "Backup or create bkpSettings.ini"
         Top             =   405
         Width           =   1485
      End
      Begin VB.CommandButton btnSaveRestart 
         Caption         =   "Save && Restart"
         Height          =   345
         Left            =   6885
         TabIndex        =   44
         ToolTipText     =   "A save and restart of Rocketdock is required when any icon changes have been made"
         Top             =   765
         Width           =   1485
      End
      Begin VB.CommandButton btnCloseCancel 
         Caption         =   " Close"
         Height          =   345
         Left            =   8385
         TabIndex        =   41
         ToolTipText     =   "Cancel the current operation and close the window"
         Top             =   765
         Width           =   1470
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help"
         Height          =   345
         Left            =   8385
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         ToolTipText     =   "Help on this utility"
         Top             =   405
         Width           =   1470
      End
      Begin VB.CheckBox chkConfirmSaves 
         Height          =   225
         Left            =   3975
         TabIndex        =   39
         ToolTipText     =   "Confirmation on saves and deletes"
         Top             =   465
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.CommandButton btnDefaultIcon 
         Caption         =   "Default Icon"
         Height          =   330
         Left            =   1830
         TabIndex        =   38
         ToolTipText     =   "Not implemented yet"
         Top             =   435
         Width           =   1725
      End
      Begin VB.Label Label9 
         Caption         =   "Toggle info. dialogs"
         Height          =   240
         Left            =   4230
         TabIndex        =   42
         ToolTipText     =   "This will turn off most of the information pop-ups"
         Top             =   450
         Width           =   1410
      End
   End
   Begin VB.Menu mnuMainOpts 
      Caption         =   "Other Options"
      Visible         =   0   'False
      Begin VB.Menu mnuOtherOpts 
         Caption         =   "GDI"
         Begin VB.Menu mnuSubOpts 
            Caption         =   "Don't Use GDI+"
            Index           =   0
         End
         Begin VB.Menu mnuSubOpts 
            Caption         =   "Use GDI+"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSaveOpts 
         Caption         =   "Save As"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using GDI+)"
            Index           =   0
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As PNG (Using zLIB)"
            Index           =   1
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Default Filter"
               Index           =   0
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use No Filters (Fastest)"
               Index           =   1
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Left Filter"
               Index           =   2
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Top Filter"
               Index           =   3
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adjacent Average Filter"
               Index           =   4
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Paeth Filter"
               Index           =   5
            End
            Begin VB.Menu mnuZlibPng 
               Caption         =   "Use Adaptive Filtering (Slowest)"
               Index           =   6
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As JPG"
            Index           =   2
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As TGA"
            Index           =   3
            Begin VB.Menu mnuTGA 
               Caption         =   "Compressed"
               Index           =   0
            End
            Begin VB.Menu mnuTGA 
               Caption         =   "Uncompressed"
               Index           =   1
            End
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As GIF"
            Index           =   4
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As BMP (Red Bkg)"
            Index           =   5
         End
         Begin VB.Menu mnuSaveAs 
            Caption         =   "Save As Rendered Example (GDI+ required)"
            Index           =   7
         End
      End
      Begin VB.Menu mnuPos 
         Caption         =   "Position"
         Enabled         =   0   'False
         Begin VB.Menu mnuPosSub 
            Caption         =   "Centered"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Top Left"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Top Right"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Bottom Left"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu mnuPosSub 
            Caption         =   "Bottom Right"
            Enabled         =   0   'False
            Index           =   4
         End
      End
   End
   Begin VB.Menu rdMapMenu 
      Caption         =   "The Map Menu"
      Visible         =   0   'False
      Begin VB.Menu menuDelete 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu menuAdd 
         Caption         =   "Add Item"
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
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About this utility"
         Index           =   1
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with paypal"
         Index           =   2
      End
      Begin VB.Menu mnuSweets 
         Caption         =   "Donate some sweets/candy with Amazon"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Utility Help"
         Index           =   4
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuLatest 
            Caption         =   "Download Latest Version"
         End
         Begin VB.Menu mnuSupport 
            Caption         =   "Contact Support"
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
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
   End
   Begin VB.Menu thumbmenu 
      Caption         =   "Thumb Menu"
      Visible         =   0   'False
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
Option Explicit
'---------------------------------------------------------------------------------------
' Form Module : rDIconConfigFrm
' Author      : Dean Beedell
' Date        : 20/06/2019
'
' Credits : LA Volpe (VB Forums) for his transparent picture handling
'           LA Volpe (VB Forums) for his manifest code
'           Shuja Ali (codeguru.com) for his settings.ini code
'           killApp code from an unknown, untraceable source, possibly on MSN
'           registry reading code from ALLAPI.COM
'           Punklabs for the original inspiration and for Rocketdock
'
'   Built using MZ-TOOLS, CodeHelp Core IDE Extender Framework & Rubberduck
'---------------------------------------------------------------------------------------

Private cImage As c32bppDIB
Private cShadow As c32bppDIB

' Note: If GDI+ is available, it is more efficient for you to
' create the token then pass the token to each class.  Not required,
' but if you don't do this, then the classes will create and destroy
' a token everytime GDI+ is used to render or modify an image.
' Passing the token can result in up to 3x faster processing overall
Private m_GDItoken As Long
'



'My Computer     ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}
'My Network Places   ::{208D2C60-3AEA-1069-A2D7-08002B30309D}
'Internet Explorer   ::{871C5380-42A0-1069-A2EA-08002B30309D}
'Recycle Bin     ::{645FF040-5081-101B-9F08-00AA002F954E}


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : The initial subroutine for the program
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim NameProcess As String
    Dim AppExists As Boolean
    Dim ans As Integer
    
    On Error GoTo Form_Load_Error

    ReDim thumbArray(12) As Integer
        
    iconChanged = False
    rdMapStart = 0
    dotCount = 0 ' a variable used on the 'working...' button
    rdIconNumber = 0
    rdIconMax = 0  ' the final icon in the registry/settings
    rdMapIndex = 0 ' the starting point for the rocketdock map
    icoSizePreset = 128
    thumbImageSize = 64
    boxSpacing = 540
    storedIndex = 99
    busyCounter = 1

    ' add any remaining types that Rocketdock's code supports
    validIconTypes = "*.jpg;*.jpeg;*.bmp;*.ico;*.png;*.tif;*.gif"
        
    ' state and position of a few manually placed controls (easier here than in the IDE)
    picRdThumbFrame.Visible = False
                              
    'if the process already exists then kill it
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "rocket1.exe"
        checkAndKill NameProcess, ans
    End If
               
    'check the state of the licence
    Call checkLicenceState
    
    ' check the Windows version and where rocketdock is installed
    Call TestWinVer
        
    ' check where rocketdock is installed
    Call checkRocketdockInstallation
    
    ' set the default path to the icons root, this will be superceded later if the user has chosen a default folder
    filesIconList.path = rdAppPath & "\Icons" ' rdAppPath is defined in driveCheck above
    textCurrentFolder.Text = rdAppPath & "\Icons"
    relativePath = "\Icons"
        
    ' read the Rocketdock settings from INI or from registry
    Call readRocketDockSettings
        
    ' read the tool settings file and do some things for the first and only time
    Call readToolSettings
    
    ' dynamically create thumbnail picboxes and sort the captions
    Call createThumbnailLayout
    
    ' dynamically create rocketdock Map thumbnail picboxes
    Call createRdMapBoxes
        
    ' set the very large icon record number displayed on the main form
    rdIconNumber = 0
    lblRdIconNumber.Caption = rdIconNumber + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
            
    ' set the filter pattern to only show the icon types supported by Rocketdock
    filesIconList.Pattern = validIconTypes
    
    ' select that file in the file list
    filesIconList.ListIndex = 0

    ' this is needed why? TODO - confirm why
    picPreview.AutoRedraw = True
        
    ' add to the treeview the folders that exist below the RD icons folder and the user-created entries to the folder list top right
    Call addRocketdockFolders
        
    ' add the extra steampunk icon folders to the treeview
    Call setSteampunkLocation
    
    ' add the user custom folder to the treeview
    Call readCustomLocation
        
    ' extract the previously selected default folder in the treeview
    ' open the app settings.ini and read the default folder for the tool to display
    Call readDefaultFolder
    
    ' display the first icon in the preview window
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)
    
    ' select the thumbnail view rather than the file list view and populate it
    Call btnThumbnailView_Click
    
    ' we signify that all changes have been lost when changes to fields are made by the program and not the user
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form rDIconConfigForm"
                
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_KeyUp
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Simple example of pasting file names on a drag drop
   On Error GoTo Form_KeyUp_Error

    If KeyCode = vbKeyV Then
        
        If (Shift And vbCtrlMask) = vbCtrlMask Then
            ' use class to load 1st file that was pasted, if any & if more than one
            ' Unicode filenames supported
            If cImage.LoadPicture_PastedFiles(1, 256, 256) = False Then
                ' couldn't load anything from the files, maybe image itself was pasted
                If cImage.LoadPicture_ClipBoard = False Then
                    MsgBox "Failed to load whatever was placed in the clipboard", vbInformation + vbOKOnly
                    Exit Sub
                End If
            End If
            
            If Not cShadow Is Nothing Then
                '
            Else
                Call refreshPicBox(picPreview, 256)
            End If
            ShowImage False, True
        
        End If
    End If

   On Error GoTo 0
   Exit Sub

Form_KeyUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyUp of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnAdd_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAdd_Click()
    Dim ans As VbMsgBoxResult
    
   On Error GoTo btnAdd_Click_Error

    Call backupSettings
    
    If chkConfirmSaves.Value = 1 Then

        ans = MsgBox(" Confirm that you wish to set this icon as the current icon " & vbCr & "in the dock.", vbYesNo)
        If ans = 6 Then
            Call picIconList_DblClick(storedIndex)
        End If
    Else
        Call picIconList_DblClick(storedIndex)
    End If

   On Error GoTo 0
   Exit Sub

btnAdd_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAdd_Click of Form rDIconConfigForm"
End Sub
'


'---------------------------------------------------------------------------------------
' Procedure : btnGenerate_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGenerate_Click()
    Dim ans As VbMsgBoxResult
    Dim xFileName As String
    
    'Call btnArrowDown_Click ' populate the dock
   On Error GoTo btnGenerate_Click_Error

    ans = MsgBox("This will COMPLETELY wipe your dock and generate a new version from the software in your system. " & _
        " Note: It will only find software that is logged in the registry with a known installation location." & vbCr & vbCr & _
        " Are you sure you want to do this?", vbYesNo)
    
    If ans = 6 Then
        ' obtain the name of the software - DisplayName
        ' the program run folder name - InstallLocation
        ' the program start folder as above -InstallLocation
        ' the binary target name - DisplayIcon
        
        xFileName = App.path & "\ins.txt"
        Dim s As String
        s = GetInstalledApps()
        
        'add the list to a RTB control on a separate form
        formSoftwareList.rtbSoftwareList.Text = s
        formSoftwareList.Show
        
        'write the data to a local file as well
        Call WriteFile(s, xFileName)
        
        ' open the known software list
        
        ' for each item in the list
        '   set the name of the software - DisplayName
        '   set the program run folder name - InstallLocation
        '   set the program start folder as above -InstallLocation
        '   if the the binary target name is an exe then
        '       set the binary target name - DisplayIcon
        '
        '   identify and compare to the list
        '       the list comprises:
        '           known software names and the suitable icons
        '           a word like photo &c that could use similar icons
        '   if there is a match then use the found list
        '   else
        '       if the the binary target name is an exe then
        '           and use the EXE to extract the 32x32 icon
        '       if the the binary target name is an ico then use it directly for the icon
    End If

   On Error GoTo 0
   Exit Sub

btnGenerate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGenerate_Click of Form rDIconConfigForm"
End Sub

Private Sub btnNext_KeyDown(KeyCode As Integer, Shift As Integer)
    Call getkeypress(KeyCode)
End Sub

Private Sub btnPrev_KeyDown(KeyCode As Integer, Shift As Integer)
    Call getkeypress(KeyCode)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRemoveFolder_Click
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRemoveFolder_Click()
    ' add the chosen folder to the treeview
    
    ' find the chosen node's parent
    ' if the parent is nothing then this is the top level
    ' if the parent exists then look above again
    Dim a As String
    Dim tNode As Node
   On Error GoTo btnRemoveFolder_Click_Error

    Set tNode = treeView.selectedItem
    
    If Not tNode.Parent Is Nothing Then
        Set tNode = tNode.Parent '  move up level
        Do Until tNode.Parent Is Nothing   ' if Nothing, then done
            If Not tNode.Next Is Nothing Then
                Set tNode = tNode.Next
                Exit Do
            End If
            Set tNode = tNode.Parent ' move up again
        Loop
    End If
    
    If tNode Is Nothing Then
        'Set treeView.selectedItem = treeView.Nodes(1).Root
        a = treeView.Nodes(1).Root
    Else
        a = tNode
        'Set treeView.selectedItem = tNode
    End If
    
    If a = "my collection" Then
            MsgBox "Cannot remove Rocketdock Enhanced Settings Utility sub-folders from the treeview."
            Exit Sub
    End If
        
    If a = "icons" Then
        MsgBox "Cannot remove Rocketdock's own sub-folders from the treeview, you have to delete the folders from Windows first then re-run this utility."
        Exit Sub
    End If
        
    If treeView.selectedItem.Key = "" Then
        Exit Sub
    End If
        
    If a = "custom folder" And Not treeView.selectedItem = "custom folder" Then
        MsgBox "Cannot remove custom sub-folders from the treeview, try again at the root."
        Exit Sub
    Else
        ' do the delete!
    End If
        
    treeView.Nodes.Remove treeView.selectedItem.Key
    
    'write the folder to the rocketdock settings file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\RocketDock", "CustomIconFolder", "?", rdSettingsFile

   On Error GoTo 0
   Exit Sub

btnRemoveFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRemoveFolder_Click of Form rDIconConfigForm"
End Sub

Private Sub busyTimer_Timer()
'    Dim busyFilename As String
'
'    picBusy.ToolTipText = "program is doing something"
'    busyCounter = busyCounter + 1
'    If busyCounter >= 7 Then busyCounter = 1
'    busyFilename = App.path & "\busy-F" & busyCounter & "-32x32x24.jpg"
'    picBusy.Picture = LoadPicture(busyFilename)
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
    
    Dim useloop As Integer
    
   On Error GoTo createThumbnailLayout_Error

    storeLeft = 165
    frmThumbLabel(0).ZOrder
    frmThumbLabel(0).BorderStyle = 0
    frmThumbLabel(0).Visible = True
        
    ' dynamically create the picture boxes for the thumbnails
    For useloop = 1 To 11 ' 0 is the template
        Load picIconList(useloop)
        Load frmThumbLabel(useloop)
        Load lblThumbName(useloop)
        
         Set lblThumbName(useloop).Container = frmThumbLabel(useloop)
    Next useloop
    
    placeThumbnailPicboxes
    
    ' the labels for the smaller thumbnail icon view
    For useloop = 0 To 11
        frmThumbLabel(useloop).Visible = False
    Next useloop

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
    Dim useloop As Integer
    
   On Error GoTo createRdMapBoxes_Error

    storeLeft = boxSpacing
    ' dynamically create more picture boxes to the maximum number of icons
    For useloop = 1 To rdIconMax
        'If useloop > 0 Then
        Load picRdMap(useloop)
        picRdMap(useloop).Width = 500
        picRdMap(useloop).Height = 500
        storeLeft = storeLeft + boxSpacing
        picRdMap(useloop).Left = storeLeft
        picRdMap(useloop).Top = 15
        picRdMap(useloop).Visible = True
        picRdMap(useloop).AutoRedraw = True
    Next useloop
    
    'if the map only has a very small number of icons we have to create a few placeholder boxes
    If rdIconMax < 17 Then
        For useloop = (rdIconMax + 1) To 17
            'If useloop > 0 Then
            Load picRdMap(useloop) ' dynamically create more empty picture boxes until the end
            picRdMap(useloop).Width = 500
            picRdMap(useloop).Height = 500
            storeLeft = storeLeft + boxSpacing
            picRdMap(useloop).Left = storeLeft
            picRdMap(useloop).Top = 15
            picRdMap(useloop).Visible = True
            picRdMap(useloop).BackColor = &HC0C0C0
        Next useloop
    End If
    

   On Error GoTo 0
   Exit Sub

createRdMapBoxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRdMapBoxes of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : readToolSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readToolSettings()
    Dim sfirst As String
 
   On Error GoTo readToolSettings_Error

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        'test to see if the tool has ever been run before
        'if it has not done so before then
        sfirst = GetINISetting("Software\RocketDockSettings", "First", toolSettingsFile)
        If sfirst = True Then
        
            sfirst = False
            
            ' insert at the final position
            ' a link with the tardis icon and the target
            ' is the app.path
            
            sFilename = "Icons\tardis.png" ' the default Rocketdock filename for a blank item
            sTitle = "Rocket Settings"
            sCommand = App.path & "\" & "rocket1.exe"
            sArguments = ""
            sWorkingDirectory = App.path
            sShowCmd = 0
            sOpenRunning = 0
            sUseContext = 0
            
            rdIconMax = rdIconMax + 1
            
            'write the rdsettings file
            writeSettingsIni (rdIconMax)

            'amend the count in both the rdSettings.ini
            PutINISetting "Software\RocketDock\Icons", "count", rdIconMax, rdSettingsFile

            'write the updated test of first run to false
            PutINISetting "Software\RocketDockSettings", "First", sfirst, toolSettingsFile
        
            'filecopy the tardis png to the rocketdock icons folder
            If FExists(App.path & "\" & "tardis.png") Then
                FileCopy App.path & "\" & "tardis.png", rdAppPath & "\icons\" & "tardis.png"
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

readToolSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readToolSettings of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readRocketDockSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRocketDockSettings()
    
    ' SETTINGS: There are three settings files
    ' the first is the RD settings file that only exists if RD is NOT using the registry
    ' the second is our tools copy of RD's settings file, we copy the original or create our own from RD's registry settings
    ' the third is the settings file for this tool to store its own preferences
        
    ' check to see if the settings file exists
    ' (Rocketdock overwrites its own settings.ini when it closes meaning that we have to work on a copy).
   On Error GoTo readRocketDockSettings_Error

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    rdSettingsFile = rdAppPath & "\rdSettings.ini" ' a copy of the settings file that we work on
        
    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
        
        Call backupSettings ' make a backup of the settings.ini file each restart
        
        ' copy the original settings file to a duplicate that we will operate upon
        FileCopy origSettingsFile, rdSettingsFile
        
        ' read the rocketdock settings.ini and find the very last icon
        rdIconMax = GetINISetting("Software\RocketDock\Icons", "count", rdSettingsFile) - 1
    Else
        chkRegistry.Value = 1
        chkSettings.Value = 0
        
        ' read the rocketdock registry and find the last icon
        rdIconMax = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count") - 1
        
        ' copy the original configs out of the registry and into a settings file that we will operate upon
        readRegistryWriteSettings
        
        ' make a backup of the rdSettings.ini after the intermediate file has been created
        Call backupSettings
        
    End If

   On Error GoTo 0
   Exit Sub

readRocketDockSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRocketDockSettings of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkRocketdockInstallation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkRocketdockInstallation()
    Dim answer As VbMsgBoxResult
    
    ' check where rocketdock is installed
   On Error GoTo checkRocketdockInstallation_Error

    RD86installed = driveCheck("Program Files (x86)\Rocketdock")
    RDinstalled = driveCheck("Program Files\Rocketdock")
    
    If RDinstalled = False And RD86installed = False Then
        answer = MsgBox(" Rocketdock has not been installed in the program files (x86) folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
         Dim ofrm As Form
         For Each ofrm In Forms
             Unload ofrm
         Next
         End
    Else
        'If RD86installed = True Then MsgBox "Rocketdock is installed in program files (x86)"
        'If RDinstalled = True Then MsgBox "Rocketdock is installed in program files"
    End If

   On Error GoTo 0
   Exit Sub

checkRocketdockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Form rDIconConfigForm"
End Sub

'check the state of the licence
'---------------------------------------------------------------------------------------
' Procedure : checkLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkLicenceState()
    Dim slicence As Integer

   On Error GoTo checkLicenceState_Error

    toolSettingsFile = App.path & "\settings.ini"
    ' read the tool's own settings file (
    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = GetINISetting("Software\RocketDockSettings", "Licence", toolSettingsFile)
        ' if the licence state is not already accepted then display the licence form
        If slicence = 0 Then
            licence.Show vbModal ' show the licence screen in VB modal mode (ie. on its own)
            ' on the licence box change the state fo the licence acceptance
        End If
    End If
    
    ' show the licence screen if it has never been run before and set it to be in focus
    If licence.Visible = True Then
        licence.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

checkLicenceState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkLicenceState of Form rDIconConfigForm"

End Sub


' add the extra steampunk icon folders to the treeview
'---------------------------------------------------------------------------------------
' Procedure : setSteampunkLocation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setSteampunkLocation()

    ' add the custom folder to the treeview
    Dim SteampunkIconFolder As String
    
   On Error GoTo setSteampunkLocation_Error

    SteampunkIconFolder = App.path & "\my collection"
    
    If DirExists(SteampunkIconFolder) Then
        ' add the chosen folder to the treeview
        treeView.Nodes.Add , , SteampunkIconFolder, SteampunkIconFolder
        Call addtotree(SteampunkIconFolder, treeView)
        treeView.Nodes(SteampunkIconFolder).Text = "my collection"
    End If

   On Error GoTo 0
   Exit Sub

setSteampunkLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setSteampunkLocation of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : placeThumbnailPicboxes
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub placeThumbnailPicboxes()
    Dim useloop As Integer
    Dim storeTop As Integer
        
    
    ' place the thumbnails picboxes where they should go
   On Error GoTo placeThumbnailPicboxes_Error

    picIconList(0).Width = 1000
    picIconList(0).Height = 1000
    picIconList(0).Left = 165
    picIconList(0).Top = 60
    
    For useloop = 1 To 11
        
        picIconList(useloop).Width = 1000
        picIconList(useloop).Height = 1000
        frmThumbLabel(useloop).BorderStyle = 0
        
        picIconList(useloop).ToolTipText = filesIconList.List(useloop)
        If useloop = 1 Then
            storeLeft = 1365
            storeTop = 30
        End If
        
        If useloop = 4 Then
            storeLeft = 165
            storeTop = 1060
        End If
                
        If useloop = 8 Then
            storeLeft = 165
            storeTop = 2100
        End If
        
        picIconList(useloop).Left = storeLeft
        picIconList(useloop).Top = storeTop
        
        frmThumbLabel(useloop).Left = storeLeft - 100
        frmThumbLabel(useloop).Top = storeTop + 800
       
        storeLeft = storeLeft + 1200

        picIconList(useloop).Visible = True
        lblThumbName(useloop).Visible = True
        frmThumbLabel(useloop).Visible = True
        
        picIconList(useloop).ZOrder
        frmThumbLabel(useloop).ZOrder
        
        picIconList(useloop).AutoRedraw = True
    Next useloop
    

   On Error GoTo 0
   Exit Sub

placeThumbnailPicboxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure placeThumbnailPicboxes of Form rDIconConfigForm"

End Sub
    'carry end


Private Sub busyStart()
        Me.MousePointer = 11
End Sub

Private Sub busyStop()
        Me.MousePointer = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnArrowDown_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'
'
Private Sub btnArrowDown_Click()
    Dim growBit As Integer
    Dim a As String
        
    'On Error GoTo btnArrowDown_MouseDown_Error
   On Error GoTo btnArrowDown_Click_Error

    
    growBit = 670
            
    btnArrowDown.Visible = False
            
    If picRdThumbFrame.Visible = False Then
        'has to do this first as redrawing errors occur otherwise

        btnWorking.Visible = True
        Call busyStart
        'If picRdMap(0).Picture = 0 Then ' only recreate the map if the array is empty
        ' we used to check the .picture property but using lavolpe's 2nd method this proprty is not set.
        ' now we check for the tooltiptext which is only set when the image is populated.
        If picRdMap(0).ToolTipText = "" Then ' only recreate the map if the array is empty
            Call populateRdMap(0) ' show the map from position zero
        End If
        
        framePreview.Top = 4545 + growBit
        frameProperties.Top = 4545 + growBit
        frameButtons.Top = 7680 + growBit
        rDIconConfigForm.Height = 9255 + growBit
        rDIconConfigForm.dllFrame.Top = 7530 + growBit
                
        btnArrowUp.Visible = True
        picRdThumbFrame.Visible = True
        
        rdMapRefresh.Visible = True
        If rdIconMax > 16 Then
            rdMapHScroll.Visible = True
        End If
        
        rdMapHScroll.max = rdIconMax
                
        ' we signify that all changes have been lost
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
        
        btnWorking.Visible = False
  
        busyStop
    End If

   On Error GoTo 0
   Exit Sub

btnArrowDown_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowDown_Click of Form rDIconConfigForm"

End Sub

Private Sub btnArrowDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        btnWorking.Visible = True
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

    If picRdThumbFrame.Visible = True Then

        framePreview.Top = 4545
        frameProperties.Top = 4545
        frameButtons.Top = 7680
        rDIconConfigForm.Height = 9255
        rDIconConfigForm.dllFrame.Top = 7530
        
        btnArrowDown.Visible = True
        btnArrowUp.Visible = False
        picRdThumbFrame.Visible = False
        btnArrowDown.ToolTipText = "Show the Rocketdock Map"
        rdMapRefresh.Visible = False
        rdMapHScroll.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

btnArrowUp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowUp_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnBackup_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Backup created of the rdSettings.ini on user request
'---------------------------------------------------------------------------------------
'
Private Sub btnBackup_Click()
   On Error GoTo btnBackup_Click_Error
    Dim ans As VbMsgBoxResult
    
    Call backupSettings
    ans = MsgBox("Created an incremental version of bksettings.ini.* " & vbCr & "in the rocketdock folder." & vbCr & "Would you like to view the backup files? ", vbYesNo)
    If ans = 6 Then
            ShellExecute 0, vbNullString, rdAppPath, vbNullString, vbNullString, 1
    End If
   On Error GoTo 0
   Exit Sub

btnBackup_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnBackup_Click of Form rDIconConfigForm"
End Sub

Private Sub btnGetMore_Click()
    ' TODO - move the link below to a right click menu as well
    Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981272/orbs-and-icons", vbNullString, App.path, 1)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnKillIcon_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Allows deletion of the selected icon in the file list
'---------------------------------------------------------------------------------------
'
Private Sub btnKillIcon_Click()
    Dim answer As VbMsgBoxResult
    On Error GoTo btnKillIcon_Click_Error

        If textCurrIconPath.Text = "" Then
            MsgBox ("Cannot perform a deletion as no icon has been selected. ")
            Exit Sub
        End If

        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox("This will delete the currently selected icon, " & vbCr & textCurrentFolder.Text & "\" & vbCr & textCurrIconPath.Text & "   -  are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
                
        'delete the icon with no confirmation
        Kill textCurrentFolder & "\" & textCurrIconPath.Text
        
        'set the selection to the current list position
        filesIconList.ListIndex = fileIconListPosition

        'refresh the file display
        filesIconList.Refresh
        populateThumbnails (thumbImageSize)

        If filesIconList.Visible = True Then
            filesIconList.SetFocus         ' return focus to the form
        Else
            picFrameThumbs.SetFocus        ' return focus to the form
        End If
        
        ' now display the current icon, the previous icon displayed now deleted
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)

   On Error GoTo 0
   Exit Sub

btnKillIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnKillIcon_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : This just saves the current icon changes
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()
    
   On Error GoTo btnSave_Click_Error

    sFilename = txtCurrentIcon.Text
    
    sTitle = lblName.Text
    sCommand = lblTarget.Text
    sArguments = lblArguments.Text
    sWorkingDirectory = lblStartIn.Text
    sShowCmd = comboRun.ListIndex
    sOpenRunning = comboOpenRunning.ListIndex
    sUseContext = checkPopupMenu.Value
        
    ' save the current fields to the settings file or registry
    If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
        ' write the rocketdock settings.ini
        writeSettingsIni (rdIconNumber) ' the settings.ini only exists when RD is set to use it
    End If
    
    ' tell the user that all has been saved
    If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
        If chkConfirmSaves.Value = 1 Then
            MsgBox "This icon change has been stored," & vbCr & "You will need to press the ""save & restart"" button " & vbCr & "to make the changes 'stick' within Rocketdock"
        End If
    End If
    
    'if the current icon has changed refresh that part of the rdMap
    If iconChanged = True Then
        'only if the rdMAp has already been displayed already do we carry out the image refresh
        If Not picRdMap(0).ToolTipText = "" Then ' check that the array has been populated already
            ' we just reload the sole picbox that has changed
            Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32)
        End If
        iconChanged = False
    End If
    
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRefresh_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : refresh the file or thumbnail display
'---------------------------------------------------------------------------------------
'
Private Sub btnRefresh_Click()
   On Error GoTo btnRefresh_Click_Error
        
        Call busyStart
        filesIconList.Refresh
        
        If picFrameThumbs.Visible = True Then
            vScrollThumbs.Value = 0
            populateThumbnails (thumbImageSize)
            picFrameThumbs.SetFocus
        Else
            filesIconList.SetFocus
        End If
        
        removeThumbHighlighting

        Call busyStop

   On Error GoTo 0
   Exit Sub

btnRefresh_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRefresh_Click of Form rDIconConfigForm"
End Sub

Function GetDirectory(path)
   GetDirectory = Left(path, InStrRev(path, "\"))
End Function
'---------------------------------------------------------------------------------------
' Procedure : btnTarget_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Private Sub btnTarget_Click()
   Dim savLblTarget As String
    Dim iconPath As String
    
   'On Error GoTo btnTarget_Click_Error
    
    'On Error GoTo l_err1
    savLblTarget = lblTarget.Text
    
    rdDialogForm.CommonDialog.DialogTitle = "Select a File" 'titlebar
    If Not lblTarget.Text = "" Then
        If FExists(lblTarget.Text) Then
            ' extract the folder name from the string
            iconPath = GetDirectory(lblTarget.Text)
            ' set the default folder to the existing reference
            rdDialogForm.CommonDialog.InitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(lblTarget.Text) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            rdDialogForm.CommonDialog.InitDir = lblTarget.Text 'start dir, might be "C:\" or so also
        Else
            rdDialogForm.CommonDialog.InitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If
    rdDialogForm.CommonDialog.FileName = "*.*"  'Something in filenamebox
    rdDialogForm.CommonDialog.CancelError = False 'allow escape key/cancel
    rdDialogForm.CommonDialog.flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    rdDialogForm.CommonDialog.ShowOpen
        
l_err1:
    If rdDialogForm.CommonDialog.FileName = "" Then
        lblTarget.Text = savLblTarget
        Exit Sub
    End If
        
    If Err <> 32755 Then    ' User didn't chose Cancel.
        If rdDialogForm.CommonDialog.FileName = "*.*" Then
            lblTarget.Text = savLblTarget
        Else
            If lblName.Text = "" Then
                lblName.Text = rdDialogForm.CommonDialog.FileTitle
            End If
            lblTarget.Text = rdDialogForm.CommonDialog.FileName
        End If
    End If

   On Error GoTo 0
   Exit Sub

btnTarget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTarget_Click of Form rDIconConfigForm"
 
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTreeView_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : switch to tree view
'---------------------------------------------------------------------------------------
'
Private Sub btnTreeView_Click()
   On Error GoTo btnTreeView_Click_Error
    Call busyStart
    If filesIconList.Visible = True Then
        picFrameThumbs.Visible = True
        filesIconList.Visible = False
        btnThumbnailView.Visible = False
        btnTreeView.Visible = True
    Else
        picFrameThumbs.Visible = False
        filesIconList.Visible = True
        btnThumbnailView.Visible = True
        btnTreeView.Visible = False
    End If
    Call busyStop
   On Error GoTo 0
   Exit Sub

btnTreeView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTreeView_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkPopupMenu_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : If the checkbox for extended rocketdock menu options is selected
'---------------------------------------------------------------------------------------
'
Private Sub checkPopupMenu_Click()
   On Error GoTo checkPopupMenu_Click_Error

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

checkPopupMenu_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkPopupMenu_Click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : chkRegistry_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Checkbox that tells the user whether the registry is being utilised or not
'---------------------------------------------------------------------------------------
'
Private Sub chkRegistry_Click()
   On Error GoTo chkRegistry_Click_Error

    chkTheRegistry
    picThingCover.Visible = True

   On Error GoTo 0
   Exit Sub

chkRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRegistry_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkSettings_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Checkbox that tells the user whether the settings.ini file is being utilised or not
'---------------------------------------------------------------------------------------
'
Private Sub chkSettings_Click()
   On Error GoTo chkSettings_Click_Error

    chkTheRegistry
    picThingCover.Visible = True

   On Error GoTo 0
   Exit Sub

chkSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkSettings_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : comboIconTypesFilter_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Selecting the type of icon that can be displayed from a dropbox
'---------------------------------------------------------------------------------------
'
Private Sub comboIconTypesFilter_Click()
    Dim Z As Integer
    Dim D As String
    Dim filterType As Integer
    Dim itemNo As Integer
    
   On Error GoTo comboIconTypesFilter_Click_Error

    Z = comboIconTypesFilter.ListIndex
    
    ' read the current filter type and display the chosen images
    D = comboIconTypesFilter.List(comboIconTypesFilter.ListIndex)        ' same result as comboIconTypesFilter.Text
    filterType = comboIconTypesFilter.ItemData(Z)

    ' set the file type filters
    
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
    
    btnRefresh_Click
        
'    filesIconList.Refresh
'    If picFrameThumbs.Visible Then
'        populateThumbnails (thumbImageSize)
'        picFrameThumbs.SetFocus
'    Else
'        filesIconList.SetFocus
'    End If
    
    If filesIconList.ListIndex <> -1 Then ' when files found of this type
        filesIconList.ListIndex = (0)
    End If

   On Error GoTo 0
   Exit Sub

comboIconTypesFilter_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comboIconTypesFilter_Click of Form rDIconConfigForm"
End Sub

Private Sub comboOpenRunning_Click()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

Private Sub comboRun_Click()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnCloseCancel_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCloseCancel_Click()
    Dim FileName As String
   On Error GoTo btnCloseCancel_Click_Error

    If btnCloseCancel.Caption = "Cancel" Then
        'if it is a good icon number then read the data
        If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
            'get the rocketdock settings.ini for this icon alone
            readSettingsIni (rdIconNumber)
        Else
            readRegistryOnce (rdIconNumber)
        End If
        
        ' if the incoming text has <quote> then replace those with a "  TODO
        txtCurrentIcon.Text = sFilename ' build the full path TODO
        
        lblName.Text = sTitle
        lblTarget.Text = sCommand
        lblArguments.Text = sArguments
        lblStartIn.Text = sWorkingDirectory
        comboRun.ListIndex = sShowCmd
        comboOpenRunning.ListIndex = sOpenRunning
        checkPopupMenu.Value = sUseContext
        
        
        ' display the icon from the alternative settings.ini config.
        FileName = rdAppPath & "\" & txtCurrentIcon.Text
        
        Call displayPreviewImage(FileName, picPreview, icoSizePreset)
        
        ' we signify that all changes have been lost
        iconChanged = False
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
    
    Else
         Dim ofrm As Form
        
         For Each ofrm In Forms
             Unload ofrm
         Next
         
         End
    End If

   On Error GoTo 0
   Exit Sub

btnCloseCancel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnCloseCancel_Click of Form rDIconConfigForm"
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
    Dim getFolder As String
    ' add the custom folder to the treeview
    Dim CustomIconFolder As String

    ' check to see if the customfolder has been previously assigned a place in the treeview
    ' read the toolSettings.ini first
        
    ' read the settings ini file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
   On Error GoTo btnAddFolder_Click_Error

    If FExists(rdSettingsFile) Then
        CustomIconFolder = GetINISetting("Software\RocketDock", "CustomIconFolder", rdSettingsFile)
    End If

    If Not CustomIconFolder = "?" Then
    ' if the customfolder has been set then remove it first from the .ini
    ' and remove it from the tree
        treeView.selectedItem.Key = CustomIconFolder
        Call btnRemoveFolder_Click
    End If
    
    
    
    savTextCurrentFolder = textCurrentFolder.Text 'save the current default folder
    
    getFolder = ChooseDir_Click ' show the dialog box to select a folder
    If getFolder = "" Then
        'textCurrentFolder.Text = savTextCurrentFolder
        Exit Sub
    End If
    If getFolder <> "" Then
        textCurrentFolder.Text = getFolder
    End If
    

    Call busyStart
    
    ' add the chosen folder to the treeview
    treeView.Nodes.Add , , textCurrentFolder.Text, textCurrentFolder.Text
    Call addtotree(textCurrentFolder.Text, treeView)
    ' "E:\dean\steampunk theme\icons\icons MKI"
    
    'write the folder to the rocketdock settings file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\RocketDock", "CustomIconFolder", "?" & textCurrentFolder.Text, rdSettingsFile

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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnSaveRestart_Click()
    Dim NameProcess As String
    Dim useloop As Integer
    Dim ans As Integer
    
    ' save the current fields to the settings file or registry
   On Error GoTo btnSaveRestart_Click_Error

    origSettingsFile = rdAppPath & "\settings.ini"
    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
                
        ' write the rocketdock settings.ini
        writeSettingsIni (rdIconNumber) ' the settings.ini only exists when RD is set to use it
        
        ' kill the rocketdock process
        NameProcess = "rocketdock.exe"
        checkAndKill NameProcess, ans
        
        ' if the rocketdock process has died then
        If ans <> 2 Then
            ' copy the duplicate settings file to the original
            FileCopy rdSettingsFile, origSettingsFile
            
            
            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, "", App.path, 1)
        End If
    Else
         ' kill the rocketdock process
        NameProcess = "rocketdock.exe"
        checkAndKill NameProcess, ans
                   
        chkRegistry.Value = 1
        chkSettings.Value = 0
        
        ' if the rocketdock process has died then
        If ans <> 2 Then
           For useloop = 0 To rdIconMax
                ' read the rocketdock alternative settings.ini
                readSettingsIni (useloop) ' the alternative settings.ini exists when RD is set to use it
            
                ' write the rocketdock registry
                writeRegistryOnce (useloop)
            Next useloop
            
            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, "", App.path, 1)
        End If
        
    End If

   On Error GoTo 0
   Exit Sub

btnSaveRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSaveRestart_Click of Form rDIconConfigForm"
            
End Sub

Private Sub Command3_Click()
    Dim getFolder As String
        
    getFolder = ChooseDir_Click ' show the dialog box to select a folder
    If getFolder <> "" Then lblStartIn.Text = getFolder
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ChooseDir_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ChooseDir_Click()
    Dim sTempDir As String
   'On Error GoTo ChooseDir_Click_Error

    On Error Resume Next
    
    sTempDir = CurDir    'Remember the current active directory
    rdDialogForm.CommonDialog.DialogTitle = "Select a directory" 'titlebar
    If Not lblStartIn.Text = "" Then
        If DirExists(lblStartIn.Text) Then
            rdDialogForm.CommonDialog.InitDir = lblStartIn.Text 'start dir, might be "C:\" or so also
        Else
            rdDialogForm.CommonDialog.InitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If
    rdDialogForm.CommonDialog.FileName = "Select a Directory"  'Something in filenamebox
    rdDialogForm.CommonDialog.flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    rdDialogForm.CommonDialog.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    rdDialogForm.CommonDialog.CancelError = True 'allow escape key/cancel ' do NOT change, causes a hang
    rdDialogForm.CommonDialog.ShowSave   'show the dialog screen

    If Err <> 32755 Then
        ChooseDir_Click = CurDir ' User didn't chose Cancel.
    Else
        ChooseDir_Click = "" ' User chose Cancel.
    End If

    ChDir sTempDir  'restore path to what it was at entering

   On Error GoTo 0
   Exit Function

ChooseDir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChooseDir_Click of Form rDIconConfigForm"
    
End Function
Private Sub btnHelp_Click()
    ' show a single help PNG with pointers as to what does what
    rdHelpForm.Show
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnNext_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNext_Click()
    'if the modification flag is set then ask before moving to the next icon
    Dim answer As VbMsgBoxResult
    Dim FileName As String
    Dim useloop As Integer
        
   On Error GoTo btnNext_Click_Error

    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'display the new icon data
    
    'increment the icon number
    rdIconNumber = rdIconNumber + 1
    'check we haven't gone too far
    If rdIconNumber > rdIconMax Then rdIconNumber = rdIconMax
    
    ' only move the map if the array has been populated,
    'Dim a As Boolean
    'a = picRdMap(0).Picture ' always empty
    If Not picRdMap(0).ToolTipText = "" Then
        ' I want to test to see if the picture property is populated but
        ' as the picture property is not being set by Lavolpe's method then we can't test for it
        ' testing the tooltip is one method of seeing if the map has been created
        ' as the program sets the tooltip just when the transparent image is set
        
        ' moves the RdMap on one position (one click) if it is already set at the rightmost screen position
        If rdIconNumber > rdMapHScroll.Value + 15 Then
            btnMapPrev_Click
        End If
    End If
    
    lblRdIconNumber.Caption = Str(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
    
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)
    
    'remove and reset the highlighting on the Rocket dock map
    For useloop = 0 To rdIconMax
        picRdMap(useloop).BorderStyle = 0
    Next useloop
    picRdMap(rdIconNumber).BorderStyle = 1

    previewFrameGotFocus = True

    ' we signify that all changes have been lost
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

btnNext_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnNext_Click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : readRegistryOnce
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryOnce(iconNumberToRead)
    ' read the settings from the registry
   On Error GoTo readRegistryOnce_Error

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

readRegistryOnce_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryOnce of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : writeRegistryOnce
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub writeRegistryOnce(iconNumberToWrite)
        
   On Error GoTo writeRegistryOnce_Error

    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", sFilename)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", sFileName2)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Title", sTitle)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Command", sCommand)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", sArguments)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", sShowCmd)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", sOpenRunning)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", sIsSeparator)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", sUseContext)
    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", sDockletFile)

   On Error GoTo 0
   Exit Sub

writeRegistryOnce_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistryOnce of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ExtractSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ExtractSuffix(ByVal strPath As String) As String
    Dim AY() As String ' string array
    Dim max As Integer
    
   On Error GoTo ExtractSuffix_Error

    AY = Split(strPath, ".")
    max = UBound(AY)
    ExtractSuffix = AY(max)

   On Error GoTo 0
   Exit Function

ExtractSuffix_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffix of Form rDIconConfigForm"
End Function
'---------------------------------------------------------------------------------------
' Procedure : displayPreviewImage
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayPreviewImage(FileName As String, targetPicBox As PictureBox, IconSize As Integer)
    Dim suffix As String
    
   On Error GoTo displayPreviewImage_Error

    If FExists(FileName) Then     ' just a final check that the chosen image file actually exists
        textCurrIconPath.Text = filesIconList.FileName
        ' find and store the indexed position of the chosen file into the global variable
        fileIconListPosition = filesIconList.ListIndex

        Set targetPicBox.Picture = Nothing ' added because the two methods of drawing an image conflict leaving an image behind

        suffix = ExtractSuffix(FileName)
        
        'suffix = Right(FileName, Len(FileName) - InStr(1, FileName, "."))
        
        ' using Lavolpe's later method as it allows for resizing of PNGs and other types
        If InStr("png,jpg,bmp,jpeg,tif", LCase(suffix)) <> 0 Then
            If targetPicBox.Name = "picPreview" Then
                targetPicBox.Left = 345
                targetPicBox.Top = 210
                targetPicBox.Width = 3450
                targetPicBox.Height = 3450
            End If

            
            Set cImage = New c32bppDIB
            cImage.LoadPicture_File FileName, IconSize, IconSize, False, 32
            Call refreshPicBox(targetPicBox, IconSize)
        Else
            ' using Lavolpe's earlier StdPictureEx method as it allows for correct display of ICOs
            ' the later method has a bug with some ICOs
            
            'because the earlier method draws the ico images from the top left of the
            'pictureBox we have to manually set the picbox to size and position for each icon size
            If IconSize = 16 Then
                If targetPicBox.Name = "picPreview" Then
                    targetPicBox.Left = 1900
                    targetPicBox.Top = 1900
                    targetPicBox.Width = 200
                    targetPicBox.Height = 200
                End If
            ElseIf IconSize = 32 Then
                If targetPicBox.Name = "picPreview" Then
                    targetPicBox.Left = 1800
                    targetPicBox.Top = 1800
                    targetPicBox.Width = 2000
                    targetPicBox.Height = 2000
                End If
            ElseIf IconSize = 64 Then
                If targetPicBox.Name = "picPreview" Then
                    targetPicBox.Left = 1450
                    targetPicBox.Top = 1450
                    targetPicBox.Width = 2000
                    targetPicBox.Height = 2000
                End If
            ElseIf IconSize = 128 Then
                If targetPicBox.Name = "picPreview" Then
                    targetPicBox.Left = 1000
                    targetPicBox.Top = 1000
                    targetPicBox.Width = 2000
                    targetPicBox.Height = 2000
                End If
            ElseIf IconSize = 256 Then
                If targetPicBox.Name = "picPreview" Then
                    targetPicBox.Left = 100 '330
                    targetPicBox.Top = 100  '270
                    targetPicBox.Width = 4000
                    targetPicBox.Height = 4000
                End If
            End If
            
            Set targetPicBox.Picture = StdPictureEx.LoadPicture(FileName, lpsCustom, , IconSize, IconSize)
        End If
    End If

   On Error GoTo 0
   Exit Sub

displayPreviewImage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayPreviewImage of Form rDIconConfigForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnHome
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHome()

    Dim answer As VbMsgBoxResult
    Dim FileName As String
    Dim useloop As Integer
    Dim ff As Long
    'if the modification flag is set then ask before moving to the next icon
   On Error GoTo btnHome_Error

    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    rdMapHScroll.Value = rdMapHScroll.Min
    
    ' we signify that all changes have been lost
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

btnHome_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHome of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnEnd
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnEnd()
    Dim answer As VbMsgBoxResult
    Dim FileName As String
    Dim useloop As Integer
    Dim ff As Long
    
    'if the modification flag is set then ask before moving to the next icon
   On Error GoTo btnEnd_Error

    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    rdMapHScroll.Value = rdMapHScroll.max

    ' we signify that all changes have been lost
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

btnEnd_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnEnd of Form rDIconConfigForm"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnPrev_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrev_Click()
    'if the modification flag is set then ask before moving to the next icon
    Dim answer As VbMsgBoxResult
    Dim FileName As String
    Dim useloop As Integer
    
   On Error GoTo btnPrev_Click_Error

    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'display the new icon data
    
    'decrement the icon number
    rdIconNumber = rdIconNumber - 1
    'check we haven't gone too far
    If rdIconNumber < 0 Then rdIconNumber = 0
    
    ' only move the map if the array has been populated,
    'Dim a As Boolean
    'a = picRdMap(0).Picture ' always empty
    If Not picRdMap(0).ToolTipText = "" Then
        ' I want to test to see if the picture property is populated but
        ' as the picture property is not being set by Lavolpe's method then we can't test for it
        ' testing the tooltip is one method of seeing if the map has been created
        ' as the program sets the tooltip just when the transparent image is set
    
        ' moves the RdMap on one position (one click) if it is already set at the rightmost screen position
        If rdIconNumber < rdMapHScroll.Value Then
            btnMapNext_Click
        End If
    End If

    lblRdIconNumber.Caption = Str(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)
        
    'remove and reset the highlighting on the Rocket dock map
    For useloop = 0 To rdIconMax
        picRdMap(useloop).BorderStyle = 0
    Next useloop
    picRdMap(rdIconNumber).BorderStyle = 1
    
    previewFrameGotFocus = True

    ' we signify that all changes have been lost
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

btnPrev_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrev_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayIconElement
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub displayIconElement(iconCount As Integer, picBox As PictureBox, icoPreset As Integer)
    Dim FileName As String
        Dim qPos As Long
        Dim filestring As String
        Dim suffix As String
    
    'if it is a good icon then read the data
   On Error GoTo displayIconElement_Error

    If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
        'get the rocketdock alternative settings.ini for this icon alone
        readSettingsIni (iconCount)
    End If

    ' if the incoming text has <quote> then replace those with a " TODO ?
    txtCurrentIcon.Text = sFilename ' build the full path
    
    lblName.Text = sTitle
    lblTarget.Text = sCommand
    lblArguments.Text = sArguments
    lblStartIn.Text = sWorkingDirectory
    
    comboRun.ListIndex = Val(sShowCmd)
    comboOpenRunning.ListIndex = Val(sOpenRunning)
    checkPopupMenu.Value = Val(sUseContext)

    ' test whether it is a valid file with a path or just a relative path
    If FExists(sFilename) Or InStr(sFilename, "?") Then
        FileName = sFilename  ' a full valid path so leave it alone
    Else
        FileName = rdAppPath & "\" & txtCurrentIcon.Text ' a relative path found as per Rocketdock
    End If

    picBox.ToolTipText = "Icon number " & iconCount + 1 & " = " & sFilename

        
    ' if the user drags an icon to the dock then RD takes a icon link of the following form:
    'FileName = "C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\vbexpress.exe?62453184"
    If Not FExists(FileName) Then  ' this should fail with the above ?62453184 style suffix
        ' does the string contain a ? if so it probably has an embedded .ICO
        qPos = InStr(1, FileName, "?")
        If qPos <> 0 Then
            ' extract the string before the ? (qPos)
            filestring = Mid$(FileName, 1, qPos - 1)
        End If
        
        ' test the resulting filestring exists
        If FExists(filestring) Then
            ' extract the suffix
            suffix = ExtractSuffix(filestring)

            'suffix = Right(filestring, Len(filestring) - InStr(1, filestring, "."))
            ' test as to whether it is an .EXE or a .DLL
            If InStr("exe,dll", LCase(suffix)) <> 0 Then
                'FileName = txtCurrentIcon.Text ' revert to the relative path which is what is expected
                Call displayEmbeddedIcons(filestring, picBox, icoPreset)
            Else
                ' the file may have a ? in the string but does not match otherwise in any useful way
                FileName = rdAppPath & "\icons\" & "help.png"
            End If
            
        Else ' the file doesn't exist in any form with ? or otherwise as a valid path
            FileName = rdAppPath & "\icons\" & "help.png"
            Call displayPreviewImage(FileName, picBox, icoPreset)
            dllFrame.Visible = False
        End If
    Else
        Call displayPreviewImage(FileName, picBox, icoPreset)
        dllFrame.Visible = False
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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnThumbnailView_Click()
   On Error GoTo btnThumbnailView_Click_Error
    Call busyStart
    populateThumbnails (thumbImageSize)
        
    picFrameThumbs.Visible = True
    filesIconList.Visible = False
    btnThumbnailView.Visible = False
    btnTreeView.Visible = True
    
    removeThumbHighlighting
    Call busyStop

   On Error GoTo 0
   Exit Sub

btnThumbnailView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnThumbnailView_Click of Form rDIconConfigForm"
    
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
'---------------------------------------------------------------------------------------
'
Private Sub populateThumbnails(imageSize As Integer)

    Dim useloop As Integer
    Dim selectedItem As Integer
    Dim FileName As String
    Dim FileCount As Integer
    Dim tooltip As String
    Dim suffix As String
    
   On Error GoTo populateThumbnails_Error

    For useloop = 0 To 11
        picIconList(useloop).Visible = False
    Next useloop

    placeThumbnailPicboxes ' in the right place
   
    ' this is a toggle
    ' change the image to a tree view icon
    ' create a matrix of 12 x image objects from file


    ' starting with the SelectedItem from filesIconList
    ' we extract the number
        
    vScrollThumbs.max = filesIconList.ListCount
        
    selectedItem = filesIconList.ListIndex
    'when there are less than a screenful of items the.ListIndex returns -1
    ' it should be the list
    If selectedItem = -1 Then
        selectedItem = 0
    End If
    
    ' populate each with an image
    For useloop = 0 To 11
        ' using the deviation from the extracted start
        ' visit the filelist at that point and extract the filename
        '  and extract the file path
'        If fileCount < 12 Then
'            FileName = filesIconList.Path & "\" & filesIconList.List(useloop)
'        Else
'            FileName = filesIconList.Path & "\" & filesIconList.List(useloop + selectedItem)
'        End If

        FileName = filesIconList.path & "\" & filesIconList.List(useloop + selectedItem)
        
        If filesIconList.List(useloop + selectedItem) <> "" Then
            picIconList(useloop).ToolTipText = filesIconList.List(useloop + selectedItem)
            If Len(picIconList(useloop).ToolTipText) > 14 Then
                Dim newString As String
                Dim leftBit As String
                Dim rightBit As String
                leftBit = Left$(picIconList(useloop).ToolTipText, 14) ' left of string
                rightBit = Mid$(picIconList(useloop).ToolTipText, 15) ' right of string
                newString = leftBit & vbCr & rightBit     ' insert vbCr
                lblThumbName(useloop).Caption = newString
            Else
                lblThumbName(useloop).Caption = picIconList(useloop).ToolTipText
            End If
            If Len(picIconList(useloop).ToolTipText) > 26 Then
                lblThumbName(useloop).Caption = Left$(picIconList(useloop).ToolTipText, 26) & "..."
                lblThumbName(useloop).Alignment = 0
            Else
                'lblThumbName(useloop).Caption = Left$(picIconList(useloop).ToolTipText, 26)
                lblThumbName(useloop).Alignment = 2
            End If
            
            ' display the image within the specified picturebox
            Call displayPreviewImage(FileName, picIconList(useloop), imageSize)
        Else
            'Exit Sub
            picIconList(useloop).ToolTipText = ""
            lblThumbName(useloop).Caption = picIconList(useloop).ToolTipText
            picIconList(useloop).Picture = LoadPicture(App.path & "\" & "blank.jpg")

            'picIconList(useloop)
        End If
        thumbArray(useloop) = useloop + selectedItem
        lblThumbName(useloop).ZOrder
        
    Next useloop
    
    For useloop = 0 To 11
        'if the thumbnail is an ico
        tooltip = picIconList(useloop).ToolTipText
        'suffix = Right(tooltip, Len(tooltip) - InStr(1, tooltip, "."))
        If tooltip <> "" Then
            suffix = ExtractSuffix(tooltip)
    
            If thumbImageSize = 32 Then
               If suffix = "ico" Then
                    frmThumbLabel(useloop).Visible = True
                    picIconList(useloop).Left = picIconList(useloop).Left + 300
                    picIconList(useloop).Top = picIconList(useloop).Top + 300
                End If
            Else
    '            If suffix = "ico" Then
    '                frmThumbLabel(useloop).Visible = False
    '                picIconList(useloop).Left = picIconList(useloop).Left - 300
    '                picIconList(useloop).Top = picIconList(useloop).Top - 300
    '            End If
            End If
            picIconList(useloop).Visible = True
        End If
    Next useloop

    For useloop = 0 To 11
        If thumbImageSize = 32 Then
                frmThumbLabel(useloop).Visible = True
                frmThumbLabel(useloop).ZOrder
        Else
                frmThumbLabel(useloop).Visible = False
        End If
    Next useloop

   On Error GoTo 0
   Exit Sub

populateThumbnails_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateThumbnails of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : populateRdMap
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : redraws the whole rdMap
'---------------------------------------------------------------------------------------
'
Private Sub populateRdMap(xDeviation As Integer)

    Dim useloop As Integer
    Dim selectedItem As Integer
    Dim FileName As String
    Dim FileCount As Integer
    
    Dim busyFilename As String
    
    Dim dotString As String
   
    On Error GoTo populateRdMap_Error

    dotString = ""
    dotCount = 0

    Refresh ' display the results prior to the for loop
    ' the above command allows the working button to show as it should
                
    ' populate each with an image
    For useloop = 0 To rdIconMax
        picRdMap(useloop).BorderStyle = 1 ' put a border around the picboxes to show an update
        
        ' using the deviation from the extracted start
        ' visit the filelist at that point and extract the filename
        '  and extract the file path
        
        ' the target picture control and the icon size
        Call displayIconElement(useloop + xDeviation, picRdMap(useloop), 32)
        picRdMap(useloop).BorderStyle = 0
        
        'do the 'working...' text on the button
        dotCount = dotCount + 1
        If dotCount = 5 Then
            'Refresh
            dotCount = 0
            dotString = dotString & "."
            btnWorking.Caption = "Working " & dotString
            If dotString = "..." Then dotString = ""
        End If
    
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        picBusy.Visible = True
        busyCounter = busyCounter + 1
        If busyCounter >= 7 Then busyCounter = 1
        busyFilename = App.path & "\busy-F" & busyCounter & "-32x32x24.jpg"
        picBusy.Picture = LoadPicture(busyFilename)
        picBusy.Visible = False

    Next useloop

   On Error GoTo 0
   Exit Sub

populateRdMap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateRdMap of Form rDIconConfigForm"

End Sub
Private Sub btnDefaultIcon_Click()
    MsgBox "Not implemented yet"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnMapNext_Click
' Author    : beededea
' Date      : 02/06/2019
' Purpose   : Scroll the RD map to the left
'---------------------------------------------------------------------------------------
'
Private Sub btnMapNext_Click()
   Dim useloop As Integer
    
    picRdMapGotFocus = True
    picFrameThumbsGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    
    If rdMapHScroll.Value >= 1 Then
        rdMapHScroll.Value = rdMapHScroll.Value - 1
    End If
     
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnMapPrev_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnMapPrev_Click()
    Dim useloop As Integer
   On Error GoTo btnMapPrev_Click_Error

    picRdMapGotFocus = True
    picFrameThumbsGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False

    If rdMapHScroll.Value < rdMapHScroll.max Then
        rdMapHScroll.Value = rdMapHScroll.Value + 1
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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_Click()
    Dim FileName As String
    Dim aa As Integer
    
    
   On Error GoTo filesIconList_Click_Error

    picPreview.AutoRedraw = True
    picPreview.AutoSize = False
    
    FileName = filesIconList.path
    If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
    FileName = FileName & filesIconList.FileName
    
    If picFrameThumbsGotFocus = True Then
        Call displayPreviewImage(FileName, picPreview, icoSizePreset)
    End If
    
    filesIconListGotFocus = True
    
    ' we signify that no changes have been made
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"


   On Error GoTo 0
   Exit Sub

filesIconList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_DblClick
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_DblClick()
    ' relative path
    ' rdAppPAth
    '\Icons
    ' takes the result from the treeview
   On Error GoTo filesIconList_DblClick_Error

    txtCurrentIcon.Text = relativePath & "\" & filesIconList.FileName
        ' we signify that no changes have been made
    btnSave.Enabled = True ' this has to be done at the end
    btnCloseCancel.Caption = "Cancel"
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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_GotFocus()
   On Error GoTo filesIconList_GotFocus_Error

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False

   On Error GoTo 0
   Exit Sub

filesIconList_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : filesIconList_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub filesIconList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo filesIconList_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If


   On Error GoTo 0
   Exit Sub

filesIconList_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_MouseDown of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : readCustomLocation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readCustomLocation()

    ' add the custom folder to the treeview
    Dim CustomIconFolder As String
    
    ' read the settings ini file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
   On Error GoTo readCustomLocation_Error

    If FExists(rdSettingsFile) Then
        CustomIconFolder = GetINISetting("Software\RocketDock", "CustomIconFolder", rdSettingsFile)
    End If
    
    If Not CustomIconFolder = "" Then
        CustomIconFolder = Mid(CustomIconFolder, 2) ' remove the question mark
        If DirExists(CustomIconFolder) Then
            ' add the chosen folder to the treeview
            treeView.Nodes.Add , , CustomIconFolder, CustomIconFolder
            Call addtotree(CustomIconFolder, treeView)
            treeView.Nodes(CustomIconFolder).Text = "custom folder"
        End If
    End If

   On Error GoTo 0
   Exit Sub

readCustomLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readCustomLocation of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : readDefaultFolder
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readDefaultFolder()
    ' extract the default folder in the treeview, this is an enhancement over the original tool
    ' it stores the last used default folder as shown in the tree view top left
    
    Dim iX As Integer
    Dim iFound As Boolean
    Dim defaultFolderNodeKey As String
        
    ' read the tool settings file
    'eg. defaultFolderNodeKey=?E:\dean\steampunk theme\icons\
   On Error GoTo readDefaultFolder_Error

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        defaultFolderNodeKey = GetINISetting("Software\RocketDockSettings", "defaultFolderNodeKey", toolSettingsFile)
    End If

    treeView.HideSelection = False ' Ensures found item highlighted

    If defaultFolderNodeKey <> "" Then
        For iX = 1 To treeView.Nodes.count
            If Trim(treeView.Nodes(iX).Key) = Trim(defaultFolderNodeKey) Then
                iFound = True
                Exit For
            End If
        Next
        If iFound Then
            ' highlight the treeview item
            treeView.Nodes(iX).EnsureVisible
            treeView.selectedItem = treeView.Nodes(iX)
            treeView.Nodes(iX).Selected = True
            treeView_Click ' click on the selected item
        Else
            'MsgBox ("String not found")
        End If
    End If

   On Error GoTo 0
   Exit Sub

readDefaultFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDefaultFolder of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : readRegistryWriteSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryWriteSettings()
    Dim useloop As Integer
    
   On Error GoTo readRegistryWriteSettings_Error

    For useloop = 0 To rdIconMax
         ' get the relevant entries from the registry
         readRegistryOnce (useloop)
         ' read the rocketdock alternative settings.ini
         writeSettingsIni (useloop) ' the alternative settings.ini exists when RD is set to use it
     Next useloop

   On Error GoTo 0
   Exit Sub

readRegistryWriteSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryWriteSettings of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : driveCheck
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function driveCheck(folder As String)
   Dim sAllDrives As String
   Dim sDrv As String
   Dim sDrives() As String
   Dim cnt As Long
   Dim folderString As String
   
  'get the list of all drives
   On Error GoTo driveCheck_Error

   sAllDrives = GetDriveString()
   
  'Change nulls to spaces, then trim.
  'This is required as using Split()
  'with Chr$(0) alone adds two additional
  'entries to the array drives at the end
  'representing the terminating characters.
   sAllDrives = Replace$(sAllDrives, Chr$(0), Chr$(32))
   sDrives() = Split(Trim$(sAllDrives), Chr$(32))
    
    For cnt = LBound(sDrives) To UBound(sDrives)
        sDrv = sDrives(cnt)
        ' on 32bit windows the folder is "Program Files\Rocketdock"
        folderString = sDrv & folder
        If DirExists(folderString) = True Then
           'test for the yahoo widgets binary
            rdAppPath = folderString
            If FExists(rdAppPath & "\rocketdock.exe") Then
                'MsgBox "YWE folder exists"
                driveCheck = True
                Exit Function
            End If
        End If
    Next

   On Error GoTo 0
   Exit Function

driveCheck_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure driveCheck of Form rDIconConfigForm"
   
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDriveString
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetDriveString() As String

  'Used by both demos
  
'  'returns string of available
'  'drives each separated by a null
'   Dim sBuff As String
'
'  'possible 26 drives, three characters
'  'each plus a trailing null for each
'  'drive letter and a terminating null
'  'for the string
'
Dim I As Long
Dim builtString As String

    '===========================
    'pure VB approach, no controls required
    'drive letters are found in positions 1-UBound(Letters)
    '"C:\ D:\ E:\ &c"
    
   On Error GoTo GetDriveString_Error

    For I = 1 To 26
        If ValidDrive(Chr(96 + I)) = True Then
            builtString = builtString + UCase(Chr(96 + I)) & ":\    "
        End If
    Next I
    
    GetDriveString = builtString

   On Error GoTo 0
   Exit Function

GetDriveString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriveString of Form rDIconConfigForm"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidDrive
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ValidDrive(D As String) As Boolean
   On Error GoTo ValidDrive_Error

  On Error GoTo driveerror
  Dim Temp As String
  
    Temp = CurDir
    ChDrive D
    
    ChDir Temp
    ValidDrive = True

  Exit Function
driveerror:

   On Error GoTo 0
   Exit Function

ValidDrive_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidDrive of Form rDIconConfigForm"
End Function

'---------------------------------------------------------------------------------------
' Procedure : addRocketdockFolders
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub addRocketdockFolders()
    Dim pathCheck As String
    
    
   On Error GoTo addRocketdockFolders_Error

    pathCheck = rdAppPath & "\icons"
        
    If Not pathCheck = "" Then
        ' add the chosen folder to the treeview
        treeView.Nodes.Add , , pathCheck, pathCheck
        
        'treeView.Nodes(pathCheck).ToolTipText = "arse"
        Call addtotree(pathCheck, treeView)
        treeView.Nodes(pathCheck).Text = "icons"
    End If

   On Error GoTo 0
   Exit Sub

addRocketdockFolders_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addRocketdockFolders of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then
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
Private Sub frameButtons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo frameButtons_MouseDown_Error

   If Button = 2 Then
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
Private Sub FrameFolders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo FrameFolders_MouseDown_Error

    If Button = 2 Then
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
Private Sub frameIcons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo frameIcons_MouseDown_Error

    If Button = 2 Then
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
Private Sub framePreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo framePreview_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

framePreview_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure framePreview_MouseDown of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : frameProperties_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub frameProperties_Click()
    Dim FileName As String
    
    ' display the icon from the alternative settings.ini config.
   On Error GoTo frameProperties_Click_Error

    FileName = rdAppPath & "\" & txtCurrentIcon.Text
    
    Call displayPreviewImage(FileName, picPreview, icoSizePreset)


   On Error GoTo 0
   Exit Sub

frameProperties_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure frameProperties_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : frameProperties_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub frameProperties_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo frameProperties_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If


   On Error GoTo 0
   Exit Sub

frameProperties_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure frameProperties_MouseDown of Form rDIconConfigForm"

End Sub

Private Sub frmThumbLabel_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    frmThumbLabel(Index).ZOrder

End Sub

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
    Dim useloop As Long
    Dim startPos As Long
    Dim maxPos As Long
    Dim rdIconMaxLong As Long
    Dim spacing As Integer
    
   On Error GoTo rdMapHScroll_Change_Error

    spacing = 540

    rdIconMaxLong = rdIconMax
    rdMapHScroll.Min = 0
    rdMapHScroll.max = rdIconMax - 16
    
    startPos = rdMapHScroll.Value - 1
    xlabel.Caption = startPos
    nLabel.Caption = (startPos * spacing)
    
    maxPos = rdIconMaxLong * spacing
    
    For useloop = 0 To rdIconMax
            picRdMap(useloop).Move ((useloop * spacing) - (startPos * spacing)), 30, 500, 500
    Next useloop

   On Error GoTo 0
   Exit Sub

rdMapHScroll_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapHScroll_Change of Form rDIconConfigForm"
    
End Sub

Private Sub lblArguments_Change()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

Private Sub lblName_Change()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

Private Sub lblStartIn_Change()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

Private Sub lblTarget_Change()
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub getkeypress(KeyCode As Integer)

'36 home
'40 is down
'38 is up
'37 is left
'39 is right
' 33 page up
' 34 page down
' 35 end
    
   On Error GoTo getkeypress_Error

    keyPressOccurred = True
    
    If KeyCode = 116 Then
        If picFrameThumbsGotFocus = True Or filesIconListGotFocus = True Then
            ' if the thumbframe has focus then refresh on f5
            btnRefresh_Click
        End If
    
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then refresh on f5
            rdMapRefresh_Click
        End If
    End If

    
    'carry start
    If KeyCode = 36 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the first icon
            btnHome
        End If
    End If
    If KeyCode = 35 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the end of the rdMap
            btnEnd
        End If
    End If
    
    'carry end
    

    If KeyCode = 38 Then
        If vScrollThumbs.Value - 4 < vScrollThumbs.Min Then
            vScrollThumbs.Value = vScrollThumbs.Min
        Else
            vScrollThumbs.Value = vScrollThumbs.Value - 4
        End If
    End If

    If KeyCode = 40 Then
        If vScrollThumbs.Value + 4 > vScrollThumbs.max Then
            vScrollThumbs.Value = vScrollThumbs.max
        Else
            vScrollThumbs.Value = vScrollThumbs.Value + 4
        End If
    End If

    If KeyCode = 37 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnPrev_Click
        Else
            If vScrollThumbs.Value - 1 < vScrollThumbs.Min Then
                vScrollThumbs.Value = vScrollThumbs.Min
            Else
                vScrollThumbs.Value = vScrollThumbs.Value - 1
            End If
        End If
    End If

    If KeyCode = 39 Then
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnNext_Click
        Else
            If vScrollThumbs.Value + 1 > vScrollThumbs.max Then
                vScrollThumbs.Value = vScrollThumbs.max
            Else
                vScrollThumbs.Value = vScrollThumbs.Value + 1
            End If
        End If
    End If
    
    '33 is page up
    If KeyCode = 33 Then
        If vScrollThumbs.Value - 12 < vScrollThumbs.Min Then
            vScrollThumbs.Value = vScrollThumbs.Min
        Else
            vScrollThumbs.Value = vScrollThumbs.Value - 12
        End If
    End If

    '34 is page down
    If KeyCode = 34 Then
        If vScrollThumbs.Value + 12 > vScrollThumbs.max Then
            vScrollThumbs.Value = vScrollThumbs.max
        Else
            vScrollThumbs.Value = vScrollThumbs.Value + 12
        End If
    
    End If


   On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Form rDIconConfigForm"

End Sub
Private Sub lblThumbName_Click(Index As Integer)
    Call picIconList_Click(Index)
End Sub
Private Sub picFrameThumbs_GotFocus()
    picFrameThumbsGotFocus = True
End Sub

Private Sub picFrameThumbs_KeyDown(KeyCode As Integer, Shift As Integer)
    Call getkeypress(KeyCode)
End Sub

Private Sub picFrameThumbs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
        Me.PopupMenu thumbmenu, vbPopupMenuRightButton
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picIconList_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picIconList_Click(Index As Integer)
    ' a click on any icon thumbnail shows a preview
    
    ' TODO - implement keypresses on the thumbnails
 
    Dim useloop As Integer
    Dim itemNo As Integer
    Dim filePath As String
    Dim FileName As String
    
    'remove the highlighting
   On Error GoTo picIconList_Click_Error

    For useloop = 0 To 11
        picIconList(useloop).BorderStyle = 0
        lblThumbName(useloop).BackColor = &HFFFFFF
        'lblThumbName(useloop).ForeColor = &H80000012
    Next useloop
    If thumbImageSize = 64 Then 'larger
        picIconList(Index).BorderStyle = 1
    ElseIf thumbImageSize = 32 Then
        lblThumbName(Index).BackColor = &HFFC0C0
    End If

    ' extract the filename from the associated array
    If Not picIconList(Index).ToolTipText = "" Then ' we use the tooltip because the .picture property is not populated
        itemNo = thumbArray(Index)
        'this next line change is meant to trigger a re-click but it does not when the index is unchanged from previous click
        filesIconList.ListIndex = (itemNo) '
         ' this next if then checks to see if the stored click is the same , if so it triggers a click on the item in the underlying file list box
        If storedIndex = Index Then
            Call filesIconList_Click
        End If
        storedIndex = Index

    End If

   On Error GoTo 0
   Exit Sub

picIconList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picIconList_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picIconList_DblClick
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picIconList_DblClick(Index As Integer)
    Dim itemNo As Integer

'    ' extract the filename from the associated array
   On Error GoTo picIconList_DblClick_Error

    itemNo = thumbArray(Index)
    filesIconList.ListIndex = (itemNo) ' this does a click the item in the underlying file list box
    filesIconList_DblClick

   On Error GoTo 0
   Exit Sub

picIconList_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picIconList_DblClick of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picIconList_GotFocus
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picIconList_GotFocus(Index As Integer)
   On Error GoTo picIconList_GotFocus_Error

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False

   On Error GoTo 0
   Exit Sub

picIconList_GotFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picIconList_GotFocus of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picIconList_KeyDown
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picIconList_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error GoTo picIconList_KeyDown_Error

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    Call getkeypress(KeyCode)

   On Error GoTo 0
   Exit Sub

picIconList_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picIconList_KeyDown of Form rDIconConfigForm"
End Sub

Private Sub picIconList_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
        Me.PopupMenu thumbmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub picIconList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    frmThumbLabel(Index).ZOrder
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picPreview_Click
' Author    : beededea
' Date      : 19/06/2019
' Purpose   : on a single click search the file list and determine the image that is being shown by filename match
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_Click()
    Dim useloop As Integer

   On Error GoTo picPreview_Click_Error

    For useloop = 1 To filesIconList.ListCount
    ' TODO - extract just the filename from txtCurrentIcon.Text
        If filesIconList.List(useloop) = GetFileNameFromPath(txtCurrentIcon.Text) Then
            filesIconList.ListIndex = useloop
            GoTo l_found_file
        End If
    Next useloop
l_found_file:

    If picFrameThumbs.Visible = True Then
            populateThumbnails (thumbImageSize)
    End If

   On Error GoTo 0
   Exit Sub

picPreview_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFileNameFromPath
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : A function to GetFileNameFromPath
'---------------------------------------------------------------------------------------
'
Function GetFileNameFromPath(strFullPath As String) As String
   On Error GoTo GetFileNameFromPath_Error

    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

   On Error GoTo 0
   Exit Function

GetFileNameFromPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFileNameFromPath of Form rDIconConfigForm"
End Function
    '

'---------------------------------------------------------------------------------------
' Procedure : picPreview_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picPreview_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnuMainOpts, vbPopupMenuRightButton
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
Private Sub picRdMap_GotFocus(Index As Integer)
   On Error GoTo picRdMap_GotFocus_Error

    picRdMapGotFocus = True
    picFrameThumbsGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    

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
Private Sub picRdMap_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error GoTo picRdMap_KeyDown_Error

    Call getkeypress(KeyCode)

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
Private Sub picRdMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim useloop As Integer
    Dim answer As VbMsgBoxResult
    Dim picX As Integer
    
   On Error GoTo picRdMap_MouseDown_Error
   
   
   
   'this is the preliminary code to allow sliding of the icons by mouse and cursor
   
'    picRdMap(Index).ZOrder 'the drag pic
'    'so it appears over the top of other controls
'    picX = X - 500
    
    'MsgBox "X = " & picRdMap(Index).Left
        
   'this is the preliminary code to allow sliding of the icons by mouse and cursor
        


    If Button = 2 Then
        rdIconNumber = Index
        Me.PopupMenu rdMapMenu, vbPopupMenuRightButton
    Else
        If btnSave.Enabled = True Then
            If chkConfirmSaves.Value = 1 Then
                answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
                If answer = vbNo Then
                    Exit Sub
                End If
            End If
        End If
       
        rdIconNumber = Index
        
        lblRdIconNumber.Caption = Str(rdIconNumber) + 1
        lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)
    
        'remove and reset the highlighting on the Rocket dock map
        For useloop = 0 To rdIconMax
            picRdMap(useloop).BorderStyle = 0
        Next useloop
        If Index <= rdIconMax Then
            picRdMap(Index).BorderStyle = 1
        End If
        
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
    End If

   On Error GoTo 0
   Exit Sub

picRdMap_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picRdMap_MouseDown of Form rDIconConfigForm"
End Sub

Private Sub picRdMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' code retained in case I want to do a graphical drag and drop of one item in the map to another

' Dim picX As Integer
'
' With picRdMap(Index)
' If Button Then
'  .Move .Left + (X) - picX
' End If
' End With

End Sub

Private Sub picRdMap_OLEDragDrop(Index As Integer, data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim a As String
  a = rdIconNumber
End Sub

Private Sub picRdMap_OLEStartDrag(Index As Integer, data As DataObject, AllowedEffects As Long)
  Dim a As String
  a = rdIconNumber
End Sub

Private Sub picThingCover_Click()
    picThingCover.Visible = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : picPreview_OLEDragDrop
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' simmple OLE drag/drop example
    
    ' use class to load 1st file that was dropped, if more than one. Unicode compatible
   On Error GoTo picPreview_OLEDragDrop_Error

    If cImage.LoadPicture_DropedFiles(data, 1, 256, 256) Then

        Call refreshPicBox(picPreview, 256)
        ShowImage False, True
    End If

   On Error GoTo 0
   Exit Sub

picPreview_OLEDragDrop_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_OLEDragDrop of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : rdMapRefresh_Click
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : the refresh button for the rocketdock map
'             it redraws the whole rdMap
'---------------------------------------------------------------------------------------
'
Private Sub rdMapRefresh_Click()
   On Error GoTo rdMapRefresh_Click_Error
        
        Call busyStart
        Call populateRdMap(0) ' show the map from position zero
        Call busyStop
        
        ' we signify that there have been no changes - this is just a refresh
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

rdMapRefresh_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdMapRefresh_Click of Form rDIconConfigForm"

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

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file

    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
    Else
        chkRegistry.Value = 1
        chkSettings.Value = 0
    End If
    
    'if they change then restart? TODO

   On Error GoTo 0
   Exit Sub

registryTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure registryTimer_Timer of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliPreviewSize_Change
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : The size slider that determines the icon size to display
'---------------------------------------------------------------------------------------
'
Private Sub sliPreviewSize_Change()
    Dim FileName As String
    
    
    ' scaling
    
'        If mnuScalePop(Index).Checked = True Then Exit Sub
'    Dim i As Integer
'    For i = mnuScalePop.LBound To mnuScalePop.UBound
'        If mnuScalePop(i).Checked = True Then
'            mnuScalePop(i).Checked = False
'            Exit For
'        End If
'    Next
'    mnuScalePop(Index).Checked = True
'    refreshPicBox(picPreview,256)
'
'
'    'positioning
'
'        If mnuPosSub(Index).Checked = True Then Exit Sub
'    Dim i As Integer
'    For i = mnuPosSub.LBound To mnuScalePop.UBound
'        If mnuPosSub(i).Checked = True Then
'            mnuPosSub(i).Checked = False
'            Exit For
'        End If
'    Next
'    mnuPosSub(Index).Checked = True
'    refreshPicBox(picPreview,256)

    
   On Error GoTo sliPreviewSize_Change_Error

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
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset)
    Else
        If filesIconList.path <> "" Then
            FileName = filesIconList.path
            If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
            FileName = FileName & filesIconList.FileName
            ' refresh the image display
            Call displayPreviewImage(FileName, picPreview, icoSizePreset)
        End If

    End If

   On Error GoTo 0
   Exit Sub

sliPreviewSize_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliPreviewSize_Change of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : treeView_MouseMove
' Author    : beededea
' Date      : 23/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub treeView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim n As Node
   On Error GoTo treeView_MouseMove_Error

  Set n = treeView.HitTest(x, y)
   If n Is Nothing Then
    treeView.ToolTipText = "Click a folder to show the icons contained within"
    ElseIf n.Text = "icons" Then
       treeView.ToolTipText = "The sub-folders within this tree are Rocketdock's own in-built icons"
    ElseIf n.Text = "custom folder" Then
       treeView.ToolTipText = "The sub-folders within this tree are the custom folders that the user can add using the + button below."
    ElseIf n.Text = "my collection" Then
       treeView.ToolTipText = "The sub-folders within this tree are the default folders that come with this enhanced settings utility."
    Else
     treeView.ToolTipText = n.Text
   End If

   On Error GoTo 0
   Exit Sub

treeView_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure treeView_MouseMove of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : treeView_Click
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : populate the file icon list or thumb view from the chosen folder
'---------------------------------------------------------------------------------------
'
Private Sub treeView_Click()
    
   On Error GoTo treeView_Click_Error
   
    Dim path As String
    Dim defaultFolderNodeKey As String
   
    path = treeView.selectedItem.FullPath
   
    Call busyStart

    On Error GoTo l_bypass_parent
    
    'strangely vb6 won't let you test for a Nothing value from a treeview selection so I've had to bodge it
    ' by dealing with the error with the goto above.
    
   'If treeView.selectedItem.Parent = Nothing Then
   '     GoTo l_bypass_parent
   'End If

   
'    If treeView.selectedItem = "icons" Then
'        path = rdAppPath & "\" & treeView.selectedItem.FullPath
'        relativePath = treeView.selectedItem.FullPath
'    Else
'        If treeView.selectedItem.Parent = "icons" Then
'             path = rdAppPath & "\" & treeView.selectedItem.FullPath
'             relativePath = treeView.selectedItem.FullPath
'         Else ' deal with the full paths using the extended selection
             path = treeView.selectedItem.Key
             relativePath = path
'         End If
'    End If

         
l_bypass_parent:
   On Error GoTo treeView_Click_Error
    
    textCurrentFolder.Text = path
    If DirExists(textCurrentFolder.Text) Then
        filesIconList.path = textCurrentFolder.Text
    End If
    
    defaultFolderNodeKey = treeView.selectedItem.Key
    'eg. defaultFolderNodeKey=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\RocketDockSettings", "defaultFolderNodeKey", defaultFolderNodeKey, toolSettingsFile
        
    If picFrameThumbs.Visible = True Then
        btnRefresh_Click
        'populateThumbnails (thumbImageSize)
    End If
        
    Call busyStop
    
   On Error GoTo 0
   Exit Sub

treeView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure treeView_Click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : addtotree
' Author    : beededea
' Date      : 17/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub addtotree(path As String, tv As treeView)
    Dim folder1 As Object
    Dim FS As Object
   On Error GoTo addtotree_Error

    Set FS = CreateObject("Scripting.FileSystemObject")
    If DirExists(path) Then
        For Each folder1 In FS.getFolder(path).SubFolders
            tv.Nodes.Add path, tvwChild, path & "\" & folder1.Name, folder1.Name
            Call addtotree(path & "\" & folder1.Name, tv)
        Next
    End If

   On Error GoTo 0
   Exit Sub

addtotree_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") a folder with the same name already exists in the tree view, choose another folder"

    
End Sub





'---------------------------------------------------------------------------------------
' Procedure : treeView_DblClick
' Author    : beededea
' Date      : 01/06/2019
' Purpose   :     'open a folder
'---------------------------------------------------------------------------------------
'
Private Sub treeView_DblClick()
    
   On Error GoTo treeView_DblClick_Error
   Dim a As String
   Dim fromNode As String
   
    If DirExists(treeView.selectedItem.Key) Then
        ShellExecute 0, vbNullString, treeView.selectedItem.Key, vbNullString, vbNullString, 1
    End If

   On Error GoTo 0
   Exit Sub

treeView_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure treeView_DblClick of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : treeView_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub treeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo treeView_MouseDown_Error

    If Button = 2 Then
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If

   On Error GoTo 0
   Exit Sub

treeView_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure treeView_MouseDown of Form rDIconConfigForm"
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

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
    
    
   On Error GoTo 0
   Exit Sub

txtCurrentIcon_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtCurrentIcon_Change of Form rDIconConfigForm"
End Sub

'Private Sub treeView_BeforeLabelEdit(Cancel As Integer)
'    ' for each click on the tree populate the folder list above
'     textCurrentFolder.Text = treeView.selectedItem
'
'End Sub

'Private Sub txtCurrentIcon_Click()
    'Dim FileName As String
    
    ' display the icon from the alternative settings.ini config.
    'FileName = rdAppPath & "\" & txtCurrentIcon.Text
    
   ' Call displayPreviewImage(FileName, picPreview, icoSizePreset)

'End Sub

'---------------------------------------------------------------------------------------
' Procedure : vScrollThumbs_Change
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Vertical scrollbar for the simulated thumbnail image box
'---------------------------------------------------------------------------------------
'
Private Sub vScrollThumbs_Change()
    Dim useloop As Integer
    
   On Error GoTo vScrollThumbs_Change_Error

    If keyPressOccurred = True Then
        keyPressOccurred = False
        picFrameThumbsGotFocus = True
    Else
        picFrameThumbsGotFocus = False
    End If
    
    If filesIconList.ListCount > 0 Then
        filesIconList.ListIndex = vScrollThumbs.Value - 1
    End If
    
    btnThumbnailView_Click

    removeThumbHighlighting

   On Error GoTo 0
   Exit Sub

vScrollThumbs_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vScrollThumbs_Change of Form rDIconConfigForm"

End Sub

Private Sub removeThumbHighlighting()
    Dim useloop As Integer
    
    'remove the highlighting
    For useloop = 0 To 11
        picIconList(useloop).BorderStyle = 0
        lblThumbName(useloop).BackColor = &HFFFFFF
    Next useloop
    'set the highlighting
    If picFrameThumbsGotFocus = True Then
        If thumbImageSize = 64 Then
            picIconList(0).BorderStyle = 1
        Else
            lblThumbName(0).BackColor = &HFF8080
        End If
    End If

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuFacebook_Click
' Author    : beededea
' Date      : 14/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFacebook_Click()
    Dim answer As VbMsgBoxResult
    'Dim hWnd As Long

    On Error GoTo mnuFacebook_Click_Error
    
    answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page.). Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.path, 1)
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
Private Sub mnuHelp_Click(Index As Integer)

    On Error GoTo mnuHelp_Click_Error

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
    Dim answer As VbMsgBoxResult
    'Dim hWnd As Long

    On Error GoTo mnuLatest_Click_Error

    answer = MsgBox("Download latest version of the program - this button opens a browser window and connects to the widget download page where you can check and download the latest zipped file). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.path, 1)
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
    Dim answer As VbMsgBoxResult
    'Dim hWnd As Long

    On Error GoTo mnuSupport_Click_Error

    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.path, 1)
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
    Dim answer As VbMsgBoxResult
    'Dim hWnd As Long

    On Error GoTo mnuSweets_Click_Error
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "http://www.amazon.co.uk/gp/registry/registry.html?ie=UTF8&id=A3OBFB6ZN4F7&type=wishlist", vbNullString, App.path, 1)
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
    Dim answer As VbMsgBoxResult
    'Dim hWnd As Long

    On Error GoTo mnuWidgets_Click_Error

    answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuWidgets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuWidgets_Click of Form quartermaster"
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
        
    about.Show
    
    If (about.WindowState = 1) Then
        about.WindowState = 0
    End If


    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of Form quartermaster"
End Sub

Private Sub mnuMoreIcons_Click()
    Call btnGetMore_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : menuLeft_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub menuLeft_Click()
    Dim storedFilename  As String
    Dim storedFileName2  As String
    Dim storedTitle  As String
    Dim storedCommand  As String
    Dim storedArguments  As String
    Dim storedWorkingDirectory  As String
    Dim storedShowCmd  As String
    Dim storedOpenRunning  As String
    Dim storedIsSeparator  As String
    Dim storedUseContext  As String
    Dim storedDockletFile  As String
        
    ' take the current icon -1 and read its details and store it
   On Error GoTo menuLeft_Click_Error

    readSettingsIni (rdIconNumber - 1)
        
    storedFilename = sFilename
    storedFileName2 = sFileName2
    storedTitle = sTitle
    storedCommand = sCommand
    storedArguments = sArguments
    storedWorkingDirectory = sWorkingDirectory
    storedShowCmd = sShowCmd
    storedOpenRunning = sOpenRunning
    storedIsSeparator = sIsSeparator
    storedUseContext = sUseContext
    storedDockletFile = sDockletFile
        
    ' take the current icon details and write it into the place of the one to the left (-1)
    readSettingsIni (rdIconNumber)
    
    writeSettingsIni (rdIconNumber - 1)
    
    sFilename = storedFilename
    sFileName2 = storedFileName2
    sTitle = storedTitle
    sCommand = storedCommand
    sArguments = storedArguments
    sWorkingDirectory = storedWorkingDirectory
    sShowCmd = storedShowCmd
    sOpenRunning = storedOpenRunning
    sIsSeparator = storedIsSeparator
    sUseContext = storedUseContext
    sDockletFile = storedDockletFile
        
    ' take the stored icon details and write it into the current location
    writeSettingsIni (rdIconNumber)
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32)
    Call displayIconElement(rdIconNumber - 1, picRdMap(rdIconNumber - 1), 32)
    
    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

menuLeft_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuLeft_Click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : menuright_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub menuright_Click()
    Dim storedFilename  As String
    Dim storedFileName2  As String
    Dim storedTitle  As String
    Dim storedCommand  As String
    Dim storedArguments  As String
    Dim storedWorkingDirectory  As String
    Dim storedShowCmd  As String
    Dim storedOpenRunning  As String
    Dim storedIsSeparator  As String
    Dim storedUseContext  As String
    Dim storedDockletFile  As String
        
    ' take the current icon plus one and read its details and store it
   On Error GoTo menuright_Click_Error

    readSettingsIni (rdIconNumber + 1)
        
    storedFilename = sFilename
    storedFileName2 = sFileName2
    storedTitle = sTitle
    storedCommand = sCommand
    storedArguments = sArguments
    storedWorkingDirectory = sWorkingDirectory
    storedShowCmd = sShowCmd
    storedOpenRunning = sOpenRunning
    storedIsSeparator = sIsSeparator
    storedUseContext = sUseContext
    storedDockletFile = sDockletFile
        
    ' take the current icon details and write it into the place of the one to the right
    readSettingsIni (rdIconNumber)
    
    writeSettingsIni (rdIconNumber + 1)
    
    sFilename = storedFilename
    sFileName2 = storedFileName2
    sTitle = storedTitle
    sCommand = storedCommand
    sArguments = storedArguments
    sWorkingDirectory = storedWorkingDirectory
    sShowCmd = storedShowCmd
    sOpenRunning = storedOpenRunning
    sIsSeparator = storedIsSeparator
    sUseContext = storedUseContext
    sDockletFile = storedDockletFile
        
    ' take the stored icon details and write it into the current location
    writeSettingsIni (rdIconNumber)
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32)
    Call displayIconElement(rdIconNumber + 1, picRdMap(rdIconNumber + 1), 32)

    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

menuright_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuright_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuAdd_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Menu option for the RD Map to add a picture item.
'---------------------------------------------------------------------------------------
'
Private Sub menuAdd_Click()
    Dim useloop As Integer
    Dim thisIcon As Integer
        
    'in the rdSettings.ini file
   On Error GoTo menuAdd_Click_Error
    
    Call busyStart

    'we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    For useloop = rdIconMax To rdIconNumber Step -1
        ' read the rocketdock alternative settings.ini
         readSettingsIni (useloop) ' the alternative settings.ini only exists when RD is set to use it
        
        ' and increment the identifier by one
         writeSettingsIni (useloop + 1) ' the settings.ini only exists when RD is set to use it
    Next useloop
    
    'write the new icon count
    rdIconMax = rdIconMax + 1
    'amend the count in both the settings and rdSettings.ini
    PutINISetting "Software\RocketDock\Icons", "count", rdIconMax, rdSettingsFile

    ' test to see if the picturebox has already been created
    If CheckControlExists(picRdMap(rdIconMax)) Then
        'do nothing
    Else
        Load picRdMap(rdIconMax) ' dynamically extend the number of picture boxes by one
        picRdMap(rdIconMax).Width = 500
        picRdMap(rdIconMax).Height = 500
        picRdMap(rdIconMax).Left = picRdMap(rdIconMax - 1).Left + boxSpacing
        picRdMap(rdIconMax).Top = 30
        picRdMap(rdIconMax).Visible = True
    End If
    
    thisIcon = useloop + 1
    
    'when we arrive at the original position then add a blank item
    sFilename = "\Icons\help.png" ' the default Rocketdock filename for a blank item
    sTitle = ""
    sCommand = ""
    sArguments = ""
    sWorkingDirectory = ""
    sShowCmd = 0
    sOpenRunning = 0
    sUseContext = 0
    
    lblName.Text = ""
    lblTarget.Text = ""
    lblArguments.Text = ""
    lblStartIn.Text = ""
    comboRun.ListIndex = 0 '"Normal"
    comboOpenRunning.ListIndex = 0 ' "Use Global Setting"
    checkPopupMenu.Value = 0
    
    writeSettingsIni (thisIcon)
    
    Call displayIconElement(thisIcon, picRdMap(thisIcon), 32)
    
    Call populateRdMap(0) ' regenerate the map from position zero
      
    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

    Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1) ' click on the picture box
    
    Call busyStop

   On Error GoTo 0
   Exit Sub

menuAdd_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuAdd_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckControlExists
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Function to see whether a control exists
'---------------------------------------------------------------------------------------
'
Public Function CheckControlExists(ctl As Object) As Boolean
   On Error GoTo CheckControlExists_Error

    CheckControlExists = (VarType(ctl) <> vbObject)

   On Error GoTo 0
   Exit Function

CheckControlExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckControlExists of Form rDIconConfigForm"
End Function

'---------------------------------------------------------------------------------------
' Procedure : menuDelete_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : a menu item for the RD Map allowing deletion of an item.
'---------------------------------------------------------------------------------------
'
Private Sub menuDelete_Click()
    Dim useloop As Integer
    Dim thisIcon As Integer
        
    'in the rdSettings.ini file
    'read the menu items from the next item
    'and decrement the identifier by one
    On Error GoTo menuDelete_Click_Error

    Call busyStart
    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    For useloop = rdIconNumber To rdIconMax
        ' read the rocketdock alternative settings.ini
         readSettingsIni (useloop + 1) ' the alternative settings.ini only exists when RD is set to use it
        
         writeSettingsIni (useloop)
    Next useloop
    
    Unload picRdMap(rdIconMax)
    storeLeft = storeLeft - boxSpacing
    'picRdMap(rdIconMax).Visible = False
    
    'write the new icon count
    rdIconMax = rdIconMax - 1
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\RocketDock\Icons", "count", rdIconMax + 1, rdSettingsFile
    
    thisIcon = rdIconNumber
            
    ' load the new icon as an image in the picturebox
    Call displayIconElement(thisIcon, picRdMap(thisIcon), 32)
    
    Call populateRdMap(0) ' regenerate the map from position zero
    
    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

    ' emulate a click on the appropriate icon in the map so that the image and details are shown
    Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1)
   
    Call busyStop

   On Error GoTo 0
   Exit Sub

menuDelete_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuDelete_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(Index As Integer)
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuCoffee_Click_Error
    
    answer = MsgBox(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hWnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form quartermaster"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkTheRegistry
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : if the original settings.ini exist then RD is not using the registry
'---------------------------------------------------------------------------------------
'
Private Sub chkTheRegistry()

   On Error GoTo chkTheRegistry_Error

    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
    Else
        chkRegistry.Value = 1
        chkSettings.Value = 0
    End If

   On Error GoTo 0
   Exit Sub

chkTheRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkTheRegistry of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : backupSettings
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Creates an incrementally named backup of the settings.ini
'---------------------------------------------------------------------------------------
'
Public Sub backupSettings()
    Dim AY() As String
    Dim suffix As String
    Dim maxBound As Integer
    Dim fileVersion As Integer
    Dim bkpSettingsFile As String
    Dim useloop As Integer
    Dim srchSettingsFile As String
    Dim versionNumberAvailable As Integer
    Dim bkpfileFound As Boolean
    
    
        ' set the name of the bkp file
   On Error GoTo backupSettings_Error

        bkpSettingsFile = rdAppPath & "\bkpSettings.ini"
                
        'check for any version of the ini file with a suffix exists
        For useloop = 1 To 32767
            srchSettingsFile = bkpSettingsFile & "." & useloop
          
            If FExists(srchSettingsFile) Then
              ' found a file
              bkpfileFound = True
            Else
              ' no file found use this entry
              GoTo l_exit_bkp_loop
            End If
        Next useloop
        
l_exit_bkp_loop:
        
        If bkpfileFound = True Then
            bkpfileFound = False
            versionNumberAvailable = useloop
            
            'if versionNumberAvailable >= 32767 then
                'versionNumberAvailable = 1
                'If FExists(bkpSettingsFile) Then
                    'delete bkpSettingsFile
                'endif
            'endif
        Else
             versionNumberAvailable = 1
        End If
        
        bkpSettingsFile = bkpSettingsFile & "." & Trim(Str(versionNumberAvailable))
        If Not FExists(bkpSettingsFile) Then
        ' copy the original settings file to a duplicate that we will keep as a safety backup
            If FExists(origSettingsFile) Then
                FileCopy origSettingsFile, bkpSettingsFile
            Else
                FileCopy rdSettingsFile, bkpSettingsFile
            End If
        End If

   On Error GoTo 0
   Exit Sub

backupSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure backupSettings of Form rDIconConfigForm"
        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : menuSmallerIcons_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : right click on the thumbnails moves the picboxes to a more appropriate position for displaying the smaller thumbnail icons
'---------------------------------------------------------------------------------------
'
Private Sub menuSmallerIcons_Click()
    Dim useloop As Integer
    Dim tooltip As String
    Dim suffix As String
    
    ' set the icon size
    
    ' the labels for the smaller thumbnail icon view
   On Error GoTo menuSmallerIcons_Click_Error

    If thumbImageSize = 64 Then ' change to 32
        thumbImageSize = 32
    End If
    
    removeThumbHighlighting
    
    'then populate them and refresh
    btnRefresh_Click

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
    Dim useloop As Integer
    Dim tooltip As String
    Dim suffix As String
    
    On Error GoTo menuLargerThumbs_Click_Error

    If thumbImageSize = 32 Then
        thumbImageSize = 64
    End If
    

       
    'then populate them, perhaps refresh?
    btnRefresh_Click

   On Error GoTo 0
   Exit Sub

menuLargerThumbs_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuLargerThumbs_Click of Form rDIconConfigForm"

End Sub

Private Sub chkBiLinear_Click()

    If chkBiLinear.Value = 1 Then
        If chkBiLinear.Tag = "" Then
            If cImage.isGDIplusEnabled = False Then
                If Not 0 Then
                    chkBiLinear.Tag = "noMsg"
                    On Error Resume Next
                    Debug.Print 1 / 0
                    If Err Then ' uncompiled
                        Err.Clear
                        MsgBox "Non-GDI+ rotation with bilinear interpolation is painfully slow in IDE." & vbCrLf & _
                            "But is acceptable when the routines are compiled", vbInformation + vbOKOnly
                    End If
                End If
            End If
        End If
    End If
    
    cImage.HighQualityInterpolation = chkBiLinear.Value
    Call refreshPicBox(picPreview, 256)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' when you create a token to be shared, you must
    ' destroy it in the Unload or Terminate event
    ' and also reset gdiToken property for each existing class
    If m_GDItoken Then
        If Not cShadow Is Nothing Then cShadow.gdiToken = 0&
        If Not cImage Is Nothing Then
            cImage.gdiToken = 0&
            cImage.DestroyGDIplusToken m_GDItoken
        End If
    End If
    
End Sub

Private Sub mnuBlend_Click()
    
'    Dim Index As Long
'
'    If mnuSubOpts(14).Tag = vbNullString Then
'        MsgBox "There are some functions that cannot be performed on the fly." & vbCrLf & _
'            "This function will permanently change the image. In order to revert back " & vbCrLf & _
'            "to the original image, it must be re-loaded.", vbInformation + vbOKOnly
'    End If
'
'    ' in this example, we will replicate the image 3 times, using different colors
'    ' But since the original is permanently changed, we will make a copy of it,
'    ' change the copy, then render the copy onto the composite image.
'
'    ' This is also a good example on how to render to multiple DIB classes
'
'    Dim cpyImage As c32bppDIB, cmpImage As c32bppDIB
'
'    ' create our composite image, blank but sized appropriately
'    Set cmpImage = New c32bppDIB
'    cmpImage.InitializeDIB cImage.Width * 2, cImage.Height * 2
'
'    For Index = 0 To 3
'        cImage.CopyImageTo cpyImage
'        Select Case Index
'        Case 0: ' do nothing, original will be in upper left
'        Case 1: cpyImage.BlendToColor vbBlue, 25
'        Case 2: cpyImage.BlendToColor vbRed, 25
'        Case 3: cpyImage.BlendToColor vbGreen, 25
'        End Select
'        ' When rendering class to class, pass the destination DIB class as one of the parameters >>>>>>
'        cpyImage.Render 0, (Index And 1) * cImage.Width, (Index \ 2) * cImage.Height, , , , , , , , , , cmpImage
'    Next
'    Set cpyImage = Nothing  ' don't need the copy any longer
'    ' optional step here:
'    cmpImage.ImageType = imgBmpPARGB
'    ' make our composite the current image
'    Set cImage = cmpImage
'    ShowImage True, True    ' display it & show msgbox explaining what you are looking at
'
'    If mnuSubOpts(14).Tag = vbNullString Then
'        mnuSubOpts(14).Tag = "msg shown"
'        MsgBox "Top Left: Original image" & vbCrLf & _
'            "Top Right: Blue blend" & vbCrLf & _
'            "Bottom Left: Red blend" & vbCrLf & _
'            "Bottom Right: Green blend", vbInformation + vbOKOnly, "Blend/Tint Sample"
'    End If

    
End Sub



Private Sub mnuPosSub_Click(Index As Integer)
'    If mnuPosSub(Index).Checked = True Then Exit Sub
'    Dim I As Integer
'    For I = mnuPosSub.LBound To mnuScalePop.UBound
'        If mnuPosSub(I).Checked = True Then
'            mnuPosSub(I).Checked = False
'            Exit For
'        End If
'    Next
'    mnuPosSub(Index).Checked = True
'    Call refreshPicBox(picPreview, 256)
End Sub

Private Sub mnuSaveAs_Click(Index As Integer)

'    Dim sFile As String
'
'    Select Case Index
'    Case 0: ' save as PNG using GDI+
'        sFile = OpenSaveFileDialog(True, "Save As", "png", True)
'        If Not sFile = vbNullString Then
'            ' to force use of GDI+, we can't have any optional PNG properties
'            cImage.PngPropertySet pngProp_ClearProps
'
'
'            If cImage.SaveToFile_PNG(sFile, False) = True Then
'                If MsgBox("PNG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                    If cImage.LoadPicture_File(sFile) = False Then
'                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                    Else
'                        ShowImage True, True
'                    End If
'                End If
'            Else
'                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'            End If
'        End If
'    Case 1 ' save using zLIB
'
'    Case 2 ' save as jpg
'        sFile = OpenSaveFileDialog(True, "Save As", "jpg", True)
'        If Not sFile = vbNullString Then
'            If cImage.SaveToFile_JPG(sFile, , vbWhite, False) = True Then
'                If MsgBox("JPG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                    If cImage.LoadPicture_File(sFile) = False Then
'                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                    Else
'                        ShowImage True, True
'                    End If
'                End If
'            Else
'                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'            End If
'        End If
'    Case 3 ' save as tga options
'
'    Case 4 ' save as GIF
'        sFile = OpenSaveFileDialog(True, "Save As", "gif", True)
'        If Not sFile = vbNullString Then
'            If cImage.SaveToFile_GIF(sFile, True, 200, False) = True Then
'                If MsgBox("GIF successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                    If cImage.LoadPicture_File(sFile) = False Then
'                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                    Else
'                        ShowImage True, True
'                    End If
'                End If
'            Else
'                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'            End If
'        End If
'    Case 5 ' save as BMP using red solid bkg (only applies to images with transparency)
'        sFile = OpenSaveFileDialog(True, "Save As", "bmp", True)
'        If Not sFile = vbNullString Then
'            If cImage.SaveToFile_Bitmap(sFile, , vbRed, False) = True Then
'                If MsgBox("BMP successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                    If cImage.LoadPicture_File(sFile) = False Then
'                        MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                    Else
'                        ShowImage True, True
'                    End If
'                End If
'            Else
'                MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'            End If
'        End If
'    Case 7 ' save as rendered example
'
'        ' There are at least 2 ways to save an image as rendered
'        ' 1. Modify the image directly by calling routines that permanently change the image
'        '   such as MakeLighterDarker, MakeGrayScale, etc. Then simply call the appropriate Save function
'        ' 2. This way, render the image to another DIB class and save that one
'
'        ' This is actually much simpler than it looks, but we will try to accomodate all the
'        ' user-selected options on the form & render this to another DIB then save that DIB
'        ' The only additional steps would be when the image is rotated. If that is the case,
'        ' then you would need to resize the new DIB appropriately. Example does this...
'
'        ' The other important thing to remember when rendering DIB to DIB is to pass the
'        ' optional destDibHost parameter in the render function. Example does this...
'
'        ' Note: This sample does not honor the shadow if it is applied.  If you want to also
'        ' include a shadow, simply adjust your image size to accomodate the shadow and then
'        ' render it first, then render the image over the shadow.
'
'        Dim rImage As c32bppDIB
'        Dim newWidth As Long, newHeight As Long
'        Dim mirrorOffsetX As Long, mirrorOffsetY As Long
'        Dim negAngleOffset As Long
'        Dim LightAdjustment As Single
'
'        sFile = OpenSaveFileDialog(True, "Save As", "png", True)
'        If Not sFile = vbNullString Then
'
'            mirrorOffsetX = 1
'            mirrorOffsetY = 1
'            If mnuSubOpts(5).Checked = True Then negAngleOffset = -1 Else negAngleOffset = 1
'            'LightAdjustment = CSng(Val(mnuLight(0).Tag))
''            Select Case True    ' scaling options from menu
''                Case mnuScalePop(1).Checked ' only scale down as needed
''                    cImage.ScaleImage picPreview.ScaleWidth, picPreview.ScaleHeight, newWidth, newHeight, scaleDownAsNeeded
''                Case mnuScalePop(2).Checked ' scale up and/or down
''                    cImage.ScaleImage picPreview.ScaleWidth, picPreview.ScaleHeight, newWidth, newHeight, ScaleToSize
''                Case mnuScalePop(3).Checked ' reduce by 1/2
''                    cImage.ScaleImage cImage.Width \ 2, cImage.Height \ 2, newWidth, newHeight, ScaleToSize
''                Case mnuScalePop(4).Checked ' enlarge by 1/2
''                    cImage.ScaleImage cImage.Width * 1.5, cImage.Height * 1.5, newWidth, newHeight, ScaleToSize
''                Case mnuScalePop(5).Checked ' enlarge by 1/2
''                    cImage.ScaleImage cImage.Width * 1.5, cImage.Height * 1.5, newWidth, newHeight, ScaleToSize
''                Case mnuScalePop(6).Checked ' actual size
''                    newWidth = cImage.Width: newHeight = cImage.Height
''                Case mnuScalePop(7).Checked ' scale to size
''                    cImage.ScaleImage picPreview.ScaleWidth, picPreview.ScaleHeight, newWidth, newHeight, ScaleToSize
''            End Select
'
'            ' the cboAngle entries are at 15 degree intervals, so we simply multiply ListIndex by 15
''            If (cboAngle.ListIndex * 15) Mod 360 Then ' rotated
''                ' rotation: size the dib to the maximum size needed to handle all rotation angles
''                newWidth = Sqr(newWidth * newWidth + newHeight * newHeight)
''                newHeight = newWidth
''            End If
'
'            ' create a new DIB & size it
'            Set rImage = New c32bppDIB
'            rImage.InitializeDIB newWidth, newHeight
'
'            ' rendering to the center (last parameter) as shown below is optional but if rendering
'            ' rotated then always render to the center of the target area
'
'            Dim ttemp As Integer
'            ttemp = -1
'            ' To correctly render DIB to DIB, always pass the target DIB as the optional destDibHost parameter
'            ' When rendering DIB to DIB, the hDC is ignored and that is why we pass zero.
'            cImage.Render 0, newWidth \ 2, newHeight \ 2, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
'                100, , , rImage, ttemp, LightAdjustment, 0 * negAngleOffset, True
'
'            rImage.TrimImage True, trimAll
'            ' ^^ if the image was rotated, you should call this to remove any transparent "borders"
'
'            If rImage.SaveToFile_PNG(sFile, False) Then
'                MsgBox "PNG successfully created", vbInformation + vbOKOnly, "Success"
'            Else
'                MsgBox "PNG failed to be created", vbExclamation + vbOKOnly, "Failure"
'            End If
'
'        End If
'
'    End Select

End Sub

Private Sub mnuScalePop_Click(Index As Integer)

'    If mnuScalePop(Index).Checked = True Then Exit Sub
'    Dim I As Integer
'    For I = mnuScalePop.LBound To mnuScalePop.UBound
'        If mnuScalePop(I).Checked = True Then
'            mnuScalePop(I).Checked = False
'            Exit For
'        End If
'    Next
'    mnuScalePop(Index).Checked = True
'    Call refreshPicBox(picPreview, 256)
End Sub


Private Sub mnuSubOpts_Click(Index As Integer)

    ' The 1st two options will be disabled if you do not have GDI+ installed
    
    Select Case Index
    Case 0: ' do not use GDI+
        If mnuSubOpts(Index).Checked = True Then Exit Sub
        cImage.isGDIplusEnabled = False
        mnuSubOpts(0).Checked = Not mnuSubOpts(0).Checked
        mnuSubOpts(1).Checked = False
        
        If m_GDItoken Then  ' when using token, we'll clean up here
            cImage.DestroyGDIplusToken m_GDItoken
            m_GDItoken = 0&
            cImage.gdiToken = m_GDItoken ' reset the token
            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
        End If
        
        Call refreshPicBox(picPreview, 256)
    
    Case 1: ' always usge GDI+.
        If mnuSubOpts(Index).Checked = True Then Exit Sub
        mnuSubOpts(0).Checked = False ' remove checkmark on "Don't Use GDI+"
        mnuSubOpts(1).Checked = True  ' show using GDI+
        cImage.isGDIplusEnabled = True
        ' verify it enabled correct and get a token to share
        If cImage.isGDIplusEnabled Then
            m_GDItoken = cImage.CreateGDIplusToken()
            cImage.gdiToken = m_GDItoken
            If Not cShadow Is Nothing Then cShadow.gdiToken = m_GDItoken
        End If
        ' tell GDI+ that we want high quality interpolation
        If chkBiLinear.Value = 0 Then chkBiLinear.Value = 1 Else Call refreshPicBox(picPreview, 256)
                
    Case 7: ' save as
            
    End Select
ExitRoutine:
End Sub

Private Sub mnuTGA_Click(Index As Integer)
    
'    Dim sFile As String
'    sFile = OpenSaveFileDialog(True, "Save As", "tga", True)
'    If Not sFile = vbNullString Then
'        If cImage.SaveToFile_TGA(sFile, (Index = 0), False, True, False) = True Then
'            If MsgBox("TGA successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                If cImage.LoadPicture_File(sFile) = False Then
'                    MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                Else
'                    ShowImage True, True
'                End If
'            End If
'        Else
'            MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'        End If
'    End If
    
End Sub


Private Sub mnuZlibPng_Click(Index As Integer)
    
'    Dim sFile As String
'    ' by setting optional parameters, class will use zLIB over GDI+
'    ' to the contrary, if no parameters are set, class uses GDI+ over zLIB
'    Select Case Index
'    Case 0: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterDefault
'    Case 1: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterNone
'    Case 2: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjLeft
'    Case 3: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjTop
'    Case 4: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdjAvg
'    Case 5: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterPaeth
'    Case 6: cImage.PngPropertySet pngProp_FilterMethod, eFilterMethods.filterAdaptive
'    End Select
'
'    sFile = OpenSaveFileDialog(True, "Save As", "png", True)
'    If Not sFile = vbNullString Then
'        If cImage.SaveToFile_PNG(sFile, False) = True Then
'            If MsgBox("PNG successfully created. Load it now?", vbQuestion + vbYesNo, "Success") = vbYes Then
'                If cImage.LoadPicture_File(sFile) = False Then
'                    MsgBox "Could not load that new image -- Error in my routines?", vbExclamation + vbOKOnly
'                Else
'                    ShowImage True, True
'                End If
'            End If
'        Else
'            MsgBox "GIF failed to be created", vbExclamation + vbOKOnly, "Failure"
'        End If
'    End If
'
'ExitRoutine:
End Sub



Private Sub refreshPicBox(picBox As PictureBox, iconSizing As Integer)

    Dim newWidth As Long, newHeight As Long
    Dim mirrorOffsetX As Long, mirrorOffsetY As Long
    Dim negAngleOffset As Long
    Dim x As Long, y As Long
    Dim ShadowOffset As Long
    Dim LightAdjustment As Single
    
    ' This one routine handles all the options of the sample form
    
    'ShadowOffset = Val(mnuShadowDepth(0).Tag) + 2   ' set shadow's blur depth as needed
    
    mirrorOffsetX = 1
    mirrorOffsetY = 1
    
    'LightAdjustment = CSng(Val(mnuLight(0).Tag))
    
'    Select Case True    ' scaling options from menu
'        Case mnuScalePop(0).Checked ' only scale down as needed
'            newWidth = 16: newHeight = 16
'        Case mnuScalePop(1).Checked ' scale up and/or down
'            newWidth = 32: newHeight = 32
'        Case mnuScalePop(2).Checked ' reduce by 1/2
'            newWidth = 64: newHeight = 64
'        Case mnuScalePop(3).Checked ' enlarge by 1/2
'            newWidth = 128: newHeight = 128
'        Case mnuScalePop(4).Checked ' 256
'            newWidth = 256: newHeight = 256
'        Case mnuScalePop(5).Checked ' scale to size
'            cImage.ScaleImage picBox.ScaleWidth, picBox.ScaleHeight, newWidth, newHeight, ScaleToSize
'        Case mnuScalePop(6).Checked ' actual size
'            newWidth = cImage.Width: newHeight = cImage.Height
'    End Select

    newWidth = iconSizing: newHeight = iconSizing
    
    ' in this sample form, to make it easier to calculate rendering X,Y coordinates,
    ' we will always pass the X,Y of where the center of the image should appear.
    ' This way, whether rotating or not, we can use the same Render call without
    ' modifying the destination X,Y and CenterOnDestXY paramters based on rotating or not
    Select Case True
        Case mnuPosSub(0).Checked   ' centered on canvas
            x = (picBox.ScaleWidth - newWidth) \ 2
            y = (picBox.ScaleHeight - newHeight) \ 2
        Case mnuPosSub(2).Checked   ' top right
            x = picBox.ScaleWidth - newWidth
        Case mnuPosSub(3).Checked   ' bottom left
            y = picBox.ScaleHeight - newHeight
        Case mnuPosSub(4).Checked   ' bottom right
            x = picBox.ScaleWidth - newWidth
            y = picBox.ScaleHeight - newHeight
        Case mnuPosSub(1).Checked   ' top left
    End Select
    
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
        cShadow.Render picBox.hDC, x + newWidth \ 2 + ShadowOffset, y + newHeight \ 2 + ShadowOffset, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
            55, , , , , LightAdjustment, 0, True
    End If
    
    Dim ttemp As Integer
    ttemp = -1
    
    cImage.Render picBox.hDC, x + newWidth \ 2, y + newHeight \ 2, newWidth * 1, newHeight * 1, , , , , _
        100, , , , -1, 0, 0, True
    
    picBox.Refresh

End Sub

Private Sub ShowImage(Optional bRefresh As Boolean = True, Optional DragDropCutPast As Boolean)

'    Dim sSource As Variant
'    Dim Cx As Long, Cy As Long
'
'    If Not DragDropCutPast Then
'        If optSource(0).Enabled = True Then
'            If optSource(0) = True Then ' from file
'
'                Select Case cboType.ListIndex
'                Case 0: sSource = "Forest.bmp"
'                Case 1: sSource = "Alpha-ARGB.bmp"
'                Case 2: sSource = "Alpha-pARGB.bmp"
'                Case 3: sSource = "Knight.gif"
'                Case 4: sSource = "Desktop.ico"
'                Case 5: sSource = "XP-Alpha.ico"
'                Case 6: sSource = "Vista-PNG.ico"
'                Case 7: sSource = "Risk.jpg"
'                Case 8: sSource = "Spider.png"
'                Case 9: sSource = "Lion.wmf"
'                Case 10: sSource = "Hand.cur"
'                Case 11: sSource = "GlobalSearch.tga"
'                Case 12: sSource = OpenSaveFileDialog(False, "Select Image")
'                End Select
'                If cboType.ListIndex < cboType.ListCount - 1 Then
'                    If Right$(App.path, 1) = "\" Then
'                        sSource = App.path & sSource
'                    Else
'                        sSource = App.path & "\" & sSource
'                    End If
'                End If
'                cImage.LoadPicture_File sSource, 256, 256, (cboType.ListIndex = 12)
'                ' ^^ the end parameters: 256,256 is just telling the class that
'                ' we want that size icon if one exists in the passed resource. If not,
'                ' then give us the one closest to it & best quality too.
'                ' The final parameter is telling the class to cache the image bytes once it is loaded.
'                ' I will use those bytes to reload the image as needed vs having the user re-select
'                ' the image from the browser. Look in this sample form for
'                ' cimage.LoadPicture_FromOrignalFormat to see how those bytes are used
'
'            Else    ' from resource
'                Select Case cboType.ListIndex
'                Case 0 ' bitmap
'                    sSource = vbResBitmap
'                Case 4 'icon
'                    sSource = vbResIcon
'                Case 10 ' cursor
'                    sSource = vbResCursor
'                Case 12 ' browse for file, n/a
'                    optSource(0) = True ' change source option & browser pop up
'                    Exit Sub
'                Case Else ' pARGB bmp, ARGB bmp, GIF, alpha icon, png icon, jpg, png, wmf, tga
'                    sSource = "Custom"
'                End Select
'                cImage.LoadPicture_Resource (cboType.ListIndex + 101) & "LaVolpe", sSource, VB.Global, 256, 256, , , 32
'                ' ^^ the last two parameters: 256,256 is just telling the class that
'                ' we want that size icon if one exists in the passed resource. If not,
'                ' then give us the one closest to it & best quality too.
'            End If
'        End If
'    End If
    
'    If Not cShadow Is Nothing Then
'        '
'    Else
'        If bRefresh Then Call refreshPicBox(picPreview, 256)
'    End If
'
'    If Me.Tag = "" Then
'        If optSource(1) = True And cboType.ListIndex = 10 Then
'            On Error Resume Next    ' only show this message in IDE
'            Debug.Print 1 / 0
'            If Err Then
'                MsgBox "Notice this is black and white." & vbCrLf & _
'                    "VB, while in IDE, forces 2 color cursors to be black & white, even though they may not be." & vbCrLf & _
'                    "When the cursor is loaded from a resource file when the project is compiled, the cursor magically shows its colors", vbInformation + vbOKOnly
'            End If
'            Me.Tag = "Message Shown"    ' only show message once
'        End If
'    End If
'
End Sub








Private Function OpenSaveFileDialog(bSave As Boolean, DialogTitle As String, Optional DefaultExt As String, Optional SingleFilter As Boolean) As String

'    ' using API version vs commondialog enables Unicode filenames to be passed to c32bppDIB classes
'    Dim ofn As OPENFILENAME
'    Dim rtn As Long
'    Dim bUnicode As Boolean
'
'    With ofn
'        .lStructSize = Len(ofn)
'        .hwndOwner = Me.hWnd
'        .hInstance = App.hInstance
'        If SingleFilter Then
'            Select Case DefaultExt
'            Case "png"
'                .lpstrFilter = "PNG" & vbNullChar & "*.png" & vbNullChar
'            Case "jpg"
'                .lpstrFilter = "JPG" & vbNullChar & "*.jpg" & vbNullChar
'            Case "tga"
'                .lpstrFilter = "TGA (Targa)" & vbNullChar & "*.tga" & vbNullChar
'            Case "gif"
'                .lpstrFilter = "GIF" & vbNullChar & "*.gif" & vbNullChar
'            Case "bmp"
'                .lpstrFilter = "Bitmap" & vbNullChar & "*.bmp" & vbNullChar
'            End Select
'        Else
'            .lpstrFilter = "Image Files" & vbNullChar & "*gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png;*.tga"
'            If cImage.isGDIplusEnabled Then
'                .lpstrFilter = .lpstrFilter & ";*.tiff"
'            End If
'            .lpstrFilter = .lpstrFilter & vbNullChar & "Bitmaps" & vbNullChar & "*.bmp" & vbNullChar & "GIFs" & vbNullChar & "*.gif" & vbNullChar & _
'                            "Icons/Cursors" & vbNullChar & "*.ico;*.cur" & vbNullChar & "JPGs" & vbNullChar & "*.jpg;*.jpeg" & vbNullChar & _
'                            "Meta Files" & vbNullChar & "*.wmf;*.emf" & vbNullChar & "PNGs" & vbNullChar & "*.png" & vbNullChar & "TGAs (Targa)" & vbNullChar & "*.tga" & vbNullChar
'            If cImage.isGDIplusEnabled Then
'                .lpstrFilter = .lpstrFilter & "TIFFs" & vbNullChar & "*.tiff" & vbNullChar
'            End If
'            .lpstrFilter = .lpstrFilter & "All Files" & vbNullChar & "*.*" & vbNullChar
'        End If
'        .lpstrDefExt = DefaultExt
'        .lpstrFile = String$(256, 0)
'        .nMaxFile = 256
'        .nMaxFileTitle = 256
'        .lpstrTitle = DialogTitle
'        .flags = OFN_LONGNAMES Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_DONTADDTORECENT _
'                Or OFN_NOCHANGEDIR
'        ' ^^ don't want to change paths otherwise VB IDE locks folder until IDE is closed
'        If bSave Then
'            .flags = .flags Or OFN_CREATEPROMPT Or OFN_OVERWRITEPROMPT
'        Else
'            .flags = .flags Or OFN_FILEMUSTEXIST
'        End If
'
'        bUnicode = Not (IsWindowUnicode(GetDesktopWindow) = 0&)
'        If bUnicode Then
'            .lpstrInitialDir = StrConv(.lpstrInitialDir, vbUnicode)
'            .lpstrFile = StrConv(.lpstrFile, vbUnicode)
'            .lpstrFilter = StrConv(.lpstrFilter, vbUnicode)
'            .lpstrTitle = StrConv(.lpstrTitle, vbUnicode)
'            .lpstrDefExt = StrConv(.lpstrDefExt, vbUnicode)
'        End If
'        .lpstrFileTitle = .lpstrFile
'    End With
'
'    If bUnicode Then
'        If bSave Then
'            rtn = GetSaveFileNameW(ofn)
'        Else
'            rtn = GetOpenFileNameW(ofn)
'        End If
'        If rtn > 0& Then
'            If bUnicode Then
'                rtn = lstrlenW(ByVal ofn.lpstrFile)
'                OpenSaveFileDialog = StrConv(Left$(ofn.lpstrFile, rtn * 2), vbFromUnicode)
'            End If
'        End If
'    Else
'        If bSave Then
'            rtn = GetSaveFileName(ofn)
'        Else
'            rtn = GetOpenFileName(ofn)
'        End If
'        If rtn > 0& Then
'            rtn = lstrlen(ofn.lpstrFile)
'            OpenSaveFileDialog = Left$(ofn.lpstrFile, rtn)
'        End If
'    End If
'
'ExitRoutine:
End Function






Private Sub vScrollThumbs_KeyDown(KeyCode As Integer, Shift As Integer)
    Call getkeypress(KeyCode)
End Sub
