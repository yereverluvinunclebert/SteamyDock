VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Object = "{FB95F7DD-5143-4C75-88F9-A53515A946D7}#2.0#0"; "CCRTreeView.ocx"
Begin VB.Form rDIconConfigForm 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rocketdock Icon Settings VB6"
   ClientHeight    =   8895
   ClientLeft      =   150
   ClientTop       =   -135
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
   Icon            =   "arse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3195
      Top             =   6345
   End
   Begin VB.Frame Frame 
      Caption         =   "Configuration"
      Height          =   600
      Index           =   0
      Left            =   2100
      TabIndex        =   56
      Top             =   4515
      Visible         =   0   'False
      Width           =   1665
      Begin VB.CheckBox chkBiLinear 
         Caption         =   "Quality Sizing"
         Height          =   240
         Left            =   90
         TabIndex        =   57
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
         TabIndex        =   58
         ToolTipText     =   "To Paste: Click on display box and press Ctrl+V"
         Top             =   5865
         Width           =   3840
      End
   End
   Begin VB.CommandButton btnWorking 
      Caption         =   "Working"
      Height          =   510
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3900
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer registryTimer 
      Interval        =   2500
      Left            =   3195
      Top             =   7065
   End
   Begin VB.PictureBox picRdThumbFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   90
      ScaleHeight     =   705
      ScaleWidth      =   9660
      TabIndex        =   48
      Top             =   4545
      Visible         =   0   'False
      Width           =   9660
      Begin VB.HScrollBar rdMapHScroll 
         Height          =   120
         Left            =   45
         Max             =   100
         TabIndex        =   52
         Top             =   540
         Visible         =   0   'False
         Width           =   9630
      End
      Begin VB.CommandButton btnMapNext 
         Height          =   450
         Left            =   9210
         Picture         =   "arse.frx":20FA
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Scroll the RD map to the right"
         Top             =   45
         Width           =   450
      End
      Begin VB.CommandButton btnMapPrev 
         Height          =   450
         Left            =   45
         Picture         =   "arse.frx":263F
         Style           =   1  'Graphical
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   63
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
      Picture         =   "arse.frx":2B8C
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
      Picture         =   "arse.frx":2EE8
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Refresh the icon map"
      Top             =   4785
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame FrameFolders 
      Caption         =   "Folders"
      Height          =   4500
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "The current list of known icon folders"
      Top             =   15
      Width           =   4005
      Begin CCRTreeView.TreeView folderTreeView 
         Height          =   3180
         Left            =   135
         TabIndex        =   76
         ToolTipText     =   "These are the icon folders available to Rocketdock"
         Top             =   630
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5609
         VisualTheme     =   1
         LineStyle       =   1
         LabelEdit       =   1
         ShowTips        =   -1  'True
         Indentation     =   38
      End
      Begin VB.ComboBox comboDockType 
         Height          =   330
         ItemData        =   "arse.frx":32F1
         Left            =   2205
         List            =   "arse.frx":32FB
         TabIndex        =   78
         Text            =   "RocketDock"
         ToolTipText     =   "Select Rocketdock or an open source dock"
         Top             =   3990
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Frame frmRegistry 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   915
         TabIndex        =   68
         Top             =   3945
         Width           =   1170
         Begin VB.CheckBox chkRegistry 
            Height          =   255
            Left            =   45
            TabIndex        =   70
            ToolTipText     =   "This will tell you whether Rocketdock is saving to the registry"
            Top             =   90
            Width           =   195
         End
         Begin VB.CheckBox chkSettings 
            Height          =   225
            Left            =   645
            TabIndex        =   69
            ToolTipText     =   "This tells you whether Rocketdock is saving to the settings.ini file"
            Top             =   105
            Width           =   210
         End
         Begin VB.Label lblReg 
            Caption         =   "Reg."
            Height          =   240
            Left            =   300
            TabIndex        =   72
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   105
            Width           =   450
         End
         Begin VB.Label lblSet 
            Caption         =   "Set."
            Height          =   240
            Left            =   870
            TabIndex        =   71
            ToolTipText     =   "These tell you whether Rocketdock is saving to the regisstry or the settings.ini file"
            Top             =   105
            Width           =   240
         End
      End
      Begin VB.TextBox textCurrentFolder 
         Height          =   330
         Left            =   135
         TabIndex        =   32
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
         TabIndex        =   31
         ToolTipText     =   "This button can remove a custom folder from the treeview above"
         Top             =   3990
         Width           =   360
      End
      Begin VB.CommandButton btnAddFolder 
         Caption         =   "+"
         Height          =   345
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   17
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
      Picture         =   "arse.frx":3317
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
      ToolTipText     =   "Thumbnail or File Viewer Window"
      Top             =   15
      Width           =   5895
      Begin VB.Frame frmNoFilesFound 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1635
         TabIndex        =   79
         Top             =   1860
         Visible         =   0   'False
         Width           =   2220
         Begin VB.Label lblNoFilesFound 
            Caption         =   "No files found"
            Height          =   285
            Left            =   525
            TabIndex        =   80
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "+"
         Height          =   270
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Set the current selected icon into the dock (double-click on the icon)"
         Top             =   240
         Width           =   270
      End
      Begin VB.CommandButton btnRefresh 
         Height          =   270
         Left            =   5085
         Picture         =   "arse.frx":35C1
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Refresh the Icon List"
         Top             =   240
         Width           =   210
      End
      Begin VB.CommandButton btnKillIcon 
         Height          =   255
         Left            =   165
         Picture         =   "arse.frx":39CA
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Delete the currently selected icon file"
         Top             =   4020
         Width           =   240
      End
      Begin VB.TextBox textCurrIconPath 
         Height          =   330
         Left            =   1035
         TabIndex        =   24
         Text            =   "textCurrIconPath"
         ToolTipText     =   "Shows the selected icon file name"
         Top             =   210
         Width           =   3660
      End
      Begin VB.ComboBox comboIconTypesFilter 
         Height          =   330
         ItemData        =   "arse.frx":3BF7
         Left            =   510
         List            =   "arse.frx":3C0D
         TabIndex        =   21
         Text            =   "All Normal Icons"
         ToolTipText     =   "Filter icon types to display"
         Top             =   3975
         Width           =   2790
      End
      Begin VB.CommandButton btnGetMore 
         Caption         =   "Get More"
         Height          =   345
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Click to install more icons"
         Top             =   3975
         Width           =   1710
      End
      Begin VB.CommandButton btnThumbnailView 
         Height          =   270
         Left            =   5355
         Picture         =   "arse.frx":3C78
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "View as thumbnails"
         Top             =   240
         Width           =   285
      End
      Begin VB.CommandButton btnFileListView 
         Height          =   270
         Left            =   5355
         Picture         =   "arse.frx":3E86
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "View as a file listing"
         Top             =   240
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picFrameThumbs 
         BackColor       =   &H00FFFFFF&
         Height          =   3210
         Left            =   120
         ScaleHeight     =   3150
         ScaleWidth      =   5490
         TabIndex        =   33
         ToolTipText     =   "Double-click an icon to set it into the dock"
         Top             =   615
         Width           =   5550
         Begin VB.Frame frmThumbLabel 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   400
            Index           =   0
            Left            =   60
            TabIndex        =   65
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
               Height          =   425
               Index           =   0
               Left            =   0
               TabIndex        =   66
               Top             =   -15
               Width           =   1095
               WordWrap        =   -1  'True
            End
         End
         Begin VB.VScrollBar vScrollThumbs 
            CausesValidation=   0   'False
            Height          =   3180
            LargeChange     =   12
            Left            =   5265
            SmallChange     =   4
            TabIndex        =   34
            Top             =   -30
            Width           =   240
         End
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
            Left            =   195
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   67
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   64
            ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
            Top             =   45
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
         Width           =   5580
      End
      Begin VB.Label lblIconName 
         Caption         =   "Icon Name:"
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.Frame frameProperties 
      Caption         =   "Properties"
      Height          =   3495
      Left            =   4230
      TabIndex        =   0
      ToolTipText     =   "The Icon Properties Window"
      Top             =   4530
      Width           =   5895
      Begin VB.TextBox txtDbg02 
         Height          =   315
         Left            =   5370
         TabIndex        =   82
         Text            =   "txtDbg01"
         Top             =   2370
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtDbg01 
         Height          =   315
         Left            =   5370
         TabIndex        =   81
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
         Picture         =   "arse.frx":4262
         ScaleHeight     =   795
         ScaleWidth      =   825
         TabIndex        =   74
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
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Sets the icon characteristics but you will need to restart to make it 'fix'"
         Top             =   3030
         Width           =   1470
      End
      Begin VB.TextBox txtCurrentIcon 
         Height          =   345
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "txtCurrentIcon"
         ToolTipText     =   "Double click on an image above to set the current icon"
         Top             =   690
         Width           =   4305
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
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Press to select a target file (or right click)"
         Top             =   1095
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
         Height          =   330
         ItemData        =   "arse.frx":4CDD
         Left            =   1395
         List            =   "arse.frx":4CEA
         TabIndex        =   6
         Text            =   "Use Global Setting"
         ToolTipText     =   "Choose what to do if the chosen app is already running"
         Top             =   2640
         Width           =   2145
      End
      Begin VB.ComboBox comboRun 
         Height          =   330
         ItemData        =   "arse.frx":4D11
         Left            =   1395
         List            =   "arse.frx":4D1E
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
         ToolTipText     =   "Add any additional arguments that the target file operation requires eg. -s -t 00 -f "
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
      Begin VB.TextBox lblName 
         Height          =   345
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "The name of the icon as it appears on the dock"
         Top             =   300
         Width           =   4305
      End
      Begin VB.Label lblCurrentIcon 
         Caption         =   "Current Icon:"
         Height          =   225
         Left            =   345
         TabIndex        =   29
         Top             =   750
         Width           =   1215
      End
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
         Left            =   4800
         TabIndex        =   26
         ToolTipText     =   "This is Rocketdock icon number one."
         Top             =   1875
         Width           =   480
      End
      Begin VB.Label lblPopUp 
         Caption         =   "Popup Menu:"
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label lblDisplaySpecialActions 
         Caption         =   "Display Special Actions"
         Height          =   225
         Left            =   1665
         TabIndex        =   14
         ToolTipText     =   "If you want extra options to appear when you right click on an icon, enable this checkbox"
         Top             =   3030
         Width           =   1965
      End
      Begin VB.Label lblArgument 
         Caption         =   "Arguments:"
         Height          =   225
         Left            =   450
         TabIndex        =   13
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label lblStartIn 
         Caption         =   "Start in:"
         Height          =   225
         Left            =   735
         TabIndex        =   12
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label lblRun 
         Caption         =   "Run:"
         Height          =   225
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblOpenRunning 
         Caption         =   "Open Running:"
         Height          =   225
         Left            =   225
         TabIndex        =   10
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label lblTarget 
         Caption         =   "Target:"
         Height          =   225
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
      Height          =   4290
      Left            =   120
      TabIndex        =   59
      ToolTipText     =   "The Preview Pane"
      Top             =   4530
      Width           =   4000
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
         TabIndex        =   62
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
         TabIndex        =   61
         ToolTipText     =   "select the next icon"
         Top             =   240
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
         Height          =   3060
         Left            =   285
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   204
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   230
         TabIndex        =   60
         ToolTipText     =   "This is the currently selected icon scaled to fit the preview box"
         Top             =   675
         Width           =   3450
      End
      Begin CCRSlider.Slider sliPreviewSize 
         Height          =   300
         Left            =   30
         TabIndex        =   83
         ToolTipText     =   "Icon Size"
         Top             =   3855
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
         Left            =   2130
         TabIndex        =   77
         Top             =   3735
         Width           =   1785
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
         Left            =   120
         TabIndex        =   75
         Top             =   3735
         Width           =   1950
      End
   End
   Begin VB.Frame frameButtons 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   210
      TabIndex        =   36
      Top             =   7680
      Width           =   10080
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate Dock"
         Enabled         =   0   'False
         Height          =   360
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Makes a whole NEW rocketdock - use with care!"
         Top             =   765
         Width           =   1755
      End
      Begin VB.CommandButton btnBackup 
         Caption         =   "Backup"
         Height          =   345
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Backup or create bkpSettings.ini"
         Top             =   405
         Width           =   1485
      End
      Begin VB.CommandButton btnSaveRestart 
         Caption         =   "Save && Restart"
         Height          =   345
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "A save and restart of Rocketdock is required when any icon changes have been made"
         Top             =   780
         Width           =   1485
      End
      Begin VB.CommandButton btnCloseCancel 
         Caption         =   " Close"
         Height          =   345
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Cancel the current operation and close the window"
         Top             =   780
         Width           =   1470
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help"
         Height          =   345
         Left            =   8430
         MousePointer    =   14  'Arrow and Question
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Help on this utility"
         Top             =   405
         Width           =   1470
      End
      Begin VB.CheckBox chkConfirmSaves 
         Height          =   225
         Left            =   4035
         TabIndex        =   38
         ToolTipText     =   "Confirmation on saves and deletes"
         Top             =   450
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.CommandButton btnDefaultIcon 
         Caption         =   "Default Icon"
         Height          =   330
         Left            =   1830
         TabIndex        =   37
         ToolTipText     =   "Not implemented yet"
         Top             =   435
         Width           =   1725
      End
      Begin VB.Label lblToggleDialogs 
         Caption         =   "Toggle info. dialogs"
         Height          =   240
         Left            =   4290
         TabIndex        =   41
         ToolTipText     =   "This will toggle on/off most of the information pop-ups"
         Top             =   435
         Width           =   2010
      End
   End
   Begin VB.Menu mnuTrgtMenu 
      Caption         =   "Target Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuTrgtSeparator 
         Caption         =   "target = Separator"
      End
      Begin VB.Menu mnuTrgtFolder 
         Caption         =   "target = Folder"
      End
      Begin VB.Menu mnuTrgtMyComputer 
         Caption         =   "target = My Computer"
      End
      Begin VB.Menu mnuTrgtShutdown 
         Caption         =   "target = Shutdown"
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
      Begin VB.Menu mnuTrgtAdministrative 
         Caption         =   "target = Administrative Tools"
      End
      Begin VB.Menu mnuTrgtRecycle 
         Caption         =   "target = Recycle Bin"
      End
      Begin VB.Menu mnuTrgtDock 
         Caption         =   "target = Dock Settings"
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
   Begin VB.Menu mnuMainOpts 
      Caption         =   "Other Options"
      Enabled         =   0   'False
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
   End
   Begin VB.Menu rdMapMenu 
      Caption         =   "The Map Menu"
      Visible         =   0   'False
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
         Begin VB.Menu mnuAddSeparator 
            Caption         =   "Add a Separator"
         End
         Begin VB.Menu mnuAddFolder 
            Caption         =   "Add Folder"
         End
         Begin VB.Menu mnuAddMyComputer 
            Caption         =   "Add My Computer"
         End
         Begin VB.Menu mnuAddShutdown 
            Caption         =   "Add Shutdown"
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
         Begin VB.Menu mnuAddAdministrative 
            Caption         =   "Add Administrative Tools"
         End
         Begin VB.Menu mnuAddRecycle 
            Caption         =   "Add Recycle Bin"
         End
         Begin VB.Menu mnuAddDock 
            Caption         =   "Add Dock Settings"
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
      Visible         =   0   'False
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
      Begin VB.Menu mnuHelp 
         Caption         =   "Utility Help"
         Index           =   4
      End
      Begin VB.Menu mnuOnline 
         Caption         =   "Online Help and other options"
         Begin VB.Menu mnuHelpPdf 
            Caption         =   "View Help (PDF)"
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
      Begin VB.Menu mnuLicence 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuRocketDock 
         Caption         =   "Set RocketDock Location"
      End
      Begin VB.Menu mnuseparator1 
         Caption         =   ""
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Turn Debugging ON"
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

Option Explicit
'--------------------------------------------------------------------------------------------------------------
' Form Module : rDIconConfigFrm
' Author      : Dean Beedell
' Date        : 20/06/2019
'
'           This is the first VB6 project that I have 'undertaken and completed' so forgive the errors in coding
'           styles and methods. Entirely self-taught and a mere hobbyist in VB6.
'
'           The reason I created it was to teach myself VB6, to get back into the 'groove'.
'
'           Back in the 90s I was programming in QB45 (then VB-DOS and VB6) but I left BASIC programming forever
'           and abandoned my main project when VB6 was deprecated. My skills were paltry then and had been picked up from the days
'           of Sinclair Zx80s, 81s and ZX Spectrums.
'
'           My aim now is to resurrect such skills that I once had and improve upon them.
'           A secondary aim is to teach myself how to code in technologies that I have encountered but never fully embraced.
'
'           When I dropped BASIC I picked up Javascript and managed to hone my programming skills to a reasonable hobbyist level
'           but I always missed VB6 and that familiar old IDE. Javascript still has no equivalent decent IDE for what I need it to do.
'           Javascript however, is very much like BASIC in so many ways and a hobbyist BASIC programming style works in Javascript too.
'
'           Returning to VB6 after so many years, it was a big surprise to me to find such inadequate native image type handling,
'           VB6 being unable to handle the various image types without the usage of a great deal of code and
'           API calls. In the process of creating this utlity I learned that VB6 can 'do' anything but it can also be hard work to
'           make it do so. VB.NET makes a lot of these things possible but it also makes programming in general a lot more painful.
'           Either way programming in VB6 or VB.NET is a hard slog.
'
'           When this project is complete my next aim is to migrate it to VB.NET through the versions to find out what
'           problems are typically encountered in a project such as this.
'
'           I could not have made this utility without the help of code from the various projects I have listed below.
'
'           I hope you enjoy the functionality this utility provides. If you think you can improve anything then please
'           feel free to do so. If you dislike my programming style then do keep those thoughts to yourself. :)
'
'           Built on a 2.5ghz core2duo Dell Latitude E5400 running Windows 7 Ultimate 64bit using VB6 SP6.
'
' Credits : LA Volpe (VB Forums) for his transparent picture handling.
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           KillApp code from an unknown, untraceable source, possibly on MSN.
'           Registry reading code from ALLAPI.COM.
'           Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.
'           Active VB Germany for information on the undocumented PrivateExtractIcons API.
'           Elroy on VB forums for his Persistent debug window
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
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
'           Open font dialog code without dependent OCX
'
'   Tested on :
'           Windows 7 Pro 32bit on Intel
'           Windows 7 Ultimate 64bit on Intel
'           Windows XP SP3 on Intel
'           Windows 10 Home 64bit on AMD and Intel
'
' Dependencies:
'           Krool's replacement for the Microsoft Windows Common Controls found in
'           mscomctl.ocx (treeview, slider) are replicated by the addition of two
'           dedicated OCX files that are shipped with this package.
'
'           CCRSlider.ocx
'           CCRTreeView.ocx
'
'           RocketDock 1.3.5 - must be installed before this tool will function.
'
' Notes:
'           Integers are retained (rather than longs) as some of these are passed to
'           library API functions in code that is not my own so I am loathe to change.
'           A lot of the code provided (by better devs than me) seems to have code quality
'           issues - I haven't gone through all their code to fix every problem but I have fixed lots...
'
'           The icons are displayed using Lavolpe's transparent DIB image code,
'           except for the .ico files which use his earlier StdPictureEx class.
'           The original ico code caused many strange visual artifacts and complete failures to show .ico files.
'           especially when other image types were displayed on screen simultaneously.
'
' Summary:
'           The program reads a default icon folder from Rocketdock's settings.ini or registry.
'           It reads the contents of the folder and sub-folders into a treeview and displays the first 12 of the
'           icons using 12 dynamically created picboxes. The icons are displayed using Lavolpe's
'           transparent DIB image code, except for the .ico files which use the earlier StdPictureEx class.
'           DLLs and EXEs with embedded icons are handled using an undocumented API named PrivateExtractIcons.
'           One selected image is extracted and displayed in larger size using the above code in the preview window.
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
'           Rather than create a manifest and bundle the OCX within a .res file for extraction on the utilities'
'           first run, I have created instead my own installer program that attempts to place the required OCX
'           file(s) in the correct location.
'
'           The reason that an SxS configuration is absent is that errors were generated as soon as comctl32 was
'           added to the manifest. I became so bored always trying to fix the many manifest errors that I gave up...
'           A finicky and therefore useless method to packaging up a VB6 app.
'
'           The font selection and file/folder dialogs are generated using Win32 APIs rather than the
'           common dialog OCX which dispensed with another OCX.
'
'           I made an attempt to replace the mscomctl.ocx with an in-built treeview replacement using
'           Win32 APIs but that was a fair bit of work so that task remains unfinished. I have that version
'           put aside and may complete it later. This will free the program of all external dependencies.
'           In the meantime I have used Krool's amazing control replacement project. The specific code for
'           just two of the controls (treeview and slider) has been incorporated rather than all 32 of
'           Krool's complete package.
'
' Missing:
'           The only component not yet functional is the 'generate dock' button. At the moment
'           it only tests the registry for certain entries in the uninstall section of the registry.
'           Eventually, it will generate a settings.ini file containing all the useful software you have in your
'           system.
'
' Licence:
'           Copyright © 2019 Dean Beedell
'
'           This program is free software; you can redistribute it and/or modify it under the terms of the
'           GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
'           License, or (at your option) any later version.
'
'           This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'           even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'           General Public Licence for more details.
'
'           You should have received a copy of the GNU General Public Licence along with this program; if not,
'           write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
'           USA
'
'           If you use this software in any way whatsoever then that implies acceptance of the licence. If you
'           do not wish to comply with the licence terms then please remove the download, binary and source code
'           from your systems immediately.
'
'--------------------------------------------------------------------------------------------------------------

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
'Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As OLE_COLOR, ByVal hPal As Long, ByRef lpColorRef As Long) As Long

Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean

Private cImage As c32bppDIB
Private cShadow As c32bppDIB

' Note: If GDI+ is available, it is more efficient for you to
' create the token then pass the token to each class.  Not required,
' but if you don't do this, then the classes will create and destroy
' a token everytime GDI+ is used to render or modify an image.
' Passing the token can result in up to 3x faster processing overall

Private m_GDItoken As Long
Private FontDlg As CommonDlgs

Private Const COLOR_BTNFACE As Long = 15

'some variables for temporarily storing the old image name
Private previousIcon As String
Private mapImageChanged As Boolean
Private thumbPos0Pressed As Boolean
Private validIconTypes As String  ' change from VB6 to scope due to replacement of filelistbox control to simple listbox






'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : The initial subroutine for the program after the graphics code has done its stuff.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
        
    On Error GoTo Form_Load_Error
    If debugflg = 1 Then DebugPrint "%" & "Form_Load"

    Dim NameProcess As String
    Dim AppExists As Boolean
    'Dim validIconTypes As String
    Dim srdMapState As String

    ReDim thumbArray(12) As Integer
        
    iconChanged = False
    dotCount = 0 ' a variable used on the 'working...' button
    rdIconNumber = 0
    rdIconMax = 0  ' the final icon in the registry/settings
    icoSizePreset = 128
    thumbImageSize = 64
    boxSpacing = 540
    storedIndex = 9999
    busyCounter = 1
    
    mapImageChanged = False
    
    ' theme variables
    classicTheme = False
    storeThemeColour = 13160660  '15790320 = Windows 'modern'  13160660 ' Windows classic

    ' add any remaining types that Rocketdock's code supports
    validIconTypes = "*.jpg;*.jpeg;*.bmp;*.ico;*.png;*.tif;*.gif"
        
    ' state and position of a few manually placed controls (easier here than in the IDE)
    picRdThumbFrame.Visible = False
    
    previewFrameGotFocus = True
                              
    'if the process already exists then kill it
    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "rocket1.exe"
        checkAndKill NameProcess
    End If
               
    ' get the location of this tool's settings file
    Call getSettingsFile
                   
    'check the state of the licence
    Call checkLicenceState
    
    ' check the Windows version and where rocketdock is installed
    Call testWinVer
        
    ' check where rocketdock is installed
    Call checkRocketdockInstallation
    
    ' set the default path to the icons root, this will be superceded later if the user has chosen a default folder
    filesIconList.path = rdAppPath & "\Icons" ' rdAppPath is defined in driveCheck above
    textCurrentFolder.Text = rdAppPath & "\Icons"
    relativePath = "\Icons"
        
    ' read the Rocketdock settings from INI or from registry
    Call readRocketDockSettings
    
    ' dynamically create thumbnail picboxes and sort the captions
    Call createThumbnailLayout
    
    ' dynamically create rocketdock Map thumbnail picboxes
    Call createRdMapBoxes
                
    ' read the tool settings file and do some things for the first and only time
    Call readToolSettings
    
    ' set the very large icon record number displayed on the main form
    rdIconNumber = 0
    lblRdIconNumber.Caption = rdIconNumber + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
            
    ' set the filter pattern to only show the icon types supported by Rocketdock
    filesIconList.Pattern = validIconTypes
    
    ' select that file in the file list
    ' filesIconList.ListIndex = 0 '- commented out as this causes an unnecessary fileiconlist click and a first display of an image

    ' set the preview pane to redraw smoothly
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
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
    
    ' set the theme colour
    Call setThemeColour
    
    ' select the thumbnail view rather than the file list view and populate it
    fileIconListPosition = 0
    Call refreshThumbnailViewPanel
    
    ' we indicate that all changes have been lost when changes to fields are made by the program and not the user
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"
    
    srdMapState = GetINISetting("Software\RocketDockSettings", "rdMapState", toolSettingsFile)
    If srdMapState <> "hidden" Then
        Call backupSettings("")
    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form rDIconConfigForm"
                
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
    
    Dim SysClr As Long
    
    On Error GoTo setThemeColour_Error
    If debugflg = 1 Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeDark
        SysClr = GetSysColor(COLOR_BTNFACE)
    Else
        'MsgBox "Windows Alternate Theme detected"
        SysClr = GetSysColor(COLOR_BTNFACE)
        If SysClr = 13160660 Then
            Call setThemeDark
        Else ' 15790320
            Call setThemeLight
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
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
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
   Dim ans As VbMsgBoxResult
   Dim itemno As Integer
    
   On Error GoTo btnAdd_Click_Error
   If debugflg = 1 Then DebugPrint "%" & "btnAdd_Click"
    
   If storedIndex <> 9999 Then
       If chkConfirmSaves.Value = 1 Then
            itemno = thumbArray(storedIndex)
    
            ans = MsgBox(" Confirm that you wish to set icon " & filesIconList.List(itemno) & " as the current icon " & vbCr & "in the dock.", vbYesNo)
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
    Dim itemno As Integer
   On Error GoTo changeMapImage_Error
   If debugflg = 1 Then Debug.Print "%changeMapImage"

    mapImageChanged = True
    itemno = thumbArray(storedIndex)
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
Private Sub btnGenerate_Click()
    Dim ans As VbMsgBoxResult
    Dim xFileName As String
    
    'Call btnArrowDown_Click ' populate the dock
    On Error GoTo btnGenerate_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnGenerate_Click"
    

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
        
        
        'formSoftwareList.rtbSoftwareList.Text = s
        formSoftwareList.Show
        
        'write the data to a local file as well
        Call WriteOutputFile(s, xFileName)
        
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
        '           a word like photo &frameProperties that could use similar icons
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

'---------------------------------------------------------------------------------------
' Procedure : btnNext_KeyDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnNext_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub btnPrev_KeyDown(KeyCode As Integer, Shift As Integer)
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
   Dim a As String
   
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
            MsgBox "Cannot remove Rocketdock Enhanced Settings Utility sub-folders from the treeview."
            Exit Sub
    End If
        
    If a = "icons" Then
        MsgBox "Cannot remove Rocketdock's own sub-folders from the treeview, you have to delete the folders from Windows first then re-run this utility."
        Exit Sub
    End If
        
    If folderTreeView.SelectedItem.Key = vbNullString Then
        Exit Sub
    End If
        
    If a = "custom folder" And Not folderTreeView.SelectedItem = "custom folder" Then
        MsgBox "Cannot remove custom sub-folders from the treeview, try again at the root."
        Exit Sub
    Else
        ' do the delete!
    End If
        
    folderTreeView.Nodes.Remove folderTreeView.SelectedItem.Key
    
    'write the folder to the rocketdock settings file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
    PutINISetting "Software\RocketDock", "CustomIconFolder", "?", rdSettingsFile

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
    
    Dim useloop As Integer
    
    On Error GoTo createThumbnailLayout_Error
    If debugflg = 1 Then DebugPrint "%" & "createThumbnailLayout"

    storeLeft = 165
    frmThumbLabel(0).ZOrder
    frmThumbLabel(0).BorderStyle = 0
    frmThumbLabel(0).Visible = True
         
    ' dynamically create the picture boxes for the thumbnails
    For useloop = 1 To 11 ' 0 is the template
        Load picThumbIcon(useloop)
        Load frmThumbLabel(useloop)
        Load lblThumbName(useloop)
        
        Set lblThumbName(useloop).Container = frmThumbLabel(useloop)
    Next useloop
    
    Call placeThumbnailPicboxes(64)
    
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
   If debugflg = 1 Then DebugPrint "%" & "createRdMapBoxes"

    storeLeft = boxSpacing
    ' dynamically create more picture boxes to the maximum number of icons
    For useloop = 1 To rdIconMax
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
' Procedure : readToolSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read this utilties' own settings.ini file and do some things using the data
'---------------------------------------------------------------------------------------
'
Private Sub readToolSettings()
    Dim sfirst As String
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedStrength As String
    Dim suppliedStyle As String

   On Error GoTo readToolSettings_Error
   If debugflg = 1 Then DebugPrint "%" & "readToolSettings"
   
    If Not FExists(toolSettingsFile) Then Exit Sub ' does the tool's own settings.ini exist?
    
    'test to see if the tool has ever been run before
    sfirst = GetINISetting("Software\RocketDockSettings", "First", toolSettingsFile)
    If sfirst = True Then
    
        sfirst = False
        
        ' insert at the final position
        ' a link with the rocketdockSettings icon and the target
        ' is the app.path
        
        sFilename = "Icons\rocketdockSettings.png" ' the default Rocketdock filename for a blank item
        sTitle = "Rocket Settings"
        sCommand = App.path & "\" & "rocket1.exe"
        sArguments = vbNullString
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
    
        'filecopy the rocketdockSettings png to the rocketdock icons folder
        If FExists(App.path & "\" & "rocketdockSettings.png") Then
            FileCopy App.path & "\" & "rocketdockSettings.png", rdAppPath & "\icons\" & "rocketdockSettings.png"
        End If
        
    End If

    If IsUserAnAdmin() = 0 And requiresAdmin = True Then
        MsgBox "This tool requires to be run as administrator on Windows 8 and above in order to function. Admin access is NOT required on Win7 and below. If you aren't entirely happy with that then you'll need to remove the software now. This is a limitation imposed by Windows itself. To enable administrator access find this tool's exe and right-click properties, compatibility - run as administrator. YOU have to do this manually, I can't do it for you."
    End If
    
    ' set the tool's default font
    suppliedFont = GetINISetting("Software\RocketDockSettings", "defaultFont", toolSettingsFile)
    suppliedSize = GetINISetting("Software\RocketDockSettings", "defaultSize", toolSettingsFile)
    suppliedStrength = GetINISetting("Software\RocketDockSettings", "defaultStrength", toolSettingsFile)
    suppliedStyle = GetINISetting("Software\RocketDockSettings", "defaultStyle", toolSettingsFile)
    
    If Not suppliedFont = "" Then
        Call changeFont(suppliedFont, suppliedSize, suppliedStrength, suppliedStyle)
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
   If debugflg = 1 Then DebugPrint "%" & "readRocketDockSettings"

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    rdSettingsFile = App.path & "\rdSettings.ini" ' a copy of the settings file that we work on
        
    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
        
        Call backupSettings("") ' make a backup of the settings.ini file each restart
        
        ' copy the original settings file to a duplicate that we will operate upon
        FileCopy origSettingsFile, rdSettingsFile
        
        ' read the rocketdock settings.ini and find the very last icon
        theCount = GetINISetting("Software\RocketDock\Icons", "count", rdSettingsFile)
        rdIconMax = theCount - 1
    Else
        chkRegistry.Value = 1
        chkSettings.Value = 0
        
        ' read the rocketdock registry and find the last icon
        theCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count")
        rdIconMax = theCount - 1
        
        ' copy the original configs out of the registry and into a settings file that we will operate upon
        readRegistryWriteSettings
        
        ' make a backup of the rdSettings.ini after the intermediate file has been created
        Call backupSettings("")
        
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
    Dim RDinstalled As Boolean
    Dim RD86installed As Boolean
    Dim chkFolder As String
    
    ' check where rocketdock is installed
    On Error GoTo checkRocketdockInstallation_Error
    If debugflg = 1 Then DebugPrint "%" & "checkRocketdockInstallation"
        
    RD86installed = driveCheck("Program Files (x86)\Rocketdock")
    RDinstalled = driveCheck("Program Files\Rocketdock")
    
    If RDinstalled = True Then mnuRocketDock.Caption = "Rocketdock location - program files - click to change"
    If RD86installed = True Then mnuRocketDock.Caption = "Rocketdock location - program files (x86) - click to change"
    

    ' read the tool settings file
    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        chkFolder = GetINISetting("Software\RocketDockSettings", "rocketDockLocation", toolSettingsFile)
        If chkFolder <> vbNullString Then
            If FExists(chkFolder & "\rocketDock.exe") Then
                rdAppPath = chkFolder
                mnuRocketDock.Caption = "Rocketdock location - " & chkFolder & " - click to change"
            End If
        End If
    End If
    
    ' get the value of the rocketdock folder location
    ' check the value exists
    ' check the location exists
    ' if the location exists do not search
    
    If rdAppPath = vbNullString Then
        answer = MsgBox(" Rocketdock has not been installed in the program files (x86) folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
         Dim ofrm As Form
         For Each ofrm In Forms
             Unload ofrm
         Next
         End
    End If
    
   On Error GoTo 0
   Exit Sub

checkRocketdockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : getSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
Private Sub getSettingsFile()
    Dim toolSettingsDir As String

    On Error GoTo getSettingsFile_Error
    If debugflg = 1 Then Debug.Print "%getSettingsFile"
    
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
        FileCopy App.path & "\settings.ini", toolSettingsFile
    End If
    
    'confirm the settings file exists, if not use the version in the app itself
    If Not FExists(toolSettingsFile) Then
        toolSettingsFile = App.path & "\settings.ini"
    End If
    
   On Error GoTo 0
   Exit Sub

getSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getSettingsFile of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : 'check the state of the licence
'---------------------------------------------------------------------------------------
'
Private Sub checkLicenceState()
    Dim slicence As Integer

    On Error GoTo checkLicenceState_Error
    If debugflg = 1 Then DebugPrint "%" & "checkLicenceState"

    ' read the tool's own settings file (
    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = GetINISetting("Software\RocketDockSettings", "Licence", toolSettingsFile)
        ' if the licence state is not already accepted then display the licence form
        If slicence = 0 Then
            Call LoadFileToTB(licence.txtLicenceTextBox, App.path & "\licence.txt", False)
            
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


'---------------------------------------------------------------------------------------
' Procedure : placeThumbnailPicboxes
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : ' place the thumbnails picboxes where they should go
'---------------------------------------------------------------------------------------
'
Private Sub placeThumbnailPicboxes(ByVal imageSize As Integer)
    Dim useloop As Integer
    Dim storeTop As Integer
        
   
    On Error GoTo placeThumbnailPicboxes_Error
    If debugflg = 1 Then DebugPrint "%" & "placeThumbnailPicboxes"

'    picThumbIcon(0).Width = 1000
'    picThumbIcon(0).Height = 1000
'    picThumbIcon(0).Left = 165
'    picThumbIcon(0).Top = 60
    
    For useloop = 0 To 11
        
        picThumbIcon(useloop).Width = 1000
        picThumbIcon(useloop).Height = 1000
        frmThumbLabel(useloop).BorderStyle = 0
        
        picThumbIcon(useloop).ToolTipText = filesIconList.List(useloop)

        If useloop = 0 Then
            If imageSize = 32 Then
                storeLeft = 165
                storeTop = -200
            Else
                storeLeft = 165
                storeTop = 30
            End If
        End If

        If useloop = 4 Then
            If imageSize = 32 Then
                storeLeft = 165
                storeTop = 880
            Else
                storeLeft = 165
                storeTop = 1060
            End If
        End If

        If useloop = 8 Then
            If imageSize = 32 Then
                storeLeft = 165
                storeTop = 1970
            Else
                storeLeft = 165
                storeTop = 2100
            End If
        End If
        
        picThumbIcon(useloop).Left = storeLeft
        picThumbIcon(useloop).Top = storeTop
        
        frmThumbLabel(useloop).Left = storeLeft - 100
        frmThumbLabel(useloop).Top = storeTop + 800
       
        storeLeft = storeLeft + 1200

        picThumbIcon(useloop).Visible = True
        lblThumbName(useloop).Visible = True
        frmThumbLabel(useloop).Visible = True
        
        picThumbIcon(useloop).ZOrder
        frmThumbLabel(useloop).ZOrder
        
        picThumbIcon(useloop).AutoRedraw = True
    Next useloop
    

   On Error GoTo 0
   Exit Sub

placeThumbnailPicboxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure placeThumbnailPicboxes of Form rDIconConfigForm"

End Sub
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
    Dim ans As VbMsgBoxResult
    Dim iconPath As String
    Dim dllPath As String
    Dim dialogInitDir As String
    Dim bkpSettingsFile As String
    Dim bkpFilename As String
    
    Const x_MaxBuffer = 256
    On Error GoTo btnBackup_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnBackup_Click"

    Call backupSettings(bkpFilename)
    ans = MsgBox("Created an incremental backup of the Rocketdock settings file - " & vbCr & vbCr & bkpFilename & vbCr & vbCr & "Would you like to review ALL the backup files? ", vbYesNo)
    If ans = 6 Then

        On Error Resume Next

        ' set the default folder to the existing reference
        If DirExists(App.path & "\backup") Then
            ' set the default folder to the existing reference
            dialogInitDir = App.path & "\backup" 'start dir, might be "C:\" or so also
        Else
            MsgBox "Backup folder " & App.path & "\backup" & " has been removed. Backup cancelled"
            Exit Sub
        End If

        With x_OpenFilename
        '    .hwndOwner = Me.hWnd
        .hInstance = App.hInstance
        .lpstrTitle = "Select a backup INI file to restore - or cancel"
        .lpstrInitialDir = dialogInitDir

        .lpstrFilter = "Ini Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        .nFilterIndex = 2

        .lpstrFile = String(x_MaxBuffer, 0)
        .nMaxFile = x_MaxBuffer - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = x_MaxBuffer - 1
        .lStructSize = Len(x_OpenFilename)
        End With
        
        Dim retFileName As String
        Dim retfileTitle As String
        Call f_GetOpenFileName(retFileName, retfileTitle)
        bkpSettingsFile = retFileName
        
        If Not bkpSettingsFile = "" Then
        
            ans = MsgBox("Do you wish to restore this file?  " & bkpSettingsFile & "? ", vbYesNo)
            If ans = 6 Then
                ' take the backup file and copy it into the app's folder
                ' refresh the map using the restored setings.ini file
                ' restart rocketdock
                FileCopy bkpSettingsFile, rdSettingsFile
                
                Call btnSaveRestart_Click
            End If
        End If

        'ShellExecute 0, vbNullString, App.path & "\backup", vbNullString, vbNullString, 1
    End If
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

    Call ShellExecute(Me.hWnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981272/orbs-and-icons", vbNullString, App.path, 1)

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
    Dim answer As VbMsgBoxResult
    On Error GoTo btnKillIcon_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnKillIcon_Click"


        If textCurrIconPath.Text = vbNullString Then
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
        
        ' using the current filelist as the start point on the list, repopulate the thumbs
        Call populateThumbnails(thumbImageSize, fileIconListPosition)

        If filesIconList.Visible = True Then
            filesIconList.SetFocus         ' return focus to the form
        Else
            picFrameThumbs.SetFocus        ' return focus to the form
        End If
        
        ' now display the current icon, the previous icon displayed now deleted
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)

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
   If debugflg = 1 Then DebugPrint "%" & "btnSave_Click"


    sFilename = txtCurrentIcon.Text
    
    sTitle = lblName.Text
    If sDockletFile = "" Then
        sCommand = txtTarget.Text
    Else
        sDockletFile = txtTarget.Text
    End If
    sArguments = txtArguments.Text
    sWorkingDirectory = txtStartIn.Text
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
    
    'if the current icon has changed by a dblclick on the file list then refresh that part of the rdMap
    If iconChanged = True Then
        'only if the rdMAp has already been displayed already do we carry out the image refresh
        If Not picRdMap(0).ToolTipText = vbNullString Then ' check that the array has been populated already
            ' we just reload the sole picbox that has changed
            Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32, True)
        End If
        iconChanged = False
    End If
    
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"
    
    If triggerRdMapRefresh = True Then
        'Call rdMapRefresh_Click
        'Call busyStart
        Call populateRdMap(0) ' show the map from position zero
        'Call busyStop

        ' we signify that there have been no changes - this is just a refresh
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
        triggerRdMapRefresh = False
    End If

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
   If debugflg = 1 Then DebugPrint "%" & "btnRefresh_Click"

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
            picThumbIcon(thumbIndexNo).BorderStyle = 1
        ElseIf thumbImageSize = 32 Then
            lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200)
        End If
    
        Call busyStop

   On Error GoTo 0
   Exit Sub

btnRefresh_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRefresh_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetDirectory
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Private Function GetDirectory(ByRef path As String) As String
   On Error GoTo GetDirectory_Error
   If debugflg = 1 Then DebugPrint "%" & "GetDirectory"

   GetDirectory = Left(path, InStrRev(path, "\") - 1)

   On Error GoTo 0
   Exit Function

GetDirectory_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDirectory of Form rDIconConfigForm"
End Function
'---------------------------------------------------------------------------------------
' Procedure : btnTarget_Click
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Private Sub btnTarget_Click()
    Dim iconPath As String
    Dim dllPath As String
    Dim dialogInitDir As String

    Const x_MaxBuffer = 256
    
    'On Error GoTo btnTarget_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnTarget_Click"
    
    'On Error GoTo l_err1
    'savLblTarget = txtTarget.Text
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not txtTarget.Text = vbNullString Then
        If FExists(txtTarget.Text) Then
            ' extract the folder name from the string
            iconPath = GetDirectory(txtTarget.Text)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(txtTarget.Text) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = txtTarget.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If
    
    If Not sDockletFile = vbNullString Then
        If FExists(sDockletFile) Then
            ' extract the folder name from the string
            dllPath = GetDirectory(sDockletFile)
            ' set the default folder to the existing reference
            dialogInitDir = dllPath 'start dir, might be "C:\" or so also
        ElseIf DirExists(sDockletFile) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = sDockletFile 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath & "\docklets" 'start dir, might be "C:\" or so also
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target for this icon to call"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With
  
  Dim retFileName As String
  Dim retfileTitle As String
  Call f_GetOpenFileName(retFileName, retfileTitle)
  If retFileName <> vbNullString Then
    txtTarget.Text = retFileName
    'fill in the file title and the start in automatically if they are empty and need filling
    If lblName.Text = vbNullString Then lblName.Text = retfileTitle
    If txtStartIn.Text = vbNullString Then txtStartIn.Text = GetDirectory(txtTarget.Text)
  End If
  
' this is the code left over from the use of the common dialog OCX, left here for reference
'
'    rdDialogForm.CommonDialog.DialogTitle = "Select a File" 'titlebar
'    If Not txtTarget.Text = vbNullString Then
'        If FExists(txtTarget.Text) Then
'            ' extract the folder name from the string
'            iconPath = GetDirectory(txtTarget.Text)
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
'            If lblName.Text = vbNullString Then
'                lblName.Text = rdDialogForm.CommonDialog.FileTitle
'            End If
'            txtTarget.Text = rdDialogForm.CommonDialog.FileName
'        End If
'    End If

   On Error GoTo 0
   
   Exit Sub

btnTarget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTarget_Click of Form rDIconConfigForm"
 
End Sub
'---------------------------------------------------------------------------------------
' Procedure : f_GetOpenFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub f_GetOpenFileName(retFileName As String, retfileTitle As String)
   On Error GoTo f_GetOpenFileName_Error
   If debugflg = 1 Then DebugPrint "%f_GetOpenFileName"

  If GetOpenFileName(x_OpenFilename) <> 0 Then
    If x_OpenFilename.lpstrFile = "*.*" Then
        'txtTarget.Text = savLblTarget
    Else
        retfileTitle = x_OpenFilename.lpstrFileTitle
        retFileName = x_OpenFilename.lpstrFile
    End If
  Else
    'The CANCEL button was pressed
    'MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

f_GetOpenFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure f_GetOpenFileName of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : f_GetSaveFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub f_GetSaveFileName()
   On Error GoTo f_GetSaveFileName_Error
   If debugflg = 1 Then DebugPrint "%f_GetSaveFileName"

  If GetSaveFileName(x_OpenFilename) <> 0 Then
    'PURPOSE: A file was selected
    MsgBox Left$(x_OpenFilename.lpstrFile, x_OpenFilename.nMaxFile)
  Else
    'PURPOSE: The CANCEL button was pressed
    MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

f_GetSaveFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure f_GetSaveFileName of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblToggleDialogs_Click
' Author    : beededea
' Date      : 12/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblToggleDialogs_Click()
   On Error GoTo lblToggleDialogs_Click_Error
   If debugflg = 1 Then Debug.Print "%lblToggleDialogs_Click"

        If chkConfirmSaves.Value = 0 Then
            chkConfirmSaves.Value = 1
        Else
            chkConfirmSaves.Value = 0
        End If

   On Error GoTo 0
   Exit Sub

lblToggleDialogs_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblToggleDialogs_Click of Form rDIconConfigForm"
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
   If debugflg = 1 Then Debug.Print "%mnuAddPreviewIcon_Click"

    'Debug.Print picPreview.Tag
    
    Call btnAdd_Click

   On Error GoTo 0
   Exit Sub

mnuAddPreviewIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPreviewIcon_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRocketDock_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   : Menu option to direct where Rocketdock may be found
'---------------------------------------------------------------------------------------
'
Private Sub mnuRocketDock_click()

    Dim getFolder As String
    Dim dialogInitDir As String
   
   On Error GoTo mnuRocketDock_click_Error
   If debugflg = 1 Then DebugPrint "%mnuRocketDock_click"

    dialogInitDir = "C:\" 'start dir, might be "C:\" or so also

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder
    If getFolder <> vbNullString Then
        If FExists(getFolder & "\rocketdock.exe") Then
            rdAppPath = getFolder & "\rocketdock.exe"
            If DirExists(getFolder) Then mnuRocketDock.Caption = "RocketDock Location - " & getFolder & " - click to change."
            
            If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                PutINISetting "Software\RocketDockSettings", "rocketDockLocation", rdAppPath, toolSettingsFile
            End If
            
        End If
    End If

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
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddSeparator_click_Error
    If debugflg = 1 Then Debug.Print "mnuAddSeparator_click"
           
    iconFileName = App.path & "\my collection" & "\separator.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If

    sIsSeparator = "1"
        
    ' general tool to add an icon
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Separator", vbNullString, vbNullString, vbNullString, vbNullString, sIsSeparator)
        
    lblName.Enabled = False
    txtCurrentIcon.Enabled = False
    txtTarget.Enabled = False
    btnTarget.Enabled = False
    txtArguments.Enabled = False
    txtStartIn.Enabled = False
    comboRun.Enabled = False
    comboOpenRunning.Enabled = False
    checkPopupMenu.Enabled = False
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
    Dim iconImage As String
    Dim iconFileName As String
    
    Dim getFolder As String
    Dim dialogInitDir As String
   
   On Error GoTo mnuaddFolder_click_Error
   If debugflg = 1 Then Debug.Print "%mnuaddFolder_click"

    If txtStartIn.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder

    If DirExists(getFolder) Then
    
        iconFileName = App.path & "\my collection" & "\folder-closed.png"
        If FExists(iconFileName) Then
            iconImage = iconFileName
        Else
            iconImage = "\Icons\help.png"
        End If
        
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call menuAddSummat(iconImage, "User Folder", getFolder, vbNullString, vbNullString, vbNullString, vbNullString)
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
    Dim iconImage As String
    Dim iconFileName As String
    
    ' check the icon exists
   On Error GoTo mnuAddMyComputer_click_Error
   If debugflg = 1 Then DebugPrint "%mnuAddMyComputer_click"

    iconFileName = App.path & "\my collection" & "\my folder.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "My Computer", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)


   On Error GoTo 0
   Exit Sub

mnuAddMyComputer_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyComputer_click of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddEnhanced_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   : Menu option to add an enhanced settings utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddEnhanced_click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddEnhanced_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddEnhanced_click"

    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\rocketdockSettings.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Enhanced Icon Settings", App.path & "\rocket1.exe", vbNullString, vbNullString, vbNullString, vbNullString)

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
Private Sub mnuAddDocklet_click()
   
    Dim dllPath As String
    Dim dialogInitDir As String

    Const x_MaxBuffer = 256
    
    On Error GoTo mnuAddDocklet_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddDocklet_click"
    
    ' set the default folder to the docklet folder under rocketdock
    dialogInitDir = rdAppPath & "\docklets"
 
    With x_OpenFilename
    '    .hwndOwner = Me.hWnd
      .hInstance = App.hInstance
      .lpstrTitle = "Select a Rocketdock Docklet DLL"
      .lpstrInitialDir = dialogInitDir
      
      .lpstrFilter = "DLL Files" & vbNullChar & "*.dll" & vbNullChar & vbNullChar
      .nFilterIndex = 2
      
      .lpstrFile = String(x_MaxBuffer, 0)
      .nMaxFile = x_MaxBuffer - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = x_MaxBuffer - 1
      .lStructSize = Len(x_OpenFilename)
    End With
          
    Dim retFileName As String
    Dim retfileTitle As String
    Call f_GetOpenFileName(retFileName, retfileTitle)
    txtTarget.Text = retFileName
    'lblName.Text = retfileTitle
      
  If txtTarget.Text <> "" Then
    ' check the folder is valid docklet folder (beneath the docklets folder)
    ' set it to the docklet image yet to be created
    ' if it is a clock docklet use a temporary clock image just as RD does without hands?
    ' if it is a weather docklet use a temporary weather image of my own making
    ' if it is a recycling docklet use a temporary recycling image of my own making
    
    ' set the icon to that used by the docklet, it a mere guess as we cannot read the docklet DLL at this stage
    ' to determine what icon image it intends to use, later it writes to the 'other' settings.ini file in docklets
    ' but that's of no use now.
    
      If InStr(GetFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
        txtCurrentIcon.Text = rdAppPath & "\icons\clock.png"
      ElseIf InStr(GetFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
        txtCurrentIcon.Text = App.path & "\my collection\recyclebin-full.png"
      Else
        txtCurrentIcon.Text = rdAppPath & "\icons\blank.png" ' has to be an icon of some sort
      End If
      
       '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
      Call menuAddSummat(txtCurrentIcon.Text, "Docklet", vbNullString, vbNullString, vbNullString, txtTarget.Text, vbNullString)
    
    ' disable the fields, only enable the target fields and use the target field as a temporary location for the docklet data
      
      lblName.Enabled = False
      txtCurrentIcon.Enabled = False
      
      sDockletFile = txtTarget.Text
      txtTarget.Enabled = True
      btnTarget.Enabled = True
      
      txtArguments.Enabled = False
      txtStartIn.Enabled = False
      comboRun.Enabled = False
      comboOpenRunning.Enabled = False
      checkPopupMenu.Enabled = False
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

    End If
    Call busyStop
   On Error GoTo 0
   Exit Sub

btnFileListView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFileListView_Click of Form rDIconConfigForm"
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
      If debugflg = 1 Then DebugPrint "%" & "checkPopupMenu_Click"
   

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
      If debugflg = 1 Then DebugPrint "%" & "chkRegistry_Click"

   

    chkTheRegistry

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
      If debugflg = 1 Then DebugPrint "%" & "chkSettings_Click"
   
   

    chkTheRegistry

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

    Dim filterType As Integer
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
    btnRefresh_Click
    
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
' Procedure : comboOpenRunning_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub comboOpenRunning_Click()
   On Error GoTo comboOpenRunning_Click_Error
   If debugflg = 1 Then DebugPrint "%comboOpenRunning_Click"

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

comboOpenRunning_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comboOpenRunning_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : comboRun_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub comboRun_Click()
   On Error GoTo comboRun_Click_Error
   If debugflg = 1 Then DebugPrint "%comboRun_Click"

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

comboRun_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comboRun_Click of Form rDIconConfigForm"
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
    If debugflg = 1 Then DebugPrint "%" & "btnCloseCancel_Click"

    If btnCloseCancel.Caption = "Cancel" Then
        If mapImageChanged = True Then
                ' now change the icon image back again
                ' the target picture control and the icon size
                Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
                mapImageChanged = False
        End If
        
        'if it is a good icon number then read the data
        If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
            'get the rocketdock settings.ini for this icon alone
            readSettingsIni (rdIconNumber)
        Else
            readRegistryOnce (rdIconNumber)
        End If
        
        ' if the incoming text has <quote> then replace those with a "  TODO
        txtCurrentIcon.Text = sFilename ' build the full path
        
        lblName.Text = sTitle
        txtTarget.Text = sCommand
        txtArguments.Text = sArguments
        txtStartIn.Text = sWorkingDirectory
        comboRun.ListIndex = sShowCmd
        comboOpenRunning.ListIndex = sOpenRunning
        checkPopupMenu.Value = sUseContext
        
        
        ' display the icon from the alternative settings.ini config.
        FileName = txtCurrentIcon.Text
        
        Call displayResizedImage(FileName, picPreview, icoSizePreset)
        
        ' we signify that all changes have been lost
        iconChanged = False
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
    Else
        Form_Unload 0
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
    If debugflg = 1 Then DebugPrint "%" & "btnAddFolder_Click"
   
    If FExists(rdSettingsFile) Then
        CustomIconFolder = GetINISetting("Software\RocketDock", "CustomIconFolder", rdSettingsFile)
    End If
    
    If CustomIconFolder = "?" Then
        ' currently do nothing here
    Else

        If folderTreeView.SelectedItem.Text = "my collection" Or folderTreeView.SelectedItem.Text = "icons" Then
            ' do nothing
        Else
            ' if the customfolder has been set then remove it first from the .ini
            ' and remove it from the tree
            ' remove the ?
            
            If Left(CustomIconFolder, 1) = "?" Then
                folderTreeView.SelectedItem.Key = Mid(CustomIconFolder, 2)
            Else
                folderTreeView.SelectedItem.Key = CustomIconFolder
            End If
            Call btnRemoveFolder_Click
        End If
    End If
    
    savTextCurrentFolder = textCurrentFolder.Text 'save the current default folder
    
    Dim dialogInitDir As String
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

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder

    'getFolder = ChooseDir_Click ' show the dialog box to select a folder
    If getFolder = vbNullString Then
        'textCurrentFolder.Text = savTextCurrentFolder
        Exit Sub
    End If
    If getFolder <> vbNullString Then
        textCurrentFolder.Text = getFolder
    End If
    

    Call busyStart
    
    ' add the chosen folder to the treeview
    folderTreeView.Nodes.Add , , textCurrentFolder.Text, textCurrentFolder.Text
    Call addtotree(textCurrentFolder.Text, folderTreeView)
    folderTreeView.Nodes.Item(textCurrentFolder.Text).Text = "custom folder"
    
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
    Dim ans As Boolean
    Dim answer As VbMsgBoxResult
     
    
    ' save the current fields to the settings file or registry
    On Error GoTo btnSaveRestart_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnSaveRestart_Click"
   
   

    origSettingsFile = rdAppPath & "\settings.ini"
    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
                
        ' write the rocketdock settings.ini
        writeSettingsIni (rdIconNumber) ' the settings.ini only exists when RD is set to use it
        
        ' kill the rocketdock process
        NameProcess = "rocketdock.exe"
        ans = checkAndKill(NameProcess)
        
        ' if the rocketdock process has died then
        If ans = True Then
            ' copy the duplicate settings file to the original
            FileCopy rdSettingsFile, origSettingsFile
            
            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.path, 1)
        Else
            answer = MsgBox("Could not find a Rocketdock process, would you like me to restart rocketdock?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If

            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.path, 1)
        
        End If
    Else
         ' kill the rocketdock process
        NameProcess = "rocketdock.exe"
        ans = checkAndKill(NameProcess)
                   
        chkRegistry.Value = 1
        chkSettings.Value = 0
        
        ' if the rocketdock process has died then
        If ans = True Then
           For useloop = 0 To rdIconMax
                ' read the rocketdock alternative settings.ini
                readSettingsIni (useloop) ' the alternative settings.ini exists when RD is set to use it
            
                ' write the rocketdock registry
                writeRegistryOnce (useloop)
            Next useloop
            '0-IsSeparator
            'now write the count to the registry
            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count", Str$(theCount))
            
            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.path, 1)
        Else
            answer = MsgBox("Could not find a Rocketdock process, would you like me to restart rocketdock?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If

            ' restart Rocketdock
            Call ShellExecute(hWnd, "Open", rdAppPath & "\" & NameProcess, vbNullString, App.path, 1)
        
        End If
        
    End If

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
    Dim getFolder As String
    Dim dialogInitDir As String
   
    On Error GoTo btnSelectStart_Click_Error
    If debugflg = 1 Then DebugPrint "%btnSelectStart_Click"
    If txtStartIn.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtStartIn.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder
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

    rdHelpForm.Show

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form rDIconConfigForm"
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
    'Dim FileName As String
    Dim useloop As Integer
        
   On Error GoTo btnNext_Click_Error
      If debugflg = 1 Then DebugPrint "%" & "btnNext_Click"
    
    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
        If mapImageChanged = True Then
            ' now change the icon image
            ' the target picture control and the icon size
            Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
            mapImageChanged = False
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
    If Not picRdMap(0).ToolTipText = vbNullString Then
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
    
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
    
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
' Purpose   : read the registry and set obtain the necessary icon data for the specific icon
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryOnce(ByVal iconNumberToRead As Integer)
    ' read the settings from the registry
   On Error GoTo readRegistryOnce_Error
   If debugflg = 1 Then DebugPrint "%" & "readRegistryOnce"
  
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
Private Sub writeRegistryOnce(ByVal iconNumberToWrite As Integer)
        
   On Error GoTo writeRegistryOnce_Error
    If debugflg = 1 Then DebugPrint "%" & "writeRegistryOnce"
   
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
    Dim Max As Integer
    
    On Error GoTo ExtractSuffix_Error
    If debugflg = 1 Then DebugPrint "%" & "ExtractSuffix"
   
    If strPath = "" Then
        ExtractSuffix = ""
        Exit Function
    End If
    
    AY = Split(strPath, ".")
    Max = UBound(AY)
    ExtractSuffix = AY(Max)

   On Error GoTo 0
   Exit Function

ExtractSuffix_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffix of Form rDIconConfigForm"
End Function
'---------------------------------------------------------------------------------------
' Procedure : displayResizedImage was previously displayPreviewImage
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Displays and places resized image onto the specified picture box
'             Uses two methods to display native and non-native image file types
'             both methods are supplied by LaVolpe
'---------------------------------------------------------------------------------------
'
Private Sub displayResizedImage(ByRef FileName As String, ByRef targetPicBox As PictureBox, ByRef IconSize As Integer)
    Dim suffix As String
    Dim picWidth As Long
    Dim picHeight As Long
                
    On Error GoTo displayResizedImage_Error
    If debugflg = 1 Then DebugPrint "%" & "displayResizedImage"
    
    If Not FExists(FileName) Then Exit Sub    ' just a final check that the chosen image file actually exists
    
    textCurrIconPath.Text = filesIconList.FileName
    ' find and store the indexed position of the chosen file into the global variable
    fileIconListPosition = filesIconList.ListIndex
    
    ' dispose of the image prior to use
    Set targetPicBox.Picture = Nothing ' added because the two methods of drawing an image conflict leaving an image behind
    
    suffix = ExtractSuffix(FileName)
    
    ' using Lavolpe's later method as it allows for resizing of PNGs and all other types
    If InStr("png,jpg,bmp,jpeg,tif,gif", LCase(suffix)) <> 0 Then
        If targetPicBox.Name = "picPreview" Then
            targetPicBox.Left = 345
            targetPicBox.Top = 210
            targetPicBox.Width = 3450
            targetPicBox.Height = 3450
        End If
        
        Set cImage = New c32bppDIB
        cImage.LoadPictureFile FileName, IconSize, IconSize, False, 32
        Call refreshPicBox(targetPicBox, IconSize)
    ElseIf InStr("ico", LCase(suffix)) <> 0 Then
        ' *.ico
        ' using Lavolpe's earlier StdPictureEx method as it allows for correct display of ICOs
        ' the later method above has a bug with some ICOs
        
        'because the earlier method draws the ico images from the top left of the
        'pictureBox we have to manually set the picbox to size and position for each icon size
        Call centrePreviewImage(targetPicBox, IconSize)
        Set targetPicBox.Picture = StdPictureEx.LoadPicture(FileName, lpsCustom, , IconSize, IconSize)
    End If


    ' display the sizes from the image types that are native to VB6
    
    ' check the size of the image and display it,
    ' unlike the .NET version, the sizing has to be done after the display of the image
    ' as it is LaVolpe's code that does the extraction of the icon count.
    'FileName = "C:\Program Files\Rocketdock\"
    If InStr("jpg,bmp,jpeg,gif", LCase(suffix)) <> 0 Then
        Call checkImageSize(FileName, picWidth, picHeight) 'check the size of the image
        lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
    ElseIf InStr("ico", LCase(suffix)) <> 0 Then
        ' captureIconCount is obtained elsewhere in Lavolpe's StdPictureEx code
        If captureIconCount = 1 Then

            On Error GoTo handleResizing_Error
            Call checkImageSize(FileName, picWidth, picHeight) 'check the size of the image
            GoTo displaySizes ' don't want to use a goto but error handling in VB6...
            
handleResizing_Error:
                
             ' if the ico file is damaged then display a blank icon
             ' an example of damage is an icon with an incorrect header count of thousands
             ' Note: Lavolpe's code will still display icons that are considered damaged by Windows.
             
             targetPicBox.ToolTipText = "This icon is damaged - " & FileName
             If FExists(App.path() & "\my collection\" & "red-X.png") Then
                 FileName = App.path() & "\my collection\" & "red-X.png"
             Else
                 FileName = rdAppPath & "\icons\" & "help.png"
             End If
             
             'display image here after the error is handled
             Set targetPicBox.Picture = Nothing

             Set cImage = New c32bppDIB
             cImage.LoadPictureFile FileName, IconSize, IconSize, False, 32
             Call refreshPicBox(targetPicBox, IconSize)

             lblWidthHeight.Caption = " This is a damaged icon." ' < must go here.

             Exit Sub

        End If
        
displaySizes:

        If InStr("ico", LCase(suffix)) <> 0 Then
            If captureIconCount > 1 Then
                lblWidthHeight.Caption = " multiple size (" & captureIconCount & ") ICO file"
            Else
                lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
            End If
        ElseIf InStr("TIFF", LCase(suffix)) <> 0 Then
            lblWidthHeight.Caption = " no sizes obtained"
        Else
            ' PNG is not a native image type for VB6
            ' There is no native method of obtaining width and height for PNG, ICO, TIFF &c
            ' so instead we use a 3rd party method.
            ' see ref point 0001 in cPNGparser.cls for PNG size extraction
            lblWidthHeight.Caption = " width " & picWidth & " height " & picHeight & " (pixels)"
        End If
            
    End If

    
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
Private Sub checkImageSize(ByRef FileName As String, ByRef picWidth As Long, ByRef picHeight As Long)
    
    If debugflg = 1 Then Debug.Print "%checkImageSize"

    'create an original size bitmap
    Dim bmpsizingImage As StdPicture
            
    ' the on error must not be activated within this routine, if it fails it goes to the calling routine
   'On Error GoTo checkImageSize_Error
   
    ' if the ico file has a corrupt header it will fail the loadpicture
    Set bmpsizingImage = LoadPicture(FileName)
    
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

    Dim answer As VbMsgBoxResult
    'Dim FileName As String
    'Dim useloop As Integer
    'Dim ff As Long
    'if the modification flag is set then ask before moving to the next icon
    On Error GoTo btnHomeRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "btnHomeRdMap"
   
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
    Dim answer As VbMsgBoxResult
    'Dim FileName As String
    'Dim useloop As Integer
    'Dim ff As Long
    
    'if the modification flag is set then ask before moving to the next icon
    On Error GoTo btnEndRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "btnEndRdMap"
   
    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    rdMapHScroll.Value = rdMapHScroll.Max

    ' we signify that all changes have been lost
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

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
    'if the modification flag is set then ask before moving to the next icon
    Dim answer As VbMsgBoxResult
    'Dim FileName As String
    Dim useloop As Integer
    
    On Error GoTo btnPrev_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnPrev_Click"

    If btnSave.Enabled = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
        If mapImageChanged = True Then
            ' now change the icon image back again
            ' the target picture control and the icon size
            Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
            mapImageChanged = False
        End If

    End If
    
    'display the new icon data
    
    'decrement the icon number
    rdIconNumber = rdIconNumber - 1
    'check we haven't gone too far
    If rdIconNumber < 0 Then rdIconNumber = 0
    
    ' only move the map if the array has been populated,
    If Not picRdMap(0).ToolTipText = vbNullString Then
        ' I want to test to see if the picture property is populated but
        ' as the picture property is not being set by Lavolpe's method then we can't test for it
        ' testing the tooltip above is one method of seeing if the map has been created
        ' as the program sets the tooltip just when the transparent image is set
    
        ' moves the RdMap on one position (one click) if it is already set at the rightmost screen position
        If rdIconNumber < rdMapHScroll.Value Then
            btnMapNext_Click
        End If
    End If

    lblRdIconNumber.Caption = Str(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
        
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
' Date      : 27/09/2019
' Purpose   : Display the icon details, text, sizing and image extracted via various methods according to source
'---------------------------------------------------------------------------------------
'
Private Sub displayIconElement(ByVal iconCount As Integer, ByRef picBox As PictureBox, ByRef icoPreset As Integer, ByVal showProperties As Boolean)
    Dim FileName As String
    Dim qPos As Long
    Dim filestring As String
    Dim suffix As String
    Dim picSize As Long
    
    'if it is a good icon then read the data
    On Error GoTo displayIconElement_Error
    If debugflg = 1 Then DebugPrint "%" & "displayIconElement"

    If FExists(rdSettingsFile) Then ' does the alternative settings.ini exist?
        'get the rocketdock alternative settings.ini for this icon alone
        readSettingsIni (iconCount)
    End If

    'showProperties = True
    If showProperties = True Then
        ' if the incoming text has <quote> then replace those with a " TODO ?
        txtCurrentIcon.Text = sFilename ' build the full path
        
        lblName.Text = sTitle
        txtTarget.Text = sCommand
        txtArguments.Text = sArguments
        txtStartIn.Text = sWorkingDirectory
        
        'If the docklet entry in the settings.ini is populated then blank off all the target, folder and image fields
        If sDockletFile <> "" Then
              lblName.Enabled = False
              txtCurrentIcon.Enabled = False
              
              'only enable the target fields and use the target field as a temporary location for the docklet data
              txtTarget.Text = sDockletFile
              txtTarget.Enabled = True
              btnTarget.Enabled = True
              
              txtArguments.Enabled = False
              txtStartIn.Enabled = False
              comboRun.Enabled = False
              comboOpenRunning.Enabled = False
              checkPopupMenu.Enabled = False
              btnSelectStart.Enabled = False
        Else
              lblName.Enabled = True
              txtCurrentIcon.Enabled = True
              txtTarget.Enabled = True
              txtArguments.Enabled = True
              txtStartIn.Enabled = True
              comboRun.Enabled = True
              comboOpenRunning.Enabled = True
              checkPopupMenu.Enabled = True
              btnTarget.Enabled = True
              btnSelectStart.Enabled = True
        End If
        
        If sIsSeparator = "1" Then
              lblName.Text = "Separator"
              lblName.Enabled = False
              txtCurrentIcon.Enabled = False
              txtTarget.Enabled = False
              btnTarget.Enabled = False
              txtArguments.Enabled = False
              txtStartIn.Enabled = False
              comboRun.Enabled = False
              comboOpenRunning.Enabled = False
              checkPopupMenu.Enabled = False
              btnSelectStart.Enabled = False
        End If
        
        comboRun.ListIndex = Val(sShowCmd)
        comboOpenRunning.ListIndex = Val(sOpenRunning)
        checkPopupMenu.Value = Val(sUseContext)
    End If
    
    'If the docklet entry in the settings.ini is populated then set a helpful tooltiptext
    If sDockletFile <> "" Then
        picBox.ToolTipText = "Icon number " & iconCount + 1 & "You can modify this docklet by selecting a new target, click on the ... button next to the target field."
    Else
        picBox.ToolTipText = "Icon number " & iconCount + 1 & " = " & sFilename
    End If
    picPreview.Tag = sFilename
    
    suffix = ExtractSuffix(LCase(sFilename))

    ' test whether it is a valid file with a path or just a relative path
    If InStr(sFilename, "?") Then
        FileName = sFilename
        lblFileInfo.Caption = ""
    ElseIf FExists(sFilename) Then
        FileName = sFilename  ' a full valid path so leave it alone
        picSize = FileLen(FileName)
        lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
    Else
        FileName = rdAppPath & "\" & sFilename ' a relative path found as per Rocketdock
        If FExists(FileName) Then
            picSize = FileLen(FileName)
            lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
            txtCurrentIcon.Text = FileName
            
            ' if the path is the relative path from the RD folder then repair it giving it a full path
            sFilename = FileName
            PutINISetting "Software\RocketDock\Icons", iconCount & "-FileName", sFilename, rdSettingsFile
        End If
    End If

    ' if the user drags an icon to the dock then RD takes a icon link of the following form:
    'FileName = "C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\vbexpress.exe?62453184"
    
    If InStr(sFilename, "?") Then ' Note: the question mark is an illegal character and test for a valid file will fail in VB.NET despite working in VB6 so we test it as a string instead
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
                picSize = FileLen(filestring)
                lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (binary)"

            Else
                ' the file may have a ? in the string but does not match otherwise in any useful way
                FileName = rdAppPath & "\icons\" & "help.png"
            End If
            
        Else ' the file doesn't exist in any form with ? or otherwise as a valid path
            If sIsSeparator = 1 Then
                FileName = App.path & "\my collection\" & "separator.png" ' change to separator
            Else
                FileName = rdAppPath & "\icons\" & "help.png"
            End If
            Call displayResizedImage(FileName, picBox, icoPreset)
            'dllFrame.Visible = False
        End If
    Else
        Call displayResizedImage(FileName, picBox, icoPreset)
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
    
    If filesIconList.ListIndex >= 0 Then
        'if the two are the same it ought not to trigger a pointless click
        If vScrollThumbs.Value <> filesIconList.ListIndex Then
            vScrollThumbs.Value = filesIconList.ListIndex
        End If
    End If
    
    picFrameThumbs.Visible = True
    filesIconList.Visible = False
    frmNoFilesFound.Visible = False
    btnThumbnailView.Visible = False
    btnFileListView.Visible = True

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
   If debugflg = 1 Then Debug.Print "%refreshThumbnailViewPanel"

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
                picThumbIcon(thumbIndexNo).BorderStyle = 1
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
'
'               When I created this I did not know that image lists and imageview controls existed
'               and if I had then I may have used them. However, creating something resembling an
'               imageview in code is arguably better than loading another OCx.
'
'---------------------------------------------------------------------------------------
'
Private Sub populateThumbnails(ByRef imageSize As Integer, startItem As Integer)

    Dim useloop As Integer
    Dim FileName As String
    Dim shortFilename As String
    Dim tooltip As String
    Dim suffix As String
    Dim busyFilename As String
    
    On Error GoTo populateThumbnails_Error
    If debugflg = 1 Then DebugPrint "%" & "populateThumbnails"

    Call placeThumbnailPicboxes(imageSize)  ' in the right place
   
    ' change the image to a tree view icon
    ' create a matrix of 12 x image objects from file

    ' starting with the startItem from filesIconList
    ' we extract the number
        
    ' changing the vScrollThumbs.Maximum causes the vScrollThumbs_changed routine to trigger in Vb.net but not in VB6 so it has the VB.NET equivalent check for a non-zero value for compatibility
    If filesIconList.ListCount - 1 > 0 Then
        vScrollThumbs.Max = filesIconList.ListCount - 1
    End If
    
    Dim newString As String
    Dim leftBit As String
    Dim rightBit As String
    
    ' populate each with an image
    For useloop = 0 To 11
    
        'startItem = filesIconList.ListIndex ' the starting point in the file list for the thumnbnails to start
        'when there are less than a screenful of items the.ListIndex returns -1
        
        If startItem = -1 Then
            startItem = 0
        End If
        
        ' .net collection can't handle going up to or beyond the count, VB6 control array copes
        ' but the count check is here for compatibility with the .NET version.
        
        If useloop + startItem < filesIconList.ListCount Then
            ' take the fileame from the underlying filelist control
            shortFilename = filesIconList.List(useloop + startItem)
            
            FileName = textCurrentFolder.Text & "\" & shortFilename ' changed from filelistbox.path to different path source for Vb.NET compatibility
            ' if any file does not exist
            If FExists(FileName) = False Then
                FileName = App.path & "\" & "blank.jpg"
                picThumbIcon(useloop).ToolTipText = vbNullString
                lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
                picThumbIcon(useloop).Picture = LoadPicture(FileName)
                
                ' display the image within the specified picturebox
                Call displayResizedImage(FileName, picThumbIcon(useloop), imageSize)
            End If
            
            If filesIconList.List(useloop + startItem) <> vbNullString Then
                ' set the tooltip to the filename
                picThumbIcon(useloop).ToolTipText = shortFilename
                
                ' synch. the label to the tooltiptext
                lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
                
                ' centre align the label
                lblThumbName(useloop).Alignment = 2
                
                ' if the label is too long for the label width then add a CR so that the wordwrap
                ' feature of the label comes into play, by default it will wrap on spaces alone
                If Len(picThumbIcon(useloop).ToolTipText) > 13 Then
                    leftBit = Left$(picThumbIcon(useloop).ToolTipText, 13) ' left of string
                    rightBit = Mid$(picThumbIcon(useloop).ToolTipText, 14) ' right of string
                    newString = leftBit & vbCr & rightBit     ' insert vbCr
                    lblThumbName(useloop).Caption = newString
                End If
                
                'if the remaining label string is longer than two lines then truncate the text
                If Len(lblThumbName(useloop).Caption) > 25 Then
                    lblThumbName(useloop).Caption = Left$(lblThumbName(useloop).Caption, 25) & "..."
                End If
                
                ' display the image within the specified picturebox
                Call displayResizedImage(FileName, picThumbIcon(useloop), imageSize)
    
            End If
            thumbArray(useloop) = useloop + startItem
            lblThumbName(useloop).ZOrder
        Else
            
            picThumbIcon(useloop).ToolTipText = vbNullString
            lblThumbName(useloop).Caption = picThumbIcon(useloop).ToolTipText
            picThumbIcon(useloop).Picture = LoadPicture(App.path & "\" & "blank.jpg")
            
        End If
  
        ' do the hourglass timer
        'the timer busy display moved from the non-functional timer to here where it works
        If displayHourglass = True Then

            picBusy.Visible = True
            busyCounter = busyCounter + 1
            If busyCounter >= 7 Then busyCounter = 1
            If classicTheme = True Then
                busyFilename = App.path & "\busy-F" & busyCounter & "-32x32x24.jpg"
            Else
                busyFilename = App.path & "\busy-A" & busyCounter & "-32x32x24.jpg"
            End If
            picBusy.Picture = LoadPicture(busyFilename)
            
            ' attempted to load using LaVolpe's method but to no avail
            
'            cImage.LoadPictureFile busyFilename, 32, 32, True, 32
'            Call refreshPicBox(picBusy, 32)
        End If
    Next useloop
    
    ' additional code to deal with ICO positioning
    ' if the thumbnail is an ico, nudge the image over a bit as it displays ICOs fromt he top left
    ' this is dealt with in the VB.NET version by padding it just after displaying it
    
    For useloop = 0 To 11
        
        ' it uses the tooltip as the filename is not easy to extract being that the image is written to
        ' the picture box in a non-standard manner. The label might not have the full file name and the
        ' filtered filenames are not stored in an array for easy access. They are stored in the tooltip.
        ' that's why it uses the tooltip to confirm the filename
        
        tooltip = picThumbIcon(useloop).ToolTipText
        If tooltip <> vbNullString Then
            suffix = ExtractSuffix(LCase(tooltip))
    
            If thumbImageSize = 32 Then
               If suffix = "ico" Then
                    frmThumbLabel(useloop).Visible = True
                    picThumbIcon(useloop).Left = picThumbIcon(useloop).Left + 300
                    picThumbIcon(useloop).Top = picThumbIcon(useloop).Top + 300
                End If
            End If
            picThumbIcon(useloop).Visible = True
        End If
    Next useloop
        
    ' hide or show the labels - best placed here
    For useloop = 0 To 11
        If thumbImageSize = 32 Then
                frmThumbLabel(useloop).Visible = True
                frmThumbLabel(useloop).ZOrder
        Else
                frmThumbLabel(useloop).Visible = False
        End If
    Next useloop
    
    picBusy.Visible = False
    
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
Private Sub populateRdMap(ByVal xDeviation As Integer)

    Dim useloop As Integer
    Dim busyFilename As String
    Dim dotString As String
   
    On Error GoTo populateRdMap_Error
    If debugflg = 1 Then DebugPrint "%" & "populateRdMap"

    dotString = vbNullString
    dotCount = 0

    'Refresh ' display the results prior to the for loop
    ' the above command allows the working button to show as it should
                
    ' populate each with an image
    For useloop = 0 To rdIconMax
        picRdMap(useloop).BorderStyle = 1 ' put a border around the picboxes to show an update
        
        ' using the deviation from the extracted start
        ' visit the filelist at that point and extract the filename
        '  and extract the file path
        
        ' the target picture control and the icon size
        Call displayIconElement(useloop + xDeviation, picRdMap(useloop), 32, False)
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
            busyFilename = App.path & "\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.path & "\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        picBusy.Visible = False

    Next useloop

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
Private Sub deleteRdMap()

    Dim useloop As Integer
    Dim busyFilename As String
    Dim dotString As String
    Dim answer As VbMsgBoxResult
    
   On Error GoTo deleteRdMap_Error
   If debugflg = 1 Then Debug.Print "%deleteRdMap"

    If picRdMapGotFocus <> True Then Exit Sub
    answer = MsgBox(" This will delete all the icons in your dock , are you sure?", vbYesNo)
    If answer = vbNo Then
        Exit Sub
    End If

    Call backupSettings("")

    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    For useloop = 1 To rdIconMax
            
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
            busyFilename = App.path & "\busy-F" & busyCounter & "-32x32x24.jpg"
        Else
            busyFilename = App.path & "\busy-A" & busyCounter & "-32x32x24.jpg"
        End If
        picBusy.Picture = LoadPicture(busyFilename)
        picBusy.Visible = False
            
    Next useloop
        
    'decrement the icon count and the maximum icon
    theCount = 1
    rdIconMax = 0
    
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\RocketDock\Icons", "count", theCount, rdSettingsFile
    
    'set the slider bar
    rdMapHScroll.Max = 1

    rdIconNumber = 0
    
    ' load the new icon as an image in the picturebox
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32, True)
    
    Call populateRdMap(0) ' regenerate the map from position zero


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
'
Private Sub filesIconList_Click()
    Dim FileName As String
    Dim picSize As Long
    Dim suffix As String
    
    On Error GoTo filesIconList_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "filesIconList_Click"

    picPreview.AutoRedraw = True
    picPreview.AutoSize = False
    
    FileName = textCurrentFolder.Text ' textCurrentFolder.Text ' changed from .path to use alternate path source to be compatible with VB.NET
    If Right$(FileName, 1) <> "\" Then
        FileName = FileName & "\"
    End If
    FileName = FileName & filesIconList.FileName
    
    If filesIconList.FileName = "" Then
        Exit Sub
    End If
    
    suffix = ExtractSuffix(FileName)
    picSize = FileLen(FileName)
    lblFileInfo.Caption = "File Size: " & Format(picSize, "###,###,###") & " bytes (" & UCase$(suffix) & ")"
    
    'If picFrameThumbsGotFocus = True Then
        
    'refresh the preview displaying the selected image
    Call displayResizedImage(FileName, picPreview, icoSizePreset)
    
    'End If
    
    filesIconListGotFocus = True
    
    picPreview.ToolTipText = FileName
    picPreview.Tag = FileName
    
    ' we signify that no changes have been made
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

filesIconList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_Click of Form rDIconConfigForm"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuFont_Click
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuFont_Click()
    Dim suppliedFont As String
    Dim suppliedSize As Integer
    Dim suppliedStrength As Boolean
    Dim suppliedStyle As Boolean
    
    On Error GoTo mnuFont_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuFont_Click"
   
    Set FontDlg = New CommonDlgs

    With FontDlg
            .DialogTitle = "Select a Font"
            .flags = cdlCFScreenFonts _
                  Or cdlCFBoth _
                  Or cdlCFEffects _
                  Or cdlCFApply _
                  Or cdlCFForceFontExist
    End With
    
    On Error Resume Next
    If FontDlg.ShowFont(hWnd, hDC) Then
        suppliedFont = FontDlg.FontName
        suppliedSize = FontDlg.FontSize
        suppliedStrength = FontDlg.FontBold
        suppliedStyle = FontDlg.FontItalic
    End If
            
l_err1:
'    If dlgFontForm.dlgFont.FontName = vbNullString Then
'        Exit Sub
'    End If
        
    If Err <> 32755 Then    ' User didn't chose Cancel.
        'suppliedFont = dlgFontForm.dlgFont.FontName
    End If
    
    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        PutINISetting "Software\RocketDockSettings", "defaultFont", suppliedFont, toolSettingsFile
        PutINISetting "Software\RocketDockSettings", "defaultSize", suppliedSize, toolSettingsFile
        PutINISetting "Software\RocketDockSettings", "defaultStrength", suppliedStrength, toolSettingsFile
        PutINISetting "Software\RocketDockSettings", "defaultStyle", suppliedStyle, toolSettingsFile
    End If

    If suppliedFont <> vbNullString Then
        Call changeFont(FontDlg.FontName, FontDlg.FontSize, FontDlg.FontBold, FontDlg.FontItalic)
    End If

   On Error GoTo 0
   Exit Sub

mnuFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFont_Click of Form rDIconConfigForm"
    
End Sub

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
    txtCurrentIcon.Text = relativePath & "\" & filesIconList.FileName
    
    ' now change the icon image
    ' the target picture control and the icon size
    Call displayResizedImage(txtCurrentIcon.Text, picRdMap(rdIconNumber), 32)
        
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
Private Sub filesIconList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo filesIconList_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "filesIconList_MouseDown"

    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If


   On Error GoTo 0
   Exit Sub

filesIconList_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filesIconList_MouseDown of Form rDIconConfigForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : readDefaultFolder
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the default folder, the folder the user previously selected
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
    If debugflg = 1 Then DebugPrint "%" & "readDefaultFolder"

    If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        defaultFolderNodeKey = GetINISetting("Software\RocketDockSettings", "defaultFolderNodeKey", toolSettingsFile)
    End If

    folderTreeView.HideSelection = False ' Ensures found item highlighted

    If defaultFolderNodeKey <> vbNullString Then
        For iX = 1 To folderTreeView.Nodes.count
            If Trim(folderTreeView.Nodes(iX).Key) = Trim(defaultFolderNodeKey) Then
                iFound = True
                Exit For
            End If
        Next
        If iFound Then
            ' highlight the treeview item
            folderTreeView.Nodes(iX).EnsureVisible
            folderTreeView.SelectedItem = folderTreeView.Nodes(iX)
            folderTreeView.Nodes(iX).Selected = True
            'folderTreeView_NodeSelect (Node)
            'folderTreeView_Click ' click on the selected item ' does not trigger a thumbnail refresh on startup
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
' Purpose   : Read the registry one line at a time and create a temporary settings file
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryWriteSettings()
    Dim useloop As Integer
    
   On Error GoTo readRegistryWriteSettings_Error
      If debugflg = 1 Then DebugPrint "%" & "readRegistryWriteSettings"
   
   

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
' Purpose   : check for the existence of the rocketdock binary
'---------------------------------------------------------------------------------------
'
Public Function driveCheck(ByVal folder As String) As Boolean
   Dim sAllDrives As String
   Dim sDrv As String
   Dim sDrives() As String
   Dim cnt As Long
   Dim folderString As String
   
  'get the list of all drives
   On Error GoTo driveCheck_Error
      If debugflg = 1 Then DebugPrint "%" & "driveCheck"
   
   

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
           'test for the Rocketdock binary
            rdAppPath = folderString
            If FExists(rdAppPath & "\rocketdock.exe") Then
                'MsgBox "Rocketdock binary exists"
                driveCheck = True
                If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                    PutINISetting "Software\RocketDockSettings", "rocketDockLocation", rdAppPath, toolSettingsFile
                End If
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
' Purpose   : Determine the number and name of drives using VB alone
'---------------------------------------------------------------------------------------
'
Private Function GetDriveString() As String

    'Used by both demos
      
    ' returns string of available
    ' drives each separated by a null
    ' Dim sBuff As String
    '
    ' possible 26 drives, three characters
    ' each plus a trailing null for each
    ' drive letter and a terminating null
    ' for the string
    
    Dim I As Long
    Dim builtString As String

    '===========================
    'pure VB approach, no controls required
    'drive letters are found in positions 1-UBound(Letters)
    '"C:\ D:\ E:\ &frameProperties"
    
    On Error GoTo GetDriveString_Error
       If debugflg = 1 Then DebugPrint "%" & "GetDriveString"

    

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
' Purpose   : Check if the drive found is a valid one
'---------------------------------------------------------------------------------------
'
Public Function ValidDrive(ByVal D As String) As Boolean
   On Error GoTo ValidDrive_Error
      If debugflg = 1 Then DebugPrint "%" & "ValidDrive"

   

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
' Purpose   : Add the rocketdock icon folders that exist below the RD icons folder
'---------------------------------------------------------------------------------------
'
Private Sub addRocketdockFolders()
    Dim pathCheck As String
    
    On Error GoTo addRocketdockFolders_Error
    If debugflg = 1 Then DebugPrint "%" & "addRocketdockFolders"

    pathCheck = rdAppPath & "\icons"
        
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
    Dim SteampunkIconFolder As String
    
   On Error GoTo setSteampunkLocation_Error
   If debugflg = 1 Then DebugPrint "%" & "setSteampunkLocation"


    SteampunkIconFolder = App.path & "\my collection"
    
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
    Dim CustomIconFolder As String
    
    ' read the settings ini file
    'eg. CustomIconFolder=?E:\dean\steampunk theme\icons\
    On Error GoTo readCustomLocation_Error
    If debugflg = 1 Then DebugPrint "%" & "CustomIconFolder"

    If FExists(rdSettingsFile) Then
        CustomIconFolder = GetINISetting("Software\RocketDock", "CustomIconFolder", rdSettingsFile)
    End If
    
    If Not CustomIconFolder = vbNullString Then
        CustomIconFolder = Mid(CustomIconFolder, 2) ' remove the question mark
        If DirExists(CustomIconFolder) Then
            ' add the chosen folder to the treeview
            folderTreeView.Nodes.Add , , CustomIconFolder, CustomIconFolder
            Call addtotree(CustomIconFolder, folderTreeView)
            folderTreeView.Nodes(CustomIconFolder).Text = "custom folder"
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo Form_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%" & "Form_MouseDown"
   
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
Private Sub frameButtons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo frameButtons_MouseDown_Error
   If debugflg = 1 Then DebugPrint "%" & "frameButtons_MouseDown"

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
Private Sub FrameFolders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub frameIcons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub framePreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
' Procedure : frameProperties_MouseDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : show the standard menu - this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub frameProperties_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo frameProperties_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "frameProperties_MouseDown"
   
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
                mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
    End If


   On Error GoTo 0
   Exit Sub

frameProperties_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure frameProperties_MouseDown of Form rDIconConfigForm"

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

   Call thumbsLostFocus

   On Error GoTo 0
   Exit Sub

picFrameThumbs_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picFrameThumbs_LostFocus of Form rDIconConfigForm"
End Sub





'Private Sub frmThumbLabel_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
'    frmThumbLabel(Index).ZOrder
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
    Dim useloop As Long
    Dim startPos As Long
    Dim maxPos As Long
    Dim rdIconMaxLong As Long
    Dim spacing As Integer
    
    On Error GoTo rdMapHScroll_Change_Error
    If debugflg = 1 Then DebugPrint "%" & "rdMapHScroll_Change"
   
    spacing = 540

    rdIconMaxLong = rdIconMax
    rdMapHScroll.Min = 0
    rdMapHScroll.Max = theCount - 15
    
    startPos = rdMapHScroll.Value - 1
    
    'xlabel.Caption = startPos
    'nLabel.Caption = (startPos * spacing)
    
    maxPos = rdIconMaxLong * spacing
    
    For useloop = 0 To rdIconMax
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


'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 20/09/2019
' Purpose   : Test to see if the system colour settings have changed due to a theme changing
'---------------------------------------------------------------------------------------
'
Public Sub themeTimer_Timer()
    Dim SysClr As Long

' This should only be required on a machine that can give the Windows classic theme to the UI
' that excludes windows 8 and 10 so this timer can be switched off on these o/s.

   On Error GoTo themeTimer_Timer_Error
   If debugflg = 1 Then Debug.Print "%themeTimer_Timer"

    SysClr = GetSysColor(COLOR_BTNFACE)
    If debugflg = 1 Then DebugPrint "COLOR_BTNFACE = " & SysClr ' generates too many debug statements in the log
    If SysClr <> storeThemeColour Then
    
        Call setThemeColour

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

    lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200) ' grey
    lblThumbName(thumbIndexNo).ForeColor = RGB(0, 0, 0) ' black

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

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

txtArguments_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtArguments_Change of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblName_Change
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblName_Change()
   On Error GoTo lblName_Change_Error
   If debugflg = 1 Then DebugPrint "%lblName_Change"

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

lblName_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblName_Change of Form rDIconConfigForm"
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

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

   On Error GoTo 0
   Exit Sub

txtStartIn_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtStartIn_Change of Form rDIconConfigForm"
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

    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"

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
            deleteRdMap
    End If
    
    'f5
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

    ' home
    If KeyCode = 36 Then
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the first icon
            btnHomeRdMap
        ElseIf picFrameThumbsGotFocus = True Then
            refreshThumbnailView = True
            triggerStartCalc = True
            thumbIndexNo = 0
            vScrollThumbs.Value = vScrollThumbs.Min
        End If
    End If
    
    ' end
    If KeyCode = 35 Then
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll to the end of the rdMap
            btnEndRdMap
        ElseIf picFrameThumbsGotFocus = True Then
            refreshThumbnailView = True
            triggerStartCalc = True
            vScrollThumbs.Value = vScrollThumbs.Max
        End If
    End If
    

    If debugflg = 1 Then DebugPrint "%" & "picFrameThumbsGotFocus= " & picFrameThumbsGotFocus

    '38 is up
    If KeyCode = 38 Then
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll up one line
        ElseIf picFrameThumbsGotFocus = True Then
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
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then scroll down one line 'TODO
        ElseIf picFrameThumbsGotFocus = True Then
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
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnPrev_Click
        ElseIf picFrameThumbsGotFocus = True Then
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
            
'            DebugPrint "getkeypress left "
'            DebugPrint thumbIndexNo
'            Sleep (1000)
        End If
    End If
    
    '39 is right
    If KeyCode = 39 Then
        
        'DebugPrint "########################### getkeypress right STARTS"
    
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then
            btnNext_Click
        ElseIf picFrameThumbsGotFocus = True Then
            refreshThumbnailView = False
            If thumbArray(thumbIndexNo + 1) <= vScrollThumbs.Max Then
                thumbIndexNo = thumbIndexNo + 1
            End If

            If thumbIndexNo > 11 Then
                thumbIndexNo = 11
                ' check if there are any icons subsequent to this
                ' if so then scroll down one line using the vertical scroll bar
                ' and select the next icon, the first icon on that line
                
                If rdIconMax > filesIconList.ListIndex Then
                    refreshThumbnailView = True
                    vScrollThumbs.Value = vScrollThumbs.Value + 1
                    vScrollThumbsGotFocus = False
                    picFrameThumbsGotFocus = True
                End If
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
'            DebugPrint "########################### getkeypress right ENDS"
'            Sleep (1000)

        End If
    End If
    
    '33 is page up
    If KeyCode = 33 Then
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then do nothing
        ElseIf picFrameThumbsGotFocus = True Then
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
        
        If picRdMapGotFocus = True Then
            ' if the rdMap has focus then do nothing
        ElseIf picFrameThumbsGotFocus = True Then
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

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblThumbName_Click
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblThumbName_Click(Index As Integer)
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
Private Sub picFrameThumbs_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo picFrameThumbs_KeyDown_Error
   If debugflg = 1 Then DebugPrint "%picFrameThumbs_KeyDown"

    Call getKeyPress(KeyCode)

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
Private Sub picFrameThumbs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
'changed from a click to a mousedown as it allows me to catch the right button press and retain the index
Private Sub picThumbIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim thumbItemNo As Integer
    
    If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_MouseDown"
    On Error GoTo picThumbIcon_MouseDown_Error
    
    If Index = 0 Then ' a temporary kludge to fix a bug
        thumbPos0Pressed = True
    Else
        thumbPos0Pressed = False
    End If
    
    If Button = 2 Then
        menuAddToDock.Caption = "Add icon at position " & rdIconNumber + 1 & " in the map"
        storedIndex = Index ' get the icon number from the array's index
        Me.PopupMenu thumbmenu, vbPopupMenuRightButton
    Else
        'Do not refresh the whole thumbnail view array
        refreshThumbnailView = False
        keyPressOccurred = True
        
        thumbIndexNo = Index ' allow other functions access to the chosen index number
    
        ' extract the filename from the associated array
        If Not picThumbIcon(Index).ToolTipText = vbNullString Then ' we use the tooltip because the .picture property is not populated
            thumbItemNo = thumbArray(Index)
            'this next line change is meant to trigger a re-click but it does not when the index is unchanged from previous click
            vScrollThumbs.Value = thumbItemNo

             ' this next if then checks to see if the stored click is the same , if so it triggers a click on the item in the underlying file list box
'            If storedIndex <> Index Or storedIndex = 9999 Then ' if the storedindex = 9999 it is the first time the icon has been pressed so it triggers
'               'Call vScrollThumbs_Change
'            End If
            storedIndex = Index
        End If
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

Private Sub picThumbIcon_DblClick(Index As Integer)
    Dim itemno As Integer

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
Private Sub picThumbIcon_GotFocus(Index As Integer)
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
Private Sub picThumbIcon_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_KeyDown"
    On Error GoTo picThumbIcon_KeyDown_Error

    picFrameThumbsGotFocus = True
    picRdMapGotFocus = False
    previewFrameGotFocus = False
    filesIconListGotFocus = False
    vScrollThumbsGotFocus = False

    Call getKeyPress(KeyCode)

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
Private Sub picThumbIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo picThumbIcon_MouseMove_Error
   'If debugflg = 1 Then DebugPrint "%" & "picThumbIcon_MouseMove"

    frmThumbLabel(Index).ZOrder

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

    Dim useloop As Integer

    On Error GoTo picPreview_DblClick_Error
    If debugflg = 1 Then DebugPrint "%" & "picPreview_DblClick"
   
    For useloop = 0 To filesIconList.ListCount - 1
    ' TODO - extract just the filename from txtCurrentIcon.Text
        If filesIconList.List(useloop) = GetFileNameFromPath(txtCurrentIcon.Text) Then
            filesIconList.ListIndex = useloop
            GoTo l_found_file ' if the file is found no need to process the whole list
        End If
    Next useloop
    MsgBox ("The icon " & GetFileNameFromPath((txtCurrentIcon.Text)) & " is not found in the currently selected folder, please select the " & GetDirectory(txtCurrentIcon.Text) & " folder")
l_found_file:

    ' using the current preview image as the start point on the list, repopulate the thumbs
    If picFrameThumbs.Visible = True Then
        Call populateThumbnails(thumbImageSize, filesIconList.ListIndex)
    
        removeThumbHighlighting

        'highlight the current thumbnail
        thumbIndexNo = 0
        If thumbImageSize = 64 Then 'larger
            picThumbIcon(thumbIndexNo).BorderStyle = 1
        ElseIf thumbImageSize = 32 Then
            lblThumbName(thumbIndexNo).BackColor = RGB(212, 208, 200)
        End If
    End If

   On Error GoTo 0
   Exit Sub

picPreview_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picPreview_DblClick of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFileNameFromPath
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : A function to GetFileNameFromPath
'---------------------------------------------------------------------------------------
'
Public Function GetFileNameFromPath(ByRef strFullPath As String) As String
   On Error GoTo GetFileNameFromPath_Error
   If debugflg = 1 Then DebugPrint "%" & "GetFileNameFromPath"
   
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
' Purpose   : show the standard menu - but add one new option for adding this icon to the map -
'             this has to be done for each area that requires a menu
'---------------------------------------------------------------------------------------
'
Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub picRdMap_GotFocus(Index As Integer)
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
Private Sub picRdMap_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
Private Sub picRdMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim useloop As Integer
    Dim answer As VbMsgBoxResult
    'Dim picX As Integer
    
   On Error GoTo picRdMap_MouseDown_Error
      If debugflg = 1 Then DebugPrint "%" & "picRdMap_MouseDown"
   
   
   'this is the preliminary code to allow sliding of the icons by mouse and cursor
   
'    picRdMap(Index).ZOrder 'the drag pic
'    'so it appears over the top of other controls
'    picX = X - 500
    
    'MsgBox "X = " & picRdMap(Index).Left
        
   'this is the preliminary code to allow sliding of the icons by mouse and cursor
        


    If Button = 2 Then
        rdIconNumber = Index ' get the icon number from the array's index
        Me.PopupMenu rdMapMenu, vbPopupMenuRightButton
    Else
        If btnSave.Enabled = True Then
            If chkConfirmSaves.Value = 1 Then
                answer = MsgBox(" This will lose your recent changes to this icon, are you sure?", vbYesNo)
                If answer = vbNo Then
                    Exit Sub
                End If
            End If
            If mapImageChanged = True Then
                ' now change the icon image back again
                ' the target picture control and the icon size
                Call displayResizedImage(previousIcon, picRdMap(rdIconNumber), 32)
                mapImageChanged = False
            End If
        End If
       
        rdIconNumber = Index
        
        lblRdIconNumber.Caption = Str(rdIconNumber) + 1
        lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
    
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

'Private Sub picRdMap_OLEDragDrop(Index As Integer, data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'' code retained in case I want to do a graphical drag and drop of one item in the map to another
'  Dim a As String
'  a = rdIconNumber
'End Sub
'
'Private Sub picRdMap_OLEStartDrag(Index As Integer, data As DataObject, AllowedEffects As Long)
'' code retained in case I want to do a graphical drag and drop of one item in the map to another
'  Dim a As String
'  a = rdIconNumber
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : picPreview_OLEDragDrop
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : drag and drop of an image to the icon preview
'---------------------------------------------------------------------------------------
'
'Private Sub picPreview_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
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
    Dim oldRdIconMax As Integer
    Dim useloop As Integer
    Dim answer As VbMsgBoxResult
   
    On Error GoTo rdMapRefresh_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "rdMapRefresh_Click"
    
    If btnSave.Enabled = True Or mapImageChanged = True Then
        If chkConfirmSaves.Value = 1 Then
            answer = MsgBox("This will lose your recent changes to the map. Proceed?", vbExclamation + vbYesNo)
            If answer = vbNo Then
                Exit Sub
            Else
                Me.Refresh ' just to clear the dialog box remnants
            End If
        End If
    End If
    
    mapImageChanged = False
    
    Call busyStart
    
    oldRdIconMax = rdIconMax
    Call readRocketDockSettings
    If rdIconMax > oldRdIconMax Then
        ' if you do a refresh and the old rdIconMax is less than the recently read
        ' then items have been added to the Rocketdock via RD itself
        ' in which case you need to create the extra slots in the RD map
        
        'loop from the old rdIConMax to the new rdiconmax and create a new slot in the map
        For useloop = oldRdIconMax To rdIconMax
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
    'If debugflg = 1 Then DebugPrint "%" & "registryTimer_Timer" ' no messages thankyou

    origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file

    If FExists(origSettingsFile) Then ' does the original settings.ini exist?
        chkRegistry.Value = 0
        chkSettings.Value = 1
    Else
        chkRegistry.Value = 1
        chkSettings.Value = 0
    End If
    
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
           
    On Error GoTo sliPreviewSize_Change_Error
    If debugflg = 1 Then DebugPrint "%" & "sliPreviewSize_Change"

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
        Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
    Else
        If textCurrentFolder.Text <> vbNullString Then ' changed from filesIconList.path to textCurrentFolder.Text for compatibility with VB.net
            FileName = textCurrentFolder.Text ' changed from filesIconList.path to textCurrentFolder.Text for compatibility with VB.net
            If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
            FileName = FileName & filesIconList.FileName
            ' refresh the image display
            Call displayResizedImage(FileName, picPreview, icoSizePreset)
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
Private Sub folderTreeView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Dim n As Node
  
  Dim n As CCRTreeView.TvwNode

  On Error GoTo folderTreeView_MouseMove_Error
   'If debugflg = 1 Then DebugPrint "%" & "folderTreeView_MouseMove" ' we don't want too many notifications in the debug log

  Set n = folderTreeView.HitTest(x, y)
   If n Is Nothing Then
    folderTreeView.ToolTipText = "Click a folder to show the icons contained within"
    ElseIf n.Text = "icons" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are Rocketdock's own in-built icons"
    ElseIf n.Text = "custom folder" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are the custom folders that the user can add using the + button below."
    ElseIf n.Text = "my collection" Then
       folderTreeView.ToolTipText = "The sub-folders within this tree are the default folders that come with this enhanced settings utility."
    Else
     folderTreeView.ToolTipText = n.Text
   End If

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
   
    Dim path As String
    Dim defaultFolderNodeKey As String
   
    displayHourglass = True
   
    Call busyStart

    On Error GoTo l_bypass_parent
    
    If Not folderTreeView.SelectedItem Is Nothing Then

         path = folderTreeView.SelectedItem.Key
         relativePath = path
    End If
     
l_bypass_parent:
   On Error GoTo folderTreeView_Click_Error
    
    If Not folderTreeView.SelectedItem Is Nothing Then
        textCurrentFolder.Text = path
        If DirExists(textCurrentFolder.Text) Then
            filesIconList.path = textCurrentFolder.Text
        End If
        
        defaultFolderNodeKey = folderTreeView.SelectedItem.Key
        'eg. defaultFolderNodeKey=?E:\dean\steampunk theme\icons\
        If FExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
                PutINISetting "Software\RocketDockSettings", "defaultFolderNodeKey", defaultFolderNodeKey, toolSettingsFile
        End If
            
        If picFrameThumbs.Visible = True Then
            btnRefresh_Click
        End If
    End If
    
    Call busyStop
    
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
Private Sub addtotree(path As String, tv As CCRTreeView.TreeView)
    Dim folder1 As Object
    Dim FS As Object
    'Dim tvwChild As tv.tvwChild
    
    On Error GoTo addtotree_Error
    If debugflg = 1 Then DebugPrint "%" & "addtotree"

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
' Procedure : folderTreeView_DblClick
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : open the selected folder in an explorer window
'---------------------------------------------------------------------------------------
'
Private Sub folderTreeView_DblClick()
    
   On Error GoTo folderTreeView_DblClick_Error
      If debugflg = 1 Then DebugPrint "%" & "folderTreeView_DblClick"
   
   
   'Dim a As String
   'Dim fromNode As String
   
    If DirExists(folderTreeView.SelectedItem.Key) Then
        ShellExecute 0, vbNullString, folderTreeView.SelectedItem.Key, vbNullString, vbNullString, 1
    End If

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
Private Sub folderTreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo folderTreeView_MouseDown_Error
    If debugflg = 1 Then DebugPrint "%" & "folderTreeView_MouseDown"
   
    If Button = 2 Then
        ' only required for VB6, the VB.NET version allows
        ' click-throughs on transparent images so that the main main menu is shown, the image itself shows the preview menu
        mnuAddPreviewIcon.Visible = False ' "add the icon to the dock" menu option
        
        Me.PopupMenu mnupopmenu, vbPopupMenuRightButton
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
    
    btnSave.Enabled = True ' tell the program that something has changed
    btnCloseCancel.Caption = "Cancel"
    
    'txtCurrentIcon.Text = savIt
    
   On Error GoTo 0
   Exit Sub

txtCurrentIcon_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtCurrentIcon_Change of Form rDIconConfigForm"
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

    On Error GoTo vScrollThumbs_Change_Error
    If debugflg = 1 Then DebugPrint "%" & "vScrollThumbs_Change"
   
    picFrameThumbsGotFocus = True
    triggerStartCalc = True
    
    keyPressOccurred = False ' TBD
    If keyPressOccurred = True Then
        picFrameThumbsGotFocus = True
'        If Me.ActiveControl = False Then
           picFrameThumbs.SetFocus
'        End If
    Else
        'If Me.ActiveControl = False Then
            Me.SetFocus
            
        'End If
        'refreshThumbnailView = True
        triggerStartCalc = True
    End If
    
    ' update the underlying file list control that determines which icon has been selected
    ' causes the preview to be refreshed
    If filesIconList.ListCount > 0 Then
        If vScrollThumbs.Value <= vScrollThumbs.Max Then
            'if they are the same it does not trigger a click
            filesIconList.ListIndex = (vScrollThumbs.Value) 'Causes a click on the window that holds the icon files listing in text mode
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
        refreshThumbnailViewPanel ' click the button switching to thumbnail view causing a thumbnail list refresh
    Else
        refreshThumbnailView = True ' the refresh flag is set back to true immediately
    End If
    
    'remove the highlighting
    removeThumbHighlighting
    
    'highlight the current thumb
    If thumbIndexNo >= 0 Then ' -1 when there are no icons as a result of an empty filter pattern
        If thumbArray(thumbIndexNo) = 0 Or (thumbArray(thumbIndexNo) And thumbArray(thumbIndexNo) <= vScrollThumbs.Max) Then
            If thumbImageSize = 64 Then 'larger
                picThumbIcon(thumbIndexNo).BorderStyle = 1
            ElseIf thumbImageSize = 32 Then
                lblThumbName(thumbIndexNo).BackColor = RGB(10, 36, 106) ' blue
                lblThumbName(thumbIndexNo).ForeColor = RGB(255, 255, 255) ' white
            End If
        End If
    End If
    
    txtDbg01.Text = vScrollThumbs.Value
    txtDbg02.Text = vScrollThumbs.Max
            
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
    Dim useloop As Integer
    
    'remove the highlighting
   On Error GoTo removeThumbHighlighting_Error
      If debugflg = 1 Then DebugPrint "%" & "removeThumbHighlighting"
    
    'remove the highlighting
    For useloop = 0 To 11
        picThumbIcon(useloop).BorderStyle = 0
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
   On Error GoTo mnuHelpPdf_click_Error
   If debugflg = 1 Then Debug.Print "%mnuHelpPdf_click"

        Call ShellExecute(Me.hWnd, "Open", App.path & "\Rocketdock Enhanced Settings.pdf", vbNullString, App.path, 1)

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
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuFacebook_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuFacebook_Click"

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
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuLatest_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuLatest_Click"

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
    If debugflg = 1 Then DebugPrint "%" & "mnuLicence_Click"
        
    Call LoadFileToTB(licence.txtLicenceTextBox, App.path & "\licence.txt", False)
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
    If debugflg = 1 Then DebugPrint "%" & "mnuSupport_Click"

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
       If debugflg = 1 Then DebugPrint "%" & "mnuSweets_Click"
    
    
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
       If debugflg = 1 Then DebugPrint "%" & "mnuWidgets_Click"
    
    

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
' Procedure : mnuDebug_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Run the runtime debugging window exectuable
'---------------------------------------------------------------------------------------
'
Private Sub mnuDebug_Click()
    Dim NameProcess As String
    Dim debugPath As String
    
    On Error GoTo mnuDebug_Click_Error
    If debugflg = 1 Then DebugPrint "%mnuDebug_Click"

    NameProcess = "PersistentDebugPrint.exe"
    debugPath = App.path() & "\" & NameProcess
    
    If debugflg = 0 Then
        debugflg = 1
        mnuDebug.Caption = "Turn Debugging OFF"
        If FExists(debugPath) Then
            Call ShellExecute(hWnd, "Open", debugPath, vbNullString, App.path, 1)
        End If
    Else
        debugflg = 0
        mnuDebug.Caption = "Turn Debugging ON"
        checkAndKill NameProcess
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
Private Sub mnuAbout_Click(Index As Integer)
    
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
    If debugflg = 1 Then DebugPrint "%" & "menuLeft_Click"
    
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
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32, True)
    Call displayIconElement(rdIconNumber - 1, picRdMap(rdIconNumber - 1), 32, True)
    
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
' Purpose   : Click event from the menu option that slides a icon in the map one step to the right
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
    If debugflg = 1 Then DebugPrint "%" & "menuright_Click"

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
    
    Call displayIconElement(rdIconNumber, picRdMap(rdIconNumber), 32, True)
    Call displayIconElement(rdIconNumber + 1, picRdMap(rdIconNumber + 1), 32, True)

    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

   On Error GoTo 0
   Exit Sub

menuright_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuright_Click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : menuAddSummat
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add something to the RD icon map, called by all the menuAdd functions that follow
'---------------------------------------------------------------------------------------
'
Private Sub menuAddSummat(thisFilename As String, thisTitle As String, thisCommand As String, thisArguments As String, thisWorkingDirectory As String, thisDocklet As String, thIsSeparator As String)
    Dim useloop As Integer
    Dim thisIcon As Integer

    On Error GoTo menuAddSummat_Error
    If debugflg = 1 Then DebugPrint "%" & "menuAddSummat"
  
    Call busyStart

    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    ' starting at the end of the rocketdock map, scroll backward and increment the number
    ' until we reach the current position.
    
    For useloop = rdIconMax To rdIconNumber Step -1
        ' read the rocketdock alternative settings.ini
         readSettingsIni (useloop) ' the settings.ini only exists when RD is set to use it
        
        ' and increment the identifier by one
         writeSettingsIni (useloop + 1)
    Next useloop
    
    'increment the new icon count
    theCount = theCount + 1
    
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\RocketDock\Icons", "count", theCount, rdSettingsFile

    rdIconMax = theCount - 1 '

    'set the slider bar to the new maximum
    rdMapHScroll.Max = theCount - 15

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
    ' with the following blank characteristics
    sFilename = thisFilename ' the default Rocketdock filename for a blank item
    
    sTitle = thisTitle
    sCommand = thisCommand
    sArguments = thisArguments
    sWorkingDirectory = thisWorkingDirectory
    sDockletFile = thisDocklet
    sIsSeparator = thIsSeparator
    
    sShowCmd = 0
    sOpenRunning = 0
    sUseContext = 0
    
    'set the fields for this icon to the correct value as supplied
    lblName.Text = sTitle
    
    If sDockletFile <> vbNullString Then
        txtTarget.Text = sDockletFile
    Else
        txtTarget.Text = sCommand
    End If
    
    txtArguments.Text = sArguments
    txtStartIn.Text = sWorkingDirectory
    
    comboRun.ListIndex = 0 '"Normal"
    comboOpenRunning.ListIndex = 0 ' "Use Global Setting"
    checkPopupMenu.Value = 0
    
    writeSettingsIni (thisIcon)
    
    Call displayIconElement(thisIcon, picRdMap(thisIcon), 32, True)
    
    Call populateRdMap(0) ' regenerate the map from position zero
      
    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

    Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1) ' click on the picture box
    
    Call busyStop

   On Error GoTo 0
   Exit Sub

menuAddSummat_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuAddSummat of Form rDIconConfigForm"
    
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
    Call menuAddSummat(sFilename, sTitle, sCommand, sArguments, sWorkingDirectory, sDockletFile, sIsSeparator)
    
   On Error GoTo 0
   Exit Sub

mnuClone_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClone_Click of Form rDIconConfigForm"
    
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
    Call menuAddSummat("\Icons\help.png", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    
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
    Dim iconImage As String
    Dim iconFileName As String
    
   On Error GoTo mnuAddShutdown_click_Error
      If debugflg = 1 Then DebugPrint "%" & "mnuAddShutdown_click"
   
   
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\shutdown.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Shutdown", "C:\Windows\System32\shutdown.exe", "/s /t 00 /f /i", vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddShutdown_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddShutdown_click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddLog_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddLog_click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddLog_click_Error
    If debugflg = 1 Then DebugPrint "%mnuAddLog_click"
    
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\padlock(log off).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Log Off", "frameProperties:\WINDOWS\system32\rundll32.exe", "user32.dll, LockWorkStation", "%windir%", vbNullString, vbNullString)

    On Error GoTo 0
    Exit Sub

mnuAddLog_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddLog_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddNetwork_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a network icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddNetwork_click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddNetwork_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddNetwork_click"
   
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\big-globe(network).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    ' thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Network", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddWorkgroup_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddWorkgroup_click"
   
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\big-globe(network).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Network", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    On Error GoTo mnuAddPrinters_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddPrinters_click"
    
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\printer.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Printers", "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddTask_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddTask_click"
    
    iconFileName = App.path & "\my collection" & "\task-manager(tskmgr).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Task Manager", "taskmgr", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddControl_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddControl_click"

    iconFileName = App.path & "\my collection" & "\control-panel(control).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Control panel", "control", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    On Error GoTo mnuAddPrograms_click_Error
       If debugflg = 1 Then DebugPrint "%" & "mnuAddPrograms_click"
    
    
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\programs and features.ico"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Programs and Features", "appwiz.cpl", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddPrograms_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPrograms_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddDock_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a dock settings icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDock_click()
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddDock_click_Error
      If debugflg = 1 Then DebugPrint "%" & "mnuAddDock_click"

    iconFileName = App.path & "\my collection" & "\dock settings.ico"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Dock Settings", "[Settings]", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddAdministrative_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddAdministrative_click"

    iconFileName = App.path & "\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Administration Tools", "compmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String
    On Error GoTo mnuAddRecycle_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddRecycle_click"
   
    ' check the icon exists
    iconFileName = App.path & "\my collection" & "\recyclebin-full.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Recycle Bin", "::{645ff040-5081-101b-9f08-00aa002f954e}", vbNullString, vbNullString, vbNullString, vbNullString)

   On Error GoTo 0
   Exit Sub

mnuAddRecycle_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddRecycle_click of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAddQuit_click
' Author    : beededea
' Date      : 19/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddQuit_click()
    Dim iconImage As String
    Dim iconFileName As String

    ' check the icon exists
    On Error GoTo mnuAddQuit_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddQuit_click"
   
    iconFileName = App.path & "\my collection" & "\quit.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Quit", "[Quit]", vbNullString, vbNullString, vbNullString, vbNullString)

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
    Dim iconImage As String
    Dim iconFileName As String

    ' check the icon exists
    On Error GoTo mnuAddProgramFiles_click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuAddProgramFiles_click"
   
    iconFileName = App.path & "\my collection" & "\hard-drive-indicator-D.png"
    If FExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = "\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    Call menuAddSummat(iconImage, "Program Files", "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", vbNullString, vbNullString, vbNullString, vbNullString)

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
   If debugflg = 1 Then Debug.Print "mnuTrgtSeparator_click"
      
    sIsSeparator = "1"
        
    lblName.Text = "Separator"
        
    ' set fields to blank
    txtCurrentIcon.Text = ""
    txtTarget.Text = ""
    txtArguments.Text = ""
    txtStartIn.Text = ""
    
    lblName.Enabled = False
    txtCurrentIcon.Enabled = False
    txtTarget.Enabled = False
    btnTarget.Enabled = False
    txtArguments.Enabled = False
    txtStartIn.Enabled = False
    comboRun.Enabled = False
    comboOpenRunning.Enabled = False
    checkPopupMenu.Enabled = False
    btnSelectStart.Enabled = False

   On Error GoTo 0
   Exit Sub

mnuTrgtSeparator_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtFolder_click of Form rDIconConfigForm"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuTrgtFolder_click
' Author    : beededea
' Date      : 27/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtFolder_click()
    
    Dim getFolder As String
    Dim dialogInitDir As String
   
   On Error GoTo mnuTrgtFolder_click_Error
   If debugflg = 1 Then Debug.Print "%mnuTrgtFolder_click"

    If txtTarget.Text <> vbNullString Then
        If DirExists(txtStartIn.Text) Then
            dialogInitDir = txtTarget.Text 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = rdAppPath 'start dir, might be "C:\" or so also
        End If
    End If

    getFolder = BrowseFolder(hWnd, dialogInitDir) ' show the dialog box to select a folder
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

    sCommand = "C:\Windows\System32\shutdown.exe"
    sArguments = "/s /t 00 /f /i"
    
    txtTarget.Text = sCommand
    txtArguments.Text = sArguments

   On Error GoTo 0
   Exit Sub

mnuTrgtShutdown_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtShutdown_click of Form rDIconConfigForm"

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

    sCommand = App.path & "\rocket1.exe"
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

    sCommand = "frameProperties:\WINDOWS\system32\rundll32.exe"
    sArguments = "user32.dll, LockWorkStation"
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
' Procedure : mnuTrgtWorkgroup_click
' Author    : beededea
' Date      : 28/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtWorkgroup_click()

   On Error GoTo mnuTrgtWorkgroup_click_Error
   If debugflg = 1 Then Debug.Print "%mnuTrgtWorkgroup_click"

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
   If debugflg = 1 Then Debug.Print "%mnuTrgtNetwork_click"

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
   If debugflg = 1 Then Debug.Print "%mnuTrgtMyComputer_click"

    sCommand = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtMyComputer_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtMyComputer_click of Form rDIconConfigForm"

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
' Procedure : mnuTrgtAdministrative_click
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTrgtAdministrative_click()
   On Error GoTo mnuTrgtAdministrative_click_Error
   If debugflg = 1 Then DebugPrint "%mnuTrgtAdministrative_click"

    sCommand = "compmgmt.msc"
    txtTarget.Text = sCommand

   On Error GoTo 0
   Exit Sub

mnuTrgtAdministrative_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTrgtAdministrative_click of Form rDIconConfigForm"
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
    Dim dialogInitDir As String

    Const x_MaxBuffer = 256
    ' set the default folder to the docklet folder under rocketdock
   On Error GoTo mnuTrgtDocklet_click_Error
   If debugflg = 1 Then Debug.Print "%mnuTrgtDocklet_click"

    dialogInitDir = rdAppPath & "\docklets"
 
    With x_OpenFilename
    '    .hwndOwner = Me.hWnd
      .hInstance = App.hInstance
      .lpstrTitle = "Select a Rocketdock Docklet DLL"
      .lpstrInitialDir = dialogInitDir
      
      .lpstrFilter = "DLL Files" & vbNullChar & "*.dll" & vbNullChar & vbNullChar
      .nFilterIndex = 2
      
      .lpstrFile = String(x_MaxBuffer, 0)
      .nMaxFile = x_MaxBuffer - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = x_MaxBuffer - 1
      .lStructSize = Len(x_OpenFilename)
    End With
      
    Dim retFileName As String
    Dim retfileTitle As String
    Call f_GetOpenFileName(retFileName, retfileTitle)
    txtTarget.Text = retFileName
    'lblName.Text = retfileTitle
    
    sDockletFile = txtTarget.Text
    
    'it chooses the icon here as with a docklet no alternative icon is allowed, the docklet determines that
    If InStr(GetFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
      txtCurrentIcon.Text = rdAppPath & "\icons\clock.png"
    ElseIf InStr(GetFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
      txtCurrentIcon.Text = App.path & "\my collection\recyclebin-full.png"
    Else
      txtCurrentIcon.Text = rdAppPath & "\icons\blank.png" ' has to be an icon of some sort
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
    Dim useloop As Integer
    Dim thisIcon As Integer
    Dim notQuiteTheTop As Integer
    Dim answer As VbMsgBoxResult
        
    On Error GoTo mnuDelete_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuDelete_Click"
    If chkConfirmSaves.Value = 1 Then
        answer = MsgBox("This will delete the currently selected entry in the Rocketdock map, " & vbCr & txtCurrentIcon & "   -  are you sure?", vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    If rdIconNumber = 0 And rdIconMax = 1 Then
        MsgBox "Cannot currently delete the last icon, one icon must always be present for Rocketdock to operate - apologies."
        Exit Sub
    End If
    
    Call busyStart
    'Note: we only write to the interim settings file
    'the write to the actual settings or registry happens when the user "saves & restarts"
    
    If rdIconNumber < rdIconMax Then 'if not the top icon loop through them all and reassign the values
        notQuiteTheTop = rdIconMax - 1
        For useloop = rdIconNumber To notQuiteTheTop
            
            ' read the rocketdock alternative rdsettings.ini one item up in the list
            readSettingsIni (useloop + 1) ' the alternative rdsettings.ini only exists when RD is set to use it
            
            'write the the new item at the current location effectively overwriting it
            writeSettingsIni (useloop)
        
        Next useloop
    End If
    
    ' to tidy up we need to overwrite the final data from the rdsettings.ini, we will write sweet nothings to it
    removeSettingsIni (rdIconMax)
    
    'clear the icon
    picRdMap(rdIconMax).BackColor = &H8000000F
'    Set picRdMap(rdIconMax).Picture = LoadPicture(vbNullString)
    Set picRdMap(rdIconMax).Picture = Nothing
    Unload picRdMap(rdIconMax)
    
    ' the picbox positioning
    storeLeft = storeLeft - boxSpacing
        
    'decrement the icon count and the maximum icon
    theCount = theCount - 1
    rdIconMax = theCount - 1
    
    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\RocketDock\Icons", "count", theCount, rdSettingsFile
    
    'set the slider bar to the new maximum
    rdMapHScroll.Max = theCount - 15

    If rdIconNumber > rdIconMax Then rdIconNumber = rdIconMax
    thisIcon = rdIconNumber
    
    ' load the new icon as an image in the picturebox
    Call displayIconElement(thisIcon, picRdMap(thisIcon), 32, True)
    
    Call populateRdMap(0) ' regenerate the map from position zero
    
    btnSave.Enabled = False ' tell the program that nothing has changed
    btnCloseCancel.Caption = "Close"

    ' emulate a click on the appropriate icon in the map so that the image and details are shown
    Call picRdMap_MouseDown(thisIcon, 1, 1, 1, 1)
   
    Call busyStop

   On Error GoTo 0
   Exit Sub

mnuDelete_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDelete_Click of Form rDIconConfigForm"
    
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
       If debugflg = 1 Then DebugPrint "%" & "mnuCoffee_Click"
    
    
    
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
      If debugflg = 1 Then DebugPrint "%" & "chkTheRegistry"
   
   

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
Public Sub backupSettings(ByRef bkpFilename As String)
    'Dim AY() As String
    'Dim suffix As String
    'Dim maxBound As Integer
    'Dim fileVersion As Integer
    Dim bkpSettingsFile As String
    Dim useloop As Integer
    Dim srchSettingsFile As String
    Dim versionNumberAvailable As Integer
    Dim bkpfileFound As Boolean
    
    
        ' set the name of the bkp file
   
   On Error GoTo backupSettings_Error
      If debugflg = 1 Then DebugPrint "%" & "backupSettings"

        bkpSettingsFile = App.path & "\backup\bkpSettings.ini"
                
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
        
        bkpFilename = bkpSettingsFile

   On Error GoTo 0
   Exit Sub

backupSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure backupSettings of Form rDIconConfigForm"
        
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
   If debugflg = 1 Then Debug.Print "%menuAddToDock"

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
    'Dim useloop As Integer
    'Dim tooltip As String
    'Dim suffix As String
    
    On Error GoTo menuLargerThumbs_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "menuLargerThumbs_Click"
    
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
Private Sub Form_Unload(Cancel As Integer)
    Dim ofrm As Form
    Dim NameProcess As String
    
    ' when you create a token to be shared, you must
    ' destroy it in the Unload or Terminate event
    ' and also reset gdiToken property for each existing class
    On Error GoTo Form_Unload_Error
    'If debugflg = 1 Then DebugPrint "%" & "Form_Unload"
    
    NameProcess = "PersistentDebugPrint.exe"
    
    If debugflg = 1 Then
        checkAndKill NameProcess
    End If
        
    If m_GDItoken Then
        If Not cShadow Is Nothing Then cShadow.gdiToken = 0&
        If Not cImage Is Nothing Then
            cImage.gdiToken = 0&
            cImage.DestroyGDIplusToken m_GDItoken
        End If
    End If
    
    For Each ofrm In Forms
        Unload ofrm
    Next
    
    End

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
Private Sub mnuSubOpts_Click(Index As Integer)

    ' The 1st two options will be disabled if you do not have GDI+ installed
    
    On Error GoTo mnuSubOpts_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "mnuSubOpts_Click"

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
'ExitRoutine:

   On Error GoTo 0
   Exit Sub

mnuSubOpts_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSubOpts_Click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : refreshPicBox
' Author    : beededea
' Date      : 14/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub refreshPicBox(ByRef picBox As PictureBox, ByVal iconSizing As Integer)

    Dim newWidth As Long
    Dim newHeight As Long
    
    Dim mirrorOffsetX As Long
    Dim mirrorOffsetY As Long
    
    Dim x As Long
    Dim y As Long
    
    Dim ShadowOffset As Long
    Dim LightAdjustment As Single
        
    On Error GoTo refreshPicBox_Error
    If debugflg = 1 Then DebugPrint "%" & "refreshPicBox"

    mirrorOffsetX = 1
    mirrorOffsetY = 1

    newWidth = iconSizing: newHeight = iconSizing
    
    x = (picBox.ScaleWidth - newWidth) \ 2
    y = (picBox.ScaleHeight - newHeight) \ 2
    
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
'Private Sub vScrollThumbs_KeyDown(KeyCode As Integer, Shift As Integer)
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

'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub changeFont(suppliedFont As String, suppliedSize As Integer, suppliedStrength As String, suppliedStyle As String)
    Dim useloop As Integer
    Dim Ctrl As Control
    
    On Error GoTo changeFont_Error
    
    If debugflg = 1 Then DebugPrint "%" & "changeFont"
      
    ' a method of looping through all the controls and identifying the labels and text boxes
    For Each Ctrl In rDIconConfigForm.Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
           If suppliedFont <> "" Then Ctrl.Font.Name = suppliedFont
           If suppliedSize > 0 Then Ctrl.Font.Size = suppliedSize
           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
        End If
    Next
    
    ' change the size of the two labels beneath the preview image
    lblFileInfo.Font.Size = 7
    lblWidthHeight.Font.Size = 7
    
    ' change the font size of the large number
    lblRdIconNumber.Font.Name = "Trebuchet MS"
    lblRdIconNumber.Font.Size = 45

    'loop through the 12 dynamic icon thumbnails, they all exist by the time this routine is called
    For useloop = 0 To 11
        picThumbIcon(useloop).Font.Name = suppliedFont 'array
        If suppliedSize > 0 Then picThumbIcon(useloop).Font.Size = suppliedSize 'array
        
        frmThumbLabel(useloop).Font.Name = suppliedFont 'array
        If suppliedSize > 0 Then frmThumbLabel(useloop).Font.Size = suppliedSize 'array
        
        lblThumbName(useloop).Font.Name = suppliedFont 'array
        If suppliedSize > 0 Then lblThumbName(useloop).Font.Size = suppliedSize 'array
    Next useloop
    
    ' then the treeview that is picky about .fontname or .font.name where the others are not.
    folderTreeView.Font.Name = suppliedFont
    If suppliedSize > 0 Then folderTreeView.Font.Size = suppliedSize
    
    ' The comboboxes all autoselect when the font is changed, we need to reset this afterwards
    
    comboIconTypesFilter.SelLength = 0
    comboDockType.SelLength = 0
    comboRun.SelLength = 0
    comboOpenRunning.SelLength = 0
   
    ' after changing the font, sometimes the filelistbox changes height arbitrarily
    filesIconList.Height = 3310
   
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : beededea
' Date      : 26/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function LoadFileToTB(TxtBox As TextBox, filePath As _
   String, Optional Append As Boolean = False) As Boolean
       
    'PURPOSE: Loads file specified by FilePath into textcontrol
    '(e.g., Text Box, Rich Text Box) specified by TxtBox
    
    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    'Returns: True if Successful, false otherwise
    
    Dim iFile As Integer
    Dim s As String
    
   On Error GoTo LoadFileToTB_Error
      If debugflg = 1 Then DebugPrint "%" & "LoadFileToTB"
   
   
   If debugflg = 1 Then DebugPrint "%" & LoadFileToTB

    If Dir(filePath) = "" Then Exit Function
    
    On Error GoTo ErrorHandler:
    s = TxtBox.Text
    
    iFile = FreeFile
    Open filePath For Input As #iFile
    s = Input(LOF(iFile), #iFile)
    If Append Then
        TxtBox.Text = TxtBox.Text & s
    Else
        TxtBox.Text = s
    End If
    
    LoadFileToTB = True
    
ErrorHandler:
    If iFile > 0 Then Close #iFile

   On Error GoTo 0
   Exit Function

LoadFileToTB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function LoadFileToTB of Form rDIconConfigForm"

End Function


'---------------------------------------------------------------------------------------
' Procedure : btnTarget_MouseDown
' Author    : beededea
' Date      : 28/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTarget_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    setThemeLight
    
   On Error GoTo 0
   Exit Sub

mnuLight_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : setThemeLight
' Author    : beededea
' Date      : 26/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setThemeLight()
    
    Dim a As Long
    Dim Ctrl As Control
    
    On Error GoTo setThemeLight_Error
    If debugflg = 1 Then DebugPrint "%setThemeLight"

    classicTheme = False
    mnuLight.Checked = True
    mnuDark.Checked = False
    
    btnArrowDown.Picture = LoadPicture(App.path & "\arrowDown10.gif")
    btnMapPrev.Picture = LoadPicture(App.path & "\leftArrow10.jpg")
    btnMapNext.Picture = LoadPicture(App.path & "\rightArrow10.jpg")
    btnArrowUp.Picture = LoadPicture(App.path & "\arrowUp10.jpg")
       
    ' RGB(240, 240, 240) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(240, 240, 240)
    ' a method of looping through all the controls that require reversion of any background colouring
    For Each Ctrl In rDIconConfigForm.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(240, 240, 240)
        End If
    Next
    
    ' these elements are normal elements that should have their styling reverted
    
    'all other buttons go here
    
    picPreview.BackColor = RGB(240, 240, 240)
    picRdThumbFrame.BackColor = RGB(240, 240, 240)
    btnRemoveFolder.BackColor = RGB(240, 240, 240)
    picCover.BackColor = RGB(240, 240, 240)
    back.BackColor = RGB(240, 240, 240)
    sliPreviewSize.BackColor = RGB(240, 240, 240)

    ' on NT6 plus using the MSCOMCTL slider with the lighter default theme, the slider
    ' fails to pick up the new theme colour fully
    ' the following lines triggers a partial colour change on the treeview that has no backcolor property
    ' this also causes a refresh of the preview pane - so don't remove it.
    ' I will have to create a new slider to overcome this - not yet tested the VB.NET version
    ' do not remove - essential
    
    'a = sliPreviewSize.Value
    'sliPreviewSize.Value = 1
    'sliPreviewSize.Value = a
    
    ' the slider has a redrawing problem after changing the theme
    
    ' the above no longer required with Krool's replacement controls

   On Error GoTo 0
   Exit Sub

setThemeLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeLight of Form rDIconConfigForm"
End Sub

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
    setThemeDark

   On Error GoTo 0
   Exit Sub

mnuDark_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setThemeDark
' Author    : beededea
' Date      : 26/09/2019
' Purpose   : if running on Win7 with the classic theme setting the theme to dark should do nothing
'             if running on any other theme then setting the theme to dark should replace the visual elements
'---------------------------------------------------------------------------------------
'
Private Sub setThemeDark()

    Dim a As Long
    Dim firstRun As Boolean
    Dim Ctrl As Control
    
    firstRun = False
    
    On Error GoTo setThemeDark_Error
    If debugflg = 1 Then DebugPrint "setThemeDark"

    classicTheme = True
    mnuLight.Checked = False
    mnuDark.Checked = True

    'these buttons must be styled as they are graphical buttons with images that conform to a classic theme
    
    btnArrowDown.Picture = LoadPicture(App.path & "\arrowDown.gif")
    btnMapPrev.Picture = LoadPicture(App.path & "\leftArrow.jpg")
    btnMapNext.Picture = LoadPicture(App.path & "\rightArrow.jpg")
    btnArrowUp.Picture = LoadPicture(App.path & "\arrowUp.jpg")
    
    ' RGB(212, 208, 199) is the background colour used by the classic theme
    
    Me.BackColor = RGB(212, 208, 199)
    ' a method of looping through all the controls that require reversion of any background colouring
    For Each Ctrl In rDIconConfigForm.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(212, 208, 199)
        End If
    Next
    
    ' these elements are normal elements that should have their styling removed on a classic theme
    
    ' we don't want all pictureboxes to be themed, only this one
    picPreview.BackColor = RGB(212, 208, 199)
    
    ' all other buttons go here, note we can colour buttons on VB6 succesfully without losing the theme,
    ' whilst VB.NET loses the bleeding theme deliberately and VB6 is superior in this respect.
    
    picRdThumbFrame.BackColor = RGB(212, 208, 199)
    btnRemoveFolder.BackColor = RGB(212, 208, 199)
    picCover.BackColor = RGB(212, 208, 199)
    back.BackColor = RGB(212, 208, 199)
    sliPreviewSize.BackColor = RGB(212, 208, 199)
    
    ' on NT6 plus using the MSCOMCTL slider with the lighter default theme, the slider
    ' fails to pick up the new theme colour fully
    ' the following lines triggers a partial colour change on the treeview that has no backcolor property
    ' this also causes a refresh of the preview pane - so don't remove it.
    ' I will have to create a new slider to overcome this - not yet tested the VB.NET version
    ' do not remove - essential

    'a = sliPreviewSize.Value
    'sliPreviewSize.Value = 1
    'sliPreviewSize.Value = a
    
    ' the above no longer required with Krool's replacement controls
    
   On Error GoTo 0
   Exit Sub

setThemeDark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeDark of Form rDIconConfigForm"

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
            Call setThemeColour
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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub menuRun_click()
    Dim testURL As String
    Dim validURL As Boolean
    Dim ans As VbMsgBoxResult
    
    On Error GoTo menuRun_click_Error
    If debugflg = 1 Then DebugPrint "%menuRun_click"

    lblRdIconNumber.Caption = Str(rdIconNumber) + 1
    lblRdIconNumber.ToolTipText = "This is Rocketdock icon number " & Str(rdIconNumber) + 1
    Call displayIconElement(rdIconNumber, picPreview, icoSizePreset, True)
    
    ' we signify that all changes have been lost so the "save this icon" will not appear when switching icons
    btnSave.Enabled = False ' this has to be done at the end
    btnCloseCancel.Caption = "Close"
    
    'now deal with the special extras
    ' contains "shutdown"
    If InStr(txtTarget.Text, "shutdown.exe") <> 0 Then
        MsgBox "I am sure you don't really want me to shutdown... test cancelled."
        Exit Sub
    End If
    
    ' is the target a URL?
    testURL = Left(txtTarget.Text, 3)
    If testURL = "htt" Or testURL = "www" Then
        validURL = True
        Call ShellExecute(hWnd, "Open", txtTarget.Text, vbNullString, vbNullString, 1)
    End If
                
    If txtTarget.Text = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
        '  my computer
        Call Shell("explorer.exe /e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNormalFocus)
        Exit Sub
    End If
    
    If txtTarget.Text = "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" Then
        ' network
        Call Shell("explorer.exe /e,::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", vbNormalFocus)
        Exit Sub
    End If
    
    If txtTarget.Text = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then
        ' network
        Call Shell("explorer.exe /e,::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNormalFocus)
        Exit Sub
    End If
    ' control panel
    If txtTarget.Text = "control" Or txtTarget.Text = "::{26EE0668-A00A-44D7-9371-BEB064C98683}\0" Then
        Call Shell("rundll32.exe shell32.dll,Control_RunDLL", vbNormalFocus)
        Exit Sub
    End If
    'printer
    If txtTarget.Text = "::{2227a280-3aea-1069-a2de-08002b30309d}" Then
        Call Shell("explorer.exe /e,::{2227a280-3aea-1069-a2de-08002b30309d}", vbNormalFocus)
        Exit Sub
    End If
    ' RD quit
    If txtTarget.Text = "[Quit]" Then
        MsgBox "I am sure you don't really want me to quit Rocketdock... test cancelled."
        Exit Sub
    End If
    ' RD settings
    If txtTarget.Text = "[Settings]" Then
        MsgBox "One cannot simply initiate the old settings screen from outside of Rocketdock as it is not a separate program that can be run from the command line or from Windows" & vbCr & "but as long as it looks like this [Settings] in the target field then it is good to go... test cancelled."
        Exit Sub
    End If
    ' program files folder
    If txtTarget.Text = "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}" Then
        Call Shell("explorer.exe /e,::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", vbNormalFocus)
        Exit Sub
    End If
     ' applications And features
    If txtTarget.Text = "appwiz.cpl" Then
        If debugflg = 1 Then DebugPrint "Shell " & "rundll32.exe shell32.dll,Control_RunDLL " & txtTarget.Text
        Call Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl", vbNormalFocus)
        Exit Sub
    End If
    ' recycle bin
    If txtTarget.Text = "::{645ff040-5081-101b-9f08-00aa002f954e}" Or txtTarget.Text = "[RecycleBin]" Then
        Call Shell("explorer.exe /e,::{645ff040-5081-101b-9f08-00aa002f954e}", vbNormalFocus)
        Exit Sub
    End If
    ' admin tools
    If txtTarget.Text = "compmgmt.msc" Then
        Call ShellExecute(hWnd, "Open", Environ$("windir") & "\SYSTEM32\COMPMGMT.MSC", 0&, 0&, 1)
        Exit Sub
    End If
    ' task manager
    If txtTarget.Text = "taskmgr" Then
        Call ShellExecute(hWnd, "Open", Environ$("windir") & "\SYSTEM32\taskmgr", 0&, 0&, 1)
        Exit Sub
    End If
    ' RocketdockEnhancedSettings.exe (the .NET version of this program)
    If GetFileNameFromPath(txtTarget.Text) = "RocketdockEnhancedSettings.exe" Then
        ans = MsgBox("It might not be a good idea to run the .NET and VB6 versions of the Rocketdock Utility at the same time. The two might conflict, and the results might not be positive. Are you sure you want me to?", vbYesNo)
        If ans = 6 Then
            Call ShellExecute(hWnd, "Open", txtTarget.Text, vbNullString, txtArguments.Text, 1)
        Else
            Exit Sub
        End If
    End If
    ' rocket1.exe (this program)
    If GetFileNameFromPath(txtTarget.Text) = "rocket1.exe" Then
        MsgBox "If you run the Rocketdock Utility, the first thing it does is to kill any existing instance, ie. this program you are running now - and I'm sure you don't really want me to do that... test cancelled."
        Exit Sub
    End If
    'anything else
    If FExists(txtTarget.Text) Then
        If debugflg = 1 Then DebugPrint "ShellExecute " & txtTarget.Text
        Call ShellExecute(hWnd, "Open", txtTarget.Text, vbNullString, txtArguments.Text, 1)
    ElseIf DirExists(txtTarget.Text) Then
        If debugflg = 1 Then DebugPrint "ShellExecute " & txtTarget.Text
        Call ShellExecute(hWnd, "Open", txtTarget.Text, vbNullString, txtArguments.Text, 1)
    ElseIf validURL = False Then
        MsgBox txtTarget.Text & " - That isn't valid as a target file or a folder so I can't run that to test it."
    End If

   On Error GoTo 0
   Exit Sub

menuRun_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuRun_click of Form rDIconConfigForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnArrowDown_Click
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub btnArrowDown_Click()
    Dim growBit As Integer
   
    On Error GoTo btnArrowDown_Click_Error
    If debugflg = 1 Then DebugPrint "%" & "btnArrowDown_Click"
    
    growBit = 670
            
    btnArrowDown.Visible = False
            
    If picRdThumbFrame.Visible = False Then
        'has to do this first as redrawing errors occur otherwise

        btnWorking.Visible = True
        Call busyStart
        'If picRdMap(0).Picture = 0 Then ' only recreate the map if the array is empty
        ' we used to check the .picture property but using lavolpe's 2nd method this proprty is not set.
        ' now we check for the tooltiptext which is only set when the image is populated.
        If picRdMap(0).ToolTipText = vbNullString Then ' only recreate the map if the array is empty
            Call populateRdMap(0) ' show the map from position zero
        End If
        
        framePreview.Top = 4545 + growBit
        frameProperties.Top = 4545 + growBit
        frameButtons.Top = 7680 + growBit
        rDIconConfigForm.Height = 9260 + growBit
                
        btnArrowUp.Visible = True
        picRdThumbFrame.Visible = True
        
        rdMapRefresh.Visible = True
        If rdIconMax > 16 Then
            rdMapHScroll.Visible = True
        End If
        
        rdMapHScroll.Max = theCount
                
        ' we signify that all changes have been lost
        btnSave.Enabled = False ' this has to be done at the end
        btnCloseCancel.Caption = "Close"
        
        btnWorking.Visible = False
  
        Call busyStop
        
        'write the visible state
        PutINISetting "Software\RocketDockSettings", "rdMapState", "visible", toolSettingsFile
    
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
Private Sub btnArrowDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo btnArrowDown_MouseUp_Error
   If debugflg = 1 Then Debug.Print "%btnArrowDown_MouseUp"

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

        framePreview.Top = 4545
        frameProperties.Top = 4545
        frameButtons.Top = 7680
        rDIconConfigForm.Height = 9305
        'rDIconConfigForm.dllFrame.Top = 7530
        
        btnArrowDown.Visible = True
        btnArrowUp.Visible = False
        picRdThumbFrame.Visible = False
        btnArrowDown.ToolTipText = "Show the Rocketdock Map"
        rdMapRefresh.Visible = False
        rdMapHScroll.Visible = False
        
        'write the hidden state
        PutINISetting "Software\RocketDockSettings", "rdMapState", "hidden", toolSettingsFile
        
   End If

   On Error GoTo 0
   Exit Sub

btnArrowUp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnArrowUp_Click of Form rDIconConfigForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : picRdMapSetFocus
' Author    : beededea
' Date      : 17/11/2019
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub picRdMapSetFocus()
    
    On Error GoTo picRdMapSetFocus_Error
    If debugflg = 1 Then Debug.Print "%picRdMapSetFocus"

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
   If debugflg = 1 Then Debug.Print "%picFrameThumbsSetFocus"

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




