VERSION 5.00
Begin VB.Form showAndTell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show and Tell..."
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "showAndTell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCloseBox 
      Caption         =   "Close"
      Height          =   495
      Left            =   10485
      TabIndex        =   1
      Top             =   5865
      Width           =   1335
   End
   Begin VB.TextBox Textbox 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5550
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "showAndTell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCloseBox_Click()
    Me.Hide
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 22/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

   On Error GoTo Form_Load_Error

    Textbox.Text = Textbox.Text + "What works in Steamydock?" & vbCrLf & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock works! It runs, it really feels like Rocketdock - and more." & vbCrLf
    Textbox.Text = Textbox.Text + "* Most configuration options now function as they did in Rocketdock." & vbCrLf
    Textbox.Text = Textbox.Text + "* It does a lot more than Rocketdock does. It has new functionality that you will discover." & vbCrLf
    Textbox.Text = Textbox.Text + "* Fully documented HTML help provided." & vbCrLf
    Textbox.Text = Textbox.Text + "* For compatibility, the dock operates and reads Rocketdock's obsolete registry/settings locations." & vbCrLf
    Textbox.Text = Textbox.Text + "* You can import your Rocketdock settings directly to SteamyDock" & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock reads/writes its own settings location that is compatible with Windows current standards in appdata." & vbCrLf
    Textbox.Text = Textbox.Text + "* It animates and displays the dock efficiently using less cpu than Rocketdock." & vbCrLf
    Textbox.Text = Textbox.Text + "* It uses almost zero cpu when idle." & vbCrLf
    Textbox.Text = Textbox.Text + "* The animation interval can be manually tweaked for different cpu types." & vbCrLf
    Textbox.Text = Textbox.Text + "* Drag and drop to the dock is functional." & vbCrLf
    Textbox.Text = Textbox.Text + "* The enhanced icon settings utility appears to be 99% complete" & vbCrLf
    Textbox.Text = Textbox.Text + "* There is now a huge and growing library of steampunk icons." & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock settings utility is 97% functional." & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock has comprehensive right-click menus to add new icon types." & vbCrLf
    Textbox.Text = Textbox.Text + "* The other right-click options all work as per Rocketdock." & vbCrLf
    Textbox.Text = Textbox.Text + "* Auto-hide implemented using Rocketdock's slide out method." & vbCrLf
    Textbox.Text = Textbox.Text + "* Auto-hide has additional options, a gentle fade and an instant fade." & vbCrLf
    Textbox.Text = Textbox.Text + "* Can run new instances of an application at will" & vbCrLf
    Textbox.Text = Textbox.Text + "* Can maximise and minimise windows, even to/from the systray." & vbCrLf
    Textbox.Text = Textbox.Text + "* Can alter the z-order of Windows, sending them to front and back." & vbCrLf
    Textbox.Text = Textbox.Text + "* New splash screen." & vbCrLf
    Textbox.Text = Textbox.Text + "* The icon quality settings now work as expected." & vbCrLf
    Textbox.Text = Textbox.Text + "* The icon and theme opacity settings now work as expected." & vbCrLf
    Textbox.Text = Textbox.Text + "* 28 new themes, some based upon the originals, some new." & vbCrLf
    Textbox.Text = Textbox.Text + "* Themes can be sized independently of the icons." & vbCrLf
    Textbox.Text = Textbox.Text + "* DIY theme creation is now much easier than in Rocketdock, based upon three simple images." & vbCrLf
    Textbox.Text = Textbox.Text + "* Additional confirmation dialog to give you a chance to say NO! to accidentally triggered functions." & vbCrLf
    Textbox.Text = Textbox.Text + "* Additional confirmation dialog for programs that provide no feedback." & vbCrLf
    Textbox.Text = Textbox.Text + "* Additional menu options for running as admin or opening an apps' default folder." & vbCrLf
    Textbox.Text = Textbox.Text + "* Most o/s commands (cpl, msc) now operate via an icon click, just as you would expect." & vbCrLf
    Textbox.Text = Textbox.Text + "* Lots of bugs fixed all the time just as they are discovered!" & vbCrLf
    Textbox.Text = Textbox.Text + "* Lots of new icons types being added all the time" & vbCrLf
    Textbox.Text = Textbox.Text + "* The running process indicators are now just as dynamic as Rocketdock's (running on timers to minimise CPU usage)." & vbCrLf
    Textbox.Text = Textbox.Text + "* Can open a existing instance of an app just as Rocketdock does." & vbCrLf
    Textbox.Text = Textbox.Text + "* Clicking on an icon for an minimised program causes the app to restore just as the Windows taskbar." & vbCrLf
    Textbox.Text = Textbox.Text + "* Very useful menu option for deleting any running instance." & vbCrLf
    Textbox.Text = Textbox.Text + "* You can send an app to the front or back." & vbCrLf
    Textbox.Text = Textbox.Text + "* Non Modal message boxes added meaning that an informational message does not always stop the dock dead." & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock will now hide on an extended timer if the user hits a predefined function key (F11)." & vbCrLf
    Textbox.Text = Textbox.Text + "* Adding a new icon to the dock via the menu now pops up the icon settings tool immediately following." & vbCrLf
    Textbox.Text = Textbox.Text + "* New confirmation dialogs can be added to an icon click, before or after the event." & vbCrLf
    Textbox.Text = Textbox.Text + "* SteamyDock responds to %userprofile% environment variables and to Windows CLSIDs." & vbCrLf
    Textbox.Text = Textbox.Text + "* Dragging from the dock now deletes the dragged icon after confirmation as per Rocketdock" & vbCrLf
    Textbox.Text = Textbox.Text + "* Dragging within the dock to re-order the icons is now operational" & vbCrLf
    Textbox.Text = Textbox.Text + "* Auto bulk generation of dock items in iconSettings is progressing..." & vbCrLf
    Textbox.Text = Textbox.Text + "* Automatic discovery of an relevant steampunk icon using an application compatibility list" & vbCrLf
    Textbox.Text = Textbox.Text + "* The dock icon bounce effect can now use an easeIN function" & vbCrLf
    Textbox.Text = Textbox.Text + "* Added quick launch functionality to run an app more quickly half way through the bounce animation." & vbCrLf
    Textbox.Text = Textbox.Text + "* Added dock automatic hiding for apps such as games that require full screen access" & vbCrLf
    Textbox.Text = Textbox.Text + "* Steamydock can fire up a secondary program if required" & vbCrLf
    Textbox.Text = Textbox.Text + "* SteamyDock can now auto-hide the dock after running a app" & vbCrLf
    Textbox.Text = Textbox.Text + "* Has a menu option to clone current item" & vbCrLf
    Textbox.Text = Textbox.Text + "* Adds a sound option when initiating an icon click." & vbCrLf
    Textbox.Text = Textbox.Text + "* Addition of new administrative icon choices." & vbCrLf
    Textbox.Text = Textbox.Text + "* Icons can be disabled to appear semi-transparent in the dock." & vbCrLf
    Textbox.Text = Textbox.Text + "* SteamyDock tests whether dock items point to valid programs, a red 'x' marks an invalid link." & vbCrLf
    Textbox.Text = Textbox.Text + "* SteamyDock can terminate any other chosen program prior to running the assigned app." & vbCrLf
    Textbox.Text = Textbox.Text + "* Allows an appplication to run elevated as an administrator directly from the dock." & vbCrLf
    Textbox.Text = Textbox.Text + "* Allows an appplication to run elevated as an administrator as an assigned characteristic of the icon." & vbCrLf
    Textbox.Text = Textbox.Text + "* Adds running process indicators above open Explorer windows." & vbCrLf
    Textbox.Text = Textbox.Text + "* Allows a restart of SteamyDock via the dock menu." & vbCrLf
    Textbox.Text = Textbox.Text + "* Dynamically handles monitor resolution changes repositioning to the bottom of the screen." & vbCrLf
    Textbox.Text = Textbox.Text + "* " & vbCrLf
    
    Textbox.Text = Textbox.Text + "* " & vbCrLf & vbCrLf
    Textbox.Text = Textbox.Text + "Things SteamyDock can't yet do:." & vbCrLf & vbCrLf
    Textbox.Text = Textbox.Text + "  " & vbCrLf
    Textbox.Text = Textbox.Text + "* Animations are missing on icon deletion, these will be done in the future." & vbCrLf
    Textbox.Text = Textbox.Text + "* Cannot yet extract embedded icons from EXEs and display them with transparent backgrounds - this is tricky as VB6 has no native PNG support. WIP," & vbCrLf
    Textbox.Text = Textbox.Text + "* Right/left dock positions are not yet implemented." & vbCrLf
    Textbox.Text = Textbox.Text + "* The additional right click context menu is not implemented as Steamydock already has a full and useful context menu of its own." & vbCrLf
    Textbox.Text = Textbox.Text + "* Only one dock animation type currently implemented - this will be implemented." & vbCrLf
    Textbox.Text = Textbox.Text + "* SteamyDock is not aware of any other language except for English. Fairly sure this will never happen! Too many languages and too much text to translate." & vbCrLf
    Textbox.Text = Textbox.Text + "* Sizing a 256x256 pixel icons on small laptop type screens generates an error." & vbCrLf
    Textbox.Text = Textbox.Text + "* Rocketdock's animations are still just - nicer. I am trying to fix this." & vbCrLf
    Textbox.Text = Textbox.Text + "* Cannot yet display on a second monitor but I'm working on it." & vbCrLf
    Textbox.Text = Textbox.Text + "* Cannot use Rocketdock's own themed backgrounds (but it can show its own similar themes)." & vbCrLf
    Textbox.Text = Textbox.Text + "* Cannot use docklets. These are obsolete and undocumented and so steamydock will not support them." & vbCrLf
    Textbox.Text = Textbox.Text + "* " & vbCrLf

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form showAndTell"

End Sub
