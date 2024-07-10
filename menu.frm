VERSION 5.00
Begin VB.Form menuForm 
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu mnuMainMenu 
      Caption         =   "mainmenu"
      Begin VB.Menu mnuRunApp 
         Caption         =   "Run this App"
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Run this App as Administrator"
      End
      Begin VB.Menu mnuRunNewApp 
         Caption         =   "Run New Instance of this App"
      End
      Begin VB.Menu mnuRunNewAppAsAdmin 
         Caption         =   "Run New Instance of this App As Admin."
      End
      Begin VB.Menu mnuBlank5 
         Caption         =   "______________________"
      End
      Begin VB.Menu mnuCloseApp 
         Caption         =   "Close Running Instances of this App"
      End
      Begin VB.Menu mnuFocusApp 
         Caption         =   "Bring Application to Front"
      End
      Begin VB.Menu mnuBackApp 
         Caption         =   "Send Application to Back"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "______________________"
      End
      Begin VB.Menu mnuIconSettings 
         Caption         =   "Edit Icon Properties"
      End
      Begin VB.Menu mnuDeleteIcon 
         Caption         =   "Delete Icon"
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "Add Icon"
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
            Caption         =   "Add My Videos"
         End
         Begin VB.Menu mnuAddShutdown 
            Caption         =   "Add Shutdown"
         End
         Begin VB.Menu mnuAddHibernate 
            Caption         =   "Add Hibernate"
         End
         Begin VB.Menu mnuAddReboot 
            Caption         =   "Add Reboot"
         End
         Begin VB.Menu mnuAddSleep 
            Caption         =   "Add Sleep"
         End
         Begin VB.Menu mnuAddLogOut 
            Caption         =   "Add Log Out"
         End
         Begin VB.Menu mnuAddLockWorkstation 
            Caption         =   "Add Lock Workstation"
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
         Begin VB.Menu mnuAddEnhanced 
            Caption         =   "Add Enhanced Icon Settings"
         End
         Begin VB.Menu mnuAddCache 
            Caption         =   "Add Clear Cache"
         End
         Begin VB.Menu mnuAddQuit 
            Caption         =   "Add Dock Quit"
         End
      End
      Begin VB.Menu mnuCloneIcon 
         Caption         =   "Clone Current Icon"
      End
      Begin VB.Menu mnuDisableIcon 
         Caption         =   "Disable This Icon"
      End
      Begin VB.Menu mnuApplicationFolder 
         Caption         =   "Open App Folder in Windows Explorer"
      End
      Begin VB.Menu mnublnk 
         Caption         =   "______________________"
      End
      Begin VB.Menu mnuDockSettings 
         Caption         =   "Dock Settings"
      End
      Begin VB.Menu mnuScreenPosition 
         Caption         =   "Screen Position"
         Begin VB.Menu mnuTop 
            Caption         =   "Top"
         End
         Begin VB.Menu mnuBottom 
            Caption         =   "Bottom"
         End
         Begin VB.Menu mnuLeft 
            Caption         =   "Left"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRight 
            Caption         =   "Right"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuAutoHide 
         Caption         =   "Auto Hide"
      End
      Begin VB.Menu mnuHideTwenty 
         Caption         =   "Hide for the next 10 minutes"
      End
      Begin VB.Menu mnuLockIcons 
         Caption         =   "Disable Drag/Drop and Icon Deletion"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "______________________"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "Other"
         Begin VB.Menu mnuAbout 
            Caption         =   "About this utility"
         End
         Begin VB.Menu mnuSplash 
            Caption         =   "Show the Splash Screen"
         End
         Begin VB.Menu mnuShowTell 
            Caption         =   "Show and Tell About SteamyDock"
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
         Begin VB.Menu mnuLicence 
            Caption         =   "Display Licence Agreement"
         End
         Begin VB.Menu mnuseparator1 
            Caption         =   ""
         End
         Begin VB.Menu mnuDevOptions 
            Caption         =   "Developer Options"
            Begin VB.Menu mnuAppFolder 
               Caption         =   "Reveal Program Location in Windows Explorer"
            End
            Begin VB.Menu mnuEditWidget 
               Caption         =   "Edit Program Using..."
            End
            Begin VB.Menu mnuDebug 
               Caption         =   "Turn Debugging ON"
            End
         End
      End
      Begin VB.Menu menuRestart 
         Caption         =   "Restart Steamydock"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit Steamydock"
      End
   End
End
Attribute VB_Name = "menuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' .01 DAEB 24/01/2021 menu.frm modified to handle the new timer name
' .02 DAEB 24/01/2021 menu.frm added disable of the autoFadeInTimer during menu operation
' .03 DAEB 02/02/2021 menu.frm Added menu option to clear the cache - mnuAddCache_Click
' .04 DAEB 03/03/2021 menu.frm New lose focus menu option
' .05 DAEB 03/03/2021 menu.frm To support new receive focus menu option
' .06 DAEB 05/03/2021 menu.frm Simplified the boolean checks and removed the cannot kill message
' .07 DAEB 07/03/2021 menu.frm Menu option to add a "my Videos" utility dock entry
' .08 DAEB 07/03/2021 menu.frm Menu option to add a "my pictures" utility dock entry
' .09 DAEB 07/03/2021 menu.frm Menu option to add a "my documents" utility dock entry
' .10 DAEB 07/03/2021 menu.frm Menu option to add a "my music" utility dock entry
' .11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
' .12 DAEB 08/04/2021 menu.frm made public so that it can be called by another routine in the dock frmMain.frm
' .13 DAEB 01/04/2021 menu.frm post addicon tasks, adding an icon now calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
' .14 DAEB 01/04/2021 menu.frm made public so that it can be called by another routine in the dock frmMain.frm
' .15 DAEB 01/04/2021 menu.frm make changes for running in the IDE
' .16 DAEB 17/11/2020 menu.frm Replaced all occurrences of rocket1.exe with iconsettings.exe
' .17 DAEB 05/05/2021 menu.frm cause the docksettings utility to reopen if it has already been initiated

Option Explicit

'Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : The main dock won't take a menu when using GDI so we have a separate form for the menu
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Width = 1  ' the menu form is made as small as possible and moved off screen so that it does not show anywhere on the
    Me.Height = 1 ' screen, the menu appearing at the cursor point when it is told to do so by the dock form mousedown.

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : menuRestart_Click
' Author    : beededea
' Date      : 11/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub menuRestart_Click()
    
   On Error GoTo 0
   
   Call restartSteamydock
   
   Exit Sub

menuRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure menuRestart_Click of Form menuForm"
End Sub





' .03 02/02/2021 DAEB Added menu option to clear the cache - mnuAddCache_Click
'---------------------------------------------------------------------------------------
' Procedure : mnuAddCache_Click
' Author    : beededea
' Date      : 02/02/2021
' Purpose   : Add menu option to clear the cache
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddCache_Click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddCache_Click_Error
    
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\recyclebin-full.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)

        Call insertNewIconDataIntoCurrentPosition(iconImage, "Clear Cache", "C:\WINDOWS\system32\rundll32.exe", "advapi32.dll , ProcessIdleTasks", "%windir%", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Clear Cache")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Clear Cache image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        'MsgBox "Unable to add Clear Cache image as it does not exist"
    End If
   On Error GoTo 0
   Exit Sub

mnuAddCache_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddCache_Click of Form menuForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : mnuAddEventViewer_Click
' Author    : beededea
' Date      : 24/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddEventViewer_Click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddEventViewer_Click_Error

    ' check the icon exists
    'If debugflg = 1 Then debugLog "%" & "mnuAddDevMgmt_Click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\event-viewer(CEventVwr.msc).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Event Viewer", "eventvwr.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Event Viewer")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Event Viewer image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If
    

    On Error GoTo 0
    Exit Sub

mnuAddEventViewer_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddEventViewer_Click of Form menuForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddMyDocuments_Click
' Author    : beededea
' Date      : 07/03/2021
' Purpose   : ' .09 DAEB 07/03/2021 menu.frm Added menu option to add a "my Documents" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyDocuments_Click()
'
    Dim iconImage As String
    Dim iconFileName As String
    
    ' initialise the vars above
    
    iconImage = vbNullString
    iconFileName = vbNullString
    
    On Error GoTo mnuAddMyDocuments_Click_Error

    'If debugflg = 1 Then debugLog "%mnuAddMyComputer_click"
    
    ' check the icon exist
    iconFileName = App.Path & "\iconSettings\my collection" & "\folder-closed.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
       
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "My Documents", "::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "My Documents")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my Documents image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add my Documents image as it does not exist"
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
' Purpose   : .10 DAEB 07/03/2021 menu.frm Added menu option to add a "my Music" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyMusic_Click()
'
    Dim iconImage As String
    Dim iconFileName As String
    Dim userprof As String
    
    ' initialise the vars above
    
    iconImage = vbNullString
    iconFileName = vbNullString
    userprof = vbNullString
    
    ' check the icon exists
    On Error GoTo mnuAddMyMusic_Click_Error

    'If debugflg = 1 Then debugLog "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\music.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If

    userprof = Environ$("USERPROFILE")
    
    If fFExists(iconImage) Then

        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        ' using the Special CLSID for the video folder this, in fact resolves to the my documents folder and not the video folder below.
        'Call insertNewIconDataIntoCurrentPosition(iconImage, "My Music", "::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", vbNullString, vbNullString, vbNullString, vbNullString)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "My Music", userprof & "\Documents\Music", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "My Music")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my Music image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '                MsgBox "Unable to add my Music image as it does not exist"
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
' Purpose   : .08 DAEB 07/03/2021 menu.frm Added menu option to add a "my Pictures" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyPictures_Click()
'
    Dim iconImage As String
    Dim iconFileName As String
    Dim userprof As String
    
    ' initialise the vars above
    
    iconImage = vbNullString
    iconFileName = vbNullString
    userprof = vbNullString
    
    ' check the icon exists
    On Error GoTo mnuAddMyPictures_Click_Error

    'If debugflg = 1 Then debugLog "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\pictures.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
       
    userprof = Environ$("USERPROFILE")

    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "My Pictures", userprof & "\Documents\Pictures", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "My Pictures")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my Pictures image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add my Pictures image as it does not exist"
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
' Purpose   : .07 DAEB 07/03/2021 menu.frm Added menu option to add a "my Videos" utility dock entry
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddMyVideos_Click()

    Dim iconImage As String
    Dim iconFileName As String
    Dim userprof As String
    
    ' initialise the vars above
    
    iconImage = vbNullString
    iconFileName = vbNullString
    userprof = vbNullString
        
    ' check the icon exists
    On Error GoTo mnuAddMyVideos_Click_Error

    'If debugflg = 1 Then debugLog "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\video-folder.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
           
    userprof = Environ$("USERPROFILE")
       
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        'Call insertNewIconDataIntoCurrentPosition(iconImage, "My Videos", "::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "My Videos", userprof & "\Documents\Videos", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "My Videos")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my Videos image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add my Videos image as it does not exist"
    End If
        

   On Error GoTo 0
   Exit Sub

mnuAddMyVideos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddMyVideos_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddPerfMon_Click
' Author    : beededea
' Date      : 24/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddPerfMon_Click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddPerfMon_Click_Error

    ' check the icon exists
    'If debugflg = 1 Then debugLog "%" & "mnuAddDevMgmt_Click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\perfmon.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Performance Monitor", "perfmon.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Performance Monitor")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Performance Monitor image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If

    On Error GoTo 0
    Exit Sub

mnuAddPerfMon_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddPerfMon_Click of Form menuForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddProgram_Click
' Author    : beededea
' Date      : 12/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddProgram_Click()
    ' the file dialog would not display when the code for the dialog was under the docl_form
    ' this may be because the dock_form is not visible at any time. Moving the file dialog form to the
    ' main dock form caused the dialog to display.
    
    Call addProgramDLLorEXE
    
   On Error GoTo 0
   Exit Sub

mnuAddProgram_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddProgram_Click of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAddServices_Click
' Author    : beededea
' Date      : 24/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddServices_Click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddServices_Click_Error

    ' check the icon exists
    'If debugflg = 1 Then debugLog "%" & "mnuAddDevMgmt_Click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Services Management", "services.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Services Management")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Services Management image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If

    On Error GoTo 0
    Exit Sub

mnuAddServices_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddServices_Click of Form menuForm"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddSleep_Click
' Author    : beededea
' Date      : 17/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddSleep_Click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddSleep_Click_Error
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\sleep.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
           
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Sleep", "C:\Windows\System32\RUNDLL32.exe", "powrprof.dll,SetSuspendState 0,1,0", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Sleep")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add sleep image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add sleep image as it does not exist"
    End If

    On Error GoTo 0
    Exit Sub

mnuAddSleep_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddSleep_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddTaskSched_Click
' Author    : beededea
' Date      : 24/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddTaskSched_Click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddTaskSched_Click_Error

    ' check the icon exists
    'If debugflg = 1 Then debugLog "%" & "mnuAddDevMgmt_Click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\glass-clipboard.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Task Scheduler", "taskschd.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Task Scheduler")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Task Scheduler image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If

    

    On Error GoTo 0
    Exit Sub

mnuAddTaskSched_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddTaskSched_Click of Form menuForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAdmin_Click
' Author    : beededea
' Date      : 10/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAdmin_Click()
    ' we already have the iconIndex set, so it is a matter of just triggering a run timer to run the command
'    dock.animateTimer.Enabled = True

    On Error GoTo mnuAdmin_Click_Error
    
    'Call readIconData(selectedIconIndex)
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", selectedIconIndex, dockSettingsFile
        
    

    Call dock.fMouseUp(1) ' performs the equivalent of a 'left' click on the dock

   On Error GoTo 0
   Exit Sub

mnuAdmin_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAdmin_Click of Form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAppFolder_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAppFolder_Click()
    Dim folderPath As String: folderPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuAppFolder_Click_Error

    folderPath = App.Path
    If fDirExists(folderPath) Then ' if it is a folder already

        execStatus = ShellExecute(Me.hwnd, "open", folderPath, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        MsgBox "Having a bit of a problem opening a folder for this widget - " & folderPath & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
        'MessageBox Me.hWnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuAppFolder_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAppFolder_Click of Form menuForm"

End Sub

Private Sub mnuApplicationFolder_Click()
    Dim folderPath As String: folderPath = vbNullString
    Dim execStatus As Long: execStatus = 0
    
    'Call readIconData(selectedIconIndex)
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", selectedIconIndex, dockSettingsFile
    
    If fDirExists(sCommand) Then ' if it is a folder already
        'If debugflg = 1 Then debugLog "ShellExecute " & sCommand
        'Call ShellExecute(hwnd, "open", sCommand, sArguments, vbNullString, 1)
        execStatus = ShellExecute(hwnd, "open", sCommand, sArguments, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
    Else
        'obtain the folder from the scommand
        folderPath = getFolderNameFromPath(sCommand)  ' extract the default folder from the batch full path
        If fDirExists(folderPath) Then
            'If debugflg = 1 Then debugLog "ShellExecute " & sCommand
            execStatus = ShellExecute(hwnd, "open", folderPath, sArguments, vbNullString, 1)
            If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
        Else
            'if that fails try and obtain the folder from the Working Directory
            If fDirExists(sWorkingDirectory) Then
                execStatus = ShellExecute(hwnd, "open", sWorkingDirectory, sArguments, vbNullString, 1)
                If execStatus <= 32 Then MsgBox "Attempt to open folder failed."
            Else
                ' if that fails, spit out an error.
                '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
                MessageBox Me.hwnd, "Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
                'MsgBox ("Having a bit of a problem opening a folder for that command - " & sCommand & " It doesn't seem to have a valid working directory set.")
            End If
        End If
    End If

End Sub



    
'---------------------------------------------------------------------------------------
' Procedure : mnuAutoHide_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAutoHide_Click()
   On Error GoTo mnuAutoHide_Click_Error

    If mnuAutoHide.Checked = False Then
        mnuAutoHide.Checked = True
        rDAutoHide = "1"
    Else
        mnuAutoHide.Checked = False
        rDAutoHide = "0"
    End If
    
'    If rDGeneralWriteConfig = "True" Then ' the 3rd option, steamydock compatibility
        PutINISetting "Software\SteamyDock\DockSettings", "AutoHide", rDAutoHide, dockSettingsFile
        
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
        PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", "9999", dockSettingsFile

'    Else ' rocketdock compatibility
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            PutINISetting "Software\RocketDock", "AutoHide", rDAutoHide, origSettingsFile
'        Else
'            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHide", rDAutoHide)
'        End If
'    End If
    
    If mnuAutoHide.Checked = True Then
        Call dock.drawSmallStaticIcons 'here deanie ' here 'also calls this when the autohide timer has done its job
        
        dock.autoHideChecker.Enabled = True
    Else
        dock.autoHideChecker.Enabled = False
        dock.autoFadeOutTimer.Enabled = False ' .01 24/01/2021 DAEB modified to handle the new timer name
        dock.autoFadeInTimer.Enabled = False  ' .02 24/01/2021 DAEB added disable of the autoFadeInTimer during menu operation
        dock.autoSlideOutTimer.Enabled = False
        dock.autoSlideInTimer.Enabled = False
        Call dock.drawSmallStaticIcons 'here deanie ' here

    End If

    
   On Error GoTo 0
   Exit Sub

mnuAutoHide_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAutoHide_Click of Form menuForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuBackApp_Click
' Author    : beededea
' Date      : 03/03/2021
' Purpose   : .04 DAEB 03/03/2021 menu.frm New lose focus menu option
'---------------------------------------------------------------------------------------
'
Private Sub mnuBackApp_Click()

   On Error GoTo mnuBackApp_Click_Error

    If userLevel <> "runas" Then userLevel = "open"
    Call dock.runCommand("back", "") ' added new parameter to allow override .68

   On Error GoTo 0
   Exit Sub

mnuBackApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuBackApp_Click of Form menuForm"

End Sub


Private Sub mnublank1_Click()
    Call mnuAbout_Click
End Sub

Private Sub mnuBlank2_Click()
    Call mnuIconSettings_Click_Event
End Sub



Private Sub mnublnk_Click()
    Call mnuDockSettings_Click
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuBottom_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuBottom_Click()
   On Error GoTo mnuBottom_Click_Error

    menuForm.mnuTop.Checked = False
    menuForm.mnuBottom.Checked = False
    menuForm.mnuLeft.Checked = False
    menuForm.mnuRight.Checked = False

    rDSide = vbBottom
    menuForm.mnuBottom.Checked = True
    dockPosition = vbBottom
    
    PutINISetting "Software\SteamyDock\DockSettings", "Side", rDSide, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", "9999", dockSettingsFile ' no icon changed

    Call dock.drawSmallStaticIcons

   On Error GoTo 0
   Exit Sub

mnuBottom_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuBottom_Click of Form menuForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCloseApp_Click
' Author    : beededea
' Date      : 28/07/2020
' Purpose   : Closes any currently running instance of this process or explorer window.
'---------------------------------------------------------------------------------------
'
Private Sub mnuCloseApp_Click()
    On Error GoTo mnuCloseApp_Click_Error
    
    Dim NameProcess As String: NameProcess = vbNullString
    
    NameProcess = sCommandArray(selectedIconIndex)
    
    ' if the open application is an Explorer window then kill it this way
    If explorerCheckArray(selectedIconIndex) = True Then
        Call CloseExplorerWindowByPath(NameProcess)
        explorerCheckArray(selectedIconIndex) = False
        initiatedExplorerArray(selectedIconIndex) = vbNullString ' removes the entry from the array that we test regularly so it isn't caught again
        Exit Sub
    Else
        ' if the open application is a standard binary then kill it this way
        If checkAndKillPutWindowBehind(NameProcess, True, True) = True Then ' .06 DAEB 05/03/2021 menu.frm Simplified the boolean checks and removed the cannot kill message
            Sleep 200 ' this ESSENTIAL small delay is required as it may take a moment or two for the system list to be updated.
            If IsRunning(NameProcess, vbNull) = False Then ' .06 DAEB 05/03/2021 menu.frm Simplified the boolean checks and removed the cannot kill message
                processCheckArray(selectedIconIndex) = False ' remove the entry from the cog array
                initiatedProcessArray(selectedIconIndex) = vbNullString ' removes the entry from the array that we test regularly so it isn't caught again
            Else
                ' .06 DAEB 05/03/2021 menu.frm Simplified the boolean checks and removed the cannot kill message
                ' sometimes the target process does not die in time and this message can be generated, I could drop this whole wait into a timer but it still would not handle
                ' the indeterminate time that processes can take to die dependant upon cpu load and delays at the time.
                'MsgBox ("Cannot kill this process - " & NameProcess)
            End If
        End If
    End If
   On Error GoTo 0
   Exit Sub

mnuCloseApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCloseApp_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDeleteIcon_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDeleteIcon_Click()
    
    Call deleteThisIcon ' .09 DAEB 30/04/2021 mdlMain.bas deleteThisIcon created by extracting from the menu form so it can be used elsewhere
    
   On Error GoTo 0
   Exit Sub

mnuDeleteIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDeleteIcon_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuCloneIcon_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuCloneIcon_Click()
    dock.Refresh
    
    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", selectedIconIndex, dockSettingsFile

    Call insertNewIconDataIntoCurrentPosition(sFilename, sTitle, sCommand, sArguments, sWorkingDirectory, sShowCmd, sOpenRunning, sIsSeparator, sDockletFile, sUseContext, sUseDialog, sUseDialogAfter, sQuickLaunch, sDisabled)
    Call menuForm.addImageToDictionaryAndCheckForRunningProcess(sFilename, sTitle)

   On Error GoTo 0
   Exit Sub

mnuCloneIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCloneIcon_Click of Form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuDisableIcon_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDisableIcon_Click()
    Dim largeKey As String: largeKey = vbNullString
    Dim smallKey As String: smallKey = vbNullString
    Dim partialStringKey As String: partialStringKey = vbNullString
    Dim imageOpacity As Integer: imageOpacity = 0
    
    dock.Refresh
        
    Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", selectedIconIndex, dockSettingsFile)
    
    If sDisabled = "1" Then
        sDisabled = "0"
        menuForm.mnuDisableIcon.Caption = "Disable This Icon"
        menuForm.mnuDisableIcon.Checked = False
        
        disabledArray(selectedIconIndex) = 0
    Else
        sDisabled = "1"
        menuForm.mnuDisableIcon.Caption = "Enable This Icon"
        menuForm.mnuDisableIcon.Checked = True
        
        ' tell SD that this icon has recently been disabled (during the lifetime of this program run)
        disabledArray(selectedIconIndex) = 1
    End If
    
    PutINISetting "Software\SteamyDock\IconSettings\Icons", selectedIconIndex & "-Disabled", sDisabled, dockSettingsFile
    
    ' triggers for telling the iconSettings tool that it needs to refresh
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", selectedIconIndex, dockSettingsFile
    
'   we cannot modify the current icon in the dictionary
'   instead, create a transparent version of the disabled icon on the fly based upon the filename
'   and store that as an additional icon at the end of the main collection and use the location array to point to that instead
        
'    create new keys for both transparent images in the collLargeIcons/collSmallIcons collections
    largeKey = dictionaryLocationArray(selectedIconIndex) & "TransparentImg" & LTrim$(Str$(iconSizeLargePxls))
    smallKey = dictionaryLocationArray(selectedIconIndex) & "TransparentImg" & LTrim$(Str$(iconSizeSmallPxls))
        
    partialStringKey = LTrim$(Str$(selectedIconIndex))
    imageOpacity = Val(rDIconOpacity)
    
    ' reduce the opacity
    imageOpacity = imageOpacity / 4
    
    ' load and cache transparent versions of the current images to the associated collections
    ' at small and large sizes.
    If fFExists(sFilename) Then
        resizeAndLoadImgToDict collLargeIcons, partialStringKey, fileNameArray(selectedIconIndex), sDisabled, (0), (0), (iconSizeLargePxls), (iconSizeLargePxls), largeKey, imageOpacity
        resizeAndLoadImgToDict collSmallIcons, partialStringKey, fileNameArray(selectedIconIndex), sDisabled, (0), (0), (iconSizeSmallPxls), (iconSizeSmallPxls), smallKey, imageOpacity
    End If

   On Error GoTo 0
   Exit Sub

mnuDisableIcon_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDisableIcon_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuFocusApp_Click
' Author    : beededea
' Date      : 03/03/2021
' Purpose   : .05 DAEB 03/03/2021 menu.frm New receive focus menu option
'---------------------------------------------------------------------------------------
'
Private Sub mnuFocusApp_Click()

    ' the runCommand is called directly when the app is already running to avoid delay, no bounce
   On Error GoTo mnuFocusApp_Click_Error

    If userLevel <> "runas" Then userLevel = "open"
    Call dock.runCommand("focus", "") ' added new parameter to allow override .68

   On Error GoTo 0
   Exit Sub

mnuFocusApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFocusApp_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHideTwenty_Click
' Author    : beededea
' Date      : 25/01/2021
' Purpose   : menu option to hide the dock, equivalent of F11
'---------------------------------------------------------------------------------------
'
Private Sub mnuHideTwenty_Click()

   On Error GoTo mnuHideTwenty_Click_Error

    If hideDockForNMinutes = True Then
        hideDockForNMinutes = False
    Else ' set the flag
        ' autohide immediately '
        Call dock.HideDockNow
        
        ' change the autohide code so that if the hidefor20 flag is set the dock does not come back when the mnouse enters the dock zone
        ' enable the timer that is running for 20 mins
    End If
    

   On Error GoTo 0
   Exit Sub

mnuHideTwenty_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHideTwenty_Click of Form menuForm"
    
End Sub



' the hidefor20 timer runs disables itself
' removes the hidefor20 flag
' shows the dock



'---------------------------------------------------------------------------------------
' Procedure : mnuLeft_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLeft_Click()

   On Error GoTo mnuLeft_Click_Error

    menuForm.mnuTop.Checked = False
    menuForm.mnuBottom.Checked = False
    menuForm.mnuLeft.Checked = False
    menuForm.mnuRight.Checked = False

    rDSide = vbLeft
    menuForm.mnuLeft.Checked = True
    dockPosition = vbLeft
    
'    If rDGeneralWriteConfig = "True" Then ' the 3rd option, steamydock compatibility
        PutINISetting "Software\SteamyDock\DockSettings", "Side", rDSide, dockSettingsFile
'    Else ' rocketdock compatibility
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            PutINISetting "Software\RocketDock", "Side", rDSide, origSettingsFile
'        Else
'            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", rDSide)
'        End If
'    End If
    
    Call dock.drawSmallStaticIcons 'here deanie ' here

   On Error GoTo 0
   Exit Sub

mnuLeft_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLeft_Click of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuLockIcons_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLockIcons_Click()
   On Error GoTo mnuLockIcons_Click_Error

    If mnuLockIcons.Checked = False Then
        mnuLockIcons.Checked = True
        rDLockIcons = 1
        mnuDeleteIcon.Enabled = False
    Else
        mnuLockIcons.Checked = False
        rDLockIcons = 0
        mnuDeleteIcon.Enabled = True
    End If
    
    
'    If rDGeneralWriteConfig = "True" Then ' the 3rd option, steamydock compatibility
        PutINISetting "Software\SteamyDock\DockSettings", "LockIcons", rDLockIcons, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", "9999", dockSettingsFile
        
'    Else ' rocketdock compatibility
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            PutINISetting "Software\RocketDock", "LockIcons", rDLockIcons, origSettingsFile
'        Else
'            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "LockIcons", rDLockIcons)
'        End If
'    End If

   On Error GoTo 0
   Exit Sub

mnuLockIcons_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLockIcons_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuQuit_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuQuit_Click()
    Dim frm As Form
    On Error GoTo mnuQuit_Click_Error
        
    Call dock.shutdwnGDI
        
'    For Each frm In Forms
'        Unload frm
'        Set frm = Nothing
'    Next
    

    
    '   If gdipFullScreenBitmap Then
    '        GdipReleaseDC gdipFullScreenBitmap, dcMemory
    '        GdipDeleteGraphics gdipFullScreenBitmap
    '    End If
    '    If iconBitmap Then GdipDisposeImage iconBitmap
    '    If lngGDI Then GdiplusShutdown lngGDI
        
    End
    
    

   On Error GoTo 0
   Exit Sub

mnuQuit_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuQuit_Click of Form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuIconSettings_Click_Event
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuIconSettings_Click() ' .14 DAEB 01/04/2021 menu.frm made public so that it can be called by another routine in the dock frmMain.frm
    Dim thisCommand As String: thisCommand = vbNullString
    Dim execStatus As Long: execStatus = 0
    Dim prg As String: prg = vbNullString
   
   On Error GoTo mnuIconSettings_Click_Error
   
    ' .15 DAEB 01/04/2021 menu.frm make changes for running in the IDE
    If Not InIDE Then
        thisCommand = App.Path & "\iconSettings\iconsettings.exe" ' .16 DAEB 17/11/2020 menu.frm Replaced all occurrences of rocket1.exe with iconsettings.exe
    Else
        prg = Environ$("ProgramFiles")
        thisCommand = prg & "\SteamyDock\iconSettings\iconsettings.exe"
    End If
    
    'If IsRunning(thisCommand, vbNull) = False Then
    
        If fFExists(thisCommand) Then
        
        ' code was added here to re-use the existing icon settings process if it was open already. However, the selectedIconIndex cannot currently be
        ' passed to an running process as there is no inter-process communication and it is required that we pass the selectedIconIndex to identify
        ' which icon to display in the utility. We can do that when starting a new process but not when re-using an existing one. So, the gentle opening
        ' of the icon settings tool will have to wait until it is all brought into one program.
        
    '        If userLevel <> "runas" Then userLevel = "open"
    '        Call dock.runCommand("focus", thisCommand)
        
            'If debugflg = 1 Then debugLog "ShellExecute " & sCommand
            If InStr(WindowsVer, "Windows XP") <> 0 Then
                execStatus = ShellExecute(Me.hwnd, "open", thisCommand, selectedIconIndex, vbNullString, 1)
                If execStatus <= 32 Then MsgBox "Attempt to open utility failed."
            Else
                execStatus = ShellExecute(hwnd, "open", thisCommand, selectedIconIndex, vbNullString, 1)
                If execStatus <= 32 Then MsgBox "Attempt to open utility failed."
            End If
        Else
             '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
             MessageBox Me.hwnd, thisCommand & " is missing", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
             '        MsgBox thisCommand & " is missing"
        End If
        
    'End If

   On Error GoTo 0
   Exit Sub

mnuIconSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuIconSettings_Click of Form menuForm"
End Sub

Public Sub mnuIconSettings_Click_Event()
    Dim thisCommand As String: thisCommand = vbNullString
    Dim execStatus As Long: execStatus = 0
    Dim prg As String: prg = vbNullString
      
    ' .15 DAEB 01/04/2021 menu.frm make changes for running in the IDE
    If Not InIDE Then
        thisCommand = App.Path & "\iconSettings\iconsettings.exe" ' .16 DAEB 17/11/2020 menu.frm Replaced all occurrences of rocket1.exe with iconsettings.exe
    Else
        prg = Environ$("ProgramFiles")
        thisCommand = prg & "\SteamyDock\iconSettings\iconsettings.exe"
    End If    ' .15 DAEB 01/04/2021 menu.frm make changes for running in the IDE
    If Not InIDE Then
        thisCommand = App.Path & "\iconSettings\iconsettings.exe" ' .16 DAEB 17/11/2020 menu.frm Replaced all occurrences of rocket1.exe with iconsettings.exe
    Else
        prg = Environ$("ProgramFiles")
        thisCommand = prg & "\SteamyDock\iconSettings\iconsettings.exe"
    End If
    
    If IsRunning(thisCommand, vbNull) = False Then
    
        If fFExists(thisCommand) Then
        
        ' code was added here to re-use the existing icon settings process if it was open already. However, the selectedIconIndex cannot currently be
        ' passed to an running process as there is no inter-process communication and it is required that we pass the selectedIconIndex to identify
        ' which icon to display in the utility. We can do that when starting a new process but not when re-using an existing one. So, the gentle opening
        ' of the icon settings tool will have to wait until it is all brought into one program.
        
    '        If userLevel <> "runas" Then userLevel = "open"
    '        Call dock.runCommand("focus", thisCommand)
        
            'If debugflg = 1 Then debugLog "ShellExecute " & sCommand
            If InStr(WindowsVer, "Windows XP") <> 0 Then
                execStatus = ShellExecute(Me.hwnd, "open", thisCommand, selectedIconIndex, vbNullString, 1)
                If execStatus <= 32 Then MsgBox "Attempt to open utility failed."
            Else
                execStatus = ShellExecute(hwnd, "open", thisCommand, selectedIconIndex, vbNullString, 1)
                If execStatus <= 32 Then MsgBox "Attempt to open utility failed."
            End If
        Else
             '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
             MessageBox Me.hwnd, thisCommand & " is missing", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
             '        MsgBox thisCommand & " is missing"
        End If
        
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuDockSettings_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub mnuDockSettings_Click()
    Dim thisCommand As String
    On Error GoTo mnuDockSettings_Click_Error

    thisCommand = App.Path & "\dockSettings\dockSettings.exe"
    
    If fFExists(thisCommand) Then
        'If debugflg = 1 Then debugLog "ShellExecute " & sCommand
                    
        ' .17 DAEB 05/05/2021 menu.frm cause the docksettings utility to reopen if it has already been initiated
        If userLevel <> "runas" Then userLevel = "open"
        Call dock.runCommand("focus", thisCommand)

'        If InStr(WindowsVer, "Windows XP") <> 0 Then ' XP doesn't like this runas
'            Call ShellExecute(hwnd, "open", thisCommand, vbNullString, vbNullString, 1)
'        Else
'            Call ShellExecute(hwnd, "runas", thisCommand, vbNullString, vbNullString, 1)
'        End If

    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, thisCommand & " is missing", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox thisCommand & " is missing"
    End If

   On Error GoTo 0
   Exit Sub

mnuDockSettings_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDockSettings_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRight_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRight_Click()
   On Error GoTo mnuRight_Click_Error

    menuForm.mnuTop.Checked = False
    menuForm.mnuBottom.Checked = False
    menuForm.mnuLeft.Checked = False
    menuForm.mnuRight.Checked = False

    rDSide = vbRight
    menuForm.mnuRight.Checked = True
    dockPosition = vbRight
    
    
'    If rDGeneralWriteConfig = "True" Then ' the 3rd option, steamydock compatibility
        PutINISetting "Software\SteamyDock\DockSettings", "Side", rDSide, dockSettingsFile
'    Else ' rocketdock compatibility
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            PutINISetting "Software\RocketDock", "Side", rDSide, origSettingsFile
'        Else
'            Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "Side", rDSide)
'        End If
'    End If
        
    Call dock.drawSmallStaticIcons 'here deanie ' here

   On Error GoTo 0
   Exit Sub

mnuRight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRight_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRunApp_Click
' Author    : beededea
' Date      : 30/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRunApp_Click()

   On Error GoTo mnuRunApp_Click_Error

    forceRunNewAppFlag = False
    Call dock.startRunTimer
    'dock.runTimer.Enabled = True

   On Error GoTo 0
   Exit Sub

mnuRunApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRunApp_Click of Form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuRunNewApp_Click
' Author    : beededea
' Date      : 03/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuRunNewApp_Click()

   On Error GoTo mnuRunNewApp_Click_Error

    forceRunNewAppFlag = True
    dock.runTimer.Enabled = True

   On Error GoTo 0
   Exit Sub

mnuRunNewApp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuRunNewApp_Click of Form menuForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuShowTell_Click
' Author    : beededea
' Date      : 30/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuShowTell_Click()
   On Error GoTo mnuShowTell_Click_Error

    showAndTell.Show

   On Error GoTo 0
   Exit Sub

mnuShowTell_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuShowTell_Click of Form menuForm"
End Sub

Private Sub mnuSplash_Click()
    splashForm.Show
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuTop_Click
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuTop_Click()

   On Error GoTo mnuTop_Click_Error

    menuForm.mnuTop.Checked = False
    menuForm.mnuBottom.Checked = False
    menuForm.mnuLeft.Checked = False
    menuForm.mnuRight.Checked = False

    rDSide = vbtop
    menuForm.mnuTop.Checked = True
    dockPosition = vbtop
    
    PutINISetting "Software\SteamyDock\DockSettings", "Side", rDSide, dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", "9999", dockSettingsFile
    
    Call dock.drawSmallStaticIcons ' here

   On Error GoTo 0
   Exit Sub

mnuTop_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuTop_Click of Form menuForm"

End Sub
  
'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click(index As Integer)
    Dim answer As VbMsgBoxResult

    On Error GoTo mnuCoffee_Click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuCoffee_Click"
    
    answer = MsgBox(" Help support the creation of more widgets like this, send us a beer! This button opens a browser window and connects to the Paypal donate page for this widget). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=info@lightquick.co.uk&currency_code=GBP&amount=2.50&return=&item_name=Donate%20a%20Beer", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of form menuForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuHelpPdf_click
' Author    : beededea
' Date      : 30/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuHelpPdf_click()
   Dim answer As VbMsgBoxResult

   On Error GoTo mnuHelpPdf_click_Error
   'If debugflg = 1 Then debugLog "%mnuHelpPdf_click"

    answer = MsgBox("This option opens a browser window and displays this tool's help. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        If fFExists(App.Path & "\help\SteamyDock.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\SteamyDock.html", vbNullString, App.Path, 1)
        Else
            MsgBox ("The help file - SteamyDock.html- is missing from the help folder.")
        End If
    End If

   On Error GoTo 0
   Exit Sub

mnuHelpPdf_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpPdf_click of form menuForm"

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
    'If debugflg = 1 Then debugLog "%" & "mnuFacebook_Click"

    answer = MsgBox("Visiting the Facebook chat page - this button opens a browser window and connects to our Facebook chat page. Proceed?", vbExclamation + vbYesNo)
    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.facebook.com/profile.php?id=100012278951649", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuFacebook_Click of form menuForm"
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
    'If debugflg = 1 Then debugLog "%" & "mnuLatest_Click"

    answer = MsgBox("Download latest version of the program - this button opens a browser window and connects to the widget download page where you can check and download the latest zipped file). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If


    On Error GoTo 0
    Exit Sub

mnuLatest_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLatest_Click of form menuForm"

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
    'If debugflg = 1 Then debugLog "%" & "mnuLicence_Click"
        
    Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
    licence.Show

    On Error GoTo 0
    Exit Sub

mnuLicence_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_Click of form menuForm"
    
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

    On Error GoTo mnuSupport_Click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuSupport_Click"

    answer = MsgBox("Visiting the support page - this button opens a browser window and connects to our contact us page where you can send us a support query or just have a chat). Proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/art/Quartermaster-VB6-Desktop-784624943", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of form menuForm"
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
       'If debugflg = 1 Then debugLog "%" & "mnuSweets_Click"
    
    
    answer = MsgBox(" Help support the creation of more widgets like this. Buy me a small item on my Amazon wishlist! This button opens a browser window and connects to my Amazon wish list page). Will you be kind and proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "http://www.amazon.co.uk/gp/registry/registry.html?ie=UTF8&id=A3OBFB6ZN4F7&type=wishlist", vbNullString, App.Path, 1)
    End If
    
    On Error GoTo 0
    Exit Sub

mnuSweets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSweets_Click of form menuForm"
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
    'If debugflg = 1 Then debugLog "%" & "mnuWidgets_Click"

    answer = MsgBox(" This button opens a browser window and connects to the Steampunk widgets page on my site. Do you wish to proceed?", vbExclamation + vbYesNo)

    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://www.deviantart.com/yereverluvinuncleber/gallery/59981269/yahoo-widgets", vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

mnuWidgets_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuWidgets_Click of form menuForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuDebug_Click
' Author    : beededea
' Date      : 26/08/2019
' Purpose   : Run the runtime debugging window exectuable
'---------------------------------------------------------------------------------------
'
Private Sub mnuDebug_Click()
    
    On Error GoTo mnuDebug_Click_Error
    'If debugflg = 1 Then Debug.Print "%mnuDebug_Click" '< must always be debug.print

    Call toggleDebugging

   On Error GoTo 0
   Exit Sub

mnuDebug_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDebug_Click of form menuForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click()
    
    On Error GoTo mnuAbout_Click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuAbout_Click"
          
     about.lblMajorVersion.Caption = App.Major
     about.lblMinorVersion.Caption = App.Minor
     about.lblRevisionNum.Caption = App.Revision
     
     about.Show
     
     If (about.windowState = 1) Then
         about.windowState = 0
     End If

    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of form menuForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : menuAddBlank_Click
' Author    : beededea
' Date      : 31/05/2019
' Purpose   : Menu option for the RD Map to add a blank picture item.
'---------------------------------------------------------------------------------------
'
Private Sub menuAddBlank_Click()
    Dim iconImage As String
    
    On Error GoTo menuAdd_Click_Error
    'If debugflg = 1 Then debugLog "%" & "menuAddBlank_Click"
          
    iconImage = App.Path & "\blank.png"
    ' when we arrive at the original position then add a blank item
    ' with the following blank characteristics
    ' App.path & "\iconSettings\Icons\help.png" ' the default Rocketdock filename for a blank item
    
    If fFExists(iconImage) Then
        ' general tool to add an icon
        Call insertNewIconDataIntoCurrentPosition(iconImage, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Separator")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add blank image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add blank image as it does not exist"
    End If
    
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
      'If debugflg = 1 Then debugLog "%" & "mnuAddShutdown_click"
   
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\shutdown.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
           
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Shutdown", "C:\Windows\System32\shutdown.exe", "/s /t 00 /f /i", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Shutdown")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add shutdown image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add shutdown image as it does not exist"
    End If
       
    On Error GoTo 0
   Exit Sub

mnuAddShutdown_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddShutdown_click of Form rDIconConfigForm"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddHibernate_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a shutdown icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddHibernate_click()
    Dim iconImage As String
    Dim iconFileName As String
    
   On Error GoTo mnuAddHibernate_click_Error
      'If debugflg = 1 Then debugLog "%" & "mnuAddHibernate_click"
   
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\shutdown.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
           
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Hibernate", "C:\Windows\System32\shutdown.exe", "/h", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Hibernate")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add hibernate image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add shutdown image as it does not exist"
    End If
       
    On Error GoTo 0
   Exit Sub

mnuAddHibernate_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddHibernate_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddReboot_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add a Reboot icon on an Icon map right click.
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddReboot_click()
    Dim iconImage As String
    Dim iconFileName As String
    
    On Error GoTo mnuAddReboot_click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuAddReboot_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\Reboot.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
           
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Reboot", "C:\Windows\System32\shutdown.exe", "/r /t 00", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "1", vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Reboot")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add Reboot image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add Reboot image as it does not exist"
    End If
       
    On Error GoTo 0
   Exit Sub

mnuAddReboot_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddReboot_click of Form rDIconConfigForm"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddLockWorkstation_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddLockWorkstation_click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddLockWorkstation_click_Error
    'If debugflg = 1 Then debugLog "%mnuAddLog_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\padlockLockWorkstation.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Lock Workstation", Environ$("windir") & "\system32\rundll32.exe", "user32.dll, LockWorkStation", "%windir%", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Lock Workstation")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add Lock Workstation image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add log off image as it does not exist"
    End If
    
    On Error GoTo 0
    Exit Sub

mnuAddLockWorkstation_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddLockWorkstation_click of Form rDIconConfigForm"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddLogOut_click
' Author    : beededea
' Date      : 18/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddLogOut_click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddLogOut_click_Error
    'If debugflg = 1 Then debugLog "%mnuAddLogOut_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\console-green-screen-logout.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
        
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Log Out", Environ$("windir") & "\system32\shutdown.exe", "/l", "%windir%", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Log Out")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add log off image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '        MsgBox "Unable to add log off image as it does not exist"
    End If
    
    On Error GoTo 0
    Exit Sub

mnuAddLogOut_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddLogOut_click of Form rDIconConfigForm"

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
    'If debugflg = 1 Then debugLog "%" & "mnuAddNetwork_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\big-globe(network).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    ' thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Network", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Network")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add network image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         '         MsgBox "Unable to add network image as it does not exist"
    End If
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddWorkgroup_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\big-globe(network).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "WorkGroup", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "WorkGroup")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add workgroup image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         'MsgBox "Unable to add log off image as it does not exist"
    End If
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddPrinters_click"
    
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\printer.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Printers", "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Printers")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add printers image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         'MsgBox "Unable to add printers image as it does not exist"
    End If
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddTask_click"
    
    iconFileName = App.Path & "\iconSettings\my collection" & "\task-manager(tskmgr).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    If Is64bit() Then
        sixtyFourBit = True
        ' if a 32 bit application on a 64bit o/s, regardless of the command, the o/s calls C:\Windows\SysWOW64\taskmgr.exe
        If fFExists(iconImage) Then
            '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
            Call insertNewIconDataIntoCurrentPosition(iconImage, "Task Manager", Environ$("windir") & "\SysWOW64\" & "taskmgr.exe", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
            Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Task Manager")
        
            ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
            If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
                Call menuForm.mnuIconSettings_Click_Event
            End If
        Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add Task Manager image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         'MsgBox "Unable to add Task Manager image as it does not exist"
        End If
    Else
        ' if a 32 bit application on a 32bit o/s, regardless of the o/s calls C:\Windows\System32\taskmgr.exe
        If fFExists(iconImage) Then
            '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
            Call insertNewIconDataIntoCurrentPosition(iconImage, "Task Manager", Environ$("windir") & "\System32\" & "taskmgr.exe", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
            Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Task Manager")
        Else
         '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
         MessageBox Me.hwnd, "Unable to add Task Manager image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
         'MsgBox "Unable to add Task Manager image as it does not exist"
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
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddControl_click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuAddControl_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\control-panel(control).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Control panel", "control.exe", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Control panel")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add control panel image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        ' MsgBox "Unable to add control panel image as it does not exist"
    End If
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
       'If debugflg = 1 Then debugLog "%" & "mnuAddPrograms_click"
    
    
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\programs and features.ico"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Programs and Features", "appwiz.cpl", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Programs and Features")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Program and Features image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '         MsgBox "Unable to add Program and Features image as it does not exist"
    End If
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
      'If debugflg = 1 Then debugLog "%" & "mnuAddDock_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\dock settings.ico"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)

    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Dock Settings", "[Settings]", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Dock Settings")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Dock Settings image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Dock Settings image as it does not exist"
    End If
    
    On Error GoTo 0
   Exit Sub

mnuAddDock_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDock_click of Form rDIconConfigForm"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddCompMgmt_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddCompMgmt_click()
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddCompMgmt_click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuAddCompMgmt_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Computer Management", "compmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Computer Management")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Computer Management image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If
   On Error GoTo 0
   Exit Sub

mnuAddCompMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddCompMgmt_click of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuAddDevMgmt_Click
' Author    : beededea
' Date      : 24/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDevMgmt_Click()
    Dim iconImage As String
    Dim iconFileName As String

    On Error GoTo mnuAddDevMgmt_Click_Error
    ' check the icon exists
    'If debugflg = 1 Then debugLog "%" & "mnuAddDevMgmt_Click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\Administrative Tools(compmgmt.msc).png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Device Management", "devmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Device Management")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Device Management image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If
    

    On Error GoTo 0
    Exit Sub

mnuAddDevMgmt_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDevMgmt_Click of Form menuForm"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAddDiscMgmt_click
' Author    : beededea
' Date      : 18/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAddDiscMgmt_click()
    Dim iconImage As String
    Dim iconFileName As String
    ' check the icon exists
    On Error GoTo mnuAddDiscMgmt_click_Error
    'If debugflg = 1 Then debugLog "%" & "mnuAddDiscMgmt_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\discMgmt.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Disc Management", "diskmgmt.msc", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Disc Management")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Disc Management image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Administration Tools image as it does not exist"
    End If
   On Error GoTo 0
   Exit Sub

mnuAddDiscMgmt_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDiscMgmt_click of Form rDIconConfigForm"
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddRecycle_click"
   
    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\recyclebin-full.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Recycle Bin", "::{645ff040-5081-101b-9f08-00aa002f954e}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Recycle Bin")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Recycle Bin image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Recycle Bin image as it does not exist"
    End If
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddQuit_click"
   
    iconFileName = App.Path & "\iconSettings\my collection" & "\quit.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
    
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Quit", "[Quit]", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Quit")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Quit image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Quit image as it does not exist"
    End If
           
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
    'If debugflg = 1 Then debugLog "%" & "mnuAddProgramFiles_click"
   
    iconFileName = App.Path & "\iconSettings\my collection" & "\hard-drive-indicator-D.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)

    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Program Files", "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Program Files")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add Program Files image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Program Files image as it does not exist"
    End If
    
   On Error GoTo 0
   Exit Sub

mnuAddProgramFiles_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddProgramFiles_click of Form rDIconConfigForm"
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
    'If debugflg = 1 Then debugLog "mnuAddSeparator_click"
           
    iconFileName = App.Path & "\separator.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If

    sIsSeparator = "1"
        
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "Separator", vbNullString, vbNullString, vbNullString, vbNullString, sIsSeparator, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Separator")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add separator image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add separator image as it does not exist"
    End If

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
   'If debugflg = 1 Then debugLog "%mnuaddFolder_click"

    dialogInitDir = App.Path 'start dir, might be "C:\" or so also

    getFolder = BrowseFolder(hwnd, dialogInitDir) ' show the dialog box to select a folder

    If fDirExists(getFolder) Then
    
        iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-dir.png"
    
        If fFExists(iconFileName) Then
            iconImage = iconFileName
        End If
            
        ' if no specific image found
        If iconImage = vbNullString Then
            iconImage = App.Path & "\nixietubelargeQ.png"
        End If
   
        If fFExists(iconImage) Then
            '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
            Call insertNewIconDataIntoCurrentPosition(iconImage, getFolder, getFolder, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
            Call addImageToDictionaryAndCheckForRunningProcess(iconImage, getFolder)
        
            ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
            If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
                Call menuForm.mnuIconSettings_Click_Event
            End If
        Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add folder image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '            MsgBox "Unable to add folder image as it does not exist"
        End If

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
   'If debugflg = 1 Then debugLog "%mnuAddMyComputer_click"

    iconFileName = App.Path & "\iconSettings\my collection" & "\my folder.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
       
    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        Call insertNewIconDataIntoCurrentPosition(iconImage, "My Computer", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "My Computer")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my computer  image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '         MsgBox "Unable to add my computer image as it does not exist"
    End If
        
        
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
    'If debugflg = 1 Then debugLog "%mnuAddEnhanced_click"

    ' check the icon exists
    iconFileName = App.Path & "\iconSettings\my collection" & "\rocketdockSettings.png"
    If fFExists(iconFileName) Then
        iconImage = iconFileName
    Else
        iconImage = App.Path & "\iconSettings\Icons\help.png"
    End If
    
    '[icons]
    '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)

    If fFExists(iconImage) Then
        '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
        ' .16 DAEB 17/11/2020 menu.frm Replaced all occurrences of rocket1.exe with iconsettings.exe

        Call insertNewIconDataIntoCurrentPosition(iconImage, "Enhanced Icon Settings", App.Path & "\iconsettings.exe", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        Call addImageToDictionaryAndCheckForRunningProcess(iconImage, "Enhanced Icon Settings")
        
        ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
        If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
            Call menuForm.mnuIconSettings_Click_Event
        End If
    Else
        '.11 DAEB 01/04/2021 menu.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hwnd, "Unable to add my Enhanced Icon Settings image as it does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Unable to add Enhanced Icon Settings image as it does not exist"
    End If
    
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
   
'    Dim dllPath As String
'    Dim dialogInitDir As String
'
'    Const x_MaxBuffer = 256
'
'    On Error GoTo mnuAddDocklet_click_Error
'    'If debugflg = 1 Then debugLog "%mnuAddDocklet_click"
'
'    ' set the default folder to the docklet folder under rocketdock
'    dialogInitDir = rdAppPath & "\docklets"
'
'    With x_OpenFilename
'    '    .hwndOwner = Me.hWnd
'      .hInstance = App.hInstance
'      .lpstrTitle = "Select a Rocketdock Docklet DLL"
'      .lpstrInitialDir = dialogInitDir
'
'      .lpstrFilter = "DLL Files" & vbNullChar & "*.dll" & vbNullChar & vbNullChar
'      .nFilterIndex = 2
'
'      .lpstrFile = String(x_MaxBuffer, 0)
'      .nMaxFile = x_MaxBuffer - 1
'      .lpstrFileTitle = .lpstrFile
'      .nMaxFileTitle = x_MaxBuffer - 1
'      .lStructSize = Len(x_OpenFilename)
'    End With
'
'    Dim retFileName As String
'    Dim retfileTitle As String
'    'Call f_GetOpenFileName(retFileName, retfileTitle)
'    'txtTarget.Text = retFileName
'    'lblName.Text = retfileTitle
'
'  If txtTarget.Text <> vbNullString Then
'    ' check the folder is valid docklet folder (beneath the docklets folder)
'    ' set it to the docklet image yet to be created
'    ' if it is a clock docklet use a temporary clock image just as RD does without hands?
'    ' if it is a weather docklet use a temporary weather image of my own making
'    ' if it is a recycling docklet use a temporary recycling image of my own making
'
'    ' set the icon to that used by the docklet, it a mere guess as we cannot read the docklet DLL at this stage
'    ' to determine what icon image it intends to use, later it writes to the 'other' settings.ini file in docklets
'    ' but that's of no use now.
'
'      If InStr(GetFileNameFromPath(txtTarget.Text), "Clock") > 0 Then
'        txtCurrentIcon.Text = rdAppPath & "\icons\clock.png"
'      ElseIf InStr(GetFileNameFromPath(txtTarget.Text), "recycle") > 0 Then
'        txtCurrentIcon.Text = App.path & "\iconSettings\my collection\recyclebin-full.png"
'      Else
'        txtCurrentIcon.Text = rdAppPath & "\iconSettings\icons\blank.png" ' has to be an icon of some sort
'      End If
'
'       '    thisFilename, thisTitle, thisCommand, thisArguments, thisWorkingDirectory)
'      Call insertNewIconDataIntoCurrentPosition(txtCurrentIcon.Text, "Docklet", vbNullString, vbNullString, vbNullString, txtTarget.Text, vbNullString)
'
'    ' disable the fields, only enable the target fields and use the target field as a temporary location for the docklet data
'
'      lblName.Enabled = False
'      txtCurrentIcon.Enabled = False
'
'      sDockletFile = txtTarget.Text
'      txtTarget.Enabled = True
'      btnTarget.Enabled = True
'
'      txtArguments.Enabled = False
'      txtStartIn.Enabled = False
'      comboRun.Enabled = False
'      comboOpenRunning.Enabled = False
'      checkPopupMenu.Enabled = False
'      btnSelectStart.Enabled = False
'    End If
    
    'triggerRdMapRefresh = True
        
   On Error GoTo 0
   Exit Sub

mnuAddDocklet_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAddDocklet_click of Form rDIconConfigForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : addImageToDictionaryAndCheckForRunningProcess
' Author    : beededeaand
' Date      : 02/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub addImageToDictionaryAndCheckForRunningProcess(newFileName As String, newName As String)

   On Error GoTo addImageToDictionaryAndCheckForRunningProcess_Error
   'If debugflg = 1 Then debugLog "%addImageToDictionaryAndCheckForRunningProcess"
        
    'Sleep 25 ' a small pause to allow the o/s time to write the registry
        
    ' add to the image dictionary
    Call addNewImageToDictionary(newFileName, newName)

    'add to the initiated ProcessArray
    Call checkDockProcessesRunning ' trigger a test of running processes in half a second
    Call checkExplorerRunning

   On Error GoTo 0
   Exit Sub

addImageToDictionaryAndCheckForRunningProcess_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImageToDictionaryAndCheckForRunningProcess of Form menuForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuEditWidget_Click
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuEditWidget_Click()
   Dim execStatus As Long: execStatus = 0
    
   On Error GoTo mnuEditWidget_Click_Error
    
    If sDDefaultEditor = vbNullString Then
        MsgBox "Select the .VBP file that is associated with the Icon Settings VB6 program."
        sDDefaultEditor = addTargetProgram("")
        If fFExists(sDDefaultEditor) Then
            PutINISetting "Software\SteamyDock\DockSettings", "defaultEditor", sDDefaultEditor, dockSettingsFile
            menuForm.mnuEditWidget.Caption = "Edit Program using " & sDDefaultEditor
        End If
    End If
    
    If fFExists(sDDefaultEditor) Then
      
        ' run the selected program
        execStatus = ShellExecute(Me.hwnd, "open", sDDefaultEditor, vbNullString, vbNullString, 1)
        If execStatus <= 32 Then MsgBox "Attempt to open the IDE for this widget failed."
    Else
        MsgBox "Having a bit of a problem opening an IDE for this widget - " & sDDefaultEditor & " It doesn't seem to have a valid working directory set.", "Panzer Earth Gauge Confirmation Message", vbOKOnly + vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

mnuEditWidget_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuEditWidget_Click of Form menuForm"
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
    If debugflg = 1 Then debugLog "%" & "addTargetProgram"
    
    'On Error GoTo l_err1
    'savLblTarget = txtTarget.Text
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not targetText = vbNullString Then
        If fFExists(targetText) Then
            ' extract the folder name from the string
            iconPath = getFolderNameFromPath(targetText)
            ' set the default folder to the existing reference
            dialogInitDir = iconPath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(targetText) Then ' this caters for the entry being just a folder name
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
    If fDirExists("c:\program files") Then
            dialogInitDir = "c:\program files"
        End If
    End If
    
    If Not sDockletFile = vbNullString Then
        If fFExists(sDockletFile) Then
            ' extract the folder name from the string
            dllPath = getFolderNameFromPath(sDockletFile)
            ' set the default folder to the existing reference
            dialogInitDir = dllPath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(sDockletFile) Then ' this caters for the entry being just a folder name
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

  addTargetProgram = retFileName

   On Error GoTo 0
   
   Exit Function

addTargetProgram_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addTargetProgram of Form rDIconConfigForm"
 
End Function
