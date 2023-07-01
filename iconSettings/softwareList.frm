VERSION 5.00
Begin VB.Form formSoftwareList 
   Caption         =   "Software Discovered on this system"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16725
   Icon            =   "softwareList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   16725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBinaryImgStore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   7350
      Picture         =   "softwareList.frx":058A
      ScaleHeight     =   600
      ScaleWidth      =   780
      TabIndex        =   21
      Tag             =   "to store .ico bitmap images"
      Top             =   6885
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picLinkImgStore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   6420
      Picture         =   "softwareList.frx":1654
      ScaleHeight     =   600
      ScaleWidth      =   780
      TabIndex        =   20
      Tag             =   "to store .ico bitmap images"
      Top             =   6885
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton btnDeselectItems 
      Caption         =   "De-select Items"
      Height          =   435
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Deselect the items in the software list above"
      Top             =   6810
      Width           =   1455
   End
   Begin VB.Timer genDragTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10845
      Top             =   8010
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   435
      Left            =   13470
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Get some help on how to operate this utility"
      Top             =   8040
      Width           =   1605
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   10695
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Clear the above list"
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton btnGenerateDock 
      Caption         =   "Generate Dock"
      Height          =   480
      Left            =   15135
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Generate the new dock"
      Top             =   7305
      Width           =   1515
   End
   Begin VB.ListBox lbxApprovedList 
      DragMode        =   1  'Automatic
      Height          =   6690
      Left            =   10680
      TabIndex        =   11
      Top             =   495
      Width           =   5970
   End
   Begin VB.TextBox txtSizeOfFiles 
      Height          =   345
      Left            =   4935
      TabIndex        =   10
      Top             =   6855
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame fraLinkSource 
      Height          =   1245
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   "Select the location to search from"
      Top             =   7230
      Width           =   1920
      Begin VB.OptionButton rdbRegistry 
         Caption         =   "Registry"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "This obtains a program list from the list of applications in the uninstall section of the registry"
         Top             =   675
         Width           =   1395
      End
      Begin VB.OptionButton rdbProgramData 
         Caption         =   "Start Menu"
         Height          =   390
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "This obtains the programs installed on this system from the start menu items belonging to the system"
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.CommandButton btnCopyItems 
      Caption         =   "Copy Items >"
      Height          =   435
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copy selected items from the program list to the  list that will be used to generate the dock"
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton btnCloseSoft 
      Caption         =   "Close"
      Height          =   435
      Left            =   15150
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close this utility"
      Top             =   8025
      Width           =   1515
   End
   Begin VB.TextBox txtNumOfFiles 
      Height          =   375
      Left            =   1110
      TabIndex        =   3
      Top             =   6840
      Width           =   3690
   End
   Begin VB.TextBox txtFileFilter 
      Height          =   360
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "This is the filter that the program will use to find program shortcuts for any programs installed on this system"
      Top             =   6840
      Width           =   900
   End
   Begin VB.TextBox txtPathToTest 
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "This is the path that the utility will use to find and identify any programs installed on this system"
      Top             =   6390
      Width           =   10455
   End
   Begin VB.ListBox lbxSoftwareList 
      DragIcon        =   "softwareList.frx":271E
      Height          =   5715
      ItemData        =   "softwareList.frx":37E8
      Left            =   75
      List            =   "softwareList.frx":37EF
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   495
      Width           =   10455
   End
   Begin VB.Label Label1 
      Caption         =   "(Dock Items that will be used to generate the dock)"
      Height          =   315
      Left            =   11955
      TabIndex        =   13
      Top             =   270
      Width           =   4380
   End
   Begin VB.Label Label 
      Caption         =   "Dock Entries"
      Height          =   345
      Index           =   2
      Left            =   10725
      TabIndex        =   18
      Top             =   270
      Width           =   1470
   End
   Begin VB.Label Label 
      Caption         =   "Installed Software"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   255
      Width           =   1470
   End
   Begin VB.Label lblInformation 
      Height          =   1095
      Left            =   2100
      TabIndex        =   16
      Top             =   7365
      Width           =   8490
   End
   Begin VB.Label lblTitle 
      Caption         =   "List of software in this location"
      Height          =   300
      Left            =   1890
      TabIndex        =   6
      Top             =   60
      Width           =   8490
   End
End
Attribute VB_Name = "formSoftwareList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This utility is used to generate a dock from scratch. It can add icons to the dock in bulk. The source of the data are the
' registry uninstall area and the start menu. These are the only places where applications are 'registered' with the o/s.
' Note: many apps do not do any form of registering, they simply exist on a hard drive somewhere. This application will NOT find those.
'
' We do not currently extract an icon from an EXE. That might happen later.
'
' When the user's list of apps is processed the program will try to identify an appropriate icon to use. It does this through
' the use of an CSV file containing lists of applications, named appIdent.csv that contains two factors used to identify an app.
' if the app corresponds then it will be assigned an icon. That list is limited and only contains a few dozen major applications.
' If the app is not found then it will be assigned a default 'link' icon.

' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files
' .02 DAEB 12/04/2021 formSoftwareList.frm added second list box and controls to add to and clear it as well
' .03 DAEB 12/04/2021 formSoftwareList.frm added controls to create a useful utility, close, help &c
' .04 DAEB 29/05/2022 formSoftwareList.frm Add the ability to turn the tooltips off in the generate dock utility as per ico. sett.
' .05 DAEB 29/05/2022 formSoftwareList.frm Add balloon tooltips to the generate dock utility
' .06 DAEB 30/05/2022 formSoftwareList.frm Add drag and drop to the generate dock utility
' .07 DAEB 30/05/2022 formSoftwareList.frm Check for a path for all entries in the registry list - if there is no path deal with it
' .08 DAEB 30/05/2022 formSoftwareList.frm dragging a binary from the registry should show the binary image instead of the link.

' CREDITS
' Rod Stephens vb-helper.com            Resize controls to fit when a form resizes
' KPD-Team 1999 http://www.allapi.net/  Recursive search
' IT researcher https://www.vbforums.com/showthread.php?784053-Get-installed-programs-list-both-32-and-64-bit-programs
'                                       For the idea of extracting the ununinstall keys from the registry

Option Explicit

' APIs for querying the registry START
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As FILETIME) As Long
' APIs for querying the registry ENDS

' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files START
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = &H20019
Private Const KEY_WOW64_64KEY As Long = &H100&

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
 
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
 
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files END

' .02 DAEB 12/04/2021 formSoftwareList.frm added code to resize the form dynamically STARTS
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single
' .02 DAEB 12/04/2021 formSoftwareList.frm added code to resize the form dynamically ENDS

Private genDragTimerCounter As Integer

'------------------------------------------------------ STARTS
' Constants for hiding/adding horizontal scrollbars to the listboxes
Private Const LB_SETHORIZONTALEXTENT As Long = &H194
Private Const SB_VERT As Long = 1

' APIs for hiding/adding horizontal scrollbars to the listboxes
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'------------------------------------------------------ ENDS



Private Sub btnDeselectItems_Click()

    Dim k As Integer: k = 0
    Dim deselectedCount As Integer: deselectedCount = 0
    
    k = 0
    deselectedCount = 0
    
    For k = 0 To lbxSoftwareList.ListCount - 1
        If lbxSoftwareList.Selected(k) Then
            lbxSoftwareList.Selected(k) = False
            deselectedCount = deselectedCount + 1
        End If
    Next
    
    If deselectedCount = 0 Then
        msgBoxA "Nothing yet to deselect. Please select items from the list above first.", vbInformation + vbOKOnly, "Dock Generation Tool"
    End If

End Sub


' .92 DAEB 26/06/2022 rDIConConfig.frm The auto generation of a dock pulling the start menu links and the registry (undelete) section.
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

   On Error GoTo Form_Load_Error

    ' set the theme colour on startup
    Call setThemeSkin(Me)
  
    txtPathToTest.Text = "C:\Users\beededea\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    txtFileFilter.Text = "*.lnk"
    
    Call SaveSizes

    ' .04 DAEB 29/05/2022 formSoftwareList.frm Add the ability to turn the tooltips off in the generate dock utility as per ico. sett.
    Call genSetToolTips
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form formSoftwareList"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Activate
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : This runs when the form is made visible
'---------------------------------------------------------------------------------------
'
Private Sub Form_Activate()
   On Error GoTo Form_Activate_Error

    formSoftwareList.Refresh
    rdbProgramData.Value = True 'this causes a rdbProgramData_Click
    
    Dim storedFont As String: storedFont = vbNullString
    
    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    Dim lLength As Long

    'storedFont = txtTextFont.Text 'TBD
    
    fntFont = SDSuppliedFont
    fntSize = SDSuppliedFontSize
    fntItalics = CBool(SDSuppliedFontItalics)
    fntColour = CLng(SDSuppliedFontColour)
    
    ' .TBD DAEB 26/05/2022 rdIconConfig.frm Call the font tool for this form
    Call changeFont(formSoftwareList, False, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
  
    ' set the theme colour on startup
    Call setThemeSkin(Me)

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Form formSoftwareList"
    
End Sub



' .02 DAEB 12/04/2021 formSoftwareList.frm  added code to resize the form dynamically
'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 16/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

    ResizeControls

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form formSoftwareList"
End Sub
Private Sub genSetToolTips()

    ' .04 DAEB 29/05/2022 formSoftwareList.frm Add the ability to turn the tooltips off in the generate dock utility as per ico. sett.
    If rDIconConfigForm.chkToggleDialogs.Value = 0 Then
        btnHelp.ToolTipText = "Get some help on how to operate this utility"
        btnClear.ToolTipText = "Clear the above list"
        btnGenerateDock.ToolTipText = "Generate the new dock"
        fraLinkSource.ToolTipText = "Select the location to search from"
        rdbRegistry.ToolTipText = "This obtains a program list from the list of applications in the uninstall section of the registry"
        rdbProgramData.ToolTipText = "This obtains the programs installed on this system from the start menu items belonging to the system"
        btnCopyItems.ToolTipText = "Copy selected items from the program list to the  list that will be used to generate the dock"
        btnCloseSoft.ToolTipText = "Close this utility"
        txtFileFilter.ToolTipText = "This is the filter that the program will use to find program shortcuts for any programs installed on this system"
        txtPathToTest.ToolTipText = "This is the path that the utility will use to find and identify any programs installed on this system"
        btnDeselectItems.ToolTipText = "Deselect the items in the software list above"
    Else
        btnHelp.ToolTipText = ""
        btnClear.ToolTipText = ""
        btnGenerateDock.ToolTipText = ""
        fraLinkSource.ToolTipText = ""
        rdbRegistry.ToolTipText = ""
        rdbProgramData.ToolTipText = ""
        btnCopyItems.ToolTipText = ""
        btnCloseSoft.ToolTipText = ""
        txtFileFilter.ToolTipText = ""
        txtPathToTest.ToolTipText = ""
        btnDeselectItems.ToolTipText = ""
    End If

End Sub



' .02 DAEB 12/04/2021 formSoftwareList.frm added second list box and controls to add to and clear it as well
'---------------------------------------------------------------------------------------
' Procedure : btnClear_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClear_Click()

   On Error GoTo btnClear_Click_Error

    If lbxApprovedList.ListCount <= 0 Then
        msgBoxA "Nothing to clear. No items yet copied.", vbInformation + vbOKOnly, "Dock Generation Tool"
    Else
        lbxApprovedList.Clear
    End If

   On Error GoTo 0
   Exit Sub

btnClear_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClear_Click of Form formSoftwareList"

End Sub

' .03 DAEB 12/04/2021 formSoftwareList.frm added controls to create a useful utility, close, help &c
'---------------------------------------------------------------------------------------
' Procedure : btnCloseSoft_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnCloseSoft_Click()
   On Error GoTo btnCloseSoft_Click_Error

    formSoftwareList.Hide

   On Error GoTo 0
   Exit Sub

btnCloseSoft_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnCloseSoft_Click of Form formSoftwareList"
End Sub


' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files
'---------------------------------------------------------------------------------------
' Procedure : fCheckStartup
' Author    :
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function fCheckStartup() As Integer
    Dim SearchPath As String: SearchPath = vbNullString
    Dim FindStr As String: FindStr = vbNullString
    Dim FileSize As Long: FileSize = 0
    Dim NumFiles As Integer: NumFiles = 0
    Dim NumDirs As Integer: NumDirs = 0
    
    On Error GoTo fCheckStartup_Error

    Screen.MousePointer = vbHourglass
    
    SearchPath = txtPathToTest.Text
    FindStr = txtFileFilter.Text
    
    ' FindFilesAPI does the recursive searching
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    
    'txtSizeOfFiles.Text = "Size of all files found " & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
    
    Screen.MousePointer = vbDefault
    
    fCheckStartup = NumFiles

   On Error GoTo 0
   Exit Function

fCheckStartup_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fCheckStartup of Form formSoftwareList"
End Function

' .02 DAEB 12/04/2021 formSoftwareList.frm added second list box and controls to add to and clear it as well
'---------------------------------------------------------------------------------------
' Procedure : btnCopyItems_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   : 'loop through all selected elements in one listbox and add them to the dock listbox
'---------------------------------------------------------------------------------------
'
Private Sub btnCopyItems_Click()

    Dim k As Long: k = 0
    Dim strItem As String: strItem = vbNullString
    Dim selectedCount As Integer: selectedCount = 0
    Dim startPoint As Integer: startPoint = 0
        
    On Error GoTo btnCopyItems_Click_Error

    'lbxApprovedList.Clear
    
    If lbxApprovedList.ListCount > 0 Then
        startPoint = lbxApprovedList.ListCount
    End If
    
    For k = 0 To lbxSoftwareList.ListCount - 1
        If lbxSoftwareList.Selected(k) Then
            strItem = lbxSoftwareList.List(k)
            lbxApprovedList.AddItem strItem, startPoint
            lbxApprovedList.ListIndex = startPoint
            lbxApprovedList.ToolTipText = lbxApprovedList.List(lbxApprovedList.ListIndex)
            selectedCount = selectedCount + 1
            lbxSoftwareList.Selected(k) = False
        End If
    Next
    
    If selectedCount <= 0 Then
        msgBoxA "Program items not yet selected, select required programs from the left hand list one by one.", vbInformation + vbOKOnly, "Dock Generation Tool"
    End If
    
   On Error GoTo 0
   Exit Sub

btnCopyItems_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnCopyItems_Click of Form formSoftwareList"

End Sub

' .02 DAEB 12/04/2021 formSoftwareList.frm added second list box and controls to add to and clear it as well
'---------------------------------------------------------------------------------------
' Procedure : btnGenerateDock_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGenerateDock_Click()
   On Error GoTo btnGenerateDock_Click_Error

    If lbxApprovedList.ListCount <= 0 Then
        msgBoxA "Dock items not yet generated, select all required programs from the left hand list, then select copy.", vbInformation + vbOKOnly, "Dock Generation Utility", False
    Else
        If lbxApprovedList.ListCount >= 66 Then
            msgBoxA "Too many dock items selected, the maximum is 65 items.", vbExclamation + vbOKOnly, "Dock Generation Utility", False
        Else
            frmConfirmDock.Show
        End If
    End If

   On Error GoTo 0
   Exit Sub

btnGenerateDock_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGenerateDock_Click of Form formSoftwareList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
   On Error GoTo btnHelp_Click_Error
   
   Dim answer As VbMsgBoxResult: answer = vbNo

    answer = msgBoxA("This option opens a browser window and displays this tool's help. Proceed?", vbQuestion + vbYesNo)
    If answer = vbYes Then
        If FExists(App.Path & "\help\Rocketdock Enhanced Settings.html") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\generate documentation.html", vbNullString, App.Path, 1)
        Else
            msgBoxA ("The help file -Rocketdock Enhanced Settings.html- is missing from the help folder."), vbExclamation + vbOKOnly, ""
        End If
    End If
   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form formSoftwareList"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readInstalledAppsRegistry
' Author    :
' Date      : 14/04/2021
' Purpose   :
            ' only process an item from the registry if it has preferably, an icon that gives the location of the binary itself
            ' secondly, if no binary, then a folder location where we attempt to derive the name of all binaries in that folder.
'---------------------------------------------------------------------------------------
' CREDIT: IT researcher https://www.vbforums.com/showthread.php?784053-Get-installed-programs-list-both-32-and-64-bit-programs from the vbforums for the idea of extracting the keys
Public Function readInstalledAppsRegistry(regLocation As Long, keyToSearch As String) As Integer
    Dim hParentKey As Long: hParentKey = 0
    Dim hSubKey As Long: hSubKey = 0
    Dim lIndex As Long: lIndex = 0
    Dim sAppID As String: sAppID = vbNullString
    Dim lAppID As Long: lAppID = 0
    Dim sAppName As String: sAppName = vbNullString
    Dim sAppLocation As String: sAppLocation = vbNullString
    Dim sAppIcon As String: sAppIcon = vbNullString
    Dim appName As String: appName = vbNullString
    Dim AppLocation As String: AppLocation = vbNullString
    Dim appIcon As String: appIcon = vbNullString
    Dim lAppName As Long: lAppName = 0
    Dim lAppLocation As Long: lAppLocation = 0
    Dim lAppIcon As Long: lAppIcon = 0
    Dim ValueType As Long: ValueType = 0
    Dim DummyTime As FILETIME
    Dim sTmp As String: sTmp = vbNullString
    Dim entryCount As Integer: entryCount = 0
    Dim textVersion As String: textVersion = vbNullString
    Dim testFileName As String: testFileName = vbNullString
    
    On Error GoTo readInstalledAppsRegistry_Error
    
    If RegOpenKeyEx(regLocation, keyToSearch, _
    0, KEY_READ Or KEY_WOW64_64KEY, hParentKey) = 0 Then  ' Here passing only KEY_READ  reads 32 bit hive
        sAppID = Space(128)
        lAppID = 128
        Do While RegEnumKeyEx(hParentKey, lIndex, sAppID, 255, 0, vbNullString, 0, DummyTime) = 0
            sAppID = Left$(sAppID, lAppID)
            
            If RegOpenKeyEx(hParentKey, sAppID, 0, KEY_QUERY_VALUE, hSubKey) = 0 Then
                lAppName = 0
                appName = vbNullString
                'DisplayName
                If RegQueryValueEx(hSubKey, "DisplayName", 0, ValueType, ByVal 0, lAppName) = 0 Then
                    If ValueType = REG_SZ Then
                        sAppName = Space(lAppName)
                        RegQueryValueEx hSubKey, "DisplayName", 0, 0, ByVal sAppName, lAppName
                        sAppName = Left$(sAppName, lAppName - 1)
            
                        If sAppName <> vbNullString Then appName = sAppName
                        sAppName = vbNullString
                    End If
                End If
                
                lAppLocation = 0
                AppLocation = vbNullString
                'InstallLocation
                If RegQueryValueEx(hSubKey, "InstallLocation", 0, ValueType, ByVal 0, lAppLocation) = 0 Then
                    If ValueType = REG_SZ Then
                        sAppLocation = Space(lAppLocation)
                        RegQueryValueEx hSubKey, "InstallLocation", 0, 0, ByVal sAppLocation, lAppLocation
                        sAppLocation = Left$(sAppLocation, lAppLocation - 1)
            
                        If sAppLocation <> vbNullString Then AppLocation = sAppLocation
                        sAppLocation = vbNullString
                    End If
                End If
                
                lAppIcon = 0
                appIcon = vbNullString
                'DisplayIcon
                If RegQueryValueEx(hSubKey, "DisplayIcon", 0, ValueType, ByVal 0, lAppIcon) = 0 Then
                    If ValueType = REG_SZ Then
                        sAppIcon = Space(lAppIcon)
                        RegQueryValueEx hSubKey, "DisplayIcon", 0, 0, ByVal sAppIcon, lAppIcon
                        sAppIcon = Left$(sAppIcon, lAppIcon - 1)
            
                        If sAppIcon <> vbNullString Then appIcon = sAppIcon
                        sAppIcon = vbNullString
                    End If
                End If
                
                RegCloseKey hSubKey
                hSubKey = 0
            End If
            lIndex = lIndex + 1
           
            sAppID = Space(128)
            lAppID = 128

            ' only process an item from the registry if it has preferably, an icon that gives the location of the binary itself
            ' secondly, if no binary, then a folder location where we attempt to derive the name of all binaries in that folder.
            
            If (AppLocation <> vbNullString Or appIcon <> vbNullString) And InStr(appIcon, "{") = 0 Then
                'sTmp = AppName & " Folder: " & AppLocation & " Binary: " & AppIcon & vbCrLf & vbCrLf
                
                ' we use the application icon field as that tells us where the icon is.
                ' most will be embedded in the binary (.exe) so we have the binary names straight away.
                ' some (a few) will have a separate .ico file and that is a problem
                If InStr(appIcon, "{") <> 0 Then
                    appIcon = ""
                End If
                
                ' strip the final ,0 from the binary filename
                If Right$(appIcon, 2) = ",0" Then
                    appIcon = Left$(appIcon, Len(appIcon) - 2)
                End If
                
                ' if the key is an .ico
                If Right$(appIcon, 3) = "ico" Then
                ' we assume that a.exe exists for an ico of the same name and we strip ico and add that
                    appIcon = Left$(appIcon, Len(appIcon) - 3) & "exe"
                    If Not FExists(appIcon) Then ' does the file exist?
                        appIcon = ""
                    End If
                End If
                
                ' strip the double quotes that occasionally surround a binary string name
                If Left$(appIcon, 1) = """" Then
                    appIcon = Replace(appIcon, """", "")
                End If
                
                ' add the binary name to the list if the appicon field is non blank
                If Not LTrim$(appIcon) = "" Then
                    If Right$(appIcon, 3) = "exe" And Not Right$(appIcon, 12) = "unins000.exe" Then
                        
                        ' .07 DAEB 30/05/2022 formSoftwareList.frm Check for a path for all entries in the registry list - if there is no path deal with it
                        'if the path is valid then it will have a backslash somewhere
                        If InStr(appIcon, "\") Then
                            lbxSoftwareList.AddItem appIcon
                        Else 'if the path is missing then it will have no backslash, we possibly have just the binary name
                            lbxSoftwareList.AddItem AppLocation & "\" & appIcon
                        End If
                        entryCount = entryCount + 1
                    End If
                Else ' appicon is blank so we need to derive the binary name
                    'open the folder location and search for all binaries in that folder, add each to the list
                    If AppLocation <> "" Then '
                        testFileName = Dir(AppLocation)
                        Do While testFileName > ""
                            ' exclude all the uninstall binaries
                            If Right$(testFileName, 3) = "exe" And Not testFileName = "unins000.exe" Then
                                lbxSoftwareList.AddItem AppLocation & testFileName

                                entryCount = entryCount + 1
                            End If
                            testFileName = Dir() 'jump to the next file
                        Loop
                    End If
                End If
                
            End If
            appName = vbNullString
        Loop
        RegCloseKey hParentKey
    End If

    readInstalledAppsRegistry = entryCount

   On Error GoTo 0
   Exit Function

readInstalledAppsRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readInstalledAppsRegistry of Form formSoftwareList"

End Function



' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files
'---------------------------------------------------------------------------------------
' Procedure : StripNulls
' Author    : beededea
' Date      : 14/04/2021
' Purpose   : KPD-Team 1999
' Credit    : E-Mail: [email]KPDTeam@Allapi.net[/email]
'             URL: [url]http://www.allapi.net/[/url]
'             strip nulls from the string
'---------------------------------------------------------------------------------------
'
Function StripNulls(OriginalStr As String) As String
   On Error GoTo StripNulls_Error

    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr

   On Error GoTo 0
   Exit Function

StripNulls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StripNulls of Form formSoftwareList"
End Function
' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files
'---------------------------------------------------------------------------------------
' Procedure : FindFilesAPI
' Author    : beededea
' Date      : 14/04/2021
' Purpose   : KPD-Team 1999
'                E-Mail: [email]KPDTeam@Allapi.net[/email]
'                URL: [url]http://www.allapi.net/[/url]
'---------------------------------------------------------------------------------------
'
Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
 
    Dim Filename As String: Filename = vbNullString ' Walking filename variable...
    Dim DirName As String: DirName = vbNullString ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer: nDir = 0 ' Number of directories in this path
    Dim i As Integer: i = 0 ' For-loop counter...
    Dim hSearch As Long: hSearch = 0 ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer: Cont = 0
    
    On Error GoTo FindFilesAPI_Error

    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                lbxSoftwareList.AddItem Path & Filename
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If

   On Error GoTo 0
   Exit Function

FindFilesAPI_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindFilesAPI of Form formSoftwareList"
End Function


Private Sub fraLinkSource_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip fraLinkSource.hwnd, "Select the location to search from, choose either the registry or the start menu links.", _
                  TTIconInfo, "Help on the Link Source radio buttons", , , , True
End Sub


' .06 DAEB 30/05/2022 formSoftwareList.frm Add drag and drop to the generate dock utility
Private Sub genDragTimer_Timer()
    
    genDragTimerCounter = genDragTimerCounter + 1

    If genDragTimerCounter >= 25 Then
        ' .59 DAEB 01/05/2022 rDIConConfig.frm Added manual drag and drop functionality
        lbxSoftwareList.Drag vbBeginDrag
        genDragTimer.Enabled = False
        genDragTimerCounter = 0
    End If
End Sub




' .06 DAEB 30/05/2022 formSoftwareList.frm Add drag and drop to the generate dock utility
Private Sub lbxApprovedList_DragDrop(Source As Control, x As Single, y As Single)
    Dim k As Long: k = 0
    Dim strItem As String: strItem = vbNullString
    Dim selectedCount As Integer: selectedCount = 0
    Dim startPoint As Integer: startPoint = 0

    
    lbxSoftwareList.Drag vbEndDrag
    
    If lbxApprovedList.ListCount > 0 Then
        startPoint = lbxApprovedList.ListCount
    End If
    
    'Screen.MousePointer = vbHourglass
    
    For k = 0 To lbxSoftwareList.ListCount - 1
        If lbxSoftwareList.Selected(k) Then
            strItem = lbxSoftwareList.List(k)
            lbxApprovedList.AddItem strItem, startPoint
            lbxApprovedList.ListIndex = startPoint
            lbxApprovedList.ToolTipText = lbxApprovedList.List(lbxApprovedList.ListIndex)
            selectedCount = selectedCount + 1
            lbxSoftwareList.Selected(k) = False
        End If
    Next
    
    'Screen.MousePointer = vbDefault
    
'    If selectedCount <= 0 Then
'        msgBoxA "Program items not yet selected, select required programs from the left hand list one by one.", vbInformation + vbOKOnly, "Dock Generation Tool"
'    End If

End Sub

Private Sub lbxApprovedList_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    lbxApprovedList.ToolTipText = lbxApprovedList.List(lbxApprovedList.ListIndex)
    'lbxSoftwareList.Drag vbEndDrag
End Sub

Private Sub lbxApprovedList_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip lbxApprovedList.hwnd, "To generate a dock full of entries, this listbox must be populated with a list of your chosen software links. Drag and drop from the lists on the left which populate from the registry or the start menu..", _
                  TTIconInfo, "Help on the Chosen Links List", , , , True
End Sub


' .02 DAEB 12/04/2021 formSoftwareList.frm added second list box and controls to add to and clear it as well
'---------------------------------------------------------------------------------------
' Procedure : lbxSoftwareList_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lbxSoftwareList_Click()
   On Error GoTo lbxSoftwareList_Click_Error

    lbxSoftwareList.ToolTipText = lbxSoftwareList.List(lbxSoftwareList.ListIndex)

   On Error GoTo 0
   Exit Sub

lbxSoftwareList_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lbxSoftwareList_Click of Form formSoftwareList"
End Sub

Private Sub lbxSoftwareList_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    
    ' .08 DAEB 30/05/2022 formSoftwareList.frm dragging a binary from the registry should show the binary image instead of the link.
    If rdbProgramData.Value = True Then
        Set lbxSoftwareList.DragIcon = picLinkImgStore.Picture ' uses two picboxes to store .ico bitmap images
    Else
        Set lbxSoftwareList.DragIcon = picBinaryImgStore.Picture
    End If
    ' .06 DAEB 30/05/2022 formSoftwareList.frm Add drag and drop to the generate dock utility
    If genDragTimer.Enabled = False Then genDragTimer.Enabled = True ' initiates the vbBeginDrag after n millisecs

End Sub

' .06 DAEB 30/05/2022 formSoftwareList.frm Add drag and drop to the generate dock utility
Private Sub lbxSoftwareList_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    ' when clicking one more than one item in the listbox it is essential to deactivate the dragIcon timer as we don't want the dragicon appearing
    ' in a way that seems willy nilly to the end user.
    
    lbxSoftwareList.Drag vbEndDrag
    genDragTimer.Enabled = False
    genDragTimerCounter = 0
End Sub

' .01 DAEB 12/04/2021 formSoftwareList.frm added code to recursively trawl through the folders and find .lnk files
'---------------------------------------------------------------------------------------
' Procedure : rdbProgramData_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   : recursively extracts all .lnk files in the Windows start menu and outputs to a listbox
'---------------------------------------------------------------------------------------
'
Private Sub rdbProgramData_Click()

    Dim userprof As String: userprof = vbNullString ' %userprofile%
    Dim ProgramData As String: ProgramData = vbNullString '
    Dim s As Integer: s = 0
    Dim totalShortsFound As Integer: totalShortsFound = 0
    
    On Error GoTo rdbProgramData_Click_Error

    ProgramData = Environ$("PROGRAMDATA")
    userprof = Environ$("USERPROFILE")
    
    lbxSoftwareList.Clear
    txtFileFilter.Text = "*.lnk"
    
    txtPathToTest.Text = ProgramData & "\Microsoft\Windows\Start Menu\Programs"
    s = fCheckStartup
    totalShortsFound = totalShortsFound + s
    
    txtPathToTest.Text = userprof & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    s = fCheckStartup
    totalShortsFound = totalShortsFound + s

    lblTitle.Caption = "STARTUP MENU LIST OF INSTALLED SOFTWARE IN %PROGRAMDATA%"
    lblInformation.Caption = "The system Program Data area and the user profile are the two locations where all " & vbCrLf & _
                             "the programs that exist in the system's start menu are located. We do a recursive " & vbCrLf & _
                             "search through all folders and files beneath these two locations looking for shortcut files."
        
    txtNumOfFiles.Text = totalShortsFound & " Shortcut Files found"
        
    On Error GoTo 0
    Exit Sub

rdbProgramData_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdbProgramData_Click of Form formSoftwareList"
End Sub

Private Sub rdbProgramData_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbProgramData.hwnd, "Click here to select the program items found within the Start Menu .", _
                  TTIconInfo, "Help on the Start Menu", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : rdbRegistry_Click
' Author    : beededea
' Date      : 14/04/2021
' Purpose   :
'
'    HKEY_LOCAL_MACHINE\Software\Classes\Installer\Products
'    HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall
'    HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall
'    HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Installer\UserData

'---------------------------------------------------------------------------------------
'
Private Sub rdbRegistry_Click()
    Dim keyToSearch As String: keyToSearch = vbNullString
    Dim locationToSearch As Long: locationToSearch = 0
    Dim totalKeysFound As Integer: totalKeysFound = 0
    Dim s As Integer: s = 0
    Dim textVersion As String: textVersion = vbNullString
    
    On Error GoTo rdbRegistry_Click_Error

    lbxSoftwareList.Clear
    txtPathToTest.Text = ""
    txtFileFilter.Text = ""
    txtNumOfFiles.Text = ""
    txtSizeOfFiles.Text = ""
    
    lblInformation.Caption = "We extract the information from the list of applications following areas of the registry." & vbCrLf & _
    "HKEY_LOCAL_MACHINE\Software\Classes\Installer\Products" & vbCrLf & _
    "HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" & vbCrLf & _
    "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall" & vbCrLf & _
    "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Installer\UserData"
    
    'xFileName = App.Path & "\ins.txt"
    
    keyToSearch = "Software\Classes\Installer\Products"
    s = readInstalledAppsRegistry(HKEY_LOCAL_MACHINE, keyToSearch)
    totalKeysFound = totalKeysFound + s
    
    keyToSearch = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    s = readInstalledAppsRegistry(HKEY_LOCAL_MACHINE, keyToSearch)
    totalKeysFound = totalKeysFound + s
    
    keyToSearch = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
    s = readInstalledAppsRegistry(HKEY_LOCAL_MACHINE, keyToSearch)
    totalKeysFound = totalKeysFound + s
    
    keyToSearch = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData"
    s = readInstalledAppsRegistry(HKEY_LOCAL_MACHINE, keyToSearch)
    totalKeysFound = totalKeysFound + s
    
    txtNumOfFiles.Text = totalKeysFound & " Valid entries found"
    
    If HKEY_LOCAL_MACHINE = -2147483646 Then
        textVersion = "HKEY_LOCAL_MACHINE"
    End If

    txtPathToTest.Text = textVersion & "\" & keyToSearch
    lblTitle.Caption = "REGISTRY LIST OF SOFTWARE " & textVersion & " and subkeys"
    txtFileFilter.Text = "*.*"
        
    'Call WriteOutputFile(s, xFileName)

   On Error GoTo 0
   Exit Sub

rdbRegistry_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdbRegistry_Click of Form formSoftwareList"
    
End Sub
' .02 DAEB 12/04/2021 formSoftwareList.frm  added code to resize the form dynamically
'---------------------------------------------------------------------------------------
' Procedure : SaveSizes
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : Resize controls to fit when a form resizes
'             Save the form's and controls' dimensions.
' Credit    : Rod Stephens vb-helper.com
'---------------------------------------------------------------------------------------
'
Private Sub SaveSizes()
    Dim i As Integer: i = 0
    Dim a As Integer: a = 0
    Dim Ctrl As Control

    ' Save the controls' positions and sizes.
    On Error GoTo SaveSizes_Error

    ReDim m_ControlPositions(1 To Controls.count)
    i = 1
    For Each Ctrl In Controls
        With m_ControlPositions(i)
            
            
            If TypeOf Ctrl Is Line Then
                .Left = Ctrl.x1
                .Top = Ctrl.y1
                .Width = Ctrl.X2 - Ctrl.x1
                .Height = Ctrl.Y2 - Ctrl.y1
'            Else
            ' .TBD DAEB 26/05/2022 rdIconConfig.frm Add all the types of controls handled - after adding a timer to the form...
            ElseIf (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
                a = 1
                .Left = Ctrl.Left
                .Top = Ctrl.Top
                .Width = Ctrl.Width
                .Height = Ctrl.Height
                On Error Resume Next
                .FontSize = Ctrl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next Ctrl

    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight

   On Error GoTo 0
   Exit Sub

SaveSizes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSizes of Form formSoftwareList"
End Sub


' .02 DAEB 12/04/2021 formSoftwareList.frm  added code to resize the form dynamically
'---------------------------------------------------------------------------------------
' Procedure : ResizeControls
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : Arrange the controls for the new size.
'---------------------------------------------------------------------------------------
'
Private Sub ResizeControls()
    Dim i As Integer: i = 0
    Dim Ctrl As Control
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
        
    ' Don't bother if we are minimized.
    On Error GoTo ResizeControls_Error

    If WindowState = vbMinimized Then Exit Sub

    ' Get the form's current scale factors.
    x_scale = ScaleWidth / m_FormWid
    y_scale = ScaleHeight / m_FormHgt

    ' Position the controls.
    i = 1
    For Each Ctrl In Controls
        With m_ControlPositions(i)
            If TypeOf Ctrl Is Line Then
                Ctrl.x1 = x_scale * .Left
                Ctrl.y1 = y_scale * .Top
                Ctrl.X2 = Ctrl.x1 + x_scale * .Width
                Ctrl.Y2 = Ctrl.y1 + y_scale * .Height
            ' .TBD DAEB 26/05/2022 rdIconConfig.frm Add all the types of controls handled - after adding a timer to the form...
            ElseIf (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
            'Else
                Ctrl.Left = x_scale * .Left
                Ctrl.Top = y_scale * .Top
                Ctrl.Width = x_scale * .Width
                If Not (TypeOf Ctrl Is ComboBox) Then
                    ' Cannot change height of ComboBoxes.
                    Ctrl.Height = y_scale * .Height
                End If
                On Error Resume Next
                Ctrl.Font.Size = y_scale * .FontSize
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next Ctrl

   On Error GoTo 0
   Exit Sub

ResizeControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ResizeControls of Form formSoftwareList"
End Sub

' .05 DAEB 29/05/2022 formSoftwareList.frm Add balloon tooltips to the generate dock utility STARTS

Private Sub rdbRegistry_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbRegistry.hwnd, "Click here to select the program items found within the Registry.", _
                  TTIconInfo, "Help on the Registry", , , , True
End Sub

Private Sub txtFileFilter_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip txtFileFilter.hwnd, "This text box contains the types of items found in either the registry or in the start menu.", _
                  TTIconInfo, "Help on the type of items found", , , , True
End Sub

Private Sub txtNumOfFiles_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip txtNumOfFiles.hwnd, "This text box contains the total number of items found in either the registry or in the start menu.", _
                  TTIconInfo, "Help on the Total Number of items found", , , , True
End Sub

Private Sub txtPathToTest_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip txtPathToTest.hwnd, "This text box contains the path that the utility will use to find and identify any programs installed on this system, located either in the registry or in the start menu.", _
                  TTIconInfo, "Help on the Software Path", , , , True
End Sub
Private Sub btnCloseSoft_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnCloseSoft.hwnd, "This button cancels the current operation and closes the window.", _
                  TTIconInfo, "Help on the Cancel and Close Button", , , , True
End Sub

Private Sub btnGenerateDock_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGenerateDock.hwnd, "This button will proceed to generate the new dock using the chosen application links above.", _
                  TTIconInfo, "Help on the Generate Dock Button", , , , True
End Sub
Private Sub btnCopyItems_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnCopyItems.hwnd, "This button will copy any selected items from the above list to the chosen list on the right. You can also use drag and drop for individual links.", _
                  TTIconInfo, "Help on the Copy Items Button", , , , True
End Sub
Private Sub btnClear_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnClear.hwnd, "This button will clear the above list of your chosen links.", _
                  TTIconInfo, "Help on the Clear Chosen Links Button", , , , True
End Sub
Private Sub btnHelp_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If rDEnableBalloonTooltips = "1" Then CreateToolTip btnHelp.hwnd, "This button opens the help page in your default browser.", _
                  TTIconInfo, "Help on the Help Button", , , , True
End Sub
Private Sub lbxSoftwareList_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip lbxSoftwareList.hwnd, "This listbox contains a complete list of the software that is recognised by your Windows installation. Ths consists of entries extracted from the registry or the start menu. Any items that you want to appear in your dock, click on them. When you have selected those you want, press the Copy Items button. Each of your choices will be placed upon the list on the right hand side. You can also drag and drop individual items.", _
                  TTIconInfo, "Help on the Available Software List", , , , True
End Sub
Private Sub btnDeselectItems_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip btnDeselectItems.hwnd, "This button removes all the selections in the box above, this avoids the chance of replication of items in the approved list.", _
                  TTIconInfo, "Help on the De-Selecting items in the software list", , , , True
End Sub

' .05 DAEB 29/05/2022 formSoftwareList.frm Add balloon tooltips to the generate dock utility ENDS

'---------------------------------------------------------------------------------------
' Procedure : generateDockInformation
' Author    : beededea
' Date      : 19/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub generateDockInformation()
    Dim useloop2 As Integer: useloop2 = 0
    Dim iconImage As String: iconImage = vbNullString
    Dim iconTitle As String: iconTitle = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim iconCommand As String: iconCommand = vbNullString
    Dim iconArguments As String: iconArguments = vbNullString
    Dim iconWorkingDirectory As String: iconWorkingDirectory = vbNullString
    Dim location As String: location = vbNullString
    Dim newMaximum As Integer: newMaximum = 0
    Dim currentIcon As Integer: currentIcon = 0
    Dim testSettingsfile As String: testSettingsfile = vbNullString
    Dim listCounter As Integer: listCounter = 0
    Dim oldIconMaximum As Integer: oldIconMaximum = 0
    Dim startIcon As Integer: startIcon = 0
    Dim endIcon As Integer: endIcon = 0

    On Error GoTo generateDockInformation_Error
    
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile

    location = "Software\SteamyDock\IconSettings"
    
    'testSettingsfile = App.Path & "\testDocksettings.ini"
    
    oldIconMaximum = rdIconMaximum
    
    msgBoxA "Backup Completed" & vbCr & "Dock Generation Starting - press OK and then please hold on for a few seconds whilst it completes.", vbExclamation + vbOKOnly, "Dock Generation Tool"
    
    ' change the settings.ini to reflect our choice
    
    ' .93 DAEB 26/06/2022 formSoftwareList.frm generate dock - overwrite dock routine
    If frmConfirmDock.rdbOverwrite = True Then ' is the option overwrite?
        Call rDIconConfigForm.deleteRdMap(False, False) ' if so then empty the dock
        ' write the new data in a loop
        For useloop2 = 0 To lbxApprovedList.ListCount - 1 ' the dock is zero-based but a list is 1-based

            Call zeroAllIconCharacteristics
            
            Call checkTheLink(useloop2, iconImage, iconTitle, iconFileName, iconCommand, iconArguments, iconWorkingDirectory)

            sFilename = iconImage
            sTitle = iconTitle
            sCommand = iconCommand
            sArguments = iconArguments
            sWorkingDirectory = iconWorkingDirectory

            ' write the settings.ini
            Call writeIconSettingsIni("Software\SteamyDock\IconSettings" & "\Icons", useloop2, interimSettingsFile)
    
        Next useloop2
        
        rdIconMaximum = lbxApprovedList.ListCount - 1
        
        'amend the count to one more than the max as 0- is a valid icon
        PutINISetting location & "\Icons", "count", rdIconMaximum + 1, interimSettingsFile
    
    ElseIf frmConfirmDock.rdbAppend = True Then ' is the option append? If so, write the new icons to the end.
        
        ' read the dock maximum count and increase that to include the newly selected items
        newMaximum = rdIconMaximum + lbxApprovedList.ListCount
        
        listCounter = 0 ' the approved list starts at zero
        For useloop2 = rdIconMaximum + 1 To newMaximum
            Call zeroAllIconCharacteristics
            Call checkTheLink(listCounter, iconImage, iconTitle, iconFileName, iconCommand, iconArguments, iconWorkingDirectory)

            sFilename = iconImage
            sTitle = iconTitle
            sCommand = iconCommand
            sArguments = iconArguments
            sWorkingDirectory = iconWorkingDirectory

            ' write the alternative settings.ini
            Call writeIconSettingsIni(location & "\Icons", useloop2, interimSettingsFile)
            listCounter = listCounter + 1
            
         Next useloop2

        rdIconMaximum = newMaximum

        'amend the count to one more than the max as 0- is a valid icon
        PutINISetting location & "\Icons", "count", rdIconMaximum + 1, interimSettingsFile
    
    ElseIf frmConfirmDock.rdbPrepend = True Then ' is the option prepend?

        newMaximum = rdIconMaximum + lbxApprovedList.ListCount
        
        ' read the old icons one at a time from the end to the beginning and write them at their new location
        For useloop2 = rdIconMaximum To 0 Step -1
            
            ' write the alternative settings.ini
            Call readIconSettingsIni(location & "\Icons", useloop2, interimSettingsFile)

            ' write the alternative settings.ini
            Call writeIconSettingsIni(location & "\Icons", useloop2 + lbxApprovedList.ListCount, interimSettingsFile)

        Next useloop2

        ' write the new icons to the dock at the beginning
        For useloop2 = 0 To lbxApprovedList.ListCount - 1

            Call zeroAllIconCharacteristics
            Call checkTheLink(useloop2, iconImage, iconTitle, iconFileName, iconCommand, iconArguments, iconWorkingDirectory)

            sFilename = iconImage
            sTitle = iconTitle
            sCommand = iconCommand
            sArguments = iconArguments
            sWorkingDirectory = iconWorkingDirectory

            
            ' write the settings.ini
            Call writeIconSettingsIni(location & "\Icons", useloop2, interimSettingsFile)

        Next useloop2

        rdIconMaximum = newMaximum

        ' amend the count to one more than the max as 0- is a valid icon
        PutINISetting location & "\Icons", "count", rdIconMaximum + 1, interimSettingsFile
    
    ElseIf frmConfirmDock.rdbCurrent = True Then     ' is the option at current icon?

        currentIcon = rdIconNumber
        startIcon = currentIcon + 1 ' do not overwrite the current icon. Start one icon box to the right.
        endIcon = startIcon + lbxApprovedList.ListCount - 1
        
        newMaximum = rdIconMaximum + lbxApprovedList.ListCount
        For useloop2 = rdIconMaximum To startIcon Step -1

        '   read the old icons from the current dock position one at a time from the end to the current position.
            Call readIconSettingsIni(location & "\Icons", useloop2, interimSettingsFile)

            ' write them at their new location
            Call writeIconSettingsIni(location & "\Icons", useloop2 + lbxApprovedList.ListCount, interimSettingsFile)

        Next useloop2

        listCounter = 0
        For useloop2 = startIcon To endIcon
            Call zeroAllIconCharacteristics
            Call checkTheLink(listCounter, iconImage, iconTitle, iconFileName, iconCommand, iconArguments, iconWorkingDirectory)

            sFilename = iconImage
            sTitle = iconTitle
            sCommand = iconCommand
            sArguments = iconArguments
            sWorkingDirectory = iconWorkingDirectory


            ' write the approved icon list to the settings.ini
            Call writeIconSettingsIni(location & "\Icons", useloop2, interimSettingsFile)
            listCounter = listCounter + 1
            
        Next useloop2

        rdIconMaximum = newMaximum

        'amend the count
        PutINISetting location & "\Icons", "count", rdIconMaximum + 1, interimSettingsFile

    End If

     ' close the frmConfirmDock and the generate form, clear the app. list
    frmConfirmDock.Hide
    formSoftwareList.Hide
    lbxApprovedList.Clear
    
    ' hide the form during the generation, looks a bit messy whilst refreshing  from one form to another
    rDIconConfigForm.Visible = False
     
     ' refresh the map in the main form, the oldIconMaximum set here causes a recreation of the extra icon positions in the map.
    Call rDIconConfigForm.recreateTheMap(oldIconMaximum)
    
    rDIconConfigForm.Visible = True
    
    msgBoxA "Dock generation done.", vbExclamation + vbOKOnly, "Dock Generation Tool"
                              
    On Error GoTo 0
    Exit Sub

generateDockInformation_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generateDockInformation of Form formSoftwareList"
            Resume Next
          End If
    End With
                  
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkTheLink
' Author    : beededea
' Date      : 17/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkTheLink(ByVal listNo As Integer _
    , ByRef iconImage As String _
    , ByRef iconTitle As String _
    , ByRef iconFileName As String _
    , ByRef iconCommand As String _
    , ByRef iconArguments As String _
    , ByRef iconWorkingDirectory As String)
    
    Dim thisLink As String: thisLink = vbNullString
    Dim suffix As String: suffix = vbNullString
    Dim nname As String: nname = vbNullString
    Dim npath As String: npath = vbNullString
    Dim ndesc As String: ndesc = vbNullString
    Dim nwork As String: nwork = vbNullString
    Dim nargs As String: nargs = vbNullString
    Dim thisShortcut As Link

    On Error GoTo checkTheLink_Error
        
    thisLink = lbxApprovedList.List(listNo)
    suffix = LCase$(ExtractSuffixWithDot(thisLink))

    iconCommand = thisLink ' set the command for all file types

    ' If it is a shortcut we have some code to investigate the shortcut for the link details
    If suffix = ".lnk" Then
        If FExists(iconCommand) Then
            ' if it is a short cut then you can use two methods, the first is currently limited to only
            ' producing the path alone but it does avoid using the shell method that causes FPs to occur in a/v tools

            Call GetShortcutInfo(iconCommand, thisShortcut) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry

            iconTitle = getFileNameFromPath(thisShortcut.Filename)

            If Not thisShortcut.Filename = "" Then
                iconCommand = LCase$(thisShortcut.Filename)
            End If
            iconArguments = thisShortcut.Arguments
            iconWorkingDirectory = thisShortcut.RelPath

            ' Use a call to the older function to identify an icon using the shell object
            'if the icontitle and command are blank then this is user-created link that only provides the relative path
            If iconTitle = "" And thisShortcut.Filename = "" And Not iconWorkingDirectory = "" Then
                Call GetShellShortcutInfo(iconCommand, nname, npath, ndesc, nwork, nargs)

                iconTitle = nname
                iconCommand = npath
                iconArguments = nargs
                iconWorkingDirectory = nwork
            End If

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
    Else ' it is a binary
        ' take the name from the filename
        If FExists(iconCommand) Then
            iconTitle = getFileNameFromPath(iconCommand)
            iconArguments = ""
            iconWorkingDirectory = ""
            ' extract the icon from the exe or dll
            ' we can already extract an PNG or similar icon and we can write it to a pictureBox
            ' it should be possible to capture the stream and feed it to GDI+ or write it to a file
             iconFileName = identifyAppIcons(iconCommand)

             If FExists(iconFileName) Then
                 iconImage = iconFileName
             Else
                 iconImage = App.Path & "\my collection\steampunk icons MKVI" & "\document-lnk.png"
             End If
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkTheLink_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkTheLink of Form formSoftwareList"
            Resume Next
          End If
    End With
End Sub
