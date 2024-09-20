Attribute VB_Name = "common"

' .01 DAEB 23/01/2021 common.bas calls twipsperpixelsX/Y function when determining the twips for high DPI screens
' .02 DAEB 25/01/2021 common.bas Moved from mdlmain.bas to common to ensure the checkSteamyDockInstalled subroutine can be run from anywhere, specifically for the variable sdAppPath
' .03 DAEB 31/01/2021 common.bas Added new checkbox to determine if a post initiation dialog should appear
' .04 DAEB 06/03/2021 common.bas Moved from main code form to common to ensure the locateDockSettingsFile subroutine is common to all
' .05 DAEB 01/04/2021 common.bas Added declaration to allow replacement of some modal msgbox with the non-modal versions
' .06 DAEB 19/04/2021 common.bas moved to the common area so that it can be used by each of the utilities
' .07 DAEB 26/04/2021 common.bas changed to use pixels alone, removed all unnecessary twip conversion
' .08 DAEB 11/05/2021 common.bas Added function to pad a string similar to the VB.NET padRight & padLeft functions.
' .09 DAEB 11/05/2021 common.bas Added function to align and centre a string so it can appear in a msgbox neatly.
' .10 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
' .12 DAEB 20/05/2021 common.bas Added new check box to allow autohide of the dock after launch of the chosen app
 
Option Explicit

'------------------------------------------------------------
' common.bas
'
' Public procedures that appear in all three programs as an included module common.bas,
'
' Note: If you make a change here it affects all three programs dynamically
'------------------------------------------------------------

' APIs and variables for querying processes START
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private uProcess   As PROCESSENTRY32
Private hSnapshot As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long ' Alias "CreateToolhelp32Snapshot"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' APIs for querying processes END

' functions to determine 64bitness start
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long
' functions to determine 64bitness END

' enumerate variables for folder values start
Public Enum eSpecialFolders
  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
End Enum
' enumerate variables for folder values END


'API Function to read/write information from INI File start
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'API Function to read/write information from INI File start

' APIs, constants defined for querying the registry STARTS
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
' APIs, constants defined for querying the registry ENDS

' APIs and structures for opening a common dialog box to select files without OCX dependencies STARTS
Private Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum

Public Type OPENFILENAME
    lStructSize As Long    'The size of this struct (Use the Len function)
    hWndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" ( _
    lpofn As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" ( _
    lpofn As OPENFILENAME) As Long
    
Public OF As OPENFILENAME
Public x_OpenFilename As OPENFILENAME

Private Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long 'LPCITEMIDLIST
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long  'BFFCALLBACK
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolderA Lib "shell32.dll" (binfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl&, ByVal szPath$) As Long
Private Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' APIs and structures for opening a common dialog box to select files without OCX dependencies STARTS

' Rocketdock compatible icon global variables START
Public sFilename As String
Public sFileName2 As String
Public sTitle  As String
Public sCommand  As String
Public sArguments  As String
Public sWorkingDirectory  As String
Public sShowCmd  As String
Public sOpenRunning  As String
Public sRunElevated  As String

Public sIsSeparator  As String
Public sUseContext  As String
Public sDockletFile  As String
Public sUseDialog  As String
Public sUseDialogAfter  As String ' .03 DAEB 31/01/2021 common.bas Added new checkbox to determine if a post initiation dialog should appear
Public sQuickLaunch  As String ' .10 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
Public sAutoHideDock  As String ' .12 DAEB 20/05/2021 common.bas Added new check box to allow autohide of the dock after launch of the chosen app
Public sSecondApp  As String ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
Public sRunSecondAppBeforehand  As String

Public sAppToTerminate As String
Public sDisabled  As String

' Rocketdock icon global variables END


Public usedMenuFlag As Boolean

Public dockSettingsFile As String
Public toolSettingsFile  As String

'Public origSettingsFile As String
Public RDinstalled As String
Public RD86installed As String
Public RDregistryPresent As Boolean
Public rocketDockInstalled As Boolean
Public rdAppPath As String

' .02 STARTS DAEB 25/01/2021 Moved from mdlmain.bas to common to ensure the checkSteamyDockInstalled subroutine can be run from anywhere, specifically for the variable sdAppPath
Public sdAppPath As String
Public SDinstalled As String
Public SD86installed As String
Public dockAppPath As String
Public steamyDockInstalled As Boolean
Public defaultDock As Integer
' .02 ENDS DAEB 25/01/2021 Moved from mdlmain.bas to common to ensure the checkSteamyDockInstalled subroutine can be run from anywhere, specifically for the variable sdAppPath

Public rdIconCount As Integer
Public requiresAdmin As Boolean

Public rDRunAppInterval As String
Public rDAlwaysAsk As String
Public rDGeneralReadConfig As String
Public rDGeneralWriteConfig As String
Public rDSkinTheme As String
Public rDDefaultDock As String

Public rDLockIcons As String
Public rDRetainIcons As String ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
Public rDOpenRunning As String
Public rDShowRunning As String
Public rDManageWindows As String
Public rDDisableMinAnimation As String

Public sDDockSettingsDefaultEditor As String
Public sDIconSettingsDefaultEditor As String
Public sDDockDefaultEditor As String

Public rDDebugFlg As String

Public sixtyFourBit As Boolean
Public rDCustomIconFolder As String
Public classicThemeCapable As Boolean

Private lstDevices(1, 25) As String
Private lstDevicesListCount As Integer
Public sAllDrives As String



' Steamydock global configuration variables END

' APIs for useful functions START
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
' APIs for useful functions END

' APIs and variables for querying running processes' paths START
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const API_NULL As Long = 0

Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
' APIs and variables for querying running processes' paths ENDS



' APIs and variables for querying running processes' paths ENDS
Private Declare Function QueryDosDeviceW Lib "kernel32.dll" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Private Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
' APIs and variables for querying running processes' paths ENDS


Public storeWindowHwnd As Long '.nn

' .05 DAEB 01/04/2021 common.bas Added declaration to allow replacement of some modal msgbox with the non-modal versions
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

' Flag for debug mode '.06 DAEB 19/04/2021 common.bas moved to the common area so that it can be used by each of the utilities
Private mbDebugMode As Boolean ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without



'------------------------------------------------------ STARTS
Private Const TIME_ZONE_ID_DAYLIGHT As Integer = 2

' Types for determining the timezone
Private Type SYSTEMTIME
    wYear                   As Integer
    wMonth                  As Integer
    wDayOfWeek              As Integer
    wDay                    As Integer
    wHour                   As Integer
    wMinute                 As Integer
    wSecond                 As Integer
    wMilliseconds           As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias                    As Long
    StandardName(63)        As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(63)        As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type

' APIs for determining the timezone
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long ' Always ignore the returned value, it's useless.
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'------------------------------------------------------ ENDS

Public msgBoxOut As Boolean
Public msgLogOut As Boolean

Public windowsVersionString As String
Public sDShowIconSettings As String ' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock


'------------------------------------------------------ STARTS
' For determining dir existence
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Constants for playing sounds
Public Const SND_ASYNC As Long = &H1         '  play asynchronously
Public Const SND_FILENAME As Long = &H20000     '  name is a file name

' APIs for playing sounds
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'------------------------------------------------------ ENDS

'
'---------------------------------------------------------------------------------------
' Procedure : checkLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Sub checkLicenceState()
    Dim slicence As String: slicence = "0"

    On Error GoTo checkLicenceState_Error
    If debugflg = 1 Then debugLog "%" & " sub checkLicenceState"

    'toolSettingsFile = App.Path & "\settings.ini"
    ' read the tool's own settings file (
    If fFExists(toolSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = GetINISetting("Software\DockSettings", "Licence", toolSettingsFile)
        ' if the licence state is not already accepted then display the licence form
        
        If slicence = "0" Then
            Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
            
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkLicenceState of Form common"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : beededea
' Date      : 26/08/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function LoadFileToTB(ByRef TxtBox As Object, ByVal FilePath As String, Optional ByVal Append As Boolean = False) As Boolean
       
    'PURPOSE: Loads file specified by FilePath into textcontrol
    '(e.g., Text Box, Rich Text Box) specified by TxtBox
    
    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    'Returns: True if Successful, false otherwise
    
    Dim iFile As Integer: iFile = 0
    Dim s As String: s = vbNullString
    
   On Error GoTo LoadFileToTB_Error
      'If debugFlg = 1 Then debugLog "%" & "LoadFileToTB"
   
   
   'If debugFlg = 1 Then debugLog "%" & LoadFileToTB

    If Dir$(FilePath) = vbNullString Then Exit Function
    
    On Error GoTo ErrorHandler:
    s = TxtBox.Text
    
    iFile = FreeFile
    Open FilePath For Input As #iFile
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFileToTB of Form common"

End Function

'---------------------------------------------------------------------------------------
' Procedure : savestring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savestring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long: keyhand = 0
    Dim R As Long: R = 0
    
    On Error GoTo savestring_Error

    R = RegCreateKey(hKey, strPath, keyhand)
    R = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    R = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

savestring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savestring of Module Common"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getstring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long: keyhand = 0
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim rvar As Integer: rvar = 0
    
    'in .NET the variant type will need to be replaced by object? This code will go altogether as .NET has native functions to read the registry

    Dim lValueType As Variant ' cannot initialise

    On Error GoTo getstring_Error

    rvar = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

getstring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getstring of Module Common"
End Function
'----------------------------------------
'Name: testWindowsVersion
'Description:
'----------------------------------------
Public Sub testWindowsVersion(ByRef classicThemeCapable As Boolean)

    '=================================
    '2000 / XP / NT / 7 / 8 / 10
    '=================================
    On Error GoTo testWindowsVersion_Error

    ' variables declared
    
    Dim ProgramFilesDir As String: ProgramFilesDir = vbNullString
    Dim strString As String: strString = vbNullString
    Dim prg As String: prg = vbNullString
    
    ' other variable assignments
    classicThemeCapable = False
    windowsVersionString = vbNullString

    strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    windowsVersionString = strString
    requiresAdmin = False
    
    ' ******  note that when running in compatibility mode the o/s will respond with "Windows XP"
    ' ******  The IDE runs in compatibility mode so it will report the wrong version and thence the incorrect working folder
    
    'MsgBox windowsVersionString
    
    If debugflg = 1 Then debugLog "%" & " sub classicThemeCapable"

    'Get the value of "ProgramFiles", or "ProgramFilesDir"
    
    Select Case windowsVersionString
    Case "Microsoft Windows NT4"
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows 2000"
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows XP"
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows 2003"
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Vista"
        requiresAdmin = True
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft 7"
        requiresAdmin = True
        classicThemeCapable = True
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case Else ' Windows 8/10/11+
        requiresAdmin = True
        classicThemeCapable = False
        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    End Select

    'MsgBox strString
    
    prg = Environ$("ProgramFiles")

    ProgramFilesDir = strString
    If ProgramFilesDir = vbNullString Then ProgramFilesDir = prg ' 64bit systems
    If Not fDirExists(ProgramFilesDir) Then
        ProgramFilesDir = "c:\program files" ' 32 bit systems
    End If
    
    'If debugFlg = 1 Then debugLog "%" & "ProgramFilesDir = " & ProgramFilesDir
    


    '======================================================
    'END routine error handler
    '======================================================

   
    On Error GoTo 0: Exit Sub

testWindowsVersion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testWindowsVersion of Module Common"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : GetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo GetINISetting_Error
    Const cparmLen = 500 ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen ' not going to initialise this with a 500 char string
    Dim sDefault As String * cparmLen
    Dim lLength As Long: lLength = 0

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    GetINISetting = Mid$(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

GetINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetINISetting of Module Common"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : PutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save INI Setting in the File
'---------------------------------------------------------------------------------------
'
Public Sub PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo PutINISetting_Error

    Dim aLength As Long: aLength = 0
    
    aLength = WritePrivateProfileString(sHeading, sKey _
            , sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

PutINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutINISetting of Module Common"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByRef OrigFile As String) As Boolean
    'Dim FS As Object ' not going to initialise an object here
    
    On Error GoTo fFExists_Error
   'If debugFlg = 1 Then debugLog "%fFExists"

'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fFExists = FS.FileExists(OrigFile)
    
    ' test to see if a file exists
    Const INVALID_HANDLE_VALUE = -1&
    fFExists = Not (GetFileAttributesW(StrPtr(OrigFile)) = INVALID_HANDLE_VALUE)

   On Error GoTo 0
   Exit Function

fFExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByRef OrigFile As String) As Boolean
    'Dim FS As Object ' not going to initialise an object here
    On Error GoTo fDirExists_Error
   'If debugFlg = 1 Then debugLog "%fDirExists"

'    Set FS = CreateObject("Scripting.FileSystemObject")
'    fDirExists = FS.FolderExists(OrigFile)

   fDirExists = (GetFileAttributes(OrigFile) And vbDirectory + vbVolume) = vbDirectory
   
   On Error GoTo 0
   Exit Function

fDirExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of Module Common"
End Function
'---------------------------------------------------------------------------------------
' Procedure : SpecialFolder
' Author    :  si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SpecialFolder(pFolder As eSpecialFolders) As String
'Returns the path to the specified special folder (AppData etc)

Dim objShell  As Object  ' not going to initialise an object here
Dim objFolder As Object

   On Error GoTo SpecialFolder_Error
   'If debugFlg = 1 Then debugLog "%SpecialFolder"

  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.NameSpace(CLng(pFolder))

  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.Path

  Set objFolder = Nothing
  Set objShell = Nothing

  If SpecialFolder = vbNullString Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"

   On Error GoTo 0
   Exit Function

SpecialFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SpecialFolder of Module Common"

End Function

'---------------------------------------------------------------------------------------
' Procedure : checkAndKill
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : Find and kill any given process name
'           : This routine is an analog of checkAndKillPutWindowBehind. It is more or less identical and you should keep them in synch.
'             This version does NOT have calls to routines that require additional API calls
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for DockSettings and Enhance Icon Settings to run that it would not share with the other utilities.
'---------------------------------------------------------------------------------------
'
Public Function checkAndKill(ByRef NameProcess As String, ByVal checkForFolder As Boolean, ByVal confirmEachProcessKill As Boolean) As Boolean

    ' variables declared
    Dim AppCount As Integer: AppCount = 0
    Dim RProcessFound As Long: RProcessFound = 0
    Dim SzExename As String: SzExename = vbNullString
    Dim MyProcess As Long: MyProcess = 0
    Dim i As Integer: i = 0
    Dim binaryName As String: binaryName = vbNullString
    Dim folderName As String: folderName = vbNullString
    Dim procId As Long: procId = 0
    Dim runningProcessFolder As String: runningProcessFolder = vbNullString
    Dim processToKill As Long: processToKill = 0
    Dim ExitCode As Long: ExitCode = 0
    
    On Error GoTo checkAndKill_Error
    'If debugFlg = 1 Then debugLog "%checkAndKill"

    checkAndKill = False
    MyProcess = GetCurrentProcessId()
    
    If NameProcess <> vbNullString Then
          AppCount = 0
          
          binaryName = getFileNameFromPath(NameProcess)
          If binaryName = vbNullString Then Exit Function ' catchall to prevent closure of unknown processes if the name is malformed
           
          folderName = getFolderNameFromPath(NameProcess)
          
          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

          'hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            'WinDirEnv = Environ("Windir") + "\"
            'WinDirEnv = LCase$(WinDirEnv)

            If Right$(SzExename, Len(binaryName)) = LCase$(binaryName) Then

                    AppCount = AppCount + 1
                    processToKill = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                    If uProcess.th32ProcessID = MyProcess Then
                       'MsgBox "hmmm" & MyProcess ' we never want to kill our own process...
                    Else
                        If checkForFolder = True Then ' only check the process actual run folder when killing an app from the dock
                            procId = uProcess.th32ProcessID ' actual PID
                            runningProcessFolder = getFolderNameFromPath(getExePathFromPID(procId))
                            If LCase$(runningProcessFolder) = LCase$(folderName) Then
                                ' checkAndKill = TerminateProcess(processToKill, ExitCode)
                                ' Call CloseHandle(processToKill)
                                checkAndKill = confirmEachKill(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                            End If
                        Else ' just go ahead and kill whatever process I say must go
                            ' checkAndKill = TerminateProcess(processToKill, ExitCode)
                            ' Call CloseHandle(processToKill)
                            checkAndKill = confirmEachKill(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                        End If
                    End If
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
            
          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
    End If


   On Error GoTo 0
   Exit Function

checkAndKill_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkAndKill of Module Common"

End Function


'---------------------------------------------------------------------------------------
' Procedure : getExePathFromPID
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : getting the full path of a running process is not as easy as you'd expect
'---------------------------------------------------------------------------------------
'
Public Function getExePathFromPID(ByVal idProc As Long) As String
    Dim sBuf As String:  sBuf = vbNullString
    Dim sChar As Long: sChar = 0
    Dim useloop As Integer: useloop = 0
    Dim hProcess As Long: hProcess = 0
    
    On Error GoTo getExePathFromPID_Error

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
    If hProcess Then
        sBuf = String$(260, vbNullChar)
        sChar = GetProcessImageFileName(hProcess, sBuf, 260)
        If sChar Then
            sBuf = NoNulls(sBuf)
            ' this loop replaces the internal windows volume name with the legacy naming convention, ie. C:\, D:\ &c
            For useloop = 1 To lstDevicesListCount
                If InStr(1, sBuf, lstDevices(1, useloop)) > 0 Then
                    sBuf = Replace(sBuf, lstDevices(1, useloop), Chr$(lstDevices(0, useloop)) & ":")
                    Exit For
                End If
            Next useloop
            getExePathFromPID = sBuf
        End If
        CloseHandle hProcess
    End If

   On Error GoTo 0
   Exit Function

getExePathFromPID_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getExePathFromPID of Module common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : NoNulls
' Author    : beededea
' Date      : 25/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NoNulls(ByVal Strng As String) As String
    Dim i As Integer: i = 0
    On Error GoTo NoNulls_Error

    If Len(Strng) > 0 Then
        i = InStr(Strng, vbNullChar)
        Select Case i
            Case 0
                NoNulls = Strng
            Case 1
                NoNulls = vbNullString
            Case Else
                NoNulls = Left$(Strng, i - 1)
        End Select
    End If

   On Error GoTo 0
   Exit Function

NoNulls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NoNulls of Module common"
End Function
'---------------------------------------------------------------------------------------
' Procedure : Is64bit
' Author    : Spider Harper
' Date      : 04/07/2020
' Purpose   : is this program running on a 64bit system or not?
'---------------------------------------------------------------------------------------
'
Public Function Is64bit() As Boolean
    
    ' variables declared
    Dim handle As Long: handle = 0
    Dim bolFunc As Boolean: bolFunc = False
        
    ' Assume initially that this is not a Wow64 process
    On Error GoTo Is64bit_Error

    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

   On Error GoTo 0
   Exit Function

Is64bit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Is64bit of Module Common"

End Function








'---------------------------------------------------------------------------------------
' Procedure : ExtractSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ExtractSuffix(ByVal strPath As String) As String

    ' variables declared
    Dim AY() As String ' string array
    Dim Max As Integer: Max = 0
    
    On Error GoTo ExtractSuffix_Error
    'If debugFlg = 1 Then debugLog "%" & "ExtractSuffix"
   
    If strPath = vbNullString Then
        ExtractSuffix = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        AY = Split(strPath, ".")
        Max = UBound(AY)
        ExtractSuffix = AY(Max)
    Else
        ExtractSuffix = ""
    End If

   On Error GoTo 0
   Exit Function

ExtractSuffix_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffix of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFolderNameFromPath
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function getFolderNameFromPath(ByRef Path As String) As String

   On Error GoTo getFolderNameFromPath_Error
   'If debugFlg = 1 Then debugLog "%" & "getFolderNameFromPath"

    If InStrRev(Path, "\") = 0 Then
        getFolderNameFromPath = vbNullString
        Exit Function
    End If
    getFolderNameFromPath = Left$(Path, InStrRev(Path, "\") - 1)

   On Error GoTo 0
   Exit Function

getFolderNameFromPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFolderNameFromPath of Module Common"
End Function



'
'---------------------------------------------------------------------------------------
' Procedure : getFileNameFromPath
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : A function to getFileNameFromPath
'
'---------------------------------------------------------------------------------------
'
Public Function getFileNameFromPath(ByRef strFullPath As String) As String
   On Error GoTo getFileNameFromPath_Error
      
   ' returns the remainder of the path from the final backslash which can be a file or a folder
   If Not fFExists(strFullPath) Then ' tests to see if a file or a folder of the same name in the same location
        getFileNameFromPath = ""    ' if a file does not exist then what remains must be a folder
        Exit Function               ' if a file does exist get its name below
   End If
   getFileNameFromPath = Right$(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

   On Error GoTo 0
   Exit Function

getFileNameFromPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFileNameFromPath of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExtractSuffixWithDot
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ExtractSuffixWithDot(ByVal strPath As String) As String

    ' variables declared
    Dim AY() As String ' string array
    Dim Max As Integer:    Max = 0
    
    On Error GoTo ExtractSuffixWithDot_Error
    'If debugFlg = 1 Then debugLog "%" & "ExtractSuffixWithDot"
   
    If strPath = vbNullString Then
        ExtractSuffixWithDot = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        AY = Split(strPath, ".")
        Max = UBound(AY)
        ExtractSuffixWithDot = "." & AY(Max)
    Else
        ExtractSuffixWithDot = vbNullString
    End If

   On Error GoTo 0
   Exit Function

ExtractSuffixWithDot_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffixWithDot of Module Common"
End Function




'---------------------------------------------------------------------------------------
' Procedure : driveCheck
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check for the existence of the dock binary
'---------------------------------------------------------------------------------------
'
Public Function driveCheck(ByRef folder As String, Filename As String) As String
   
   ' variables declared
   Dim sDrv As String: sDrv = vbNullString
   Dim sDrives() As String: ' sDrives() = vbNullString
   Dim cnt As Long: cnt = 0
   Dim folderString As String: folderString = vbNullString
   Dim testAppPath As String: testAppPath = vbNullString
   
  'get the list of all drives
   On Error GoTo driveCheck_Error
   'If debugFlg = 1 Then debugLog "%" & "driveCheck"

   'sAllDrives = GetDriveString() ' redundant call - now happens using getdrives at form init
    
  'Change nulls to spaces, then trim.
  'This is required as using Split()
  'with Chr$(0) alone adds two additional
  'entries to the array drives at the end
  'representing the terminating characters.
   sAllDrives = Replace$(sAllDrives, Chr$(0), Chr$(32))
   sDrives() = Split(Trim$(sAllDrives), Chr$(32))
    
    For cnt = LBound(sDrives) To UBound(sDrives)
        sDrv = sDrives(cnt)
        ' on 32bit windows the folder is "Program Files\steamydock"
        folderString = sDrv & folder
        If fDirExists(folderString) = True Then
           'test for the steamydock binary
            testAppPath = folderString
            If fFExists(testAppPath & "\" & Filename) Then
                'MsgBox "steamydock binary exists"
                driveCheck = testAppPath
                Exit Function
            End If
        End If
    Next

   On Error GoTo 0
   Exit Function

driveCheck_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure driveCheck of Module Common"
   
End Function



'these functions need to be in a BAS module and not a form or the AddressOf does not work.


'---------------------------------------------------------------------------------------
' Procedure : BrowseCallbackProc
' Author    : beededea
' Date      : 20/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function BrowseCallbackProc(ByVal hwnd&, ByVal Msg&, ByVal lp&, ByVal initDir$) As Long
   Const BFFM_INITIALIZED As Long = 1
   Const BFFM_SETSELECTION As Long = &H466
   On Error GoTo BrowseCallbackProc_Error

   If (Msg = BFFM_INITIALIZED) And (initDir <> vbNullString) Then
      Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal initDir$)
   End If
   BrowseCallbackProc = 0

   On Error GoTo 0
   Exit Function

BrowseCallbackProc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BrowseCallbackProc of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAddress
' Author    : beededea
' Date      : 20/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetAddress(ByVal Addr As Long) As Long
   On Error GoTo GetAddress_Error

   GetAddress = Addr

   On Error GoTo 0
   Exit Function

GetAddress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetAddress of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : BrowseFolder
' Author    : beededea
' Date      : 20/08/2020
' Purpose   : create a folder window using APIs
'---------------------------------------------------------------------------------------
'
Public Function BrowseFolder(ByVal hWndOwner&, ByVal DefFolder As String) As String
   Dim bi As BROWSEINFO
   Dim pidl As Long: pidl = 0
   Dim newPath As String: newPath = vbNullString

   On Error GoTo BrowseFolder_Error

   bi.hWndOwner = hWndOwner
   bi.lpfn = GetAddress(AddressOf BrowseCallbackProc)
   bi.lParam = StrPtr(DefFolder)
   pidl = SHBrowseForFolderA(bi)
   If (pidl) Then
      newPath = String$(260, 0)
      If SHGetPathFromIDListA(pidl, newPath) Then
         newPath = Left$(newPath, InStr(1, newPath, Chr(0)) - 1)
         BrowseFolder = newPath
      End If
      Call CoTaskMemFree(ByVal pidl&)
   End If

   On Error GoTo 0
   Exit Function

BrowseFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BrowseFolder of Module Common"
End Function



'---------------------------------------------------------------------------------------
' Procedure : writeIconSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub writeIconSettingsIni(ByVal location As String, ByVal iconNumberToWrite As Integer, settingsFile As String)
'                                                   ^^^^^ not byval on the programs - so possible DEBUG
    'Writes an .INI File (SETTINGS.INI)
    
   On Error GoTo writeIconSettingsIni_Error
   'If debugFlg = 1 Then debugLog "%writeIconSettingsIni"


        PutINISetting location, iconNumberToWrite & "-FileName", sFilename, settingsFile
        PutINISetting location, iconNumberToWrite & "-FileName2", sFileName2, settingsFile
        PutINISetting location, iconNumberToWrite & "-Title", sTitle, settingsFile
        PutINISetting location, iconNumberToWrite & "-Command", sCommand, settingsFile
        PutINISetting location, iconNumberToWrite & "-Arguments", sArguments, settingsFile
        PutINISetting location, iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, settingsFile
        PutINISetting location, iconNumberToWrite & "-ShowCmd", sShowCmd, settingsFile
        PutINISetting location, iconNumberToWrite & "-OpenRunning", sOpenRunning, settingsFile
        PutINISetting location, iconNumberToWrite & "-RunElevated", sRunElevated, settingsFile
        
        PutINISetting location, iconNumberToWrite & "-IsSeparator", sIsSeparator, settingsFile
        PutINISetting location, iconNumberToWrite & "-UseContext", sUseContext, settingsFile
        PutINISetting location, iconNumberToWrite & "-DockletFile", sDockletFile, settingsFile
       
        'If defaultDock = 1 Then
        PutINISetting location, iconNumberToWrite & "-UseDialog", sUseDialog, settingsFile
        PutINISetting location, iconNumberToWrite & "-UseDialogAfter", sUseDialogAfter, settingsFile ' .03 DAEB 31/01/2021 common.bas Added new checkbox to determine if a post initiation dialog should appear
        PutINISetting location, iconNumberToWrite & "-QuickLaunch", sQuickLaunch, settingsFile ' .10 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
        PutINISetting location, iconNumberToWrite & "-AutoHideDock", sAutoHideDock, settingsFile  ' .12 DAEB 20/05/2021 common.bas Added new check box to allow autohide of the dock after launch of the chosen app
        PutINISetting location, iconNumberToWrite & "-SecondApp", sSecondApp, settingsFile  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
 
        PutINISetting location, iconNumberToWrite & "-RunSecondAppBeforehand", sRunSecondAppBeforehand, settingsFile
        PutINISetting location, iconNumberToWrite & "-AppToTerminate", sAppToTerminate, settingsFile
        PutINISetting location, iconNumberToWrite & "-Disabled", sDisabled, settingsFile  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
        
        
        
       On Error GoTo 0
   Exit Sub

writeIconSettingsIni_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Common"
    
End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : readIconSettingsIni
'' Author    : beededea
'' Date      : 15/06/2022
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub readIconSettingsIni(location As String, ByVal iconNumberToWrite As Integer, settingsFile As String)
''
'    On Error GoTo readIconSettingsIni_Error
'
'            sFilename = GetINISetting(location, iconNumberToWrite & "-FileName", settingsFile)
'            sFileName2 = GetINISetting(location, iconNumberToWrite & "-FileName2", settingsFile)
'            sTitle = GetINISetting(location, iconNumberToWrite & "-Title", settingsFile)
'            sCommand = GetINISetting(location, iconNumberToWrite & "-Command", settingsFile)
'            sArguments = GetINISetting(location, iconNumberToWrite & "-Arguments", settingsFile)
'            sWorkingDirectory = GetINISetting(location, iconNumberToWrite & "-WorkingDirectory", settingsFile)
'            sShowCmd = GetINISetting(location, iconNumberToWrite & "-ShowCmd", settingsFile)
'            sOpenRunning = GetINISetting(location, iconNumberToWrite & "-OpenRunning", settingsFile)
'            sIsSeparator = GetINISetting(location, iconNumberToWrite & "-IsSeparator", settingsFile)
'            sUseContext = GetINISetting(location, iconNumberToWrite & "-UseContext", settingsFile)
'            sDockletFile = GetINISetting(location, iconNumberToWrite & "-DockletFile", settingsFile)
'
'            sUseDialog = GetINISetting(location, iconNumberToWrite & "-UseDialog", settingsFile)
'            sUseDialogAfter = GetINISetting(location, iconNumberToWrite & "-UseDialogAfter", settingsFile)
'            sQuickLaunch = GetINISetting(location, iconNumberToWrite & "-QuickLaunch", settingsFile)
'            sAutoHideDock = GetINISetting(location, iconNumberToWrite & "-AutoHideDock", settingsFile)
'            sSecondApp = GetINISetting(location, iconNumberToWrite & "-SecondApp", settingsFile)
'
'    On Error GoTo 0
'    Exit Sub
'
'readIconSettingsIni_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconSettingsIni of Module common"
'            Resume Next
'          End If
'    End With
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkRocketdockInstallation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : we check to see if rocketdock is installed in order to know the location of the settings.ini file used by Rocketdock
'             *** also sets rdAppPath during the drivecheck ***
'---------------------------------------------------------------------------------------

Public Sub checkRocketdockInstallation()
    RD86installed = vbNullString
    RDinstalled = vbNullString
    
    ' check where rocketdock is installed
    On Error GoTo checkRocketdockInstallation_Error
    If debugflg = 1 Then debugLog "% sub checkRocketdockInstallation"

    RD86installed = driveCheck("Program Files (x86)\Rocketdock", "RocketDock.exe")
    RDinstalled = driveCheck("Program Files\Rocketdock", "RocketDock.exe")

    If RDinstalled = vbNullString And RD86installed = vbNullString Then
        rocketDockInstalled = False
    Else
        rocketDockInstalled = True
        If RDinstalled <> vbNullString Then
            rdAppPath = RDinstalled
        End If
        'the one in the x86 folder has precedence
        If RD86installed <> vbNullString Then
            rdAppPath = RD86installed
        End If
    End If
    
    ' If rocketdock Is Not installed Then test the registry
    ' if the registry settings are not located then remove them as a source.
    
    ' you might think this stuff is better placed in the docksettings utility
    ' but it has to be here as well as SteamyDock is the component that is most likely to br run first.
    
    ' rocketDockInstalled = False ' debug
    
    ' read selected random entries from the registry, if each are false then the RD registry entries do not exist.
    If rocketDockInstalled = False Then
        rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIcons")
        'rDRetainIcons unused by Rocketdock
        rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunning")
        rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunning")
        rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindows")
        rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimation")
        If rDLockIcons = vbNullString And rDOpenRunning = vbNullString And rDShowRunning = vbNullString And rDManageWindows = vbNullString And rDDisableMinAnimation = vbNullString Then
            ' rocketdock registry entries do not exist so RD has never been installed or it has been wiped entirely.
            RDregistryPresent = False
        Else
            RDregistryPresent = True 'rocketdock HAS been installed in the past as the registry entries are still present
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

checkRocketdockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Module Common"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : readRegistryIconValues
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryIconValues(ByVal iconNumberToRead As Integer)
    ' read the settings from the registry
    On Error GoTo readRegistryOnce_Error
    'If debugFlg = 1 Then debugLog "%" & "readRegistryOnce"

    sFilename = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName")
    sFileName2 = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-FileName2")
    sTitle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Title")
    sCommand = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Command")
    sArguments = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Arguments")
    sWorkingDirectory = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-WorkingDirectory")
    sShowCmd = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-ShowCmd")
    sOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-OpenRunning")
    sRunElevated = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-RunElevated")
    
    sIsSeparator = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-IsSeparator")
    sUseContext = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseContext")
    sDockletFile = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-DockletFile")

    'If defaultDock = 1 Then
    sUseDialog = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseDialog")
    sUseDialogAfter = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-UseDialogAfter")
    sQuickLaunch = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-QuickLaunch") ' .10 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
    sAutoHideDock = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-AutoHideDock")   ' .12 DAEB 20/05/2021 common.bas Added new check box to allow autohide of the dock after launch of the chosen app
    sSecondApp = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-SecondApp")   ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
    
    sRunSecondAppBeforehand = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-RunSecondAppBeforehand")
    sAppToTerminate = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-AppToTerminate")

    sDisabled = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToRead & "-Disabled")
    
   On Error GoTo 0
   Exit Sub

readRegistryOnce_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryOnce of Module Common"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : getAllDriveNames
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : put the device names in an accessible list so that they can be mapped later
'             used especially to obtain the unexpectedly hard-to-extract default folder name of a process in the function isRunning
'---------------------------------------------------------------------------------------
'
Public Sub getAllDriveNames(sDriveStrings As String)
    Dim vDrive As Variant ' probably handling this already in .NET
    Dim sDeviceName As String: sDeviceName = vbNullString
    Dim thiskey As String: thiskey = vbNullString
    Dim driveCount As Integer: driveCount = 0
    
    
    
    On Error GoTo getAllDriveNames_Error
    
    If debugflg = 1 Then debugLog "% sub sDriveStrings"

    For Each vDrive In GetDrives(sDriveStrings) ' getdrives is a collection of drive name strings C:\, D:\ &c
        sDeviceName = GetNtDeviceNameForDrive(vDrive) ' \Device\HarddiskVolume1 are the default naming conventions for Windows drives
        driveCount = driveCount + 1

        lstDevices(0, driveCount) = Asc(Mid$(vDrive, 1, 1))
        lstDevices(1, driveCount) = sDeviceName
        
    Next
    
    lstDevicesListCount = driveCount ' global variable

   On Error GoTo 0
   Exit Sub

getAllDriveNames_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getAllDriveNames of Module common"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDrives
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : getdrives returns a collection of drive name strings C:\, D:\ &c
'---------------------------------------------------------------------------------------
'
Public Function GetDrives(ByRef sDriveStrings As String) As Collection
    
    Dim colDrives As New Collection
    Dim lSize As Long: lSize = 0
    Dim lR As Long: lR = 0
    Dim iLastPos As Long: iLastPos = 0
    Dim iPos As Long: iPos = 0
    Dim sDrive As String: sDrive = vbNullString

   On Error GoTo GetDrives_Error

   lSize = GetLogicalDriveStringsA(0, ByVal 0&)
   sDriveStrings = String$(lSize + 1, 0)
   lR = GetLogicalDriveStringsA(lSize, ByVal sDriveStrings)
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, sDriveStrings, vbNullChar)
      If Not (iPos = 0) Then
         sDrive = Mid$(sDriveStrings, iLastPos, iPos - iLastPos)
         iLastPos = iPos + 1
      Else
         sDrive = Mid$(sDriveStrings, iLastPos)
      End If
      If Len(sDrive) > 0 Then
         colDrives.Add sDrive
      End If
   Loop While Not (iPos = 0)
   Set GetDrives = colDrives

   On Error GoTo 0
   Exit Function

GetDrives_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDrives of Module common"
   
End Function
    
'---------------------------------------------------------------------------------------
' Procedure : GetNtDeviceNameForDrive
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : \Device\HarddiskVolume1 are the default naming conventions for Windows drives
'---------------------------------------------------------------------------------------
'
Public Function GetNtDeviceNameForDrive(ByVal sDrive As String) As String
    
    Dim bDrive() As Byte
    Dim bResult() As Byte
    Dim lR As Long: lR = 0
    Dim sDeviceName As String: sDeviceName = vbNullString

   On Error GoTo GetNtDeviceNameForDrive_Error

   If Right$(sDrive, 1) = "\" Then
      If Len(sDrive) > 1 Then
         sDrive = Left$(sDrive, Len(sDrive) - 1)
      End If
   End If
   bDrive = sDrive
   
   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
   ReDim bResult(0 To 260 * 2 + 1) As Byte
   
   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bResult(0)), 260)
   If (lR > 2) Then
      sDeviceName = bResult
      sDeviceName = Left$(sDeviceName, lR - 2)
      GetNtDeviceNameForDrive = sDeviceName
   End If

   On Error GoTo 0
   Exit Function

GetNtDeviceNameForDrive_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetNtDeviceNameForDrive of Module common"
   
End Function




' 26/10/2020 .01 rocket1 DAEB Moved function isRunning from Steamydock mdlMain(mdlMain.bas) to a shared common.bas module so that more than just one program can utilise it
'---------------------------------------------------------------------------------------
' Procedure : IsRunning
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : determines if a process is running or not
'---------------------------------------------------------------------------------------
'
Public Function IsRunning(ByRef NameProcess As String, ByRef processID As Long) As Boolean

    Dim AppCount As Integer: AppCount = 0
    Dim RProcessFound As Long: RProcessFound = 0
    Dim SzExename As String: SzExename = vbNullString
    Dim ExitCode As Long: ExitCode = 0
    Dim procId As Long: procId = 0
    Dim a As Integer: a = 0
    Dim i As Integer: i = 0
    Dim binaryName As String: binaryName = vbNullString
    Dim folderName As String: folderName = vbNullString
    Dim runningProcessFolder As String: runningProcessFolder = vbNullString

    On Error GoTo IsRunning_Error
    'If debugFlg = 1 Then debugLog "%IsRunning"
    
    ' ignore a Windows binary that can persist
    If InStr(LCase$(NameProcess), "rundll32.exe") > 0 Then
        IsRunning = False
        Exit Function
    End If

    If NameProcess <> vbNullString Then
            AppCount = 0
                     
            If InStr(NameProcess, "::{") > 0 Then
                IsRunning = False
                Exit Function  ' the target is a CLSID so invalid
            End If

            binaryName = getFileNameFromPath(NameProcess)
            'If binaryName = vbNullString Then Exit Function
            
            folderName = getFolderNameFromPath(NameProcess) ' folder name of the binary in the stored process array
            If binaryName = "" Then
                IsRunning = False
                Exit Function  ' the target is a folder so also invalid
            End If
            
            uProcess.dwSize = Len(uProcess)
            hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
            RProcessFound = ProcessFirst(hSnapshot, uProcess)
            Do
                i = InStr(1, uProcess.szexeFile, Chr$(0))
                SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    
                If Right$(SzExename, Len(binaryName)) = LCase$(binaryName) Then

                        AppCount = AppCount + 1
                        procId = uProcess.th32ProcessID

                        runningProcessFolder = getFolderNameFromPath(getExePathFromPID(procId))
                        
                        ' some processes can only be interrogated when running with admin
                        If runningProcessFolder = vbNullString Then
                                IsRunning = True
                                processID = procId
                        Else
                            If LCase$(runningProcessFolder) = LCase$(folderName) Then
                                IsRunning = True
                                processID = procId
                            Else
                                'MsgBox runningProcessFolder & " " & binaryName
                                IsRunning = False
                            End If
                        End If
                        
'                        If NameProcess = "C:\Program Files\CPUID\CPU-Z\cpuz.exe" Then
'                            If runningProcessFolder = vbNullString Then
'                                MsgBox "error obtaining runningProcessFolder from " & aa
'                            End If
'                        End If
                        
                        'If runningProcessFolder <> vbNullString Then
'                        If LCase$(runningProcessFolder) = LCase$(folderName) Then
'                            IsRunning = True
'                            processID = procId
'                        Else
'                            'MsgBox runningProcessFolder & " " & binaryName
'                            IsRunning = False
'                        End If
                        
                        Exit Function
                End If
                RProcessFound = ProcessNext(hSnapshot, uProcess)
    
            Loop While RProcessFound
            Call CloseHandle(hSnapshot)
    End If


   On Error GoTo 0
   Exit Function

IsRunning_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsRunning of Module common"

End Function

' .02 STARTS DAEB 25/01/2021 Moved from mdlmain.bas to common to ensure the checkSteamyDockInstalled subroutine can be run from anywhere, specifically for the variable sdAppPath
'---------------------------------------------------------------------------------------
' Procedure : checkSteamyDockInstallation
' Author    : beededea
' Date      : 20/06/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub checkSteamyDockInstallation()
        
    ' variables declared
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    SD86installed = vbNullString
    SDinstalled = vbNullString
    
    ' check where SteamyDock is installed
    On Error GoTo checkSteamyDockInstallation_Error
    
    If debugflg = 1 Then debugLog "% sub checkSteamyDockInstallation"

    SD86installed = driveCheck("Program Files (x86)\SteamyDock", "steamyDock.exe")
    SDinstalled = driveCheck("Program Files\SteamyDock", "steamyDock.exe")
    
    If SDinstalled = vbNullString And SD86installed = vbNullString Then ' if both are not found
        steamyDockInstalled = False
        
        'answer = msgBoxAA(" SteamyDock has not been installed in the program files (x86) nor the program files folder on any of the drives on this system, can you please install into the correct folder and retry?", vbYesNo)
        Exit Sub
    Else
        steamyDockInstalled = True

        If SDinstalled <> vbNullString Then
            'MsgBox "SteamyDock is installed in " & SDinstalled
            sdAppPath = SDinstalled
        End If
        'the one in the x86 folder has precedence
        If SD86installed <> vbNullString Then
            'MsgBox "SteamyDock is installed in " & SD86installed
            sdAppPath = SD86installed
        End If
        
        dockAppPath = sdAppPath
        defaultDock = 1
    End If
    

   On Error GoTo 0
   Exit Sub

checkSteamyDockInstallation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkSteamyDockInstallation of Form common"
End Sub
' .02 ENDS DAEB 25/01/2021 Moved from mdlmain.bas to common to ensure the checkSteamyDockInstalled subroutine can be run from anywhere, specifically for the variable sdAppPath

' .04 DAEB 06/03/2021 Moved from main code form to common to ensure the locateDockSettingsFile subroutine is common to all STARTS
'---------------------------------------------------------------------------------------
' Procedure : locateDockSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file
'---------------------------------------------------------------------------------------
'
Public Sub locateDockSettingsFile()
        
    ' variables declared
    Dim dockSettingsDir As String: dockSettingsDir = vbNullString
    
    Dim inputData As String:  inputData = vbNullString
    Dim outputData As String: outputData = vbNullString
    Dim s As String: s = 0
        
    On Error GoTo locateDockSettingsFile_Error
    If debugflg = 1 Then debugLog "% sub locateDockSettingsFile"
    
    ' dock Settings main docksettings.ini
    dockSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\steamyDock" ' just for this user alone
    dockSettingsFile = dockSettingsDir & "\docksettings.ini" ' the third config option for steamydock alone

    'if the folder does not exist then create the folder
    If Not fDirExists(dockSettingsDir) Then
        MkDir dockSettingsDir
    End If
    
    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(dockSettingsFile) Then
    '    if it does not exist
    '    it will read the defaultDocksettings.ini line by line and create the new one, changing any occurrence of [defaultDockLocation]
    '    with the updated actual location

        If fFExists(App.Path & "\defaultDockSettings.ini") Then
            ' read the defaultDocksettings.ini line by line
            
            Open App.Path & "\defaultDockSettings.ini" For Input As #1
            Open dockSettingsFile For Output As #2
            
            Do While Not EOF(1)
                Line Input #1, inputData
                ' change any occurrence of [defaultDockLocation] to sdAppPath
                If InStr(inputData, "[defaultDockLocation]") Then
                    s = Replace(inputData, "[defaultDockLocation]", sdAppPath)
                End If
                Write #2, outputData     ' write the line to the new docksettings.ini
                'debugLog outputData
            Loop
            
            Close #1
            Close #2
            
            'FileCopy App.Path & "\defaultDockSettings.ini", dockSettingsFile
        End If
    End If
    
    'confirm the settings file exists, if not use the version in the app itself
    If Not fFExists(dockSettingsFile) Then
            dockSettingsFile = App.Path & "\settings.ini"
    End If

   On Error GoTo 0
   Exit Sub

locateDockSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure locateDockSettingsFile of Form common"

End Sub
' .04 DAEB 06/03/2021 Moved from main code form to common to ensure the locateDockSettingsFile subroutine is common to all ENDS


'---------------------------------------------------------------------------------------
' Procedure : InIDE
' Author    :
' Date      : 09/02/2021
' Purpose   : checks whether the code is running in the VB6 IDE or not
'---------------------------------------------------------------------------------------
'
Public Function InIDE() As Boolean

   On Error GoTo InIDE_Error

    ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        InIDE = True
    End If

   On Error GoTo 0
   Exit Function

InIDE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InIDE of Form dock"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InDebugMode of Form dock"
End Function


' .08 DAEB 11/05/2021 common.bas Added function to pad a string similar to the VB.NET padRight & padLeft functions.
'---------------------------------------------------------------------------------------
' Procedure : padRight
' Author    : beededea
' Date      : 11/05/2021
' Purpose   : Provides the VB.NET padRight function
'---------------------------------------------------------------------------------------
'
Public Function padRight(ByVal stringToPad As String, ByVal AmountToPad As Integer, ByVal padString As String) As String

    On Error GoTo padRight_Error

    If stringToPad = vbNullString Then padRight = vbNullString: Exit Function
    If AmountToPad = 0 Then padRight = stringToPad: Exit Function
    If padString = vbNullString Then padString = " "
    
    If Len(stringToPad) >= AmountToPad Then
        padRight = stringToPad
        Exit Function
    End If

    ' Pad on right.
    padRight = Left$(stringToPad & String$(AmountToPad, padString), AmountToPad)

    On Error GoTo 0
    Exit Function

padRight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure padRight of Module common"
End Function

' .08 DAEB 11/05/2021 common.bas Added function to pad a string similar to the VB.NET padRight & padLeft functions.
'---------------------------------------------------------------------------------------
' Procedure : padLeft
' Author    : beededea
' Date      : 11/05/2021
' Purpose   : Provides the VB.NET padLeft function
'---------------------------------------------------------------------------------------
'
Public Function padLeft(ByVal stringToPad As String, ByVal AmountToPad As Integer, ByVal padString As String) As String

    On Error GoTo padLeft_Error
    
    If stringToPad = vbNullString Then padLeft = vbNullString: Exit Function
    If AmountToPad = 0 Then padLeft = stringToPad: Exit Function
    If padString = vbNullString Then padString = " "
    
    If Len(stringToPad) >= AmountToPad Then
        padLeft = stringToPad
        Exit Function
    End If

    ' Pad on right.
    padLeft = Left$(stringToPad & String$(AmountToPad, padString), AmountToPad)

    On Error GoTo 0
    Exit Function

padLeft_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure padLeft of Module common"
End Function

' .09 DAEB 11/05/2021 common.bas Added function to align and centre a string so it can appear in a msgbox neatly.
'---------------------------------------------------------------------------------------
' Procedure : align
' Author    : beededea
' Date      : 11/05/2021
' Purpose   : Provides an align and centre function by padding it to the required amount
'---------------------------------------------------------------------------------------
'
Public Function align(ByVal stringToPad As String, ByVal AmountToPad As Integer, ByVal padString As String, ByVal padstyle As String) As String
    Dim Result As String: Result = vbNullString
    Dim result1 As String: result1 = vbNullString
    Dim useloop As Integer: useloop = 0
    Dim paddingString As String: paddingString = vbNullString

    On Error GoTo align_Error
    
    If stringToPad = vbNullString Then align = vbNullString: Exit Function
    If AmountToPad = 0 Then align = stringToPad: Exit Function
    If padString = vbNullString Then padString = " "
    If padstyle = vbNullString Then padstyle = "both"
    
    If Len(stringToPad) >= AmountToPad Then
        align = stringToPad
        Exit Function
    End If

    If padstyle = "right" Then
        ' Pad on left.
        Result = Right$(String(AmountToPad, padString) & stringToPad, AmountToPad)
    End If
    
    If padstyle = "left" Then
        ' Pad on right.
        Result = Left$(stringToPad & String$(AmountToPad, padString), AmountToPad)
    End If
    
    If padstyle = "both" Then
        Result = stringToPad
        paddingString = String$(1, padString)
        ' Pad on both sides, ie. align
        For useloop = 1 To AmountToPad
           result1 = Right$(paddingString & Result, Len(Result) + useloop)
           If Len(Result) >= AmountToPad Then
                Exit For
           End If
           Result = Left$(result1 & paddingString, Len(result1) + useloop)
        Next useloop
    End If
    
    align = Result

    On Error GoTo 0
    Exit Function

align_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure align of Module common"
End Function
'---------------------------------------------------------------------------------------
' Procedure : Append
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Append(ByVal Text As String, ByVal TextToAppend As String, Optional ByVal Delimeter As String = "|") As String
    On Error GoTo Append_Error

    Text = TrimS(Text, Delimeter)
    TextToAppend = TrimS(TextToAppend, Delimeter)
    If Len(Text) = 0 Then
        Append = TextToAppend
    ElseIf Len(TextToAppend) > 0 Then
        Append = Text & Delimeter & TextToAppend
    Else
        Append = Text
    End If

    On Error GoTo 0
    Exit Function

Append_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Append of Module common"
            Resume Next
          End If
    End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimS
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function TrimS(ByVal Text As String, ByRef TextToTrim As String) As String
    On Error GoTo TrimS_Error

    TrimS = Trim$(Text)
    If StartsWith(TrimS, TextToTrim) Then TrimS = Right$(TrimS, Len(TrimS) - Len(TextToTrim))
    If EndsWith(TrimS, TextToTrim) Then TrimS = Left$(TrimS, Len(TrimS) - Len(TextToTrim))

    On Error GoTo 0
    Exit Function

TrimS_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TrimS of Module common"
            Resume Next
          End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Procedure : StartsWith
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function StartsWith(ByVal Text As String, ByVal ToLookFor As String) As Boolean
    On Error GoTo StartsWith_Error

    StartsWith = InStr(Left$(Text, Len(ToLookFor)), ToLookFor) = 1

    On Error GoTo 0
    Exit Function

StartsWith_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StartsWith of Module common"
            Resume Next
          End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Procedure : EndsWith
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EndsWith(ByVal Text As String, ByVal ToLookFor As String) As Boolean
    On Error GoTo EndsWith_Error

    EndsWith = Right$(Text, Len(ToLookFor)) = ToLookFor

    On Error GoTo 0
    Exit Function

EndsWith_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EndsWith of Module common"
            Resume Next
          End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Procedure : debugLog
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub debugLog(ByVal inputStr As String, Optional ByVal msgBoxOutOverride As Boolean)

    Dim FN As Integer: FN = 0
    Dim timestamp As String: timestamp = vbNullString

    On Error GoTo debugLog_Error

    FN = FreeFile

    If msgBoxOut = True And Not msgBoxOutOverride = False Then MsgBox inputStr
    
    timestamp = fGetDateInUniversalFormat
    
    ' write the error to the log file
    If msgLogOut = True Then

        Open App.Path & "\SDDebugOutput.log" For Append As FN
        Print #FN, timestamp & " " & inputStr
        Close FN
        
    End If

    On Error GoTo 0
    Exit Sub

debugLog_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure debugLog of Module common"
            Resume Next
          End If
    End With
End Sub

' universal time format is required for unix systems that we may be chatting with
'fnGetDateInUniversalFormat Austin Hickl http://computer-programming-forum.com/66-vb-controls/6dff1bae05df0a6e.htm
'- formats date in form "1998.12.31 23:59:59.456
'---------------------------------------------------------------------------------------
' Procedure : fGetDateInUniversalFormat
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fGetDateInUniversalFormat() As String
  Dim TimeZoneInfo As TIME_ZONE_INFORMATION
  Dim currentLocaltime As SYSTEMTIME

  'Windows returns the inverse of the bias we need
'  If GetTimeZoneInformation(TimeZoneInfo) = TIME_ZONE_ID_DAYLIGHT Then
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.DaylightBias)
'  Else
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.StandardBias)
'  End If

    On Error GoTo fGetDateInUniversalFormat_Error

  GetSystemTime currentLocaltime
  

  With currentLocaltime
    fGetDateInUniversalFormat = Format$(.wYear, "0000") & "-" & Format(.wMonth, "00") & "-" & Format(.wDay, "00") & " " & Format$(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00") & "." & Right$(Format(.wMilliseconds, "000"), 3) '& " " & FormatTimezoneOffset(currentBias)
  End With

    On Error GoTo 0
    Exit Function

fGetDateInUniversalFormat_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetDateInUniversalFormat of Module common"
            Resume Next
          End If
    End With
End Function 'fnGetDateInUniversalFormat


            



'---------------------------------------------------------------------------------------
' Procedure : confirmEachKill
' Author    : beededea
' Date      : 20/12/2022
' Purpose   : This routine is an analog of confirmEachKillPutWindowBehind. It is more or less identical and you should keep them in synch.
'             This version does NOT have calls to routines that require additional API calls
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for DockSettings and Enhance Icon Settings to run that it would not share with the other utilities.
'---------------------------------------------------------------------------------------
'
Public Function confirmEachKill(ByVal binaryName As String, ByVal procId As Long, ByVal processToKill As String, ByVal confirmEachProcessKill As Boolean, ByRef ExitCode As Long) As Boolean
    Dim goAheadAndKill As Boolean: goAheadAndKill = False
    Dim rmessage As String: rmessage = ""
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim a As Long

    On Error GoTo confirmEachKill_Error

    If confirmEachProcessKill = True Then
        rmessage = "A matching process has been found. Kill this application? - " & binaryName & " with process ID " & procId
        'nswer = MsgBox(rmessage, vbYesNo)
        answer = msgBoxA(rmessage, vbYesNo, "Killing this application", True, "confirmEachKill")

        If answer = vbNo Then
            goAheadAndKill = False
        Else
            goAheadAndKill = True
        End If
    Else
        goAheadAndKill = True
    End If
    
    If goAheadAndKill = True Then
        confirmEachKill = TerminateProcess(processToKill, ExitCode)
        Call CloseHandle(processToKill)
    End If

    On Error GoTo 0
    Exit Function

confirmEachKill_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure confirmEachKill of Module common"
            Resume Next
          End If
    End With
End Function





' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen, see Steamydock for a potential replacement?
'---------------------------------------------------------------------------------------
' Procedure : msgBoxA
' Author    : beededea
' Date      : 20/05/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function msgBoxA(ByVal msgBoxPrompt As String, Optional ByVal msgButton As VbMsgBoxResult, Optional ByVal msgTitle As String, Optional ByVal msgShowAgainChkBox As Boolean = False, Optional ByRef msgContext As String = "none") As Integer
     
    ' set the defined properties of a form
    On Error GoTo msgBoxA_Error

    frmMessage.propMessage = msgBoxPrompt
    frmMessage.propTitle = msgTitle
    frmMessage.propShowAgainChkBox = msgShowAgainChkBox
    frmMessage.propButtonVal = msgButton
    frmMessage.propMsgContext = msgContext
    frmMessage.Display ' run a subroutine in the form that displays the form

    msgBoxA = frmMessage.propReturnedValue

    On Error GoTo 0
    Exit Function

msgBoxA_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure msgBoxA of Module mdlMain"
            Resume Next
          End If
    End With

End Function


' .89 DAEB 13/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public
'---------------------------------------------------------------------------------------
' Procedure : getFileNameAndTitle
' Author    : beededea
' Date      : 02/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub getFileNameAndTitle(ByRef retFileName As String, ByRef retfileTitle As String)
   On Error GoTo getFileNameAndTitle_Error
   If debugflg = 1 Then debugLog "%getFileNameAndTitle"

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

getFileNameAndTitle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getFileNameAndTitle of Form rDIconConfigForm"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : addTargetProgram
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Public Function addTargetProgram(ByVal targetText As String)
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
