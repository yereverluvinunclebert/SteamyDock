Attribute VB_Name = "mdlDeadCode"
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_USERS = &H80000003
'Public Const REG_SZ = 1                          ' Unicode nul terminated string

'Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'
'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long



'Public Enum eSpecialFolders
'  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
'  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
'  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
'  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
'End Enum


'Public rDIconQuality As String
'Public rDIconOpacity As String
'Public rDZoomOpaque      As String
'Public rDIconMin      As String
'Public rDHoverFX      As String
'Public rdIconMax      As String
'Public rDZoomWidth      As String
'Public rDZoomTicks      As String
'
'Public rDMonitor      As String
'Public rDSide      As String
'Public rDzOrderMode      As String
'Public rDOffset      As String
'Public rDvOffset      As String
'
'Public rDtheme      As String
'Public rDThemeOpacity      As String
'Public rDHideLabels      As String
'Public rDFontName      As String
'Public rDFontColor      As String
'
'Public rDFontSize As String
'Public rDFontCharSet  As String
'Public rDFontFlags      As String
'
''Public rDFontStrength      As Boolean
''Public rDFontItalics      As Boolean
'
'Public rDFontShadowColor      As String
'Public rDFontOutlineColor      As String
'Public rDFontOutlineOpacity      As String
'Public rDFontShadowOpacity      As String
'
'Public rDIconActivationFX     As String
'Public rDAutoHide     As String
'Public rDAutoHideTicks     As String
'Public rDAutoHideDelay     As String
'Public rDMouseActivate     As String
'Public rDPopupDelay     As String
'
'Public rDVersion As String
'Public rDCustomIconFolder As String
'Public rDHotKeyToggle As String
'Public rDLangID As String

'Public rDAnimationInterval As String
'Public rDSkinSize As String

'Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)
'Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Public dockSettingsFile As String
'Public origSettingsFile As String
'Public toolSettingsFile  As String
'Public rdAppPath As String
'Public RDinstalled As String
'Public RD86installed As String
'Public rocketDockInstalled As Boolean
'Public RDregistryPresent As Boolean
'Public rdIconCount As Integer
'Public requiresAdmin As Boolean
'Public rDLockIcons As String
'Public rDOpenRunning As String
'Public rDShowRunning As String
'Public rDManageWindows As String
'Public rDDisableMinAnimation As String
'Public rDOptionsTabIndex As String
'Public rDStartupRun As String
'Public rdStartupRunString As String






'Public sFilename  As String
'Public sFileName2  As String
'Public sTitle  As String
'Public sCommand  As String
'Public sArguments  As String
'Public sWorkingDirectory  As String
'Public sShowCmd  As String
'Public sOpenRunning  As String
'Public sIsSeparator  As String
'Public sUseContext  As String
'Public sDockletFile  As String


''FIXIT: Declare 'getstring' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
''----------------------------------------
''Name: getstring
''Description:
''----------------------------------------
'Public Function getstring(hKey As Long, strPath As String, strvalue As String)
'
'
'    ' variables declared
'    Dim keyhand As Long
'    'Dim datatype As Long
'    Dim lResult As Long
'    Dim strBuf As String
'    Dim lDataBufSize As Long
'    Dim intZeroPos As Integer
'    Dim rvar As Integer
'    'in .NET the variant type will need to be replaced by object?
'
'    'FIXIT: Declare 'lValueType' with an early-bound data type                                 FixIT90210ae-R1672-R1B8ZE
'    Dim lValueType As Variant
'
'    'initialise the dimensioned variables
'
'    rvar = RegOpenKey(hKey, strPath, keyhand)
'    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
'    If lValueType = REG_SZ Then
'        strBuf = String$(lDataBufSize, " ")
'        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
'        Dim ERROR_SUCCESS As Variant
'        If lResult = ERROR_SUCCESS Then
'            intZeroPos = InStr(strBuf, Chr$(0))
'            If intZeroPos > 0 Then
'                getstring = Left$(strBuf, intZeroPos - 1)
'            Else
'                getstring = strBuf
'            End If
'        End If
'    End If
'End Function
'
''----------------------------------------
''Name: savestring
''Description:
''----------------------------------------
'Public Sub savestring(hKey As Long, strPath As String, strvalue As String, strData As String)
'
'
'    ' variables declared
'    Dim keyhand As Long
'    Dim R As Long
'
'    'initialise the dimensioned variables
'
'
'    R = RegCreateKey(hKey, strPath, keyhand)
'    R = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
'    R = RegCloseKey(keyhand)
'End Sub

''----------------------------------------
''Name: TestWinVer
''Description:
''----------------------------------------
'Public Sub testWinVer()
'
'    '=================================
'    '2000 / XP / NT / 7 / 8 / 10
'    '=================================
'    On Error GoTo TestWinVer_Error
'
'    ' variables declared
'
'    Dim ProgramFilesDir As String
'    Dim WindowsVer As String
'    Dim strString As String
'    Dim classicThemeCapable As Boolean
'
'    'initialise the dimensioned variables
'    strString = ""
'    classicThemeCapable = False
'    WindowsVer = ""
'    ProgramFilesDir = ""
'
'    ' other variable assignments
'    strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
'    WindowsVer = strString
'    requiresAdmin = False
'
'    ' note that when running in compatibility mode the o/s will respond with "Windows XP"
'    ' The IDE runs in compatibility mode so it may report the wrong working folder
'
'    'MsgBox WindowsVer
'
'    'Get the value of "ProgramFiles", or "ProgramFilesDir"
'
'    Select Case WindowsVer
'    Case "Microsoft Windows NT4"
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    Case "Microsoft Windows 2000"
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    Case "Microsoft Windows XP"
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
'    Case "Microsoft Windows 2003"
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    Case "Microsoft Vista"
'        requiresAdmin = True
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    Case "Microsoft 7"
'        requiresAdmin = True
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    Case Else
'        requiresAdmin = True
'        classicThemeCapable = False
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
'    End Select
'
'    'MsgBox strString
'
'
'    ProgramFilesDir = strString
'    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems
'    If Not DirExists(ProgramFilesDir) Then
'        ProgramFilesDir = "c:\program files" ' 32 bit systems
'    End If
'
'
'
'    '======================================================
'    'END routine error handler
'    '======================================================
'
'
'    On Error GoTo 0: Exit Sub
'
'TestWinVer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TestWinVer of Module WinModule"
'
'End Sub



''----------------------------------------
''Name: FExists
''----------------------------------------
'Public Function FExists(OrigFile As String)
'
'    ' variables declared
'    Dim FS As Object
'
'    'initialise the dimensioned variables
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    FExists = FS.FileExists(OrigFile)
'End Function
'
''----------------------------------------
''Name: DirExists
''----------------------------------------
'Public Function DirExists(OrigFile As String)
'
'    ' variables declared
'    Dim FS As Object
'
'    'initialise the dimensioned variables
'
'
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    DirExists = FS.FolderExists(OrigFile)
'End Function



''API Function to read information from INI File
'Public Declare Function GetPrivateProfileString Lib "kernel32" _
'    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
'    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
'    , ByVal lpFileName As String) As Long
'
''API Function to write information to the INI File
'Private Declare Function WritePrivateProfileString Lib "kernel32" _
'    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
'    , ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    

''Get the INI Setting from the File
''---------------------------------------------------------------------------------------
'' Procedure : GetINISetting
'' Author    : beededea
'' Date      : 10/05/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, sINIFileName) As String
'    Const cparmLen = 500
'
'    ' variables declared
'    Dim sReturn As String * cparmLen
'    Dim sDefault As String * cparmLen
'    Dim lLength As Long
'
'    'initialise the dimensioned variables
'
'   On Error GoTo GetINISetting_Error
'   If debugflg = 1 Then Debug.Print "%GetINISetting"
'
'    lLength = GetPrivateProfileString(sHeading, sKey _
'            , sDefault, sReturn, cparmLen, sINIFileName)
'    GetINISetting = Mid(sReturn, 1, lLength)
'
'   On Error GoTo 0
'   Exit Function
'
'GetINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetINISetting of Module Module2"
'End Function

''Save INI Setting in the File
''---------------------------------------------------------------------------------------
'' Procedure : PutINISetting
'' Author    : beededea
'' Date      : 10/05/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, sINIFileName) As Boolean
'   On Error GoTo PutINISetting_Error
'   If debugflg = 1 Then Debug.Print "%PutINISetting"
'
'    On Error GoTo HandleError
'    Const cparmLen = 500
'
'    ' variables declared
'    Dim sReturn As String * cparmLen
'    Dim sDefault As String * cparmLen
'    Dim aLength As Long
'
'    'initialise the dimensioned variables
'
'    aLength = WritePrivateProfileString(sHeading, sKey _
'            , sSetting, sINIFileName)
'    PutINISetting = True
'    Exit Function
'
'HandleError:
'    DebugPrint Err.Number & " " & Err.Description
'
'   On Error GoTo 0
'   Exit Function
'
'PutINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutINISetting of Module Module2"
'End Function


'
''---------------------------------------------------------------------------------------
'' Procedure : writeIconSettingsIni
'' Author    : beededea
'' Date      : 10/05/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Sub writeIconSettingsIni(location As String, iconNumberToWrite As Integer, settingsFile As String)
'    'Writes an .INI File (SETTINGS.INI)
'
'    'E:\Program Files (x86)\RocketDock\Icons\Steampunk_Clockwerk_Kubrick
'    ' determine relative path TODO
'    ' Icons\Steampunk_Clockwerk_Kubrick
'
'   On Error GoTo writeIconSettingsIni_Error
'   If debugflg = 1 Then Debug.Print "%writeIconSettingsIni"
'
'        PutINISetting location, iconNumberToWrite & "-FileName", sFilename, settingsFile
'        PutINISetting location, iconNumberToWrite & "-FileName2", sFileName2, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Title", sTitle, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Command", sCommand, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Arguments", sArguments, settingsFile
'        PutINISetting location, iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, settingsFile
'        PutINISetting location, iconNumberToWrite & "-ShowCmd", sShowCmd, settingsFile
'        PutINISetting location, iconNumberToWrite & "-OpenRunning", sOpenRunning, settingsFile
'        PutINISetting location, iconNumberToWrite & "-IsSeparator", sIsSeparator, settingsFile
'        PutINISetting location, iconNumberToWrite & "-UseContext", sUseContext, settingsFile
'        PutINISetting location, iconNumberToWrite & "-DockletFile", sDockletFile, settingsFile
'
'    '
'    'Change the above setting to this one
'    'PutINISetting "SQLSERVER", "SERVER", "MyNewSQLServer", SettingsFile
'
'   On Error GoTo 0
'   Exit Sub
'
'writeIconSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Module2"
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : driveCheck
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Function driveCheck(folder As String, filename As String)
'
'   ' variables declared
'   Dim sAllDrives As String
'   Dim sDrv As String
'   Dim sDrives() As String
'   Dim cnt As Long
'   Dim folderString As String
'   Dim testAppPath As String
'
'   'initialise the dimensioned variables
'   sAllDrives = ""
'   sDrv = ""
'   'sDrives() = ""
'   cnt = 0
'   folderString = ""
'   testAppPath = ""
'
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure driveCheck of Form dockSettings"
'
'End Function


''---------------------------------------------------------------------------------------
'' Procedure : GetDriveString
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Function GetDriveString() As String
'
'  'Used by both demos
'
''  'returns string of available
''  'drives each separated by a null
''   Dim sBuff As String
''
''  'possible 26 drives, three characters
''  'each plus a trailing null for each
''  'drive letter and a terminating null
''  'for the string
''
'
'    ' variables declared
'    Dim I As Long
'    Dim builtString As String
'
'    'initialise the dimensioned variables
'    I = 0
'    builtString = ""
'
'
'    '===========================
'    'pure VB approach, no controls required
'    'drive letters are found in positions 1-UBound(Letters)
'    '"C:\ D:\ E:\ &c"
'
'   On Error GoTo GetDriveString_Error
'
'    For I = 1 To 26
'        If ValidDrive(Chr(96 + I)) = True Then
'            builtString = builtString + UCase(Chr(96 + I)) & ":\    "
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriveString of Form dockSettings"
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : ValidDrive
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Function ValidDrive(d As String) As Boolean
'   On Error GoTo ValidDrive_Error
'
'  On Error GoTo driveerror
'
'    ' variables declared
'    Dim Temp As String
'
'    ' initialise the dimensioned variables
'    Temp = ""
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidDrive of Form dockSettings"
'End Function


''---------------------------------------------------------------------------------------
'' Procedure : readRegistry
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : start the reading of the various areas of the registry
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistry()
'
'    ' variables declared
'    Dim useloop As Integer
'
'    'initialise the dimensioned variables
'    useloop = 0
'
'    On Error GoTo readRegistry_Error
'
'        ' current open tab on the dock settings
'         rDOptionsTabIndex = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OptionsTabIndex")
'
'         ' get items from the registry that are required to 'default' the dock but aren't controlled by the dock settings utility
'         rdIconCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Count")
'
'         rDVersion = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Version")
'         rDCustomIconFolder = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "CustomIconFolder")
'         rDHotKeyToggle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "HotKeyToggle")
'
'         ' get the relevant entries from the registry
'         readRegistryGeneral ' to populate the general tab &c &c
'         readRegistryIcons
'         readRegistryBehaviour
'         readRegistryStyle
'         readRegistryPosition
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistry_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistry of Form dockSettings"
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : readRegistryGeneral
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryGeneral()
'    ' read the settings from the registry
'   On Error GoTo readRegistryGeneral_Error
'
'   'general items
'
''   LockIcons ' lock items
''   OpenRunning 'Open Running Application Instance
''   ShowRunning 'Running Application Indicators
''   ManageWindows' Minimise Windows to the Dock
''   DisableMinAnimation 'Disable Minimise Animations
'
''HKEY_USERS\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run
'' 02 00 00 00 00 00 00
'
'    'Dim rdStartupRunString As String
'
'    'general panel
'
'    rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIcons")
'    rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunning")
'    rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunning")
'    rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindows")
'    rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimation")
'
'    Call validateRegistryGeneral
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryGeneral_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryGeneral of Form dockSettings"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryGeneral
'' Author    : beededea
'' Date      : 13/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryGeneral()
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'   On Error GoTo validateRegistryGeneral_Error
'
'    If Val(rDLockIcons) < 0 And Val(rDLockIcons) > 1 Then rDLockIcons = "1" '
'    If Val(rDOpenRunning) < 0 And Val(rDOpenRunning) > 1 Then rDOpenRunning = "1" '
'    If Val(rDShowRunning) < 0 And Val(rDShowRunning) > 1 Then rDShowRunning = "1" '
'    If Val(rDManageWindows) < 0 And Val(rDManageWindows) > 1 Then rDManageWindows = "1" '
'    If Val(rDDisableMinAnimation) < 0 And Val(rDDisableMinAnimation) > 1 Then rDDisableMinAnimation = "1" '
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryGeneral_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryGeneral of Form dockSettings"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryIcons
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryIcons()
'    ' read the settings from the registry
'   On Error GoTo readRegistryIcons_Error
'
'   'icon items
'
''   IconQuality ' Icon Quality
''   IconOpacity ' Icon Opacity
''   ZoomOpaque  ' Zoom Opaque
''   IconMin     ' Size
''   HoverFX     ' Hover Effect
''   IconMax     ' Zoom
''   ZoomWidth   ' Zoom Width
''   ZoomTicks   ' Duration
'
'    rDIconQuality = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconQuality")
'    rDIconOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconOpacity")
'    rDZoomOpaque = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomOpaque")
'    rDIconMin = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconMin")
'    rDHoverFX = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "HoverFX")
'    rdIconMax = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconMax")
'    rDZoomWidth = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomWidth")
'    rDZoomTicks = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomTicks")
'
'    Call validateRegistryIcons
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryIcons_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryIcons of Form dockSettings"
'
'End Sub
'Private Sub validateRegistryIcons()
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'    If Val(rDMonitor) < 0 Or Val(rDMonitor) > 10 Then rDMonitor = "1" 'monitor 1
'    If Val(rDIconOpacity) < 50 Or Val(rDIconOpacity) > 100 Then rDIconOpacity = "100" 'fully opaque
'    If Val(rDZoomOpaque) < 0 Or Val(rDZoomOpaque) > 1 Then rDZoomOpaque = "1" 'zooms opaque
'    If Val(rDIconMin) < 16 Or Val(rDIconMin) > 128 Then rDIconMin = "16" 'small
'    If Val(rDHoverFX) < 0 Or Val(rDHoverFX) > 3 Then rDHoverFX = "1" 'bounce
'
'    If defaultDock = 0 Then ' rocketdock
'        If Val(rdIconMax) < 1 Or Val(rdIconMax) > 128 Then rdIconMax = "128" 'largest size for Rocketdock
'    Else
'        If Val(rdIconMax) < 1 Or Val(rdIconMax) > 256 Then rdIconMax = "256" 'largest size for SteamyDock
'    End If
'
'    If Val(rDZoomWidth) < 2 Or Val(rDZoomWidth) > 10 Then rDZoomWidth = "4" ' just a few expanded
'    If Val(rDZoomTicks) < 100 Or Val(rDZoomTicks) > 500 Then rDZoomTicks = "100" ' 100ms
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : readRegistryPosition
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryPosition()
'    ' read the settings from the registry
'   On Error GoTo readRegistryPosition_Error
'
'   'style items
'
''   Monitor 'Monitor
''   Side    ' Side
''   zOrderMode  ' zOrderMode
''   Offset  ' Offset
''   vOffset ' vOffset
'
'
'    rDMonitor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Monitor")
'    rDSide = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Side")
'    rDzOrderMode = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "zOrderMode")
'    rDOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Offset")
'    rDvOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "vOffset")
'
'    Call validateRegistryPosition
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryPosition_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryPosition of Form dockSettings"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryPosition
'' Author    : beededea
'' Date      : 01/08/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryPosition()
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'   On Error GoTo validateRegistryPosition_Error
'
'    If Val(rDMonitor) < 0 Or Val(rDMonitor) > 10 Then rDMonitor = "1" 'monitor 1
'    If Val(rDSide) < 0 Or Val(rDSide) > 3 Then rDSide = "1" ' bottom
'    If Val(rDzOrderMode) < 1 Or Val(rDzOrderMode) > 10 Then rDzOrderMode = "0" ' always on top
'    If Val(rDOffset) < -100 Or Val(rDOffset) > 100 Then rDOffset = "0" ' in the middle
'    If Val(rDvOffset) < -15 Or Val(rDvOffset) > 128 Then rDvOffset = "0" ' at the bottom edge
'
'
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryPosition_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryPosition of Form dockSettings"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryStyle
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :' read the settings from the registry
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryStyle()
'
'    On Error GoTo readRegistryStyle_Error
'
'   'style items
'
''   Theme 'Theme
''   ThemeOpacity
''   HideLabels
''   FontName
''   FontShadowColor
''   FontOutlineColor
''   FontOutlineOpacity
''   FontShadowOpacity
'
'    rDtheme = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Theme")
'
'    rDThemeOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ThemeOpacity")
'    rDHideLabels = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "HideLabels")
'    rDFontName = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontName")
'    rDFontColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontColor")
'
'    rDFontSize = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontSize")
'    rDFontCharSet = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontCharSet")
'    rDFontFlags = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontFlags")
'
'    rDFontShadowColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowColor")
'    rDFontOutlineColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineColor")
'
'    rDFontOutlineOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineOpacity")
'    rDFontShadowOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowOpacity")
'
'    Call validateRegistryStyle
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryStyle_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryStyle of Form dockSettings"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryStyle
'' Author    : beededea
'' Date      : 13/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryStyle()
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'    ' read the skins available from the rocketdock folder
'
'
'    ' variables declared
'    'Dim MyFile As String
'    Dim MyPath  As String
'    'Dim themePresent As Boolean
'    'Dim myName As String
'
'    'initialise the dimensioned variables
'    'MyFile = ""
'    MyPath = ""
'    'themePresent = False
'    'myName = ""
'
'    On Error GoTo validateRegistryStyle_Error
'
'
'
''    rDFontColor - difficult to check validity of a colour but some code is coming to ensure no corruption *1
'
'    If Val(rDThemeOpacity) < 1 Or Val(rDThemeOpacity) > 100 Then rDThemeOpacity = "100" '
'    If Val(rDHideLabels) < 0 Or Val(rDHideLabels) > 1 Then rDHideLabels = "0" '
'
'
'    ' variables declared
'    Dim I As Integer
'    Dim fontPresent As Boolean
'
'    'initialise the dimensioned variables
'
'
'
'    fontPresent = False
'    For I = 0 To Screen.FontCount - 1 ' Determine number of fonts.
'        If rDFontName = Screen.Fonts(I) Then fontPresent = True
'    Next I
'    If fontPresent = False Then rDFontName = "Times New Roman" '
'
'    If Abs(Val(rDFontSize)) < 2 Or Abs(Val(rDFontSize)) > 29 Then rDFontSize = "-29" '
'
'    ' rDFontCharSet = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontCharSet")
'    ' how to validate a character set? - not supported
'
'
'
'    ' validate font flags
'    ' 0 - no qualifiers or alterations
'    ' 1 - bold
'    ' 2 - light italics
'    ' 3 - bold italics
'    ' 4 - strikeout & light
'    ' 6 - underline and italics
'    ' 7 - bold, italics & underline
'    ' 10 - strikeout & italics
'    ' 11 - bold, italics & strikeout
'    ' 13 - strikeout & italics
'    ' 14 - underline, strikeout and italics
'    ' 15 - bold, underline, strikeout and italics
'
'    If rDFontFlags < 0 Or rDFontFlags > 15 Then rDFontFlags = 0
'
'
'    If Not IsNumeric(rDFontShadowColor) Then
'        rDFontShadowColor = 0
'    End If
'
'    If Not IsNumeric(rDFontOutlineColor) Then
'        rDFontOutlineColor = 0
'    End If
'
'    ' how to validate colour?
'
'    If Val(rDFontOutlineOpacity) < 0 Or Val(rDFontOutlineOpacity) > 100 Then rDFontOutlineOpacity = "100" '
'    If Val(rDFontShadowOpacity) < 0 Or Val(rDFontShadowOpacity) > 100 Then rDFontShadowOpacity = "100" '
'
'    'validation ends
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryStyle_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryStyle of Form dockSettings"
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : validateInputs
'' Author    : beededea
'' Date      : 13/06/2020
'' Purpose   : validate the relevant entries from whichever source
''---------------------------------------------------------------------------------------
''
'Private Sub validateInputs()
'         '
'   On Error GoTo validateInputs_Error
'
'         validateRegistryGeneral
'         validateRegistryIcons
'         validateRegistryBehaviour
'         validateRegistryStyle
'         validateRegistryPosition
'
'   On Error GoTo 0
'   Exit Sub
'
'validateInputs_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of Form dockSettings"
'End Sub

'Public Function IsValidOleColor(ByVal nColor As Long) As Boolean
'  Select Case nColor
'    Case 0& To &H100FFFF, &H2000000 To &H2FFFFFF
'         IsValidOleColor = True
'    Case &H80000000 To &H80FF0018
'         IsValidOleColor = (nColor And &HFFFF&) <= &H18
'  End Select
'End Function
    
'---------------------------------------------------------------------------------------
' Procedure : FindComboIndex
' Author    : beededea
' Date      : 02/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Sub FindComboIndex(ByVal cmbDealerName As ComboBox, ByVal result As String)
'    Dim i As Integer
'   On Error GoTo FindComboIndex_Error
'   If debugflg = 1 Then Debug.Print "%FindComboIndex"
'
'    For i = 0 To cmbDealerName.ListCount - 1
'        If result = cmbDealerName.List(i) Then
'            cmbDealerName.ListIndex = i
'            Exit Sub
'        End If
'    Next i
'
'   On Error GoTo 0
'   Exit Sub
'
'FindComboIndex_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindComboIndex of Form dockSettings"
'End Sub


'
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryBehaviour
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryBehaviour()
'    ' read the settings from the registry
'   On Error GoTo readRegistryBehaviour_Error
'
'   'general items
'
''   IconActivationFX ' Icon Activation Effect
''   AutoHide         ' AutoHide
''   AutoHideTicks    ' AutoHide Duration
''   AutoHideDelay    ' AutoHide Delay
''   MouseActivate    ' Pop-up on Mouseover
''   PopupDelay       ' PopupDelay
'
'
'    rDIconActivationFX = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconActivationFX")
'    rDAutoHide = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHide")
'    rDAutoHideTicks = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHideTicks")
'    rDAutoHideDelay = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutoHideDelay")
'    rDMouseActivate = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "MouseActivate")
'    rDPopupDelay = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "PopupDelay")
'
'    Call validateRegistryBehaviour
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryBehaviour_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryBehaviour of Form dockSettings"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryBehaviour
'' Author    : beededea
'' Date      : 13/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryBehaviour()
'
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'   On Error GoTo validateRegistryBehaviour_Error
'
'    If Val(rDIconActivationFX) < 0 And Val(rDIconActivationFX) > 2 Then rDIconActivationFX = "2"
'    If Val(rDAutoHide) < 0 And Val(rDAutoHide) > 1 Then rDAutoHide = "1"
'    If Val(rDAutoHideTicks) < 0 And Val(rDAutoHideTicks) > 1000 Then rDAutoHideTicks = "200"
'    If Val(rDAutoHideDelay) < 0 And Val(rDAutoHideDelay) > 2000 Then rDAutoHideDelay = "200"
'    If Val(rDMouseActivate) < 0 And Val(rDMouseActivate) > 1 Then rDMouseActivate = "1"
'    If Val(rDPopupDelay) < 0 And Val(rDPopupDelay) > 1000 Then rDPopupDelay = "100"
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryBehaviour_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryBehaviour of Form dockSettings"
'
'    End Sub
' If Rocketdock has been installed we will also write to the registry to keep the two in synch.
' note: Rocketdock does NOT seem do this - it either keeps the Registry updated or the settings.ini file
' this will only operate correctly if this tool has administrator access.

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
'    '(e.g., txtGeneralRdLocation Box, Rich txtGeneralRdLocation Box) specified by TxtBox
'
'    'If Append = true, then loaded text is appended to existing
'    ' contents else existing contents are overwritten
'
'    'Returns: True if Successful, false otherwise
'
'
'    ' variables declared
'    Dim iFile As Integer
'    Dim s As String
'
'    'initialise the dimensioned variables
'    iFile = 0
'    s = ""
'
'    On Error GoTo LoadFileToTB_Error
'    If debugflg = 1 Then DebugPrint "%" & "LoadFileToTB"
'
'    If Dir(FilePath) = "" Then Exit Function
'
'    On Error GoTo ErrorHandler:

''---------------------------------------------------------------------------------------
'' Procedure : readSettingsFile
'' Author    : beededea
'' Date      : 12/05/2020
'' Purpose   : read
''---------------------------------------------------------------------------------------
''
'Private Sub readSettingsFile(location As String, settingsFile As String)
'
'    'SteamyDock settings only
'    On Error GoTo readSettingsFile_Error
'    If debugflg = 1 Then Debug.Print "%readSettingsFile"
'
'    If FExists(dockSettingsFile) Then
'        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
'        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
'        rDRunAppInterval = GetINISetting("Software\SteamyDock\DockSettings", "RunAppInterval", dockSettingsFile)
'        rDAlwaysAsk = GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile)
'        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
'        rDAnimationInterval = GetINISetting("Software\SteamyDock\DockSettings", "AnimationInterval", dockSettingsFile)
'        rDSkinSize = GetINISetting("Software\SteamyDock\DockSettings", "SkinSize", dockSettingsFile)
'    End If
'
'    'if the above settings do not exist in the older RD settings file then no error is thrown so it works for both
'
'    'RocketDock settings only
'    rDVersion = GetINISetting(location, "Version", settingsFile)
'    rDHotKeyToggle = GetINISetting(location, "HotKey-Toggle", settingsFile)
'    rDtheme = GetINISetting(location, "Theme", settingsFile)
'    rDThemeOpacity = GetINISetting(location, "ThemeOpacity", settingsFile)
'    rDIconOpacity = GetINISetting(location, "IconOpacity", settingsFile)
'    rDFontSize = GetINISetting(location, "FontSize", settingsFile)
'    rDFontFlags = GetINISetting(location, "FontFlags", settingsFile)
'    rDFontName = GetINISetting(location, "FontName", settingsFile)
'    rDFontColor = GetINISetting(location, "FontColor", settingsFile)
'    rDFontCharSet = GetINISetting(location, "FontCharSet", settingsFile)
'    rDFontOutlineColor = GetINISetting(location, "FontOutlineColor", settingsFile)
'    rDFontOutlineOpacity = GetINISetting(location, "FontOutlineOpacity", settingsFile)
'    rDFontShadowColor = GetINISetting(location, "FontShadowColor", settingsFile)
'    rDFontShadowOpacity = GetINISetting(location, "FontShadowOpacity", settingsFile)
'    rDIconMin = GetINISetting(location, "IconMin", settingsFile)
'    rdIconMax = GetINISetting(location, "IconMax", settingsFile)
'    rDZoomWidth = GetINISetting(location, "ZoomWidth", settingsFile)
'    rDZoomTicks = GetINISetting(location, "ZoomTicks", settingsFile)
'    rDAutoHide = GetINISetting(location, "AutoHide", settingsFile)
'    rDAutoHideTicks = GetINISetting(location, "AutoHideTicks", settingsFile)
'    rDAutoHideDelay = GetINISetting(location, "AutoHideDelay", settingsFile)
'    rDPopupDelay = GetINISetting(location, "PopupDelay", settingsFile)
'    rDIconQuality = GetINISetting(location, "IconQuality", settingsFile)
'    rDLangID = GetINISetting(location, "LangID", settingsFile)
'    rDHideLabels = GetINISetting(location, "HideLabels", settingsFile)
'    rDZoomOpaque = GetINISetting(location, "ZoomOpaque", settingsFile)
'    rDLockIcons = GetINISetting(location, "LockIcons", settingsFile)
'    rDManageWindows = GetINISetting(location, "ManageWindows", settingsFile)
'    rDDisableMinAnimation = GetINISetting(location, "DisableMinAnimation", settingsFile)
'    rDShowRunning = GetINISetting(location, "ShowRunning", settingsFile)
'    rDOpenRunning = GetINISetting(location, "OpenRunning", settingsFile)
'    rDHoverFX = GetINISetting(location, "HoverFX", settingsFile)
'    rDzOrderMode = GetINISetting(location, "zOrderMode", settingsFile)
'    rDMouseActivate = GetINISetting(location, "MouseActivate", settingsFile)
'    rDIconActivationFX = GetINISetting(location, "IconActivationFX", settingsFile)
'    rDMonitor = GetINISetting(location, "Monitor", settingsFile)
'    rDSide = GetINISetting(location, "Side", settingsFile)
'    rDOffset = GetINISetting(location, "Offset", settingsFile)
'    rDvOffset = GetINISetting(location, "vOffset", settingsFile)
'    rDOptionsTabIndex = GetINISetting(location, "OptionsTabIndex", settingsFile)
'    '= GetINISetting("Software\SteamyDock\DockSettings\WindowFilters", "Count", 0, settingsFile)
'
'   On Error GoTo 0
'   Exit Sub
'
'readSettingsFile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Form dockSettings"
'
'End Sub
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function LoadFileToTB of Form dockSettings"
'
'End Function

