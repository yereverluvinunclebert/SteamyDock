Attribute VB_Name = "Module1"
Option Explicit



'
'Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
'Public debugflg As Integer
'Public toolSettingsFile  As String
'
'Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'Private Declare Function EnumProcesses Lib "psapi.dll" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
'Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
'Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hmodule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'Private Declare Function QueryFullProcessImageName Lib "Kernel32.dll" Alias "QueryFullProcessImageNameA" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As String, lpdwSize As Long) As Long
     
'Private Const PROCESS_VM_READ = &H10
'Private Const PROCESS_QUERY_INFORMATION = &H400

' declare global arrays to make the settings data open to everyone
' they are declmnuOtherOptsared but undimensioned as we need to declare them with
' a size controlled by the rdIconMax variable above, that is done
' during the form_load when the rdIconMax value is defined

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
'Public rdSettingsFile As String
'Public usedMenuFlag As Boolean

'
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
    
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const REG_SZ = 1                          ' Unicode nul terminated string
'
'' main APIs, constants defined for querying the registry
'' some global variables and a few local subroutines/functions
'' pertaining to the main form.
'
'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long

'' functions to determine 64bitness start
'Private Declare Function GetProcAddress Lib "kernel32" _
'    (ByVal hmodule As Long, _
'    ByVal lpProcName As String) As Long
'
'Private Declare Function GetModuleHandle Lib "kernel32" _
'    Alias "GetModuleHandleA" _
'    (ByVal lpModuleName As String) As Long
'
''Private Declare Function GetCurrentProcess Lib "kernel32" _
''    () As Long 'already declared
'
'Private Declare Function IsWow64Process Lib "kernel32" _
'    (ByVal hProc As Long, _
'    bWow64Process As Boolean) As Long
'' functions to determine 64bitness END


'Public WindowsVer As String
'Public requiresAdmin As Boolean
'Public rdAppPath As String
'Public origSettingsFile As String
'Public rdIconMaximum As Integer
'Public theCount As Integer
'Public dockOpacity As Integer
'Public userLevel As String
'Public namesListArray() As String
'Public sCommandArray() As String
'Public autoHideTimerCount As Integer
'Public animationFlg As Boolean
'Public dockLoweredTime As Date
'Public dockHidden As Boolean
'
'Public rocketDockInstalled As Boolean
'Public RDinstalled As String
'Public RD86installed As String
'Public RDregistryPresent As Boolean
'
'Public dockSettingsFile As String
'
'Public rDRunAppInterval As String
'Public rDAlwaysAsk As String
'Public rDGeneralReadConfig As String
'Public rDGeneralWriteConfig As String
'
'Public rDSkinTheme As String
'Public rDDefaultDock As String
'
'
'
'Public lngBitmap As Long
'Public lngImage As Long
'Public lngGDI As Long
'Public lngReturn As Long


'Public Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
'                                             ByVal hProcess As Long, ByVal hmodule As Long, _
'                                             ByVal moduleName As String, ByVal nSize As Long) As Long





'
''---------------------------------------------------------------------------------------
'' Procedure : LoadFileToTB
'' Author    : beededea
'' Date      : 26/08/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function LoadFileToTB(TxtBox As Object, FilePath As _
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
'      If debugflg = 1 Then DebugPrint "%" & "LoadFileToTB"
'
'
'   If debugflg = 1 Then DebugPrint "%" & LoadFileToTB
'
'    If Dir$(FilePath) = vbNullString Then Exit Function
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFileToTB of Form dock"
'
'End Function





''
''---------------------------------------------------------------------------------------
'' Procedure : GetINISetting
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   : Get the INI Setting from the File
''---------------------------------------------------------------------------------------
''
'Public Function GetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
'   On Error GoTo GetINISetting_Error
'    Const cparmLen = 500 ' maximum no of characters allowed in the returned string
'    Dim sReturn As String * cparmLen
'    Dim sDefault As String * cparmLen
'    Dim lLength As Long
'
'    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
'    GetINISetting = Mid(sReturn, 1, lLength)
'
'   On Error GoTo 0
'   Exit Function
'
'GetINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetINISetting of Module Module2"
'End Function

''
''---------------------------------------------------------------------------------------
'' Procedure : PutINISetting
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   : Save INI Setting in the File
''---------------------------------------------------------------------------------------
''
'Public Sub PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)
'
'   On Error GoTo PutINISetting_Error
'
'    Dim aLength As Long
'
'    aLength = WritePrivateProfileString(sHeading, sKey _
'            , sSetting, sINIFileName)
'
'   On Error GoTo 0
'   Exit Sub
'
'PutINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutINISetting of Module Module2"
'End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : readIconSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub readIconSettingsIni(location As String, ByVal iconNumberToRead As Integer, settingsFile As String)
'    'Reads an .INI File (SETTINGS.INI)
'
'   On Error GoTo readIconSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%readIconSettingsIni"
'                                '"Software\RocketDock\Icons"
'        sFilename = GetINISetting(location, iconNumberToRead & "-FileName", settingsFile)
'        sFileName2 = GetINISetting(location, iconNumberToRead & "-FileName2", settingsFile)
'        sTitle = GetINISetting(location, iconNumberToRead & "-Title", settingsFile)
'        sCommand = GetINISetting(location, iconNumberToRead & "-Command", settingsFile)
'        sArguments = GetINISetting(location, iconNumberToRead & "-Arguments", settingsFile)
'        sWorkingDirectory = GetINISetting(location, iconNumberToRead & "-WorkingDirectory", settingsFile)
'        sShowCmd = GetINISetting(location, iconNumberToRead & "-ShowCmd", settingsFile)
'        sOpenRunning = GetINISetting(location, iconNumberToRead & "-OpenRunning", settingsFile)
'        sIsSeparator = GetINISetting(location, iconNumberToRead & "-IsSeparator", settingsFile)
'        sUseContext = GetINISetting(location, iconNumberToRead & "-UseContext", settingsFile)
'        sDockletFile = GetINISetting(location, iconNumberToRead & "-DockletFile", settingsFile)
'
'   On Error GoTo 0
'   Exit Sub
'
'readIconSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconSettingsIni of Module Module2"
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : writeIconSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
''Public Sub writeIconSettingsIni(ByVal iconNumberToWrite As Integer)
'Sub writeIconSettingsIni(location As String, ByVal iconNumberToWrite As Integer, settingsFile As String)
'
'    'Writes an .INI File (SETTINGS.INI)
'
'    'E:\Program Files (x86)\RocketDock\Icons\Steampunk_Clockwerk_Kubrick
'    ' determine relative path TODO
'    ' Icons\Steampunk_Clockwerk_Kubrick
'
'   On Error GoTo writeIconSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%writeIconSettingsIni"
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
'       On Error GoTo 0
'   Exit Sub
'
'writeIconSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Module2"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : removeSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub removeSettingsIni(ByVal iconNumberToWrite As Integer)
'
'    'removes data from the ini file at the given location
'
'   On Error GoTo removeSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%removeSettingsIni"
'
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", vbNullString, rdSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", vbNullString, rdSettingsFile
'
'   On Error GoTo 0
'   Exit Sub
'
'removeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeSettingsIni of Module Module2"
'
'End Sub

''----------------------------------------
''Name: TestWinVer
''Description:
''----------------------------------------
'Public Sub testWinVer()
'
'    Dim ProgramFilesDir As String
'    Dim classicThemeCapable As Boolean
'
'    '=================================
'    '2000 / XP / NT / 7 / 8 / 10
'    '=================================
'    On Error GoTo TestWinVer_Error
'
'    If debugflg = 1 Then DebugPrint "%" & "TestWinVer"
'
'    On Error Resume Next
'    Dim strString As String
'    strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
'    WindowsVer = strString
'
'    ' note that when running in compatibility mode the o/s will always respond with "Windows XP"
'    ' The IDE runs in compatibility mode so it will report the wrong working folder
'
'    'MsgBox WindowsVer
'
'    'Get the value of "ProgramFiles", or "ProgramFilesDir"
'
'    strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
'    classicThemeCapable = False
'    requiresAdmin = False
'
'    If InStr(WindowsVer, "Windows NT4") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    End If
'    If InStr(WindowsVer, "Windows 2000") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    End If
'    If InStr(WindowsVer, "Windows XP") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
'    End If
'    If InStr(WindowsVer, "Windows 2003") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    End If
'    If InStr(WindowsVer, "Windows Vista") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'    End If
'    If InStr(WindowsVer, "Windows 7") <> 0 Then
'        classicThemeCapable = True
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'        'MsgBox "%" & "WindowsVer = " & windowsVer & " strString = " & strString
'    End If
'    If InStr(WindowsVer, "Windows 2008") <> 0 Then
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'    End If
'    If InStr(WindowsVer, "Windows 8") <> 0 Then
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'    End If
'    If InStr(WindowsVer, "Windows 2012") <> 0 Then
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'    End If
'    If InStr(WindowsVer, "Windows 10") <> 0 Then
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'    End If
'    If InStr(WindowsVer, "Windows 2016") <> 0 Then
'        strString = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
'        requiresAdmin = True
'    End If
'
'    If debugflg = 1 Then DebugPrint "%" & "WindowsVer = " & WindowsVer
'
'    ProgramFilesDir = strString
'    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems use this
'    If Not DirExists(ProgramFilesDir) Then
'        ProgramFilesDir = "c:\program files" ' 32 bit systems use this
'    End If
'
'    If debugflg = 1 Then DebugPrint "%" & "ProgramFilesDir = " & ProgramFilesDir
'
'    'MsgBox "%" & "WindowsVer = " & windowsVer & " " & ProgramFilesDir
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

''---------------------------------------------------------------------------------------
'' Procedure : getstring
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function getstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String
'
'    Dim keyhand As Long
'    'Dim datatype As Long
'    Dim lResult As Long
'    Dim strBuf As String
'    Dim lDataBufSize As Long
'    Dim intZeroPos As Integer
'    Dim rvar As Integer
'    'in .NET the variant type will need to be replaced by object?
'
'    Dim lValueType As Variant
'
'   On Error GoTo getstring_Error
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
'
'   On Error GoTo 0
'   Exit Function
'
'getstring_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getstring of Module Module1"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : savestring
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub savestring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)
'
'    Dim keyhand As Long
'    Dim R As Long
'   On Error GoTo savestring_Error
'
'    R = RegCreateKey(hKey, strPath, keyhand)
'    R = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
'    R = RegCloseKey(keyhand)
'
'   On Error GoTo 0
'   Exit Sub
'
'savestring_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savestring of Module Module1"
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : checkRocketdockInstallation
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : we check to see if rocketdock is installed in order to know the location of the settings.ini file used by Rocketdock
''             *** also sets rdAppPath during the drivecheck ***
''---------------------------------------------------------------------------------------
'
'Public Sub checkRocketdockInstallation()
'    RD86installed = vbNullString
'    RDinstalled = vbNullString
'
'    ' check where rocketdock is installed
'    On Error GoTo checkRocketdockInstallation_Error
'    If debugflg = 1 Then DebugPrint "%" & "checkRocketdockInstallation"
'
'    RD86installed = driveCheck("Program Files (x86)\Rocketdock", "RocketDock.exe")
'    RDinstalled = driveCheck("Program Files\Rocketdock", "RocketDock.exe")
'
'    If RDinstalled = vbNullString And RD86installed = vbNullString Then
'        rocketDockInstalled = False
'    Else
'        rocketDockInstalled = True
'        If RDinstalled <> vbNullString Then
'            rdAppPath = RDinstalled
'        End If
'        'the one in the x86 folder has precedence
'        If RD86installed <> vbNullString Then
'            rdAppPath = RD86installed
'        End If
'    End If
'
'    ' If rocketdock Is Not installed Then test the registry
'    ' if the registry settings are not located then remove them as a source.
'
'    ' you might think this stuff is better placed in the docksettings utility
'    ' but it has to be here as well as SteamyDock is the component that is most likely to br run first.
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
'        If rDLockIcons = vbNullString And rDOpenRunning = vbNullString And rDShowRunning = vbNullString And rDManageWindows = vbNullString And rDDisableMinAnimation = vbNullString Then
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkRocketdockInstallation of Form dock"
'End Sub


'
''---------------------------------------------------------------------------------------
'' Procedure : driveCheck
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : check for the existence of the rocketdock binary
''---------------------------------------------------------------------------------------
''
'Public Function driveCheck(ByVal folder As String, filename As String) As String
'   Dim sAllDrives As String
'   Dim sDrv As String
'   Dim sDrives() As String
'   Dim cnt As Long
'   Dim folderString As String
'   Dim testAppPath As String
'
'  'get the list of all drives
'   On Error GoTo driveCheck_Error
'   If debugflg = 1 Then DebugPrint "%" & "driveCheck"
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
'           'test for the Rocketdock binary
'            testAppPath = folderString
'            If FExists(testAppPath & "\" & filename) Then
'                'MsgBox "Rocketdock binary exists"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure driveCheck of Form dock"
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
'    'pure VB approach, no controls required Gary Beene
'    'drive letters are found in positions 1-UBound(Letters)
'    '"C:\ D:\ E:\ &frameProperties"
'
'    On Error GoTo GetDriveString_Error
'       If debugflg = 1 Then DebugPrint "%" & "GetDriveString"
'
'
'
'    For I = 1 To 26
'        If ValidDrive(Chr$(96 + I)) = True Then
'            builtString = builtString + UCase$(Chr$(96 + I)) & ":\    "
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriveString of Form dock"
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : ValidDrive
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Check if the drive found is a valid one
''---------------------------------------------------------------------------------------
''
'Public Function ValidDrive(ByVal d As String) As Boolean
'  On Error GoTo ValidDrive_Error
'  If debugflg = 1 Then DebugPrint "%" & "ValidDrive"
'  On Error GoTo driveerror
'  Dim Temp As String
'
'    Temp = CurDir$
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidDrive of Form dock"
'End Function






''---------------------------------------------------------------------------------------
'' Procedure : readIconRegistryWriteSettings
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Read the registry one line at a time and create a temporary settings file
''---------------------------------------------------------------------------------------
''
'Private Sub readIconRegistryWriteSettings(settingsFile As String)
'    Dim useloop As Integer
'
'    On Error GoTo readIconRegistryWriteSettings_Error
'    If debugflg = 1 Then DebugPrint "%" & "readIconRegistryWriteSettings"
'
'    For useloop = 0 To rdIconMaximum
'         ' get the relevant entries from the registry
'         readRegistryOnce (useloop)
'         ' write the rocketdock alternative settings.ini
'         Call writeIconSettingsIni("Software\RocketDock\Icons", useloop, settingsFile) ' the alternative settings.ini exists when RD is set to use it
'     Next useloop
'
'   On Error GoTo 0
'   Exit Sub
'
'readIconRegistryWriteSettings_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconRegistryWriteSettings of Form dock"
'End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryOnce
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryOnce(ByVal iconNumberToRead As Integer)
'    ' read the settings from the registry
'    On Error GoTo readRegistryOnce_Error
'    If debugflg = 1 Then DebugPrint "%" & "readRegistryOnce"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryOnce of Form dock"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistryOnce of Form dock"
'End Sub






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
'    If debugflg = 1 Then DebugPrint "%" & "ExtractSuffixWithDot"
'
'    If strPath = vbNullString Then
'        ExtractSuffixWithDot = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        AY = Split(strPath, ".")
'        Max = UBound(AY)
'        ExtractSuffixWithDot = "." & AY(Max)
'    Else
'        ExtractSuffixWithDot = strPath
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'ExtractSuffixWithDot_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffixWithDot of Form dock"
'End Function

'
'
''---------------------------------------------------------------------------------------
'' Procedure : ExtractSuffix
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function ExtractSuffix(ByVal strPath As String) As String
'    Dim AY() As String ' string array
'    Dim Max As Integer
'
'    On Error GoTo ExtractSuffix_Error
'    If debugflg = 1 Then DebugPrint "%" & "ExtractSuffix"
'
'    If strPath = vbNullString Then
'        ExtractSuffix = vbNullString
'        Exit Function
'    End If
'
'    If InStr(strPath, ".") <> 0 Then
'        AY = Split(strPath, ".")
'        Max = UBound(AY)
'        ExtractSuffix = AY(Max)
'    Else
'        ExtractSuffix = strPath
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'ExtractSuffix_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractSuffix of Form dock"
'End Function







''---------------------------------------------------------------------------------------
'' Procedure : Is64bit
'' Author    : beededea
'' Date      : 04/07/2020
'' Purpose   : 'Spider Harper
''---------------------------------------------------------------------------------------
''
'Public Function Is64bit() As Boolean
'    Dim Handle As Long
'    Dim bolFunc As Boolean
'
'    ' Assume initially that this is not a Wow64 process
'    On Error GoTo Is64bit_Error
'
'    bolFunc = False
'
'    ' Now check to see if IsWow64Process function exists
'    Handle = GetProcAddress(GetModuleHandle("kernel32"), _
'                   "IsWow64Process")
'
'    If Handle > 0 Then ' IsWow64Process function exists
'        ' Now use the function to determine if
'        ' we are running under Wow64
'        IsWow64Process GetCurrentProcess(), bolFunc
'    End If
'
'    Is64bit = bolFunc
'
'   On Error GoTo 0
'   Exit Function
'
'Is64bit_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Is64bit of Module Module1"
'
'End Function


''---------------------------------------------------------------------------------------
'' Procedure : readDockSettingsFile
'' Author    : beededea
'' Date      : 12/05/2020
'' Purpose   : read
''---------------------------------------------------------------------------------------
''
'Private Sub readDockSettingsFile(location As String, settingsFile As String)
'
'    'SteamyDock settings only
'    On Error GoTo readDockSettingsFile_Error
'    If debugflg = 1 Then DebugPrint "%readDockSettingsFile"
'
'    If FExists(dockSettingsFile) Then
'        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
'        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
'        rDRunAppInterval = GetINISetting("Software\SteamyDock\DockSettings", "RunAppInterval", dockSettingsFile) ' dean
'        rDAlwaysAsk = GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile)  ' dean
'        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
'        rDAnimationInterval = GetINISetting("Software\SteamyDock\DockSettings", "AnimationInterval", dockSettingsFile)
'        rDSkinSize = GetINISetting("Software\SteamyDock\DockSettings", "SkinSize", dockSettingsFile)
'    End If
'
'    sDSkinSize = Val(rDSkinSize)
'
'    'if the above settings do not exist in the older RD settings file then no error is thrown so it works for both
'
'    'RocketDock settings only
'    rDVersion = GetINISetting(location, "Version", settingsFile) ' dean
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
'readDockSettingsFile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockSettingsFile of Form dockSettings"
'
'End Sub


' did some thinking on making the cogs disappear more quickly
' perhaps maintaining a list of process initiated from the dock then touching them to see if they exist
' the function below was to do that, it works but the VB6 IDE was terminating unexpectedly when this code was run
' not yet tested in a compiled version.

'    MsgBox IsModuleRunning("explorer.exe")

'Public Function IsModuleRunning(ByVal theModuleName As String) As Boolean
'    Dim aProcessess(1 To 1024)  As Long ' up to 1024 processess?'
'    Dim bytesNeeded             As Long
'    Dim i                       As Long
'    Dim nProcesses              As Long
'    Dim hProcess                As Long
'    Dim found                   As Boolean
'
'    EnumProcesses aProcessess(1), UBound(aProcessess), bytesNeeded
'    nProcesses = bytesNeeded / 4
'    For i = 1 To nProcesses
'
'        hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, aProcessess(i))
'        If (hProcess) Then
'            Dim hmodule(1 To 1024)  As Long ' no more than 1024 modules per process?'
'            bytesNeeded = 0
'            If EnumProcessModules(hProcess, hmodule(1), 1024 * 4, bytesNeeded) Then
'                Dim nModules    As Long
'                Dim j           As Long
'                Dim moduleName  As String
'
'                moduleName = Space(1024)   ' module name should have less than 1024 bytes'
'
'                nModules = bytesNeeded / 4
'                For j = 1 To nModules
'                    Dim fileNameLen As Long
'                    fileNameLen = GetModuleFileNameExA(hProcess, hmodule(j), moduleName, 1024)
'                    moduleName = Left(moduleName, fileNameLen)
'                    If Right(LCase(moduleName), Len(theModuleName)) = LCase(theModuleName) Then
'                        found = True
'                        Exit For
'                    End If
'                Next
'            End If
'        End If
'        CloseHandle hProcess
'        If found Then Exit For
'    Next
'    IsModuleRunning = found
'End Function
''---------------------------------------------------------------------------------------
'' Procedure : isProcessInTaskList
'' Author    : beededea
'' Date      : 11/04/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function isProcessInTaskList(ByVal sProcess As String) As Boolean
'    Const MAX_PATH As Long = 260
'    Dim lProcesses() As Long
'    Dim lModules() As Long
'    Dim N As Long
'    Dim lRet As Long
'    Dim hProcess As Long
'
'    Dim sName As String
'
'   On Error GoTo isProcessInTaskList_Error
'
'    sProcess = UCase$(sProcess)
'
'    ReDim lProcesses(1023) As Long
'    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
'        For N = 0 To (lRet \ 4) - 1
'            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
'            If hProcess Then
'                ReDim lModules(1023)
'                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
'                    sName = String$(MAX_PATH, vbNullChar)
'                    'GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
'
'                    QueryFullProcessImageName hProcess, 0, sName, MAX_PATH
'                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
'                    If Len(sName) = Len(sProcess) Then
'                        If sProcess = UCase$(sName) Then
'                            isProcessInTaskList = True
'                            Exit Function
'                        End If
'                    End If
'                End If
'            End If
'            CloseHandle hProcess
'        Next N
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'isProcessInTaskList_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure isProcessInTaskList of Module Module1"
'End Function



