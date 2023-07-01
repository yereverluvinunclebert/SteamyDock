Attribute VB_Name = "mdlDeadCode"
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const REG_SZ = 1                          ' Unicode nul terminated string

'Public interimSettingsFile As String
'Public dockSettingsFile As String
'Public origSettingsFile As String
'Public toolSettingsFile  As String
'Public WindowsVer As String
'Public requiresAdmin As Boolean
'Public rdAppPath As String
'Public RDinstalled As String
'Public RD86installed As String
'Public rocketDockInstalled As Boolean
'Public RDregistryPresent As Boolean
'Public rDCustomIconFolder As String ' .NET

'Public rDGeneralReadConfig As String
'Public rDGeneralWriteConfig As String

'Public Enum eSpecialFolders
'  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
'  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
'  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
'  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
'End Enum

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
    
'Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

'Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

'Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
'Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
    
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

''Get the INI Setting from the File
''---------------------------------------------------------------------------------------
'' Procedure : GetINISetting
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   :
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
'    GetINISetting = mid$(sReturn, 1, lLength)
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
'' Date      : 05/07/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)
'
'   On Error GoTo PutINISetting_Error
'
'    Dim aLength As Long
'
'    aLength = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)
'
'   On Error GoTo 0
'   Exit Sub
'
'PutINISetting_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutINISetting of Module Module2"
'End Sub
'Sub findIconMax()
'
'    rdIconMaximum = GetINISetting("Software\RocketDock\Icons", "count", rdAppPath & "\SETTINGS.INI")
'
'    'Reads a INI File (SETTINGS.INI)
'    'For useloop = 0 To 500 ' the current maximum
'        'sFileName(useloop) = GetINISetting("Software\RocketDock\Icons", useloop & "-FileName", rdAppPath & "\SETTINGS.INI")
'        'If sFileName(useloop) = "" Then
'        '    Exit Sub
'        '    rdIconMaximum = useloop ' obtain the number of the last icon in the settings file
'        'End If
'
'End Sub
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
'   On Error GoTo writeIconSettingsIni_Error
'   If debugflg = 1 Then Debug.Print "%writeIconSettingsIni"
'
'        sFilenameCheck = sFilename  ' debug 01
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
'        'test to see if the icon path has been truncated
'        sFilename = GetINISetting(location, iconNumberToWrite & "-FileName", settingsFile)
'        If sFilenameCheck <> "" Then
'            If sFilename <> sFilenameCheck Then
'                MsgBox " that strange truncated filename bug encountered, check " & settingsFile & " now and look for " & sFilenameCheck
'            End If
'        End If
'
'   On Error GoTo 0
'   Exit Sub
'
'writeIconSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Module2"
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : writeSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Sub writeSettingsIni(ByVal iconNumberToWrite As Integer)
'    'Writes an .INI File (SETTINGS.INI)
'
'    'E:\Program Files (x86)\RocketDock\Icons\Steampunk_Clockwerk_Kubrick
'    ' determine relative path TODO
'    ' Icons\Steampunk_Clockwerk_Kubrick
'
'   On Error GoTo writeSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%writeSettingsIni"
'
'        sFilenameCheck = sFilename  ' debug 01
'
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", sFilename, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", sFileName2, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", sTitle, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", sCommand, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", sArguments, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", sShowCmd, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", sOpenRunning, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", sIsSeparator, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", sUseContext, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", sDockletFile, interimSettingsFile
'
'
'        sFilename = GetINISetting("Software\RocketDock\Icons", iconNumberToWrite & "-FileName", interimSettingsFile)
'        If sFilenameCheck <> "" Then
'            If sFilename <> sFilenameCheck Then
'                MsgBox " that strange filename bug encountered, check rdSettings.ini now and look for " & sFilenameCheck
'            End If
'        End If
'
'   On Error GoTo 0
'   Exit Sub
'
'writeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeSettingsIni of Module Module2"
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
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-FileName2", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Title", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Command", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-Arguments", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-WorkingDirectory", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-ShowCmd", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-OpenRunning", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-IsSeparator", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-UseContext", vbNullString, interimSettingsFile
'        PutINISetting "Software\RocketDock\Icons", iconNumberToWrite & "-DockletFile", vbNullString, interimSettingsFile
'
'   On Error GoTo 0
'   Exit Sub
'
'removeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeSettingsIni of Module Module2"
'
'End Sub



' .89 DAEB 13/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public
'---------------------------------------------------------------------------------------
' Procedure : f_GetSaveFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Sub f_GetSaveFileName()
'   On Error GoTo f_GetSaveFileName_Error
'   If debugflg = 1 Then DebugPrint "%f_GetSaveFileName"
'
'  If GetSaveFileName(x_OpenFilename) <> 0 Then
'    'PURPOSE: A file was selected
'    MsgBox Left$(x_OpenFilename.lpstrFile, x_OpenFilename.nMaxFile)
'  Else
'    'PURPOSE: The CANCEL button was pressed
'    MsgBox "Cancel"
'  End If
'
'   On Error GoTo 0
'   Exit Sub
'
'f_GetSaveFileName_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure f_GetSaveFileName of Form rDIconConfigForm"
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : readSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub readSettingsIni(ByVal iconNumberToRead As Integer)
'    'Reads an .INI File (SETTINGS.INI)
'
'   On Error GoTo readSettingsIni_Error
'   If debugflg = 1 Then DebugPrint "%readSettingsIni"
'
'        sFilename = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-FileName", interimSettingsFile)
'        sFileName2 = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-FileName2", interimSettingsFile)
'        sTitle = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-Title", interimSettingsFile)
'        sCommand = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-Command", interimSettingsFile)
'        sArguments = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-Arguments", interimSettingsFile)
'        sWorkingDirectory = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-WorkingDirectory", interimSettingsFile)
'        sShowCmd = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-ShowCmd", interimSettingsFile)
'        sOpenRunning = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-OpenRunning", interimSettingsFile)
'        sIsSeparator = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-IsSeparator", interimSettingsFile)
'        sUseContext = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-UseContext", interimSettingsFile)
'        sDockletFile = GetINISetting("Software\RocketDock\Icons", iconNumberToRead & "-DockletFile", interimSettingsFile)
'
'
'   On Error GoTo 0
'   Exit Sub
'
'readSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsIni of Module Module2"
'End Sub
''FIXIT: Declare 'getstring' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
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
'    'FIXIT: Declare 'lValueType' with an early-bound data type                                 FixIT90210ae-R1672-R1B8ZE
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

''----------------------------------------
''Name: testWindowsVersion
''Description:
''----------------------------------------
'Public Sub testWindowsVersion(classicThemeCapable As Boolean)
'
'    '=================================
'    '2000 / XP / NT / 7 / 8 / 10
'    '=================================
'    On Error GoTo testWindowsVersion_Error
'
'    ' variables declared
'
'    Dim ProgramFilesDir As String
'    Dim WindowsVer As String
'    Dim strString As String
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
'    If debugflg = 1 Then DebugPrint "%" & "ProgramFilesDir = " & ProgramFilesDir
'
'    ' turn on the timer that tests every 10 secs whether the visual theme has changed
'    ' only on those o/s versions that need it
'
'    If classicThemeCapable = True Then
'        rDIconConfigForm.mnuAuto.Caption = "Auto Theme Disable"
'        rDIconConfigForm.themeTimer.Enabled = True
'    Else
'        rDIconConfigForm.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"
'        rDIconConfigForm.themeTimer.Enabled = False
'    End If
'
'    '======================================================
'    'END routine error handler
'    '======================================================
'
'
'    On Error GoTo 0: Exit Sub
'
'testWindowsVersion_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testWindowsVersion of Module WinModule"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : FExists
'' Author    : beededea
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function FExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'   On Error GoTo FExists_Error
'   If debugflg = 1 Then Debug.Print "%FExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    FExists = FS.FileExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'FExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FExists of Module Module1"
'End Function
'
'
''---------------------------------------------------------------------------------------
'' Procedure : DirExists
'' Author    : beededea
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function DirExists(ByRef OrigFile As String) As Boolean
'    Dim FS As Object
'   On Error GoTo DirExists_Error
'   If debugflg = 1 Then DebugPrint "%DirExists"
'
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    DirExists = FS.FolderExists(OrigFile)
'
'   On Error GoTo 0
'   Exit Function
'
'DirExists_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DirExists of Module Module1"
'End Function




''---------------------------------------------------------------------------------------
'' Procedure : SpecialFolder
'' Author    :  si_the_geek vbforums
'' Date      : 17/10/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function SpecialFolder(pFolder As eSpecialFolders) As String
''Returns the path to the specified special folder (AppData etc)
'
'Dim objShell  As Object
'Dim objFolder As Object
'
'   On Error GoTo SpecialFolder_Error
'   If debugflg = 1 Then DebugPrint "%SpecialFolder"
'
'  Set objShell = CreateObject("Shell.Application")
'  Set objFolder = objShell.NameSpace(CLng(pFolder))
'
'  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.path
'
'  Set objFolder = Nothing
'  Set objShell = Nothing
'
'  If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"
'
'   On Error GoTo 0
'   Exit Function
'
'SpecialFolder_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SpecialFolder of Module Module1"
'
'End Function
'
'
'Public Sub EnsureFormIsInsideMonitor(frm As Form, Optional RefForm As Form)
'
'    'typically used to determine if the previously saved position/size of a Form needs adjustment re the current monitor layout
'    ' if a Form is mapped to a disconnected monitor the Form is remapped to the Primary monitor
'
'    'typical usage:
'    'Private Sub Form_Load()
'    '   retrieve previous left and top positions from file and apply them to Me.Left and Me.Top Properties
'    '   EnsureFormIsInsideMonitor Me
'    'End Sub
'
'    'If Reform is not specified Frm is positioned to be displayed entirely within the monitor on which most of it is currently mapped
'    'If Reform is specified Frm is positioned to be displayed entirely on the same monitor as that on which most of Reform is mapped
'
'    'adjusts Frm Left and Top so that all the borders of Frm are contained within the same Monitor
'
'    ' if Frm.Width or Height exceed monitor.width or height Frm is positioned at Left/ Top of monitor and
'    '  Width/ Height of Frm may be adjusted if Frm is Sizable
'
'    Dim VFlag As Boolean: VFlag = False
'    Dim HFlag As Boolean: HFlag = False
'    Dim cMonitor As UDTMonitor
'
'    If RefForm Is Nothing Then Set RefForm = frm
'
'    cMonitor = monitorProperties(RefForm, screenTwipsPerPixelX, screenTwipsPerPixelY)
'
'    With frm
'        If .Width > cMonitor.WorkWidth Then
'            If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
'                .Width = cMonitor.WorkWidth
'            Else
'                .Left = cMonitor.WorkLeft: HFlag = True
'            End If
'        End If
'        If .Height > cMonitor.WorkHeight Then
'            If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
'                .Height = cMonitor.WorkHeight
'            Else
'                .Top = cMonitor.WorkTop: VFlag = True
'            End If
'        End If
'
'        If Not HFlag Then
'            If .Left < cMonitor.WorkLeft Then .Left = cMonitor.WorkLeft
'            If (.Left + .Width) > cMonitor.WorkRight Then .Left = cMonitor.WorkRight - .Width
'        End If
'        If Not VFlag Then
'            If .Top < cMonitor.WorkTop Then .Top = cMonitor.WorkTop
'            If (.Top + .Height) > cMonitor.Workbottom Then .Top = cMonitor.Workbottom - .Height
'        End If
'    End With
'
'End Sub
