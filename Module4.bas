Attribute VB_Name = "mdlDeadCode"

'------------------------------------------------------------
' Module4.bas
'
' Module to contain code that is now unused or copied to another shared module
' code kept for reference purposes.
'
'------------------------------------------------------------

'Public Enum FileOpenConstants
'    'ShowOpen, ShowSave constants.
'    cdlOFNAllowMultiselect = &H200&
'    cdlOFNCreatePrompt = &H2000&
'    cdlOFNExplorer = &H80000
'    cdlOFNExtensionDifferent = &H400&
'    cdlOFNFileMustExist = &H1000&
'    cdlOFNHideReadOnly = &H4&
'    cdlOFNLongNames = &H200000
'    cdlOFNNoChangeDir = &H8&
'    cdlOFNNoDereferenceLinks = &H100000
'    cdlOFNNoLongNames = &H40000
'    cdlOFNNoReadOnlyReturn = &H8000&
'    cdlOFNNoValidate = &H100&
'    cdlOFNOverwritePrompt = &H2&
'    cdlOFNPathMustExist = &H800&
'    cdlOFNReadOnly = &H1&
'    cdlOFNShareAware = &H4000&
'End Enum
'
'' APIs and structures for opening a common dialog box to select files without OCX dependencies
'
'Public Type OPENFILENAME
'    lStructSize As Long    'The size of this struct (Use the Len function)
'    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
'    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
'    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
'    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
'    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
'    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
'    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
'    nMaxFile As Long             'The length of lpstrFile + 1
'    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
'    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
'    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
'    lpstrTitle As String         'The caption of the dialog.
'    flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
'    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
'    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
'    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
'    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
'    lpfnHook As Long             'Pointer to the hook procedure.
'    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
'End Type
'
'Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" ( _
'    lpofn As OPENFILENAME) As Long
'
'Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" ( _
'    lpofn As OPENFILENAME) As Long
'
'Public OF As OPENFILENAME
'Public x_OpenFilename As OPENFILENAME
'
'Public Type BROWSEINFO
'    hwndOwner As Long
'    pidlRoot As Long 'LPCITEMIDLIST
'    pszDisplayName As String
'    lpszTitle As String
'    ulFlags As Long
'    lpfn As Long  'BFFCALLBACK
'    lParam As Long
'    iImage As Long
'End Type
'Public Declare Function SHBrowseForFolderA Lib "Shell32.dll" (binfo As BROWSEINFO) As Long
'Public Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl&, ByVal szPath$) As Long
'Public Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



'Public Declare Function QueryFullProcessImageName Lib "Kernel32.dll" Alias "QueryFullProcessImageNameA" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As String, ByVal lpdwSize As Long) As Long



'Private Declare Function GetModuleHandle Lib "kernel32" _
'    Alias "GetModuleHandleA" ( _
'    ByVal lpModuleName As Long) As Long

'Private Declare Function GetModuleBaseName Lib "psapi" _
'    Alias "GetModuleBaseNameA" ( _
'    ByVal hProcess As Long, _
'    ByVal hModule As Long, _
'    ByVal BaseName As String, _
'    ByVal nSize As Long) As Long

'Private Declare Function GetModuleFileNameEx Lib "psapi" _
'    Alias "GetModuleFileNameExA" ( _
'    ByVal hProcess As Long, _
'    ByVal hModule As Long, _
'    ByVal FileName As String, _
'    ByVal nSize As Long) As Long
    
'Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
'Private Declare Function GetModuleFileNameEx Lib "psapi" _
'    Alias "GetModuleFileNameExA" ( _
'    ByVal hProcess As Long, _
'    ByVal hModule As Long, _
'    ByVal FileName As String, _
'    ByVal nSize As Long) As Long


' APIs for querying running processes START
'Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
'Private Declare Function ProcessFirst Lib "Kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
'Private Declare Function ProcessNext Lib "Kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
'Private Declare Function CreateToolhelpSnapshot Lib "Kernel32.dll" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByRef lProcessId As Long) As Long
'Private Declare Function TerminateProcess Lib "Kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
'Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
' APIs for querying running processes END

' 26/10/2020 .03 mdlMain.bas dock DAEB removed declarations required by IsRunning since the move of this function to common.bas STARTS.
' variables for querying running processes START
'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
'Private Const TH32CS_SNAPPROCESS As Long = 2&
'Private uProcess   As PROCESSENTRY32
'Private hSnapshot As Long
' variables for querying running processes ENDS

' APIs for querying running processes START
'Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
'Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
'Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
'Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long '  Alias "CreateToolhelp32Snapshot"
'Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long
''Private Declare Function TerminateProcess Lib "Kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
' APIs for querying running processes END
' 26/10/2020 .03 mdlMain.bas dock DAEB removed declarations required by IsRunning since the move of this function to common.bas ENDS.


' .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
'Public appSystrayTypes As String ' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list

'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Const WM_POWERBROADCAST = &H218


'msgCnt = 0
'
''.nn
'Enum enPowerBroadcastType
'     PBT_APMQUERYSUSPEND = &H0
'     PBT_APMQUERYSTANDBY = &H1
'     PBT_APMQUERYSUSPENDFAILED = &H2
'     PBT_APMQUERYSTANDBYFAILED = &H3
'     PBT_APMSUSPEND = &H4
'     PBT_APMSTANDBY = &H5
'     PBT_APMRESUMECRITICAL = &H6
'     PBT_APMRESUMESUSPEND = &H7
'     PBT_APMRESUMESTANDBY = &H8
'End Enum

'Private Enum InterpolationMode
'    InterpolationModeDefault = &H0
'    InterpolationModeLowQuality = &H1
'    InterpolationModeHighQuality = &H2
'    InterpolationModeBilinear = &H3
'    InterpolationModeBicubic = &H4
'    InterpolationModeNearestNeighbor = &H5
'    InterpolationModeHighQualityBilinear = &H6
'    InterpolationModeHighQualityBicubic = &H7
'End Enum


''these functions need to be in a BAS module and not a form or the AddressOf does not work.



'
'
''---------------------------------------------------------------------------------------
'' Procedure : BrowseCallbackProc
'' Author    : beededea
'' Date      : 20/08/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Function BrowseCallbackProc(ByVal hWnd&, ByVal msg&, ByVal lp&, ByVal InitDir$) As Long
'   Const BFFM_INITIALIZED As Long = 1
'   Const BFFM_SETSELECTION As Long = &H466
'   On Error GoTo BrowseCallbackProc_Error
'
'   If (msg = BFFM_INITIALIZED) And (InitDir <> vbNullString) Then
'      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal InitDir$)
'   End If
'   BrowseCallbackProc = 0
'
'   On Error GoTo 0
'   Exit Function
'
'BrowseCallbackProc_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BrowseCallbackProc of Module Module4"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : GetAddress
'' Author    : beededea
'' Date      : 20/08/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Function GetAddress(ByVal Addr As Long) As Long
'   On Error GoTo GetAddress_Error
'
'   GetAddress = Addr
'
'   On Error GoTo 0
'   Exit Function
'
'GetAddress_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetAddress of Module Module4"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : BrowseFolder
'' Author    : beededea
'' Date      : 20/08/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function BrowseFolder(ByVal hwndOwner&, DefFolder$) As String
'   Dim bi As BROWSEINFO
'Dim pidl&
'Dim newPath$
'
'   On Error GoTo BrowseFolder_Error
'
'   bi.hwndOwner = hwndOwner
'   bi.lpfn = GetAddress(AddressOf BrowseCallbackProc)
'   bi.lParam = StrPtr(DefFolder)
'   pidl = SHBrowseForFolderA(bi)
'   If (pidl) Then
'      newPath = String(260, 0)
'      If SHGetPathFromIDListA(pidl, newPath) Then
'         newPath = Left(newPath, InStr(1, newPath, Chr(0)) - 1)
'         BrowseFolder = newPath
'      End If
'      Call CoTaskMemFree(ByVal pidl&)
'   End If
'
'   On Error GoTo 0
'   Exit Function
'
'BrowseFolder_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BrowseFolder of Module Module4"
'End Function





''---------------------------------------------------------------------------------------
'' Procedure : validateInputs
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateInputs()
'
'   On Error GoTo validateInputs_Error
'
'    If Val(rDRunAppInterval) * 1000 >= 65536 Then rDRunAppInterval = "65"
'
'    ' validate the relevant entries from whichever source
'    validateRegistryGeneral
'    validateRegistryIcons
'    validateRegistryBehaviour
'    validateRegistryStyle
'    validateRegistryPosition
'
'   On Error GoTo 0
'   Exit Sub
'
'validateInputs_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of Module Module1"
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : readRegistry
'' Author    : beededea
'' Date      : 09/05/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistry()
'     Dim useloop As Integer
'
'   On Error GoTo readRegistry_Error
'   If debugflg = 1 Then DebugPrint "%readRegistry"
'
'     'Dean Debug - this reading from the registry has to stop!
'     rDOptionsTabIndex = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OptionsTabIndex")
'
'    ' get items from the registry that are required to 'default' the dock but aren't controlled by the dock settings utility
'    rdIconCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Count")
'
'    rDVersion = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Version")
'    rDCustomIconFolder = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "CustomIconFolder")
'    rDHotKeyToggle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "HotKeyToggle")
'
'     ' get the relevant entries from the registry
'     readRegistryGeneral
'     readRegistryIcons
'     readRegistryBehaviour
'     readRegistryStyle
'     readRegistryPosition
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistry_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistry of Module Module1"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : readRegistryGeneral
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryGeneral()
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
'   On Error GoTo readRegistryGeneral_Error
'
'    rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "RocketDock")
'    If rdStartupRunString <> vbNullString Then
'        rDStartupRun = "1"
'    End If
'
'    rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIcons")
'    rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunning")
'    rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunning")
'    rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindows")
'    rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimation")
'
'    Call validateRegistryGeneral
'
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryGeneral_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryGeneral of Module Module1"
'
'End Sub
''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryGeneral
'' Author    : beededea
'' Date      : 17/06/2020
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
'    If Val(rDLockIcons) <= 0 And Val(rDLockIcons) > 1 Then rDLockIcons = "1" '
'    If Val(rDOpenRunning) <= 0 And Val(rDOpenRunning) > 1 Then rDOpenRunning = "1" '
'    If Val(rDShowRunning) <= 0 And Val(rDShowRunning) > 1 Then rDShowRunning = "1" '
'    If Val(rDManageWindows) <= 0 And Val(rDManageWindows) > 1 Then rDManageWindows = "1" '
'    If Val(rDDisableMinAnimation) <= 0 And Val(rDDisableMinAnimation) > 1 Then rDDisableMinAnimation = "1" '
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryGeneral_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryGeneral of Module Module1"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : readRegistryIcons
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryIcons()
'
'    ' read the icon configuration settings from the registry
'
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
'   On Error GoTo readRegistryIcons_Error
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryIcons of Module Module1"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryIcons
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryIcons()
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'   On Error GoTo validateRegistryIcons_Error
'
'    If Val(rDMonitor) <= 0 Or Val(rDMonitor) > 10 Then rDMonitor = "1" 'monitor 1
'    If Val(rDIconOpacity) < 50 Or Val(rDIconOpacity) > 100 Then rDIconOpacity = "100" 'fully opaque
'    If Val(rDZoomOpaque) <= 0 Or Val(rDZoomOpaque) > 1 Then rDZoomOpaque = "1" 'zooms opaque
'    If Val(rDIconMin) < 16 Or Val(rDIconMin) > 128 Then rDIconMin = "16" 'small
'    If Val(rDHoverFX) <= 0 Or Val(rDHoverFX) > 3 Then rDHoverFX = "1" 'bounce
'
'    If Val(rdIconMax) < 1 Or Val(rdIconMax) > 256 Then rdIconMax = "256" 'largest size
'    'MsgBox "icnomax = " & rdIconMax
'
'    If Val(rDZoomWidth) < 2 Or Val(rDZoomWidth) > 10 Then rDZoomWidth = "4" ' just a few expanded
'    If Val(rDZoomTicks) < 100 Or Val(rDZoomTicks) > 500 Then rDZoomTicks = "100" ' 100ms
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryIcons_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryIcons of Module Module1"
'
'End Sub


''
''---------------------------------------------------------------------------------------
'' Procedure : readRegistryPosition
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryPosition()
'
'   'style items
'
''   Monitor 'Monitor
''   Side    ' Side
''   zOrderMode  ' zOrderMode
''   Offset  ' Offset
''   vOffset ' vOffset
'
'   On Error GoTo readRegistryPosition_Error
'
'    rDMonitor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Monitor")
'    rDSide = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Side")
'    rDzOrderMode = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "zOrderMode")
'    rDOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Offset")
'    rDvOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "vOffset")
'
'    Call validateRegistryPosition
'
'
'
'   On Error GoTo 0
'   Exit Sub
'
'readRegistryPosition_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryPosition of Module Module1"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryPosition
'' Author    : beededea
'' Date      : 17/06/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub validateRegistryPosition()
'
'
'    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
'    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
'    ' when Rocketdock is restarted
'
'   On Error GoTo validateRegistryPosition_Error
'
'    If Val(rDMonitor) <= 0 Or Val(rDMonitor) > 10 Then rDMonitor = "1" 'monitor 1
'    If Val(rDSide) <= 0 Or Val(rDSide) > 3 Then rDSide = "1" ' bottom
'    If Val(rDzOrderMode) < 1 Or Val(rDzOrderMode) > 10 Then rDzOrderMode = "0" ' always on top
'    If Val(rDOffset) < -100 Or Val(rDOffset) > 100 Then rDOffset = "0" ' in the middle
'    If Val(rDvOffset) < -15 Or Val(rDvOffset) > 128 Then rDvOffset = "0" ' at the bottom edge
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryPosition_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryPosition of Module Module1"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : readRegistryStyle
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : read the settings from the registry
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryStyle()
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
'    rDFontOutlineOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineOpacity")
'    rDFontShadowOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowOpacity")
'
'    Call validateRegistryStyle
'
'
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryStyle
'' Author    : beededea
'' Date      : 17/06/2020
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
'    'Dim MyFile As String
'    Dim MyPath  As String
'    'Dim themePresent As Boolean
'    'Dim myName As String
'
'    On Error GoTo validateRegistryStyle_Error
'
'    'dean - this needs to be addressed when SD obtains its own skinning ability
'
'    MyPath = rdAppPath & "\Skins\" '"E:\Program Files (x86)\RocketDock\Skins\"
'    'themePresent = False
'
'    If Not DirExists(MyPath) Then
'        MsgBox "WARNING - The skins folder is not present in the correct location " & rdAppPath
'    End If
'
''    rDFontColor - difficult to check validity of a colour but some code is coming to ensure no corruption *1
'
'    If Val(rDThemeOpacity) < 1 Or Val(rDThemeOpacity) > 100 Then rDThemeOpacity = "100" '
'    If Val(rDHideLabels) < 0 Or Val(rDHideLabels) > 1 Then rDHideLabels = "0" '
'
'    Dim I As Integer
'    Dim fontPresent As Boolean
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
'    If rDFontFlags <= 0 Or rDFontFlags > 15 Then rDFontFlags = 0
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
'    If Val(rDFontOutlineOpacity) <= 0 Or Val(rDFontOutlineOpacity) > 100 Then rDFontOutlineOpacity = "100" '
'    If Val(rDFontShadowOpacity) <= 0 Or Val(rDFontShadowOpacity) > 100 Then rDFontShadowOpacity = "100" '
'
'    'validation ends
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryStyle_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryStyle of Module Module1"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : readRegistryBehaviour
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub readRegistryBehaviour()
'    ' read the settings from the registry
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
'    'dock.autoHideChecker.Interval = Val(rDAutoHideDelay)
'
'End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : validateRegistryBehaviour
'' Author    : beededea
'' Date      : 17/06/2020
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
'    If Val(rDIconActivationFX) <= 0 And Val(rDIconActivationFX) > 2 Then rDIconActivationFX = "2"
'    If Val(rDAutoHide) <= 0 And Val(rDAutoHide) > 1 Then rDAutoHide = "1"
'    If Val(rDAutoHideTicks) <= 0 And Val(rDAutoHideTicks) > 1000 Then rDAutoHideTicks = "200"
'    If Val(rDAutoHideDelay) <= 0 And Val(rDAutoHideDelay) > 2000 Then rDAutoHideDelay = "200"
'    If Val(rDMouseActivate) <= 0 And Val(rDMouseActivate) > 1 Then rDMouseActivate = "1"
'    If Val(rDPopupDelay) <= 0 And Val(rDPopupDelay) > 1000 Then rDPopupDelay = "100"
'
'    If Val(rDAnimationInterval) <= 0 And Val(rDAnimationInterval) > 20 Then rDAnimationInterval = "1"
'
'   On Error GoTo 0
'   Exit Sub
'
'validateRegistryBehaviour_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryBehaviour of Module Module1"
'
'    End Sub


''---------------------------------------------------------------------------------------
'' Procedure : ExeFileName
'' Author    : Bob77 https://stackoverflow.com/questions/24591408/vb6-dll-get-calling-application-path
'' Date      : 24/08/2020
'' Purpose   : unused, keep until tested under XP
''---------------------------------------------------------------------------------------
''
'Public Function ExeFileName(idProc As Long) As String
'    Dim Size As Long
'
'   On Error GoTo ExeFileName_Error
'
'    ExeFileName = Space$(256)
'    'idProc = GetCurrentProcess()
'    Size = GetModuleFileNameEx(idProc, _
'                               GetModuleHandle(API_NULL), _
'                               ExeFileName, _
'                               256)
'    ExeFileName = Left$(ExeFileName, Size)
'
'   On Error GoTo 0
'   Exit Function
'
'ExeFileName_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExeFileName of Module mdlMain"
'End Function

'Private Function ExePathFromProcID(idProc As Long) As String
'
'    Dim S As String
'    Dim c As Long
'    Dim hModule As Long
'    Dim ProcHndl As Long
'
'    S = String$(MAX_PATH, 0)
'    ProcHndl = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
'
'    If ProcHndl Then
'        If EnumProcessModules(ProcHndl, hModule, 4, c) <> 0 Then c = GetModuleFileNameEx(ProcHndl, hModule, S, MAX_PATH)
'        If c Then ExePathFromProcID = Left$(S, c)
'        CloseHandle ProcHndl
'    End If
'
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : GetDriveString
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Determine the number and name of drives using VB alone
''             redundant code - now happens using getdrives at form init
''---------------------------------------------------------------------------------------
''
'Public Function GetDriveString() As String
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
'    ' variables declared
'    Dim i As Long
'    Dim builtString As String
'
'    'initialise the dimensioned variables
'    i = 0
'    builtString = ""
'
'    '===========================
'    'pure VB approach, no controls required Gary Beene
'    'drive letters are found in positions 1-UBound(Letters)
'    '"C:\ D:\ E:\ &frameProperties"
'
'    On Error GoTo GetDriveString_Error
'    If debugflg = 1 Then DebugPrint "%" & "GetDriveString"
'
'    For i = 1 To 26
'        If ValidDrive(Chr$(96 + i)) = True Then
'            builtString = builtString + UCase$(Chr$(96 + i)) & ":\    "
'        End If
'    Next i
'
'    GetDriveString = builtString
'
'   On Error GoTo 0
'   Exit Function
'
'GetDriveString_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriveString of Module Common"
'End Function

''---------------------------------------------------------------------------------------
'' Procedure : ValidDrive
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Check if the drive found is a valid one
''             redundant code - now happens using getdrives at form init
''---------------------------------------------------------------------------------------
''
'Public Function ValidDrive(ByVal d As String) As Boolean
'    ' variables declared
'    Dim Temp As String
'
'    'initialise the dimensioned variables
'    Temp = ""
'
'    On Error GoTo ValidDrive_Error
'    If debugflg = 1 Then DebugPrint "%" & "ValidDrive"
'    On Error GoTo driveerror
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidDrive of Module Common"
'End Function
''---------------------------------------------------------------------------------------
'' Procedure : setNewStartPoint
'' Author    : beededea
'' Date      : 07/04/2020
'' Purpose   : sets the start point for the expanded, animated dock
''---------------------------------------------------------------------------------------
''
'Private Sub setNewStartPoint(dockWidth As Long)
'    Dim screenWidthPixels As Integer
'    Dim proportionalOffset As Integer
'    Dim hOffsetPxls As Integer
'    Dim centreOnScreen As Boolean
'    Dim a As Integer
'
'    Dim iconleftposition As Long
'    Dim smalliconstotheleft As Long
'    Dim smalliconsSize As Long
'    Dim widthOfLeftexpandingIcon As Long
'    Dim leftIconSizeinTwips As Long
'
'    On Error GoTo setNewStartPoint_Error
'
'    'initialise all declared variables
'    screenWidthPixels = 0
'    proportionalOffset = 0
'    hOffsetPxls = 0
'    centreOnScreen = False
'    a = 0
'
'    ' it is currently calculating the width of the new dock and placing the start point so it evenly spaces itself out from the middle of the screen
'    ' this causes the wrong icon to appear expanded in the dock as the animation places an expanded icon over due to current cursor position and not expand the first icon you chose...
'
'    ' instead, it needs to calculate where the dock start point needs to be in relation to the currently expanded icon.
'    ' ((this icon centre position (cursor position)) minus (this icon width, ie. 128 / 2) or (this icon left position)) + current width of the expanding icon to the left + (the number of small icons to the left x small icon size)
'
'    If centreOnScreen = True Then
'        screenWidthPixels = (dockWidth / screenTwipsPerPixelX)
'        hOffsetPxls = ((screenWidthPixels - dockWidth) / 2) ' dockwidth is calculated as the combined width of the expanded areas plus the smaller icon sizes combined
'        proportionalOffset = hOffsetPxls + (hOffsetPxls * (rDOffset / 100))
'        'iconLeftmostPointPxls = proportionalOffset
'    Else
'        '     6645                         11     (extended task mgr)
'        iconleftposition = iconPosLeftTwips(iconIndex) 'this icon left position
'        '     2580                 172                 15
'        leftIconSizeinTwips = leftIconSize * screenTwipsPerPixelX ' left expanded icon size
'        '       10                  11
'        smalliconstotheleft = iconIndex - 1 ' the number of small icons to the left ' assuming the number of icons animated is just three
'        '
'        '      3900               10                  26              15
'        smalliconsSize = smalliconstotheleft * iconSizeSmallPxls * screenTwipsPerPixelX
'
'        ' (this icon left position) + leftIconSize + (the number of small icons to the left x small icon size)
'        '                     5527                  2625                             3900
'        iconLeftmostPointPxls = iconleftposition - (leftIconSizeinTwips + smalliconsSize)
'
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'setNewStartPoint_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setNewStartPoint of Form dock"
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetShortcutBinaryName
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : This is a method of getting the path ONLY from a LNK file without using shell scripting - unused
'---------------------------------------------------------------------------------------
'
'Private Function GetShortcutBinaryName(ByVal full_name As String, ByRef Path As String) As String
'
'    Dim strBuff As String
'    Dim strArr() As String
'    Dim lngIdx As Long
'
'   On Error GoTo GetShortcutBinaryName_Error
'
'    Open full_name For Binary As #1
'        strBuff = Space$(LOF(1))
'        Get #1, , strBuff
'    Close #1
'
'    strArr = Split(strBuff, Chr(0))
'
'    ' this is the current bodge
'    For lngIdx = 0 To UBound(strArr)
'        If InStr(1, strArr(lngIdx), ".exe", vbTextCompare) <> 0 Then
'            MsgBox strArr(lngIdx)
'        End If
'    Next
'
'   On Error GoTo 0
'   Exit Function
'
'GetShortcutBinaryName_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShortcutBinaryName of Form dock"
'
'End Function

'
'Function ReadFileBytes(FileName As String) As Byte()
'Dim FNr&: FNr = FreeFile
'  Open FileName For Binary Access Read As FNr
'    ReDim ReadFileBytes(0 To LOF(FNr) - 1)
'    Get FNr, , ReadFileBytes
'  Close FNr
'End Function
