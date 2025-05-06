Attribute VB_Name = "common2"
'------------------------------------------------------------
' common2.bas
'
' Public procedures that appear in just two of the programs as an included module common2.bas,
' specifically, dock settings and steamy dock itself.
'
' Note: If you make a change here it affects the two programs dynamically
'------------------------------------------------------------

' .01 common2.bas DAEB 27/01/2021 Changed validation of the popup delay parameter used for fading the dock back in, now 1 second.
' .02 DAEB 01/02/2021 common2.bas always use the dockAppPath so it works on both docks
' .03 DAEB 03/03/2021 common2.bas bugfix - bottom position 0 is top
' .04 DAEB 11/03/2021 common2 added validation for the continuous hide value
' .05 DAEB 12/07/2021 common2.bas Add the BounceZone as a configurable variable.

Option Explicit

' Rocketdock global configuration variables START
Public rDOptionsTabIndex As String
Public rDStartupRun As String
Public rdStartupRunString As String
Public rDIconQuality As String
Public rDIconOpacity As String
Public rDZoomOpaque      As String
Public rDIconMin      As String
Public rDHoverFX      As String
Public rdIconMax      As String
Public rDZoomWidth      As String
Public rDZoomTicks      As String
Public rDMonitor      As String
Public rDSide      As String
Public rDzOrderMode      As String
Public rDOffset      As String
Public rDvOffset      As String
Public rDtheme      As String
Public rDWallpaper      As String
Public rDWallpaperStyle      As String
Public rDAutomaticWallpaperChange As String
Public rDWallpaperTimerIntervalIndex As String
Public rDWallpaperTimerInterval As String
Public rDWallpaperLastTimeChanged As String

Public rDMoveWinTaskbar As String

Public rDThemeOpacity      As String
Public rDHideLabels      As String
Public rDFontName      As String
Public rDFontColor      As String
Public rDFontSize As String
Public rDFontCharSet  As String
Public rDFontFlags      As String

'Public rDFontStrength      As Boolean
'Public rDFontItalics      As Boolean

Public rDFontShadowColor      As String
Public rDFontOutlineColor      As String
Public rDFontOutlineOpacity      As String
Public rDFontShadowOpacity      As String

Public rDIconActivationFX     As String
Public rDSoundSelection As String
Public rDAutoHide     As String
Public rDAutoHideTicks     As String
Public rDAutoHideDelay     As String
Public rDMouseActivate     As String
Public rDPopupDelay     As String
Public rDVersion As String
'Public rDCustomIconFolder As String
Public rDHotKeyToggle As String
Public rDLangID As String
Public rDAnimationInterval As String
Public rDSkinSize As String

Public sDSkinSize As Long ' the Steamydock version
Public sDSplashStatus As String

Public sDFontOpacity As String
Public sDAutoHideType As String
Public sDShowLblBacks As String ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files

Public sDContinuousHide As String 'nn
Public sDBounceZone As String ' .05 DAEB 12/07/2021 common2.bas Add the BounceZone as a configurable variable.
    
' development
Public sDDebugFlg As String
Public sDDefaultEditor As String

Public Const SM_CMONITORS = 80

'API to test the system, specifically the number of monitors
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'------------------------------------------------------ STARTS
' Wallpaper changing functions and vars

'Retrieves or sets the value of one of the system-wide parameters
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPIF_SENDWININICHANGE = &H2 'Send Change Message
Public Const SPIF_UPDATEINIFILE = &H1 'Update INI File
Public Const SPI_SETDESKWALLPAPER = 20 'Change Wallpaper
'------------------------------------------------------ ENDS

' Rocketdock global configuration variables END


'---------------------------------------------------------------------------------------
' Procedure : readDockSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read
'---------------------------------------------------------------------------------------
'
Public Sub readDockSettingsFile(ByVal location As String, ByVal settingsFile As String)
    
    'SteamyDock settings only
    On Error GoTo readDockSettingsFile_Error
    If debugflg = 1 Then debugLog "% sub readDockSettingsFile"

    If fFExists(dockSettingsFile) Then
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
        rDRunAppInterval = GetINISetting("Software\SteamyDock\DockSettings", "RunAppInterval", dockSettingsFile)
'        rDAlwaysAsk = GetINISetting("Software\SteamyDock\DockSettings", "AlwaysAsk", dockSettingsFile)
        rDDefaultDock = GetINISetting("Software\SteamyDock\DockSettings", "DefaultDock", dockSettingsFile)
        rDAnimationInterval = GetINISetting("Software\SteamyDock\DockSettings", "AnimationInterval", dockSettingsFile)
        rDSkinSize = GetINISetting("Software\SteamyDock\DockSettings", "SkinSize", dockSettingsFile)
        sDSplashStatus = GetINISetting("Software\SteamyDock\DockSettings", "SplashStatus", dockSettingsFile)
        sDShowIconSettings = GetINISetting("Software\SteamyDock\DockSettings", "ShowIconSettings", dockSettingsFile) '' .14 DAEB 01/05/2021 docksettings added checkbox and values to show icon settings utility when adding an icon to the dock
        
        sDFontOpacity = GetINISetting("Software\SteamyDock\DockSettings", "FontOpacity", settingsFile)
        sDAutoHideType = GetINISetting("Software\SteamyDock\DockSettings", "AutoHideType", settingsFile)
        sDShowLblBacks = GetINISetting("Software\SteamyDock\DockSettings", "ShowLblBacks", settingsFile) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files
        sDContinuousHide = GetINISetting("Software\SteamyDock\DockSettings", "ContinuousHide", settingsFile) ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files ' nn
        sDBounceZone = GetINISetting("Software\SteamyDock\DockSettings", "BounceZone", settingsFile) ' .05 DAEB 12/07/2021 common2.bas Add the BounceZone as a configurable variable.

        ' development
        sDDefaultEditor = GetINISetting(location, "dockDefaultEditor", settingsFile)
        sDDebugFlg = GetINISetting(location, "debugFlg", settingsFile)
        debugflg = Val(sDDebugFlg)
    End If
        
    sDSkinSize = Val(rDSkinSize)
        
    'if the above settings do not exist in the older RD settings file then no error is thrown so it works for both

    'RocketDock compatible settings only
    rDVersion = GetINISetting(location, "Version", settingsFile)
    rDHotKeyToggle = GetINISetting(location, "HotKey-Toggle", settingsFile)
            
    rDtheme = GetINISetting(location, "Theme", settingsFile)
    rDWallpaper = GetINISetting(location, "Wallpaper", settingsFile)
    rDWallpaperStyle = GetINISetting(location, "WallpaperStyle", settingsFile)
    rDAutomaticWallpaperChange = GetINISetting(location, "AutomaticWallpaperChange", settingsFile)
    rDWallpaperTimerIntervalIndex = GetINISetting(location, "WallpaperTimerIntervalIndex", settingsFile)
    rDWallpaperTimerInterval = GetINISetting(location, "WallpaperTimerInterval", settingsFile)
    rDWallpaperLastTimeChanged = GetINISetting(location, "WallpaperLastTimeChanged", settingsFile)
    
    rDMoveWinTaskbar = GetINISetting(location, "MoveWinTaskbar", settingsFile)
    
    rDThemeOpacity = GetINISetting(location, "ThemeOpacity", settingsFile)
    rDIconOpacity = GetINISetting(location, "IconOpacity", settingsFile)
    rDFontSize = GetINISetting(location, "FontSize", settingsFile)
    rDFontFlags = GetINISetting(location, "FontFlags", settingsFile)
    rDFontName = GetINISetting(location, "FontName", settingsFile)
    rDFontColor = GetINISetting(location, "FontColor", settingsFile)
    rDFontCharSet = GetINISetting(location, "FontCharSet", settingsFile)
    rDFontOutlineColor = GetINISetting(location, "FontOutlineColor", settingsFile)
    rDFontOutlineOpacity = GetINISetting(location, "FontOutlineOpacity", settingsFile)
    rDFontShadowColor = GetINISetting(location, "FontShadowColor", settingsFile)
    rDFontShadowOpacity = GetINISetting(location, "FontShadowOpacity", settingsFile)
    rDIconMin = GetINISetting(location, "IconMin", settingsFile)
    rdIconMax = GetINISetting(location, "IconMax", settingsFile)
    rDZoomWidth = GetINISetting(location, "ZoomWidth", settingsFile)
    rDZoomTicks = GetINISetting(location, "ZoomTicks", settingsFile)
    rDAutoHide = GetINISetting(location, "AutoHide", settingsFile) '  26/10/2020 docksettings .03 DAEB fixed a previous find/replace bug causing the autohide setting to fail to both save and read
    rDAutoHideTicks = GetINISetting(location, "AutoHideTicks", settingsFile)
    rDAutoHideDelay = GetINISetting(location, "AutoHideDelay", settingsFile)
    rDPopupDelay = GetINISetting(location, "PopupDelay", settingsFile)
    rDIconQuality = GetINISetting(location, "IconQuality", settingsFile)
    rDLangID = GetINISetting(location, "LangID", settingsFile)
    rDHideLabels = GetINISetting(location, "HideLabels", settingsFile)
    rDZoomOpaque = GetINISetting(location, "ZoomOpaque", settingsFile)
    rDLockIcons = GetINISetting(location, "LockIcons", settingsFile)
    rDRetainIcons = GetINISetting(location, "RetainIcons", settingsFile) ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    
    rDManageWindows = GetINISetting(location, "ManageWindows", settingsFile)
    rDDisableMinAnimation = GetINISetting(location, "DisableMinAnimation", settingsFile)
    rDShowRunning = GetINISetting(location, "ShowRunning", settingsFile)
    rDOpenRunning = GetINISetting(location, "OpenRunning", settingsFile)
    rDHoverFX = GetINISetting(location, "HoverFX", settingsFile)
    rDzOrderMode = GetINISetting(location, "zOrderMode", settingsFile)
    rDMouseActivate = GetINISetting(location, "MouseActivate", settingsFile)
    rDIconActivationFX = GetINISetting(location, "IconActivationFX", settingsFile)
    rDSoundSelection = GetINISetting(location, "SoundSelection", settingsFile)
    
    rDMonitor = GetINISetting(location, "Monitor", settingsFile)
    rDSide = GetINISetting(location, "Side", settingsFile)
    rDOffset = GetINISetting(location, "Offset", settingsFile)
    rDvOffset = GetINISetting(location, "vOffset", settingsFile)
    rDOptionsTabIndex = GetINISetting("Software\DockSettings", "OptionsTabIndex", toolSettingsFile)
    '= GetINISetting("Software\SteamyDock\DockSettings\WindowFilters", "Count", 0, settingsFile)
    
   On Error GoTo 0
   Exit Sub

readDockSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockSettingsFile of Module common2"

End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : readRegistryBehaviour
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryBehaviour()
    ' read the settings from the registry

   'general items

'   IconActivationFX ' Icon Activation Effect
'   AutoHide         ' AutoHide
'   AutoHideTicks    ' AutoHide Duration
'   AutoHideDelay    ' AutoHide Delay
'   MouseActivate    ' Pop-up on Mouseover
'   PopupDelay       ' PopupDelay


   On Error GoTo readRegistryBehaviour_Error

    rDIconActivationFX = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "IconActivationFX")
    rDAutoHide = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHide")
    rDAutoHideTicks = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideTicks")
    rDAutoHideDelay = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "AutoHideDelay")
    rDMouseActivate = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "MouseActivate")
    rDPopupDelay = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "PopupDelay")

    Call validateRegistryBehaviour
    
    'dock.autoHideChecker.Interval = Val(rDAutoHideDelay)

   On Error GoTo 0
   Exit Sub

readRegistryBehaviour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryBehaviour of Module common2"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : validateRegistryBehaviour
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateRegistryBehaviour()

    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
    ' when Rocketdock is restarted

   On Error GoTo validateRegistryBehaviour_Error

    If Val(rDIconActivationFX) <= 0 And Val(rDIconActivationFX) > 2 Then rDIconActivationFX = "2"
    If Val(rDAutoHide) <= 0 And Val(rDAutoHide) > 1 Then rDAutoHide = "1"
    If Val(rDAutoHideTicks) <= 0 And Val(rDAutoHideTicks) > 5000 Then rDAutoHideTicks = "200"
    If Val(rDAutoHideDelay) <= 0 And Val(rDAutoHideDelay) > 5000 Then rDAutoHideDelay = "200"
    If Val(rDMouseActivate) <= 0 And Val(rDMouseActivate) > 1 Then rDMouseActivate = "1"
    If Val(rDPopupDelay) <= 0 And Val(rDPopupDelay) > 5000 Then rDPopupDelay = "1000" ' ' .01 STARTS DAEB 27/01/2021 Changed validation of the popup delay parameter used for fading the dock back in, now 1 second.
    
    If Val(rDAnimationInterval) <= 0 And Val(rDAnimationInterval) > 20 Then rDAnimationInterval = "1"
    If Val(sDAutoHideType) <= 0 And Val(sDAutoHideType) > 2 Or sDAutoHideType = vbNullString Then sDAutoHideType = "0"
    
    If rDHotKeyToggle = vbNullString Then rDHotKeyToggle = "F11" ' .02 STARTS DAEB 27/01/2021 Added validation of the function key used for fading the dock back in.
    
    
    If Val(sDContinuousHide) <= 1 And Val(sDContinuousHide) >= 120 Then sDContinuousHide = "10" '.04 DAEB 11/03/2021 common2 added validation for the continuous hide value
    If Val(sDBounceZone) <= 1 And Val(sDBounceZone) >= 120 Then sDBounceZone = "75" ' .05 DAEB 12/07/2021 common2.bas Add the BounceZone as a configurable variable.

    
    
   On Error GoTo 0
   Exit Sub

validateRegistryBehaviour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryBehaviour of Module common2"
    
    End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error

    If Val(rDRunAppInterval) * 1000 >= 65536 Then rDRunAppInterval = "65"
        
    If rDWallpaper = "" Then rDWallpaper = "none selected"
    If rDWallpaperStyle = "" Then rDWallpaperStyle = "Centre"
    If rDAutomaticWallpaperChange = "" Then rDAutomaticWallpaperChange = "0"
    If rDWallpaperTimerIntervalIndex = "" Then rDWallpaperTimerIntervalIndex = "4" ' 1 hour
    If rDWallpaperTimerInterval = "" Then rDWallpaperTimerIntervalIndex = "60" ' 1 hour
    
    If rDWallpaperLastTimeChanged = "" Then rDWallpaperLastTimeChanged = Now()
    
    If rDMoveWinTaskbar = "" Then rDMoveWinTaskbar = "1"
        
    ' validate the relevant entries from whichever source
    validateRegistryGeneral
    validateRegistryIcons
    validateRegistryBehaviour
    validateRegistryStyle
    validateRegistryPosition


    
'    sDSplashStatus = "1"
'    chkSplashStatus.Value = Val(sDSplashStatus)

   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of Module common2"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : validateRegistryStyle
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateRegistryStyle()
    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
    ' when Rocketdock is restarted

    ' read the skins available from the rocketdock folder

    'Dim MyFile As String
    'Dim themePresent As Boolean
    'Dim myName As String
    
    Dim MyPath As String: MyPath = ""
    Dim I As Integer: I = 0
    Dim fontPresent As Boolean: fontPresent = False

    On Error GoTo validateRegistryStyle_Error
    
    ' .02 DAEB 01/02/2021 common2.bas always use the dockAppPath so it works on both docks
    MyPath = dockAppPath & "\Skins\" '"E:\Program Files (x86)\RocketDock\Skins\"
    'themePresent = False

    If Not fDirExists(MyPath) Then
        MsgBox "WARNING - The skins folder is not present in the correct location " & dockAppPath
    End If

'    rDFontColor - difficult to check validity of a colour but some code is coming to ensure no corruption *1

    If Val(rDThemeOpacity) < 1 Or Val(rDThemeOpacity) > 100 Then rDThemeOpacity = "100" '
    If Val(rDHideLabels) < 0 Or Val(rDHideLabels) > 1 Then rDHideLabels = "0" '
    If Val(sDShowLblBacks) < 0 Or Val(sDShowLblBacks) > 1 Then sDShowLblBacks = "0" ' 25/10/2020 docksettings .02 DAEB add the logic for saving/reading icon label background string to configuration files


    fontPresent = False
    For I = 0 To Screen.FontCount - 1 ' Determine number of fonts.
        If rDFontName = Screen.Fonts(I) Then fontPresent = True
    Next I
    If fontPresent = False Then rDFontName = "Times New Roman" '

    If Abs(Val(rDFontSize)) < 2 Or Abs(Val(rDFontSize)) > 29 Then rDFontSize = "-29" '

    ' rDFontCharSet = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontCharSet")
    ' how to validate a character set? - not supported

    ' validate font flags
    ' 0 - no qualifiers or alterations
    ' 1 - bold
    ' 2 - light italics
    ' 3 - bold italics
    ' 4 - strikeout & light
    ' 6 - underline and italics
    ' 7 - bold, italics & underline
    ' 10 - strikeout & italics
    ' 11 - bold, italics & strikeout
    ' 13 - strikeout & italics
    ' 14 - underline, strikeout and italics
    ' 15 - bold, underline, strikeout and italics

    If rDFontFlags <= 0 Or rDFontFlags > 15 Then rDFontFlags = 0

    If Not IsNumeric(rDFontShadowColor) Then
        rDFontShadowColor = 0
    End If

    If Not IsNumeric(rDFontOutlineColor) Then
        rDFontOutlineColor = 0
    End If

    ' how to validate colour?

    If Val(rDFontOutlineOpacity) <= 0 Or Val(rDFontOutlineOpacity) > 100 Then rDFontOutlineOpacity = "100" '
    If Val(rDFontShadowOpacity) <= 0 Or Val(rDFontShadowOpacity) > 100 Then rDFontShadowOpacity = "100" '

    If Val(sDFontOpacity) <= 0 Or Val(sDFontOpacity) > 100 Then sDFontOpacity = "100" '


    'validation ends

   On Error GoTo 0
   Exit Sub

validateRegistryStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryStyle of Module common2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : readRegistryStyle
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : read the settings from the registry
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryStyle()

   'style items

'   Theme 'Theme
'   ThemeOpacity
'   HideLabels
'   FontName
'   FontShadowColor
'   FontOutlineColor
'   FontOutlineOpacity
'   FontShadowOpacity

   On Error GoTo readRegistryStyle_Error

    rDtheme = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Theme")
'    rDWallpaper = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Wallpaper")
'    rDWallpaperStyle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "WallpaperStyle")
'    rDAutomaticWallpaperChange = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "AutomaticWallpaperChange")
'    rdWallpaperTimerIntervalIndex = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "WallpaperTimerIntervalIndex")

    rDThemeOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ThemeOpacity")
    rDHideLabels = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "HideLabels")
    rDFontName = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontName")
    rDFontColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontColor")

    rDFontSize = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontSize")
    rDFontCharSet = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontCharSet")
    rDFontFlags = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontFlags")

    rDFontShadowColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowColor")
    rDFontOutlineColor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineColor")
    rDFontOutlineOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontOutlineOpacity")
    rDFontShadowOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "FontShadowOpacity")
    
    Call validateRegistryStyle


   On Error GoTo 0
   Exit Sub

readRegistryStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryStyle of Module common2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateRegistryPosition
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateRegistryPosition()


    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
    ' when Rocketdock is restarted

   On Error GoTo validateRegistryPosition_Error

    If Val(rDMonitor) < 0 Or Val(rDMonitor) > 10 Then rDMonitor = "0" 'monitor 0 is the default meaning the first monitor
    If Val(rDSide) < 0 Or Val(rDSide) > 3 Then rDSide = "1" ' .03 DAEB 03/03/2021 common2.bas bugfix - bottom position 0 is top
    If Val(rDzOrderMode) < 1 Or Val(rDzOrderMode) > 10 Then rDzOrderMode = "0" ' always on top
    If Val(rDOffset) < -100 Or Val(rDOffset) > 100 Then rDOffset = "0" ' in the middle
    If Val(rDvOffset) < -15 Or Val(rDvOffset) > 128 Then rDvOffset = "0" ' at the bottom edge

   On Error GoTo 0
   Exit Sub

validateRegistryPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryPosition of Module common2"

End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : readRegistryPosition
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryPosition()

   'style items

'   Monitor 'Monitor
'   Side    ' Side
'   zOrderMode  ' zOrderMode
'   Offset  ' Offset
'   vOffset ' vOffset

   On Error GoTo readRegistryPosition_Error

    rDMonitor = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Monitor")
    rDSide = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Side")
    rDzOrderMode = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "zOrderMode")
    rDOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "Offset")
    rDvOffset = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "vOffset")

    Call validateRegistryPosition

    

   On Error GoTo 0
   Exit Sub

readRegistryPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryPosition of Module common2"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateRegistryIcons
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateRegistryIcons()
    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
    ' when Rocketdock is restarted

   On Error GoTo validateRegistryIcons_Error

    'If Val(rDMonitor) < 0 Or Val(rDMonitor) > 10 Then rDMonitor = "0" 'monitor 1
    If Val(rDIconOpacity) < 50 Or Val(rDIconOpacity) > 100 Then rDIconOpacity = "100" 'fully opaque
    If Val(rDZoomOpaque) <= 0 Or Val(rDZoomOpaque) > 1 Then rDZoomOpaque = "1" 'zooms opaque
    If Val(rDIconMin) < 16 Or Val(rDIconMin) > 128 Then rDIconMin = "16" 'small
    If Val(rDHoverFX) <= 0 Or Val(rDHoverFX) > 4 Then rDHoverFX = "1" 'bounce ' .14 DAEB 29/04/2021 docksettings Set the default zoom types available to the type of dock selected

    If Val(rdIconMax) < 1 Or Val(rdIconMax) > 256 Then rdIconMax = "256" 'largest size
    'MsgBox "icnomax = " & rdIconMax
    
    If Val(rDZoomWidth) < 2 Or Val(rDZoomWidth) > 10 Then rDZoomWidth = "4" ' just a few expanded
    If Val(rDZoomTicks) < 100 Or Val(rDZoomTicks) > 500 Then rDZoomTicks = "100" ' 100ms

   On Error GoTo 0
   Exit Sub

validateRegistryIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryIcons of Module common2"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readRegistryIcons
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryIcons()

    ' read the icon configuration settings from the registry


   'icon items

'   IconQuality ' Icon Quality
'   IconOpacity ' Icon Opacity
'   ZoomOpaque  ' Zoom Opaque
'   IconMin     ' Size
'   HoverFX     ' Hover Effect
'   IconMax     ' Zoom
'   ZoomWidth   ' Zoom Width
'   ZoomTicks   ' Duration

   On Error GoTo readRegistryIcons_Error

    rDIconQuality = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconQuality")
    rDIconOpacity = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconOpacity")
    rDZoomOpaque = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomOpaque")
    rDIconMin = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconMin")
    rDHoverFX = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "HoverFX")
    rdIconMax = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "IconMax")
    rDZoomWidth = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomWidth")
    rDZoomTicks = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ZoomTicks")

    Call validateRegistryIcons

   On Error GoTo 0
   Exit Sub

readRegistryIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryIcons of Module common2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateRegistryGeneral
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub validateRegistryGeneral()
    ' testing and adjusting the values to the ranges allowed, preventing corrupt values
    ' this is required as running the program from within the IDE without admin rights results in corrupt data from the registry
    ' when Rocketdock is restarted
    
   On Error GoTo validateRegistryGeneral_Error

    If Val(rDLockIcons) <= 0 And Val(rDLockIcons) > 1 Then rDLockIcons = "1" '
    If Val(rDRetainIcons) <= 0 And Val(rDRetainIcons) > 1 Then rDRetainIcons = "1" ' .18 DAEB 07/09/2022 docksettings save and restore the chkRetainIcons checkbox value
    If Val(rDOpenRunning) <= 0 And Val(rDOpenRunning) > 1 Then rDOpenRunning = "1" '
    If Val(rDShowRunning) <= 0 And Val(rDShowRunning) > 1 Then rDShowRunning = "1" '
    If Val(rDManageWindows) <= 0 And Val(rDManageWindows) > 1 Then rDManageWindows = "1" '
    If Val(rDDisableMinAnimation) <= 0 And Val(rDDisableMinAnimation) > 1 Then rDDisableMinAnimation = "1" '

    ' development
    If sDDebugFlg = vbNullString Then sDDebugFlg = "0"
    If sDDefaultEditor = vbNullString Then sDDefaultEditor = vbNullString
        
   On Error GoTo 0
   Exit Sub

validateRegistryGeneral_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateRegistryGeneral of Module common2"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : readRegistryGeneral
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistryGeneral()
   'general items
   
'   LockIcons ' lock items
'   OpenRunning 'Open Running Application Instance
'   ShowRunning 'Running Application Indicators
'   ManageWindows' Minimise Windows to the Dock
'   DisableMinAnimation 'Disable Minimise Animations

'HKEY_USERS\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run
' 02 00 00 00 00 00 00

    'Dim rdStartupRunString As String
    
   On Error GoTo readRegistryGeneral_Error

    rdStartupRunString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "RocketDock")
    If rdStartupRunString <> vbNullString Then
        rDStartupRun = "1"
    End If

    rDLockIcons = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "LockIcons")
    'rDRetainIcons unused by Rocketdock
    rDOpenRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "OpenRunning")
    rDShowRunning = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ShowRunning")
    rDManageWindows = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "ManageWindows")
    rDDisableMinAnimation = getstring(HKEY_CURRENT_USER, "Software\RocketDock\", "DisableMinAnimation")

    Call validateRegistryGeneral
    

   On Error GoTo 0
   Exit Sub

readRegistryGeneral_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryGeneral of Module common2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : readRegistry
' Author    : beededea
' Date      : 09/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub readRegistry()
    Dim useloop As Integer: useloop = 0

   On Error GoTo readRegistry_Error
   'If debugflg = 1 Then debugLog "%readRegistry"

     
     rDOptionsTabIndex = getstring(HKEY_CURRENT_USER, "Software\RocketDock", "OptionsTabIndex")
         
    ' get items from the registry that are required to 'default' the dock but aren't controlled by the dock settings utility
    rdIconCount = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Count")
    
    rDVersion = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "Version")
    rDCustomIconFolder = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "CustomIconFolder")
    rDHotKeyToggle = getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "HotKeyToggle")
         
     ' get the relevant entries from the registry
     readRegistryGeneral
     readRegistryIcons
     readRegistryBehaviour
     readRegistryStyle
     readRegistryPosition

   On Error GoTo 0
   Exit Sub

readRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistry of Module common2"
         
End Sub


'---------------------------------------------------------------------------------------
' Procedure : changeWallpaper
' Author    : HanneSThEGreaT https://forums.codeguru.com/showthread.php?497353-VB6-How-Do-I-Change-The-Windows-WallPaper
' Date      : 07/04/2025
' Purpose   : Routine to change the windows wallpaper
'---------------------------------------------------------------------------------------
'
Public Sub changeWallpaper(ByVal SelectedWallpaper As String, ByVal WallpaperStyle As String)

    Dim lReturn As Long: lReturn = 0 'Return of SysParInfo API

    On Error GoTo changeWallpaper_Error
    
    'Determine default WallPaper 'Style', ie. positioning
    If WallpaperStyle <> "Centre" And WallpaperStyle <> "Tile" And WallpaperStyle <> "Stretch" Then
        WallpaperStyle = "Stretch"
    End If
    
    'Write to the registry to allow Windows to determine wallpaper placement
    If WallpaperStyle = "Centre" Then
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"
    ElseIf WallpaperStyle = "Tile" Then
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1"
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"
    ElseIf WallpaperStyle = "Stretch" Then
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
        savestring HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2"
    End If
    
    'Set the WallPaper and trigger the system to apply it to the desktop
    lReturn = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, SelectedWallpaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

   On Error GoTo 0
   Exit Sub

changeWallpaper_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeWallpaper of Form dockSettings"

End Sub
