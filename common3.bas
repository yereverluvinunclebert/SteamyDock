Attribute VB_Name = "common3"
' .01 DAEB 31/01/2021 common3.bas Added new checkbox to determine if a post initiation dialog should appear
' .02 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app

Option Explicit

'------------------------------------------------------------
' common3.bas
'
' Public procedures that appear in just two of the programs as an included module common3.bas,
' specifically, enhanced icon settings and steamy dock itself.
'
' Note: If you make a change here it affects the two programs dynamically
'------------------------------------------------------------

Public gblOutsideDock As Boolean
Public iconLeftmostPointPxls As Single
Public iconRightmostPointPxls As Single
Public dockYEntrancePoint As Integer
'Public glbStartRecord As Integer
'Public gblRecordsToCommit As Integer

Public Type iconRecordTYPE
       iconRecordNumber As Integer
       iconFilename As String * 255
       iconFileName2 As String * 255
       iconTitle As String * 255
       iconCommand As String * 255
       iconArguments As String * 40
       iconWorkingDirectory As String * 255
       iconShowCmd As String * 1
       iconOpenRunning As String * 1
       iconIsSeparator As String * 1
       iconUseContext As String * 1
       iconDockletFile As String * 255
       iconUseDialog As String * 1
       iconUseDialogAfter As String * 1
       iconQuickLaunch As String * 1
       iconAutoHideDock As String * 1
       iconSecondApp As String * 255
       iconRunElevated As String * 1
       iconRunSecondAppBeforehand As String * 1
       iconAppToTerminate As String * 255
       iconDisabled As String * 1
End Type
 
Public iconVar As iconRecordTYPE
Public iconData As String

Public rdIconUpperBound As Integer
Public rdIconLowerBound As Integer
Public iconArrayUpperBound As Integer
Public iconArrayLowerBound As Integer

'
'---------------------------------------------------------------------------------------
' Procedure : putIconSettings
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save icon values to random access data file
'---------------------------------------------------------------------------------------
'
Public Sub putIconSettings(ByVal thisRecordNumber As Integer)

    On Error GoTo putIconSettings_Error
   
    Dim recordNumberToWrite As Integer: recordNumberToWrite = 0
   
    ' previously, records were written to a INI file starting at record number 0
    ' in a random access data file, record number 0 would generate an error
    ' to prevent a record number 0 bad record we increment the supplied record number
    recordNumberToWrite = thisRecordNumber
    
    ' always set the characteristsics of the first and last uneditable blank icons
    Call setFirstLastIcons(recordNumberToWrite)
    
    ' set the icon values into the binary data
    iconVar.iconRecordNumber = thisRecordNumber
    iconVar.iconFilename = sFilename
    iconVar.iconFileName2 = sFileName2
    iconVar.iconTitle = sTitle
    iconVar.iconCommand = sCommand
    iconVar.iconArguments = sArguments
    iconVar.iconWorkingDirectory = sWorkingDirectory
    iconVar.iconShowCmd = Val(sShowCmd)
    iconVar.iconOpenRunning = Val(sOpenRunning)
    iconVar.iconIsSeparator = Val(sIsSeparator)
    iconVar.iconUseContext = Val(sUseContext)
    iconVar.iconDockletFile = sDockletFile
    iconVar.iconUseDialog = Val(sUseDialog)
    iconVar.iconUseDialogAfter = Val(sUseDialogAfter)
    iconVar.iconQuickLaunch = Val(sQuickLaunch)
    iconVar.iconAutoHideDock = Val(sAutoHideDock)
    iconVar.iconSecondApp = sSecondApp
    iconVar.iconRunElevated = Val(sRunElevated)
    iconVar.iconRunSecondAppBeforehand = Val(sRunSecondAppBeforehand)
    iconVar.iconAppToTerminate = sAppToTerminate
    iconVar.iconDisabled = Val(sDisabled)

    Put #3, recordNumberToWrite, iconVar
                
   On Error GoTo 0
   Exit Sub

putIconSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure putIconSettings of Module Common"
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : getIconSettings
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Read icon values from random access data file
'---------------------------------------------------------------------------------------
'
Public Sub getIconSettings(ByVal thisRecordNumber As Integer)

    On Error GoTo getIconSettings_Error
   
    Dim recordNumberToRead As Integer: recordNumberToRead = 0
    
    ' previously, records were written to a INI file starting at record number 0
    ' in a random access data file, record number 0 would generate an error
    ' to prevent a record number 0 bad record we increment the supplied record number
    
    recordNumberToRead = thisRecordNumber
    '
    Get #3, recordNumberToRead, iconVar

    ' read the icon values from the binary data into the icon variables
    thisRecordNumber = iconVar.iconRecordNumber
    If thisRecordNumber = 0 Then thisRecordNumber = recordNumberToRead
    sFilename = RTrim$(iconVar.iconFilename)
    sFileName2 = RTrim$(iconVar.iconFileName2)
    sTitle = RTrim$(iconVar.iconTitle)
    sCommand = RTrim$(iconVar.iconCommand)
    sArguments = RTrim$(iconVar.iconArguments)
    sWorkingDirectory = RTrim$(iconVar.iconWorkingDirectory)
    sShowCmd = CStr(iconVar.iconShowCmd)
    sOpenRunning = CStr(iconVar.iconOpenRunning)
    sIsSeparator = CStr(iconVar.iconIsSeparator)
    sUseContext = CStr(iconVar.iconUseContext)
    sDockletFile = RTrim$(iconVar.iconDockletFile)
    sUseDialog = CStr(iconVar.iconUseDialog)
    sUseDialogAfter = CStr(iconVar.iconUseDialogAfter)
    sQuickLaunch = CStr(iconVar.iconQuickLaunch)
    sAutoHideDock = CStr(iconVar.iconAutoHideDock)
    sSecondApp = RTrim$(iconVar.iconSecondApp)
    sRunElevated = CStr(iconVar.iconRunElevated)
    sRunSecondAppBeforehand = CStr(iconVar.iconRunSecondAppBeforehand)
    sAppToTerminate = RTrim$(iconVar.iconAppToTerminate)
    sDisabled = CStr(iconVar.iconDisabled)
    
    ' always set the characteristsics of the first and last uneditable blank icons
    Call setFirstLastIcons(recordNumberToRead)
    
    On Error GoTo 0
   Exit Sub

getIconSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getIconSettings of Module Common"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setFirstLastIcons
' Author    : beededea
' Date      : 24/07/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setFirstLastIcons(ByVal thisRecordNumber As Integer)
    
    On Error GoTo setFirstLastIcons_Error

    ' first icon is always blank and non-editable
    If thisRecordNumber = 1 Then
        Call zeroAllIconCharacteristics
        If fFExists(App.Path & "\gog.png") Then
            sFilename = App.Path & "\gog.png"
        End If
        
        sTitle = "dockMinimum"
    End If
    
    ' the very last icon is always blank and non-editable
    ' we add 2 to the recordNumberToWrite, 1 for the conversion from settings.ini to a random access data and then another position above the maximum for a blank icon
    If thisRecordNumber = iconArrayUpperBound Then
        Call zeroAllIconCharacteristics
        sFilename = App.Path & "\blank.png"
        sTitle = "dockMaximum"
    End If

   On Error GoTo 0
   Exit Sub

setFirstLastIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFirstLastIcons of Module common3"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : writeIconSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub writeIconSettingsIni(ByVal iconNumberToWrite As Integer, Optional ByVal writeArray As Boolean)
    
    Dim errCnt As Integer: errCnt = 0
    Static writeDBFlag As Integer
        
    On Error GoTo writeIconSettingsIni_Error
   'If debugFlg = 1 Then debugLog "%writeIconSettingsIni"

'   If writeArray = False Then
   
        ' write the icon data to random access data file
        'Call putIconSettings(iconNumberToWrite)
        
        ' check to see if there have been any prior errors reading records from the db, if so, don't read any more records
        If writeDBFlag > 0 Then Exit Sub
        
         ' write the icon data to the SQLite database with error check returned
        errCnt = putIconSettingsIntoDatabase(iconNumberToWrite)
        writeDBFlag = writeDBFlag + errCnt
        
        ' the array cache was used for all the variables to speed up access when reading/writing the settings file
        ' this was due to using a settings.ini file using Windows APIs to read/write (very slow indeed)
        ' then the random access data file was implemented and the speed increased dramatically.
        
'        PutINISetting location, iconNumberToWrite & "-FileName", sFilename, settingsFile
'        PutINISetting location, iconNumberToWrite & "-FileName2", sFileName2, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Title", sTitle, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Command", sCommand, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Arguments", sArguments, settingsFile
'        PutINISetting location, iconNumberToWrite & "-WorkingDirectory", sWorkingDirectory, settingsFile
'        PutINISetting location, iconNumberToWrite & "-ShowCmd", sShowCmd, settingsFile
'        PutINISetting location, iconNumberToWrite & "-OpenRunning", sOpenRunning, settingsFile
'        PutINISetting location, iconNumberToWrite & "-RunElevated", sRunElevated, settingsFile
'
'        PutINISetting location, iconNumberToWrite & "-IsSeparator", sIsSeparator, settingsFile
'        PutINISetting location, iconNumberToWrite & "-UseContext", sUseContext, settingsFile
'        PutINISetting location, iconNumberToWrite & "-DockletFile", sDockletFile, settingsFile
'
'        'If defaultDock = 1 Then
'        PutINISetting location, iconNumberToWrite & "-UseDialog", sUseDialog, settingsFile
'        PutINISetting location, iconNumberToWrite & "-UseDialogAfter", sUseDialogAfter, settingsFile ' .03 DAEB 31/01/2021 common.bas Added new checkbox to determine if a post initiation dialog should appear
'        PutINISetting location, iconNumberToWrite & "-QuickLaunch", sQuickLaunch, settingsFile ' .10 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
'        PutINISetting location, iconNumberToWrite & "-AutoHideDock", sAutoHideDock, settingsFile  ' .12 DAEB 20/05/2021 common.bas Added new check box to allow autohide of the dock after launch of the chosen app
'        PutINISetting location, iconNumberToWrite & "-SecondApp", sSecondApp, settingsFile  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
'
'        PutINISetting location, iconNumberToWrite & "-RunSecondAppBeforehand", sRunSecondAppBeforehand, settingsFile
'        PutINISetting location, iconNumberToWrite & "-AppToTerminate", sAppToTerminate, settingsFile
'        PutINISetting location, iconNumberToWrite & "-Disabled", sDisabled, settingsFile  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
        
'    Else
'        sFileNameArray(iconNumberToWrite) = sFilename
'        sFileName2Array(iconNumberToWrite) = sFileName2
'        sTitleArray(iconNumberToWrite) = sTitle
'        sCommandArray(iconNumberToWrite) = sCommand
'        sArgumentsArray(iconNumberToWrite) = sArguments
'        sWorkingDirectoryArray(iconNumberToWrite) = sWorkingDirectory
'        sShowCmdArray(iconNumberToWrite) = sShowCmd
'        sOpenRunningArray(iconNumberToWrite) = sOpenRunning
'        sIsSeparatorArray(iconNumberToWrite) = sIsSeparator
'        sUseContextArray(iconNumberToWrite) = sUseContext
'        sDockletFileArray(iconNumberToWrite) = sDockletFile
'        sUseDialogArray(iconNumberToWrite) = sUseDialog
'        sUseDialogAfterArray(iconNumberToWrite) = sUseDialogAfter
'        sQuickLaunchArray(iconNumberToWrite) = sQuickLaunch
'        sAutoHideDockArray(iconNumberToWrite) = sAutoHideDock
'        sSecondAppArray(iconNumberToWrite) = sSecondApp
'        sRunElevatedArray(iconNumberToWrite) = sRunElevated
'        sRunSecondAppBeforehandArray(iconNumberToWrite) = sRunSecondAppBeforehand
'        sAppToTerminateArray(iconNumberToWrite) = sAppToTerminate
'        sDisabledArray(iconNumberToWrite) = sDisabled
'    End If
            
    On Error GoTo 0
   Exit Sub

writeIconSettingsIni_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeIconSettingsIni of Module Common"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : readIconSettingsIni
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : Reads an .INI File (SETTINGS.INI) using the icon number as a reference.
'             Then it assigns the value to a variable and an array cache
'             Or, alternatively read it directly from an array cache.
'---------------------------------------------------------------------------------------
'
Public Sub readIconSettingsIni(ByVal iconNumberToRead As Integer, Optional ByVal readArray As Boolean)

    Dim errCnt As Integer: errCnt = 0
    Static readDBFlag As Integer
        
    On Error GoTo readIconSettingsIni_Error
    
   ' If readArray = False Then
    
        ' obtain the icon data from random access data file
        'Call getIconSettings(iconNumberToRead) 'retained here for testing
        
        ' check to see if there have been any prior errors reading records from the db, if so, don't read any more records
        If readDBFlag > 0 Then Exit Sub
        
         ' obtain the icon data from the SQLite database with error check returned
        errCnt = getIconSettingsFromDatabase(iconNumberToRead)
        readDBFlag = readDBFlag + errCnt

        ' the array cache was used for all the variables to speed up access when reading/writing the settings file
        ' this was due to using a settings.ini file using Windows APIs to read/write (very slow indeed)
        ' then the random access data file was implemented and the speed increased dramatically.
    
        
        ' now write it straight away into the array cache
        
'        sFileNameArray(iconNumberToRead) = sFilename
'        sFileName2Array(iconNumberToRead) = sFileName2
'        sTitleArray(iconNumberToRead) = sTitle
'        sCommandArray(iconNumberToRead) = sCommand
'        sArgumentsArray(iconNumberToRead) = sArguments
'        sWorkingDirectoryArray(iconNumberToRead) = sWorkingDirectory
'        sShowCmdArray(iconNumberToRead) = sShowCmd
'        sOpenRunningArray(iconNumberToRead) = sOpenRunning
'        sIsSeparatorArray(iconNumberToRead) = sIsSeparator
'        sUseContextArray(iconNumberToRead) = sUseContext
'        sDockletFileArray(iconNumberToRead) = sDockletFile
'        sUseDialogArray(iconNumberToRead) = sUseDialog
'        sUseDialogAfterArray(iconNumberToRead) = sUseDialogAfter
'        sQuickLaunchArray(iconNumberToRead) = sQuickLaunch
'        sAutoHideDockArray(iconNumberToRead) = sAutoHideDock
'        sSecondAppArray(iconNumberToRead) = sSecondApp
'        sRunElevatedArray(iconNumberToRead) = sRunElevated
'        sRunSecondAppBeforehandArray(iconNumberToRead) = sRunSecondAppBeforehand
'        sAppToTerminateArray(iconNumberToRead) = sAppToTerminate
'        sDisabledArray(iconNumberToRead) = sDisabled
'
'    Else ' alternatively read data from the array cache as it is much faster to read
'
'        sFilename = sFileNameArray(iconNumberToRead)
'        sFileName2 = sFileName2Array(iconNumberToRead)
'        sTitle = sTitleArray(iconNumberToRead)
'        sCommand = sCommandArray(iconNumberToRead)
'        sArguments = sArgumentsArray(iconNumberToRead)
'        sWorkingDirectory = sWorkingDirectoryArray(iconNumberToRead)
'        sShowCmd = sShowCmdArray(iconNumberToRead)
'        sOpenRunning = sOpenRunningArray(iconNumberToRead)
'        sIsSeparator = sIsSeparatorArray(iconNumberToRead)
'        sUseContext = sUseContextArray(iconNumberToRead)
'        sDockletFile = sDockletFileArray(iconNumberToRead)
'        sUseDialog = sUseDialogArray(iconNumberToRead)
'        sUseDialogAfter = sUseDialogAfterArray(iconNumberToRead)
'        sQuickLaunch = sQuickLaunchArray(iconNumberToRead)
'        sAutoHideDock = sAutoHideDockArray(iconNumberToRead)
'        sSecondApp = sSecondAppArray(iconNumberToRead)
'        sRunElevated = sRunElevatedArray(iconNumberToRead)
'        sRunSecondAppBeforehand = sRunSecondAppBeforehandArray(iconNumberToRead)
'        sAppToTerminate = sAppToTerminateArray(iconNumberToRead)
'        sDisabled = sDisabledArray(iconNumberToRead)
'    End If
        
   On Error GoTo 0
   Exit Sub

readIconSettingsIni_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconSettingsIni of Module Module2"
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : readIconRegistryWriteSettings
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   : Read the registry one line at a time and create a temporary settings file
''---------------------------------------------------------------------------------------
''
'Public Sub readIconRegistryWriteSettings(settingsFile As String)
'    Dim useloop As Integer: useloop = 0
'
'    On Error GoTo readIconRegistryWriteSettings_Error
'
'    ' write to the dockSettingsFile letting the dock know who wrote the last update to the settings
'    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", dockSettingsFile
'
'    'If debugFlg = 1 Then debugLog "%" & "readIconRegistryWriteSettings"
'
'    For useloop = 0 To rdIconUpperBound
'         ' get the relevant entries from the registry
'         readRegistryIconValues (useloop)
'         ' write the rocketdock alternative settings.ini
'         Call writeIconSettingsIni(useloop, settingsFile)
'     Next useloop
'
'
'   On Error GoTo 0
'   Exit Sub
'
'readIconRegistryWriteSettings_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconRegistryWriteSettings of Module common3"
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : writeRegistryOnce
'' Author    : beededea
'' Date      : 20/06/2019
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub writeRegistryOnce(ByVal iconNumberToWrite As Integer)
'
'   On Error GoTo writeRegistryOnce_Error
'    'If debugFlg = 1 Then debugLog "%" & "writeRegistryOnce"
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
'
'    'If defaultDock = 1 Then
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialog", sUseDialog)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-UseDialog", sUseDialogAfter) ' .01 DAEB 31/01/2021 common3.bas Added new checkbox to determine if a post initiation dialog should appear
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-QuickLaunch", sQuickLaunch) ' .02 DAEB 20/05/2021 common.bas Added new check box to allow a quick launch of the chosen app
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-AutoHideDock", sAutoHideDock) ' .12 DAEB 20/05/2021 common3.bas Added new check box to allow autohide of the dock after launch of the chosen app
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-SecondApp", sSecondApp)  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-RunSecondAppBeforehand", sRunSecondAppBeforehand)
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-AppToTerminate", sAppToTerminate)
'
'    Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", iconNumberToWrite & "-Disabled", sDisabled)  ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
'
'   On Error GoTo 0
'   Exit Sub
'
'writeRegistryOnce_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistryOnce of Module common3"
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : removeSettingsIni
'' Author    : beededea
'' Date      : 21/09/2019
'' Purpose   : 'effectively removes data from the ini file at the given location by writing nulls to each value
''---------------------------------------------------------------------------------------
''
'Public Sub removeSettingsIni(ByVal iconNumberToWrite As Integer)
'
'   On Error GoTo removeSettingsIni_Error
'   'If debugFlg = 1 Then debugLog "%removeSettingsIni"
'
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-FileName", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-FileName2", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-Title", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-Command", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-Arguments", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-WorkingDirectory", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-ShowCmd", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-OpenRunning", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-IsSeparator", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-UseContext", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-DockletFile", vbNullString, dockSettingsFile
'
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-UseDialog", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-UseDialogAfter", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-QuickLaunch", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-AutoHideDock", vbNullString, dockSettingsFile
'        PutINISetting "Software\SteamyDock\IconSettings\", iconNumberToWrite & "-SecondApp", vbNullString, dockSettingsFile
'
'   On Error GoTo 0
'   Exit Sub
'
'removeSettingsIni_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeSettingsIni of Module common3"
'
'End Sub




' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
'---------------------------------------------------------------------------------------
' Procedure : identifyAppIcons
' Author    : beededea
' Date      : 19/04/2021
' Purpose   : identify an icon to assign to the entry
'---------------------------------------------------------------------------------------
'
Public Function identifyAppIcons(iconCommand As String) As String
    Dim iconFilename As String: iconFilename = ""
    Dim identFileName As String: identFileName = ""
    Dim sDataLine As String: sDataLine = ""
    Dim strDelimiter As String: strDelimiter = ""
    Dim AppName As String: AppName = ""
    Dim appIdent1 As String: appIdent1 = ""
    Dim appIdent2  As String: appIdent2 = ""
    Dim appIcon  As String: appIcon = ""
    Dim appIdent1Bool As Boolean: appIdent1Bool = False
    Dim appIdent2Bool As Boolean: appIdent2Bool = False
    Dim fileH As Long: fileH = 0
    
    On Error GoTo identifyAppIcons_Error

    
    identFileName = sdAppPath & "\appIdent.csv"

    
    strDelimiter = ","
    If fFExists(identFileName) Then
      fileH = FreeFile() 'get the next free file handle
      ' open the identFileName file
      Open identFileName For Input As #fileH
      ' loop through the identFileName file
      Do While Not EOF(fileH)
          appIdent1Bool = False
          appIdent2Bool = False
         
          ' extract the line from the appIdent file
          Input #fileH, AppName, appIdent1, appIdent2, appIcon ' read the four values
          ' set the first two factors to a unlikely value to avoid matching on a ""
          If appIdent1 = vbNullString Then appIdent1 = "XXXXXXXXXXXXXXXXXXXXXX"
          If appIdent2 = vbNullString Then appIdent2 = "XXXXXXXXXXXXXXXXXXXXXX"
          
          ' search for these two factors in the icon command
          If InStr(LCase$(iconCommand), LCase$(appIdent1)) > 0 Then
              appIdent1Bool = True
          End If
          If InStr(LCase$(iconCommand), LCase$(appIdent2)) > 0 Then
              appIdent2Bool = True
          End If
          ' if there is a match then read the icon location
          If appIdent1Bool = True And appIdent2Bool = True Then
              ' set that icon as the iconFileName to use
              
              
                  iconFilename = sdAppPath & "\" & appIcon
 
              
                'iconFileName = App.Path & appIcon
                Exit Do ' now found exit the loop
          End If
      Loop
      Close #fileH
    End If
    
    If Not iconFilename = vbNullString Then
       ' check the icon exists


        If fFExists(iconFilename) Then
            identifyAppIcons = iconFilename
        End If
    End If

   On Error GoTo 0
   Exit Function

identifyAppIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyAppIcons of Form dock"

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetShortcutInfo
' Author    : Jacques Lebrun
' Date      : 19/04/2021
' Purpose   : Resolve a shortcut
'---------------------------------------------------------------------------------------
' Not that you must add a reference to the COM library "Microsoft Shell Controls and Automation."
' CREDIT http://www.vb-helper.com/howto_get_shortcut_info.html
Public Function GetShortcutInfo(Path As String, Shortcut As Link) As Boolean

    Dim FileNo As Integer: FileNo = 0
    Dim LongValue As Long: LongValue = 0
    Dim IntValue As Integer: IntValue = 0
    Dim LinkFlags As Long: LinkFlags = 0
    Dim NextPtr As Long: NextPtr = 0
    Dim Ptr(6) As Long ': Text = 0
    Dim Idx As Integer: Idx = 0
    Dim PtrBasePath As Long: PtrBasePath = 0
    Dim PtrNetworkVolumeInfo As Long: PtrNetworkVolumeInfo = 0
    Dim PtrFilename As Long: PtrFilename = 0
    Dim Str As String: Str = ""
    
    ' Initialise link results
    On Error GoTo GetShortcutInfo_Error

    With Shortcut
        .FileName = vbNullString
        .Description = vbNullString
        .RelPath = vbNullString
        .WorkingDir = vbNullString
        .Arguments = vbNullString
        .CustomIcon = vbNullString
    End With
    For Idx = 0 To 6
        Ptr(Idx) = 0
    Next
    
    ' Open file with .lnk extension
    FileNo = FreeFile
    Str = Path
    If Right$(Str, 4) <> ".lnk" Then Str = Str & ".lnk"
    Open Str For Binary Access Read As FileNo
    
    ' First double-word of link file must be 'L'
    Get FileNo, 1, LongValue
    If LongValue = 76 Then
        ' Skip 16 bytes file GUID and get File Flags
        Get FileNo, 21, LinkFlags
        
        ' Read File Attributes
        Get FileNo, , Shortcut.Attributes
        
        ' Check if ID List section is defined (Ignored)
        NextPtr = 77
        If LinkFlags And 1 Then
            ' Position pointer to next block
            Get FileNo, NextPtr, IntValue
            NextPtr = NextPtr + IntValue + 2
        End If
        
        ' Check if Filename section is defined (no longer mandatory as user created links may not have this defined - see bottom)
        If LinkFlags And 2 Then
            Get FileNo, NextPtr + 16, PtrBasePath
            Get FileNo, , PtrNetworkVolumeInfo
            Get FileNo, , PtrFilename
            
            ' Read base path
            If PtrBasePath Then
                Shortcut.FileName = ReadSingleString(FileNo, NextPtr + PtrBasePath)
            ' Or network path
            ElseIf PtrNetworkVolumeInfo Then
                Shortcut.FileName = ReadSingleString(FileNo, NextPtr + PtrNetworkVolumeInfo + &H14)
            End If
            
            ' Read remaining filename
            If PtrFilename Then
                Str = ReadSingleString(FileNo, NextPtr + PtrFilename)
                If Str <> vbNullString Then
                    If Right$(Shortcut.FileName, 1) <> "\" Then
                        Shortcut.FileName = Shortcut.FileName & "\"
                    End If
                    Shortcut.FileName = Shortcut.FileName & Str
                End If
            End If
            
            ' Position pointer to next block
            Get FileNo, NextPtr, IntValue
            NextPtr = NextPtr + IntValue
        End If
    End If
        
    ' Check if Description section is defined (Optional)
    If LinkFlags And 4 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.Description = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if Relative Path section is defined (Optional)
    If LinkFlags And 8 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.RelPath = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if Working Directory section is defined (Optional)
    If LinkFlags And 16 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.WorkingDir = ReadDoubleString(FileNo, IntValue)
    End If
    
     ' Check if Arguments section is defined (Optional)
    If LinkFlags And 32 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.Arguments = ReadDoubleString(FileNo, IntValue)
    End If
    
    ' Check if CustomIcon section is defined (Optional)
    If LinkFlags And 64 Then
        ' Read string length followed by double-byte string
        Get FileNo, NextPtr, IntValue
        NextPtr = NextPtr + IntValue * 2 + 2
        Shortcut.CustomIcon = ReadDoubleString(FileNo, IntValue)
    End If
        
    Close FileNo
    'GetShortcutInfo = (Shortcut.FileName <> vbNullString) 'this line has been disabled to ensure that even shortcuts where we cannot extract a filename can be utilised, ie. no longer mandatory.

   On Error GoTo 0
   Exit Function

GetShortcutInfo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShortcutInfo of Module mdlMain"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadSingleString
' Author    : Jacques Lebrun
' Date      : 28/05/2021
' Purpose   :  Read a single-byte string from the file
'---------------------------------------------------------------------------------------
'
Private Function ReadSingleString(FileNo As Integer, Offset As Long) As String

    Dim Str As String: Str = ""
    Dim ByteValue As Byte: ByteValue = 0
    
    On Error GoTo ReadSingleString_Error

    Seek FileNo, Offset
    Get FileNo, , ByteValue
    Str = vbNullString
    
    Do While ByteValue <> 0
        Str = Str & ChrW$(ByteValue)
        Get FileNo, , ByteValue
    Loop
    
    ReadSingleString = Str

    On Error GoTo 0
    Exit Function

ReadSingleString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadSingleString of Module mdlMain"

End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadDoubleString
' Author    : Jacques Lebrun
' Date      : 28/05/2021
' Purpose   : Read a double-byte string value preceded by its length
'---------------------------------------------------------------------------------------
'
Private Function ReadDoubleString(FileNo As Integer, StrLen As Integer) As String

    Dim IntValue As Integer: IntValue = 0
    Dim Str As String: Str = ""
    
    On Error GoTo ReadDoubleString_Error

    Str = vbNullString
    Do While StrLen > 0
        Get FileNo, , IntValue
        Str = Str & ChrW$(IntValue)
        StrLen = StrLen - 1
    Loop
    
    ReadDoubleString = Str

    On Error GoTo 0
    Exit Function

ReadDoubleString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadDoubleString of Module mdlMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetShellShortcutInfo
' Author    : beededea
' Date      : 16/04/2021
' Purpose   : This is our old standard method of getting the path from a LNK file using shell scripting that can cause a/v tools to FP.
'---------------------------------------------------------------------------------------
' CREDIT http://www.vb-helper.com/howto_get_shortcut_info.html
Public Function GetShellShortcutInfo(ByVal full_name As String, _
    ByRef Name As String, ByRef Path As String, ByVal descr _
    As String, ByRef working_dir As String, ByRef args As _
    String) As String

    Dim shl As Shell32.Shell
    Dim shortcut_path As String: shortcut_path = ""
    Dim shortcut_name As String: shortcut_name = ""

    Dim shortcut_folder As Shell32.folder
    Dim folder_item As Shell32.FolderItem
    Dim lnk As Shell32.ShellLinkObject

    'On Error GoTo GetShellShortcutInfo_Error

    ' Make a Shell object.
    
    Set shl = New Shell32.Shell

    ' Get the shortcut's folder and name.
    shortcut_path = Left$(full_name, InStrRev(full_name, _
        "\"))
    shortcut_name = Mid$(full_name, InStrRev(full_name, _
        "\") + 1)
    If Not Right$(shortcut_name, 4) = ".lnk" Then _
        shortcut_name = shortcut_name & ".lnk"

    ' Get the shortcut's folder.
    Set shortcut_folder = shl.NameSpace(shortcut_path)

    ' Get the shortcut's file.
    Set folder_item = _
        shortcut_folder.Items.Item(shortcut_name)
    If folder_item Is Nothing Then
        GetShellShortcutInfo = "Cannot find shortcut file '" & _
            full_name & "'"
    ElseIf Not folder_item.IsLink Then
        ' It's not a link.
        GetShellShortcutInfo = "File '" & full_name & "' isn't a " & _
            "shortcut."
    Else
        ' Display the shortcut's information.
        Set lnk = folder_item.GetLink
        Name = folder_item.Name
        descr = lnk.Description
        Path = lnk.Path
        working_dir = lnk.WorkingDirectory
        args = lnk.Arguments
        GetShellShortcutInfo = vbNullString
    End If
    

   On Error GoTo 0
   Exit Function

GetShellShortcutInfo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShellShortcutInfo of Form dock"
End Function


'---------------------------------------------------------------------------------------
' Procedure : zeroAllIconCharacteristics
' Author    : beededea
' Date      : 27/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub zeroAllIconCharacteristics()

   On Error GoTo zeroAllIconCharacteristics_Error

    sFilename = vbNullString
    sFileName2 = vbNullString
    sTitle = vbNullString
    sCommand = vbNullString
    sArguments = vbNullString
    sWorkingDirectory = vbNullString
    sOpenRunning = "0"
    sIsSeparator = "0"
    sUseContext = "0"
    sDockletFile = "0"
    sUseDialog = "0"
    sUseDialogAfter = "0"
    sQuickLaunch = "0"
    sDisabled = "0"
    sAutoHideDock = "0"
    sSecondApp = vbNullString
    sRunSecondAppBeforehand = "0"
    sAppToTerminate = vbNullString
    sRunElevated = "0"

   On Error GoTo 0
   Exit Sub

zeroAllIconCharacteristics_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure zeroAllIconCharacteristics of Module common3"
            
End Sub

