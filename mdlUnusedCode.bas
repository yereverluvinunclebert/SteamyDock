Attribute VB_Name = "mdlUnusedCode"

'' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
''---------------------------------------------------------------------------------------
'' Procedure : removeImageFromDictionary
'' Author    : beededea
'' Date      : 18/06/2020
'' Purpose   : only used when a single icon is to be added to the dock
''             this routine is a workaround to the memory leakage problem in resizeAndLoadImgToDict
''             where if run twice the RAM usage doubled as the vars are not clearing their contents when
''             the routine ends
''
'' When an icon is added it should no longer call the routine to recreate the arrays and collections
'' instead it calls this routine, previously there was one dictionary.
''
'' there is now a separate dictionary for the smaller icons
'' there is another dictionary for the larger icons
'' there is a third temporary dictionary that is used as temporary storage whilst resizing the above
'' when a new icon is added to the dock
''
'' we use the existing resizeAndLoadImgToDict to read the larger icon format
'' the icons to the left are written to the 3rd temporary dictionary with existing keys, the new icon is then written using the current location as part of the key
'' the icons to the right are then read from the old dictionary and then written to the new temporary dictionary with updated keys
'' the larger image dictionary is cleared down readied for population
'' the temporary dictionary is used to repopulate the larger image dictionary, a clone
'' the temporary dictionary is cleared down, ready for re-use
'
'' then we do the same for the smaller icon format images
''---------------------------------------------------------------------------------------
''
'Public Sub oldRemoveImageFromDictionary()
'
'    Dim useloop As Integer
'    Dim thiskey As String
'    Dim newKey As String
'
'    On Error GoTo removeImageFromDictionary_Error
'
''    If debugflg = 1 Then debugLog "%" & "removeImageFromDictionary"
'
'    'resize all arrays used for storing icon information
'    ReDim fileNameArray(rdIconMaximum) As String ' the file location of the original icons
'    ReDim namesListArray(rdIconMaximum) As String ' the name assigned to each icon
'    ReDim sCommandArray(rdIconMaximum) As String ' the command assigned to each icon
'    ReDim targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
'    ReDim processCheckArray(rdIconMaximum) As String ' the process name assigned to each icon
'    ReDim initiatedProcessArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
'                                                         ' but I feel that it does not really matter so I am going to not bother at the moment, this is something that could be done later!
'
'    ' assuming that the details have already been written to the configuration file
'    ' extract filenames from Rocketdock registry, settings.ini or user data area
'    ' we reload the arrays that store pertinent icon information
'    For useloop = 0 To rdIconMaximum
'        'readIconData (useloop)
'        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
'        ' read the two main icon variables into arrays, one for each
'        fileNameArray(useloop) = sFilename
'        namesListArray(useloop) = sTitle
'        sCommandArray(useloop) = sCommand
'        targetExistsArray(useloop) = 0
'
'        ' check to see if each process is running and store the result away
'        'processCheckArray(useloop) = isProcessInTaskList(sCommand)
'        processCheckArray(useloop) = IsRunning(sCommand, vbNull)
'
'    Next useloop
'
'    'redimension the array that is used to store all of the icon current positions in pixels
'    ' preserves the data in the existing array when changing the size of only the last dimension.
'    ReDim Preserve iconStoreLeftPixels(rdIconMaximum + 1) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'    ReDim Preserve iconStoreRightPixels(rdIconMaximum + 1) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    ReDim Preserve iconStoreTopPixels(rdIconMaximum + 1) ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'    ReDim Preserve iconStoreBottomPixels(rdIconMaximum + 1) ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'
'    iconArrayUpperBound = rdIconMaximum '<*
'
'    ' populate the array element containing the final icon position
'    'iconPosLeftTwips(rdIconMaximum) = iconPosLeftTwips(rdIconMaximum - 1) + (iconWidthPxls * screenTwipsPerPixelX) '< this may need revisiting if you add left and right positions
'
'    ' re-order the large icons in the collLargeIcons dictionary collection
'    Call decrementCollection(collLargeIcons, iconSizeLargePxls)
'
'    ' re-order the small icons in the collSmallIcons dictionary collection
'    Call decrementCollection(collSmallIcons, iconSizeSmallPxls)
'
'    Call loadAdditionalImagestoDictionary ' the additional images need to be re-added back to the dictionary
'
'   On Error GoTo 0
'   Exit Sub
'
'removeImageFromDictionary_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure removeImageFromDictionary of module mdlMain.bas"
'
'End Sub



' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : decrementCollection
' Author    : beededea
' Date      : 18/06/2020
' Purpose   : Removes icon from the appropriate dictionary big or small
'---------------------------------------------------------------------------------------
'
Private Sub decrementCollection(ByRef thisCollection As Object, ByVal thisByteSize As Byte)
    Dim useloop As Integer
    Dim thiskey As String
    Dim newKey As String
    Dim partialStringKey As String: partialStringKey = ""
    
    On Error GoTo decrementCollection_Error

    ' .60 DAEB 29/04/2021 frmMain.frm Improved the speed of the deletion of icons from the dictionary collections
    ' the icons to the left of the current icon are not read nor touched.
    ' we delete the current icon from the collection
    thiskey = selectedIconIndex & "ResizedImg" & LTrim$(Str$(thisByteSize))
    thisCollection.Remove thiskey
        
    ' the icons to the right are then read from the old dictionary and then written one key down
    For useloop = selectedIconIndex + 1 To rdIconMaximum + 1 ' change this at your peril
        newKey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
        thiskey = useloop - 1 & "ResizedImg" & LTrim$(Str$(thisByteSize))
        thisCollection(thiskey) = thisCollection(newKey)
    Next useloop
    
    ' OLD METHOD (SLOW)
    ' the icons to the left are written to the 3rd temporary dictionary with existing keys, the new icon is then written with the current location as part of the key
    
    ' A.
'    For useloop = 0 To selectedIconIndex - 1
'        thiskey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        newKey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        collTemporaryIcons(newKey) = thisCollection(thiskey)
'    Next useloop
 
    ' B.
    ' the icons to the right including the current are then read from the old dictionary and then written to the new temporary dictionary with updated incremented keys
'    For useloop = selectedIconIndex + 1 To rdIconMaximum + 1 ' change this at your peril
'        thiskey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        newKey = useloop - 1 & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        collTemporaryIcons(newKey) = thisCollection(thiskey)
'    Next useloop
    

    ' the original image dictionary is cleared down readied for repopulation
    'thisCollection.RemoveAll

    ' the temporary dictionary is used to repopulate the larger image dictionary, a clone of all elements
'    For useloop = 0 To rdIconMaximum
'        thiskey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        thisCollection(thiskey) = collTemporaryIcons(thiskey)
'    Next useloop

    ' the temporary dictionary is cleared down, ready for re-use
    'collTemporaryIcons.RemoveAll
    ' Set collTemporaryIcons = New Scripting.Dictionary ' to do the SET NEW here, support for MS scripting must be enabled in project - references
    ' emptying a dictionary or disposing of the contents does not release the memory used by the construct
    ' creating a new example removes the old version from memory and creates an unpopulated dictionary

   On Error GoTo 0
   Exit Sub

decrementCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure decrementCollection of module mdlMain.bas"

End Sub

' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : incrementCollection
' Author    : beededea
' Date      : 18/06/2020
' Purpose   : Writes a new icon to the named dictionary big or small together with all the previous icons
'             We are simply moving elements up and down a dictionary
'---------------------------------------------------------------------------------------
'Private Sub incrementCollection(ByRef thisCollection As Object, ByVal thisByteSize As Byte, ByVal newFileName As String, ByVal newName As String)
'    Dim useloop As Integer
'    Dim thiskey As String
'    Dim newKey As String
'    Dim partialStringKey As String: partialStringKey = ""
'
'    On Error GoTo incrementCollection_Error
'
''    If debugflg = 1 Then debugLog "%" & "incrementCollection "
'
'    ' .62 DAEB 29/04/2021 frmMain.frm Improved the speed of the addition of icons to the dictionary collections
'    ' the icons to the left of the current icon are not read nor touched
'    ' reads from the last icon to the current one and for each it writes it one step up
'    For useloop = rdIconMaximum To selectedIconIndex Step -1
'        thiskey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        newKey = useloop + 1 & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        If thisCollection.Exists(thiskey) Then
'            thisCollection(newKey) = thisCollection(thiskey)
'        End If
'    Next useloop
'
'    'now we add the new icon to the current position in the dictionary
'    partialStringKey = LTrim$(Str$(selectedIconIndex))
'    If FExists(newFileName) Then
'        ' we use the existing resizeAndLoadImgToDict to read the icon format
'         resizeAndLoadImgToDict thisCollection, partialStringKey, newFileName, sDisabled, (0), (0), (thisByteSize), (thisByteSize), , imageOpacity
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'incrementCollection_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure incrementCollection of module mdlMain.bas"
'
'End Sub
