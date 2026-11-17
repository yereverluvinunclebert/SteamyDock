Attribute VB_Name = "modDisplayDock"
'---------------------------------------------------------------------------------------
' Module    : modDisplay
' Author    : beededea
' Date      : 05/11/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public bumpFactor As Single
Public IconIndex As Single
Public dynamicSizeModifierPxls As Double

'---------------------------------------------------------------------------------------
' Procedure : updateDisplayFromIconCollection
' Author    : beededea
' Date      : 05/11/2026
' Purpose   : This is the new icon framework that will be called frequently by an animating timer
'             putting icon images on the main display using a graphics method, currently GDI+.
'---------------------------------------------------------------------------------------
'
Public Sub updateDisplayFromIconCollection()

    Dim Icon As cwMainIcon
    Dim leftmostResizedIcon As Integer
    Dim rightmostResizedIcon As Integer

    ' Check that the collection exists.
    On Error GoTo updateDisplayFromIconCollection_Error

    If sDockIcons Is Nothing Then Exit Sub

    ' Check that a renderer has been assigned.
    If gDockRenderer Is Nothing Then Exit Sub

    ' Tell the renderer that a new frame is starting.
    gDockRenderer.BeginFrame

    Call determineDynamicIconRangeToAnimate(leftmostResizedIcon, rightmostResizedIcon)
    
    ' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
    bumpFactor = 1.2 ' this determines the 'bumpiness' of the transition of switching from one icon to the next
    If usedMenuFlag = False Then ' only recalculate dynamicSizeModifierPxls for the bump animation when the menu has not recently been used
         ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
        
        ' this is the actual line that does the main animation
        dynamicSizeModifierPxls = ((apiMouse.X) - iconStoreLeftPixels(IconIndex)) / (bumpFactor)
    
    Else
        usedMenuFlag = False ' the menu causes the mouse to move far away from the icon centre and so icon sizing was massive
    End If

    ' Pass every icon to the renderer.
    For Each Icon In sDockIcons
    
'        ' size all small icons to the left of the main icons, each sized as per small mode, do it just ONCE, all the other icons will take that same size without recalculating
'        If useloop = iconArrayLowerBound Then Call sizeEachSmallIconToLeft(useloop, leftmostResizedIcon, showsmall)
'
'        ' if the the main icon is icon zero
'        ' the main fullsize icon
'        Call sizeFullSizeIcon(useloop, showsmall)
'
'        ' the group of icons to the left of the main icon, resized dynamically
'        Call sizeEachResizedIconToLeft(useloop, leftmostResizedIcon, showsmall)
'
'        ' the group of icons to the right of the main icon, resized dynamically
'        Call sizeEachResizedIconToRight(useloop, rightmostResizedIcon, showsmall)
'
'        ' small icons to the right shown in small mode, do it just once, all the other icons will take that same size without recalculating
'        If useloop = rightmostResizedIcon + 1 Then Call sizeEachSmallIconToRight(useloop, rightmostResizedIcon, showsmall)
'
'        ' display the icon in the dock
'        If showsmall = True Then ' display the small size icon or the red X if icon missing
'            Call updateDisplayUsingSmallIcon(useloop)
'        Else
'            Call updateDisplayUsingLargeIcon(useloop) ' display the larger size icon or the
'        End If
'
'        'now draw the icon text above the selected icon
'        Call drawTextAboveIcon(useloop, textWidth)

        gDockRenderer.DrawIcon Icon

    Next Icon

    ' Tell the renderer that the frame is complete.
    gDockRenderer.EndFrame

    On Error GoTo 0
    Exit Sub

updateDisplayFromIconCollection_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateDisplayFromIconCollection of Module modDisplay"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : determineDynamicIconRangeToAnimate
' Author    : beededea
' Date      : 19/07/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub determineDynamicIconRangeToAnimate(ByRef leftmostResizedIcon As Integer, ByRef rightmostResizedIcon As Integer)
    
   On Error GoTo determineDynamicIconRangeToAnimate_Error

    rDZoomWidth = 2 'override until the animation takes this into account
    If CBool(rDZoomWidth And 1) = False Then
        rDZoomWidth = rDZoomWidth + 1  ' must be 3,5,7,9 so convert to an odd number
    End If
     
    ' what is the group size? extract the index of the group and calculate the leftmost member
    leftmostResizedIcon = IconIndex - (rDZoomWidth - 1) / 2 '
    rightmostResizedIcon = IconIndex + (rDZoomWidth - 1) / 2

   On Error GoTo 0
   Exit Sub

determineDynamicIconRangeToAnimate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineDynamicIconRangeToAnimate of Form dock"
End Sub
