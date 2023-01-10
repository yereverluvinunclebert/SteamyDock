VERSION 5.00
Begin VB.Form dock 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   5520
   Icon            =   "frmMain.frx":0000
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   Begin VB.Timer delayRunTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   240
      Tag             =   "This is the timer that causes any secondary command to run three seconds after the main"
      Top             =   4395
   End
   Begin VB.Timer targetExistsTimer 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   2925
      Tag             =   "this routine is used to identify if the main target is valid"
      Top             =   1680
   End
   Begin VB.Timer forceHideRevealTimer 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2835
      Top             =   3960
   End
   Begin VB.Timer transitTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   255
      Top             =   2940
   End
   Begin VB.Timer bounceDownTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   240
      Tag             =   "controls the bounceDownward when the icon is clicked"
      Top             =   2385
   End
   Begin VB.Timer hourGlassTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2835
      Tag             =   "load a small rotating hourglass image into the collection, used to signify running actions"
      Top             =   4470
   End
   Begin VB.Timer sleepTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2895
      Tag             =   "stores and compares the last time to see if the PC has slept"
      Top             =   1155
   End
   Begin VB.Timer positionZTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   255
      Tag             =   "Places the dock back in the defined z-order"
      Top             =   1110
   End
   Begin VB.Timer autoSlideInTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2835
      Tag             =   "slide the dock in the Y axis"
      Top             =   6030
   End
   Begin VB.Timer nMinuteExposeTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2835
      Tag             =   "causes the dock to re-appear in its default state after N mins"
      Top             =   4995
   End
   Begin VB.Timer autoFadeInTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Tag             =   "this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha"
      Top             =   6030
   End
   Begin VB.Timer autoSlideOutTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2835
      Tag             =   "slide the dock in the Y axis"
      Top             =   5505
   End
   Begin VB.Timer initiatedProcessTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2895
      Tag             =   "Provides regular checking of only processes initiated by the dock"
      Top             =   660
   End
   Begin VB.Timer autoHideChecker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   255
      Tag             =   "checks to see if the dock needs to be hidden, if so, initiates one of the hider timers"
      Top             =   4965
   End
   Begin VB.Timer autoFadeOutTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Tag             =   "this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha"
      Top             =   5490
   End
   Begin VB.Timer processTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2895
      Tag             =   "this routine is used to identify an item in the dock as currently running even if not triggered by the dock"
      Top             =   150
   End
   Begin VB.Timer runTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Tag             =   "calls the subroutine that runs the actual command"
      Top             =   3855
   End
   Begin VB.Timer bounceUpTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   240
      Tag             =   "controls the bounceUpward when the icon is clicked"
      Top             =   1890
   End
   Begin VB.Timer responseTimer 
      Interval        =   200
      Left            =   255
      Tag             =   "Determines whetherto turn on the animate timer"
      Top             =   585
   End
   Begin VB.Timer animateTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   270
      Tag             =   "this is the X millisecond timer that does the animation for the dock icons"
      Top             =   105
   End
   Begin VB.Label Label6 
      Caption         =   "delayRunTimer"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   22
      ToolTipText     =   "This is the timer that causes any secondary command to run three seconds after the main"
      Top             =   4470
      Width           =   1425
   End
   Begin VB.Label Label18 
      Caption         =   "targetExistsTimer"
      Height          =   255
      Left            =   3450
      TabIndex        =   21
      Top             =   1725
      Width           =   1665
   End
   Begin VB.Label Label 
      Caption         =   "forceHideRevealTimer"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   20
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "transitTimer"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3015
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "bounceDownTimer"
      Height          =   255
      Left            =   945
      TabIndex        =   18
      Top             =   2460
      Width           =   1485
   End
   Begin VB.Label Label16 
      Caption         =   "hourglassTimer"
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      ToolTipText     =   "causes the dock to re-appear in its default state after 10 mins"
      Top             =   4590
      Width           =   1785
   End
   Begin VB.Label Label15 
      Caption         =   "sleepTimer"
      Height          =   255
      Left            =   3435
      TabIndex        =   16
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label Label14 
      Caption         =   "positionZTimer"
      Height          =   255
      Left            =   960
      TabIndex        =   15
      ToolTipText     =   "Placing the dock back in the defined z-order"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "autoSlideInTimer"
      Height          =   255
      Left            =   3375
      TabIndex        =   14
      ToolTipText     =   "slides the dock in the Y axis"
      Top             =   6150
      Width           =   1410
   End
   Begin VB.Label Label12 
      Caption         =   "Note: there are other timers on the splashform"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   375
      TabIndex        =   13
      Top             =   8025
      Width           =   4380
   End
   Begin VB.Label Label9 
      Caption         =   "nMinuteExposeTimer"
      Height          =   255
      Left            =   3375
      TabIndex        =   12
      ToolTipText     =   "causes the dock to re-appear in its default state after 10 mins"
      Top             =   5085
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "autoFadeInTimer"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha"
      Top             =   6135
      Width           =   1425
   End
   Begin VB.Label lblDockInfo2 
      Caption         =   $"frmMain.frx":058A
      Height          =   990
      Left            =   405
      TabIndex        =   10
      Top             =   6825
      Width           =   4380
   End
   Begin VB.Label lblDockInfo 
      Caption         =   $"frmMain.frx":068F
      Height          =   1380
      Left            =   2715
      TabIndex        =   9
      Top             =   2295
      Width           =   2370
   End
   Begin VB.Label Label11 
      Caption         =   "autoSlideOutTimer"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      ToolTipText     =   "slides the dock in the Y axis"
      Top             =   5610
      Width           =   1410
   End
   Begin VB.Label Label10 
      Caption         =   "initiatedProcessTimer"
      Height          =   255
      Left            =   3435
      TabIndex        =   7
      Top             =   735
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "autoHideChecker"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   5070
      Width           =   1410
   End
   Begin VB.Label Label7 
      Caption         =   "autoFadeOutTimer"
      Height          =   255
      Left            =   945
      TabIndex        =   5
      Top             =   5610
      Width           =   1425
   End
   Begin VB.Label Label6 
      Caption         =   "runTimer"
      Height          =   255
      Index           =   0
      Left            =   975
      TabIndex        =   4
      ToolTipText     =   "This is the timer that causes any specified command to run"
      Top             =   3945
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "bounceUpTimer"
      Height          =   255
      Left            =   945
      TabIndex        =   3
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "processTimer"
      Height          =   255
      Left            =   3435
      TabIndex        =   2
      Top             =   225
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "responseTimer"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "animateTimer"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "dock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =====================================
' Functional changes since last release
' =====================================
'
' Auto generation of dock items in iconSettings is progressing
' Automatic finding of the correct icon to use using an application compatibility list
' After a system sleep the raised dock is lowered. No, not quite yet
' Runs well on a system that has never had Rocketdock installed upon it
' If the docksettings is running it is brought to the fore rather than closed and re-opened as before.
' addition and deletion of icons is no longer a very slow operation
' The rotating hourglass timer displayed during a drag and drop operation to the dock
' Now possible to drag an icon from the dock to delete it permanently
' The dock icon bounce effect is now using an easeIN function
' A smaller icon image is displayed during a drag and delete operation from the dock
' Dragging and dropping from one part of the dock to another
' Added Quick Launch functionality to run an app more quickly half way through the bounce animation
' Added autohide dock after running an app
' Added automatic running of secondary app.
' Implement the old bounce as a separate bounce type called miserable
' Implemented an icon cache for older, slower machines
' Fixed window height issues on Win 10/11
' Added menu option to clone current item
' A misconfigured or non-existent target shows a red X over the icon
' Allows you to set a secondary app to run.
' Confirms when deleting multiple instances of a program

'========================================================================================================
' Change History
' ==============
' Changes: dates in UK DD/MM/YYYY format

' 22/10/2020 .01 DAEB frmMain.frm responsetimer fix the incorrect check of the timer state to determine the dock upper limit when entering and triggering the main animation
' 23/10/2020 .02 DAEB frmMain.frm move the dock position behind the icons 8 pixels to the left to position the icons correctly on the dock
' 26/10/2020 .03 DAEB frmMain.frm removed declarations required by IsRunning since the move of this function to common.bas
' 27/10/2020 .04 DAEB frmMain.frm alternative animations to zoom: Bubble enabled as options
'            .05 DAEB frmMain.frm null
' 17/11/2020 .06 DAEB frmMain.frm Fixed the sequentialBubbleAnimation
' 24/01/2021 .07 DAEB frmMain.frm modified to handle the new timer name
' 24/01/2021 .08 DAEB frmMain.frm removed the fade in functions from the fade out function
' 24/01/2021 .09 DAEB frmMain.frm created a separate fade in timer and function
' 25/01/2021 .10 DAEB frmMain.frm Added new parameter autoFadeInTimerCount for the new fade in timer
' 25/01/2021 .11 DAEB frmMain.frm changed the setting of the dock top to a better place
' 25/01/2021 .12 DAEB frmMain.frm Change to sdAppPath
' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
' .14 DAEB frmMain.frm 27/01/2021 Add new subroutine to make the dock transparent and shutdown timers
' .15 DAEB frmMain.frm 27/01/2021 Add new subroutine to show the dock after it has been manually hidden by the user
' .16 DAEB frmMain.frm 27/01/2021 Added the user set parameter sDContinuousHide
' .17 DAEB frmMain.frm 27/01/2021 Moved disabling admin to a separate routine
' .18 DAEB frmMain.frm 31/01/2021 reinstated checks of fade out and slide timers to set a more frequent respnse timer to improve animation
' .19 DAEB frmMain.frm 02/02/2021 added sArguments field to the confirmation dialog
' .20 DAEB frmMain.frm 02/02/2021 added variable initialisation after declaration
' .21 DAEB frmMain.frm 07/02/2021 slight improvement to the the confirmation dialog
' .22 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
' .23 DAEB frmMain.frm 08/02/2021 Changed from an array to a single var
' .24 DAEB frmMain.frm 09/02/2021 handling any potential divide by zero
' .25 DAEB frmMain.frm 10/02/2021 added API and vars to test to see if a window is zoomed
' .26 DAEB frmMain.frm 10/02/2021 added test to check window state and alter it accordingly
' .27 DAEB frmMain.frm 11/02/2021 now operates like the standard Windows dock on a click, minimising then restoring
' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions
' .29 DAEB frmMain.frm 20/02/2021 Added new mdlSysTray module containing the code required to analyse the icons in the systray
' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
' .31 DAEB 03/03/2021 frmMain.frm Check return value from any GDI++ function
' .32 DAEB 03/03/2021 frmMain.frm Placing the dock back in the defined z-order
' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
' .34 DAEB 08/02/2021 frmMain.frm - commented out the extra unwanted ShowWindow(hwnd, SW_RESTORE)
' .35 DAEB 03/03/2021 frmMain.frm check whether the prefix command required to access a Windows class ID is present
' .36 DAEB 03/03/2021 frmMain.frm check whether the prefix is present that indicates a Windows class ID is present
' .37 DAEB 03/03/2021 frmMain.frm removed the individual references to a Windows class ID
' .38 DAEB 18/03/2021 frmMain.frm utilised SetActiveWindow to give window focus without bringing it to fore
' .39 DAEB 18/03/2021 frmMain.frm utilised BringWindowToTop instead of SetWindowPos & HWND_TOP as that was used by a C program that worked perfectly.
' .40 DAEB 18/03/2021 frmMain.frm Added SWP_NOOWNERZORDER as an additional flag as that was used by a C program that worked perfectly, fixing the z-order position problems
' .41 DAEB 18/03/2021 frmMain.frm utilised ShowWindowAsync instead of ShowWindow as the C program utilised it and it seemed to make sense to do so too
' .42 DAEB 03/03/2021 frmMain.frm To support new receive focus menu option
' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
' .44 DAEB 01/04/2021 frmMain.frm put the control panel reference back
' .45 DAEB 01/04/2021 frmMain.frm Changed the logic to remove the code around a folder path existing...
' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
' .47 DAEB 01/04/2021 frmMain.frm autoSlideMode is now undefined at startup - this allowed the top position to operate as expected
' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
' .49 DAEB 01/04/2021 frmMain.frm added the vertical adjustment for sliding in and out STARTS
' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioning according to the slideIn/Out state, not done here
' .51 DAEB 08/04/2021 frmMain.frm calls mnuIconSettings_Click to start up the icon settings tools and display the properties of the new icon.
' .52 DAEB 09/04/2021 frmMain.frm add code to increase the timer to 120 minutes
' .53 DAEB 11/04/2021 frmMain.frm changed all occurrences of sCommand to thisCommand to attain more compatibility with rdIconConfigFrm menuRun_click
' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
' .55 DAEB 19/04/2021 frmMain.frm Added call to the older function to identify an icon using the shell object
' .56 DAEB 19/04/2021 frmMain.frm Added a faded red background to the current image when the drag and drop is in operation.
' .57 DAEB 19/04/2021 frmMain.frm modifedAmountToSlide renamed to xAxisModifier for clarity's sake
' .58 DAEB 21/04/2021 frmMain.frm added timer and vars to check to see if the system has just emerged from sleep
' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
' .60 DAEB 29/04/2021 frmMain.frm Improved the speed of the deletion of icons from the dictionary collections
' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
' .62 DAEB 29/04/2021 frmMain.frm Improved the speed of the addition of icons to the dictionary collections
' .63 DAEB 30/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
' .64 DAEB 30/04/2021 frmMain.frm Deleted the temporary collection, now unused.
' .65 DAEB 30/04/2021 frmMain.frm Added mouseDown event to capture the time of first press and code to simulate a drag and drop of an icon from the dock
' .66 DAEB 01/05/2021 frmMain.frm huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
' .67 DAEB 01/05/2021 frmMain.frm Added creation of Windows in the states as provided by sShowCmd value in RD.
' .68 DAEB 05/05/2021 frmMain.frm Cause the docksettings utility to reopen if it has already been initiated.
' .69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position.
' .70 DAEB 06/05/2021 frmMain.frm Removed all references to Clng() in all the occurrences of updateDisplayFromDictionary to speed up animation, no references in the code will be found
' .71 DAEB 06/05/2021 frmMain.frm Changed bounceIndex to selectedIconIndex throughput the code, no references in the code will be found
' .72 DAEB 06/05/2021 frmMain.frm Created two timers that controls the bouncing when the icon is clicked, replacing the old timers
' .73 DAEB 11/05/2021 frmMain.frm sngBottom renamed to screenBottomPxls
' .74 DAEB 12/05/2021 frmMain.frm Displays a smaller size icon at the cursor position when a drag from the dock is underway.
' .75 DAEB 12/05/2021 frmMain.frm Changed Form_MouseMove to act as the correct event to a drag and drop operating from the dock
' .76 DAEB 20/05/2021 frmMain.frm Moved from the runtimer as some of the data is required before the run begins
' .77 DAEB 20/05/2021 frmMain.frm Added new check box to allow a quick launch of the chosen app
' .78 DAEB 21/05/2021 frmMain.frm Added new field for second program to be run
' .79 DAEB 21/05/2021 frmMain.frm Disable any active bounce
' .80 DAEB 28/05/2021 frmMain.frm Keep the animateTimer and therefore the bounceTimers operating in order to run the chosen app.
' .81 DAEB 28/05/2021 frmMain.frm Refresh the running process with a cog when the process is running, this had been removed earlier
' .82 DAEB 12/07/2021 frmMain.frm Add the BounceZone as a configurable variable.
' .83 DAEB 14/07/2021 frmMain.frm Modified the BounceZone and bouncetimers to run 50% slower.
' .84 DAEB 20/07/2021 frmMain.frm Added prevention of the dock returning until the hiding application is no longer running.
' .85 DAEB 16/04/2022 frmMain.frm Added new timer to allow auto-reveal of the dock once the chosen app has closed within 1.5 secs
' .86 DAEB 08/12/2022 frmMain.frm Added new timer to inspect each target command in turn and highlight if missing.
' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
' .89 DAEB 08/12/2022 frmMain.frm Fixed duplicate run of .msc files.
' .90 DAEB 08/12/2022 frmMain.frm Added routine to check for an array that has already been initialised
' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
'   Add a sound ting option when initiating a dock click
'   The menu separators should call the utilities immediately below
'   remove all references to rocketdock
'   remove all references to the rdSettings file should no longer refer to these using only docksettings.ini
' In all functions/subroutines in the dock form module - all vars initialised - done
' In all functions/subroutines in the dock form module - all parameters to be determined as bval or byref - done
' deleting multiple instances of a program, if multiple instances exist then it should pop up a modal box that
'  requests confirmation of each process to kill.
' delay on triggering the second app - implemented
' add comment - selected index is 999 - indicating that we do not want to store the second app as a cogged item in this apps stored binary location...
' common module - all vars initialised - done
' common module - all parameters to be determined as bval or byref - done
' override on opening the running application for individual applications in runCommand.
' validation of tool components - missing folders and essential files
' common2 module - all vars to be initialised
' common3 module - all vars to be initialised
' common2 module - all parameters to be determined as bval or byref
' sequentialBubbleAnimation - rewritten to simplify and create shared code
' drawDockByCursorEntryPosition - rewritten to simplify and use shared code from the main animation routine
' drawSmallStaticIcons - rewritten to use shared code from the main animation routine
' if the icon is dragged within the dock or back to the dock then show the busytimer and the red background on the target icon.
' Quick launch of an app should mean instant. Not an incomplete partial bounce - DONE. Instant launch and subsequent bounce would require a state monitor - so NO.
' Fixed bug causing ghosting, after several minutes
' Added GDI ready and tidying within each animation to ensure no dropping of handles.
' Added logic to only call the three current animation types when the bubble animation is selected
' Prevented the general hang when SD was closed using the Quit command (unloading all forms in a loop).
' common3 module - byval/byref

' General Status
' ==============

'     1. WIP making all the recent functions work at the top of the screen as well as the bottom, specifically, background theme, cog, slideout/in and
'     2. WIP the initiation point of the dock when the cursor enters
'     3. converted the slide out routine to two separate routines WIP
'     4. find the bug in the logic that causes the slide out to cycle up and down WIP
'     5. creating new icons WIP <-----
'     6. Creation of a default DockSettings.ini for a new user of the dock where the application has never been run before.
'     8. Add known identifers to the known apps list WIP <----- list growing
'     9. The rotating hourglass timer could be added to the deletion and addition of an icon.
'    10. Drag and drop needs prettying though.
'    12.
'    13.


' Main Tasks:
' ============
'
' when icon clicked and bouncing the dock should not animate when moving the cursor left/right - option to lock? - WIP when the animation is
'   re-jigged and rewritten.

' when dragging from the dock the main icons should not animate when dragging - option to lock? - WIP when the animation is
'   re-jigged and rewritten.
'   use smallicon sub?

' dock entering and leaving RIGHT/LEFT - see below.

' When we set the opacity of the dock to 0 for hiding purposes, all well and good but we ought to do a
'  disable all click interactions and stuff could be going on, we just can't see it...
'  when the fade timers are running, the clicking is disabled

' the iconsettings tool should have some separate code that allows migration from RD to SD but the rest should use SdSettings



' Add the project to Github

' finish the icons
'   reboot icon boot and candlestick
    ' droptypes to deal with by having an associated document

    ' installation packages

    '.xpi done
    '.xar
    '.bz2
'    .bak
'    .bck
'    .pup
'    .bkp

'    .7z
'    .zl
'    .s7z
'    .sfx
'    .arc
'    .ace
'    .ufs

'    .xz
'    .gz
'    .lz

' parcel with a zip in it
'    .bzip2
'    .gzip
'    .zipx

'    .lzx
'    .lzm
'    .mint
'
' look at the custom icon tool and see which you need to recreate WIP
'
' modify the zipfile icons to the correct type above
'
' create a zipped icon



' change to shellExecuteWithDialog to allow apps to run unelevated: CREDIT - fafalone
'
' reorganise windows menu option? send all windows to front/back

'   A Disable Icon option
'
'   Messagebox msgBoxA module - ship the code to FCW to replace the native msgboxes.
'
'   picRdMap_OLEDragDrop to be updated to match the RDiconSettings code improvements.
'
'   when bouncing should take into account the time in ms and not just the speed of the cpu to accomplish the bounce.
'
'   dock icon bounce new animation height, tweak and the timer interval too

' Cogs
' =====
' Adding a cog above a folder window for explorer.exe
' when the program is determining whether a program is already loaded (to show the cog)
' it could test to see if explorer is running and whether the currently open folder matches
' the one in the dock's target folder. If so, then it shows the cog there also.
' When an icon is clicked,  if it is explorer then it tests the open folder's current directory
' and if it matches then it opens the existing folder instance (which is what it does now)
' if the two do not match then the option should exist on a right click to open a new instance
'
' the current folder path can be obtained using Faf's code here:
' https://www.vbforums.com/showthread.php?898235-VB6-twinBASIC-Code-snippet-Close-Explorer-window-by-path

' Bugs and Regressions
' =====================

' see the separate bug/task list provided by vBAdvance


' Detail of General Status
' ========================


' Drag and drop
' ==============
' when the dragged item leaves the dock area it should leave an empty place in the dock from whence it came
' when the dragged item leaves the dock area the animation timer needs to continue but the dock itself should return to small state
' when an icon is dragged to the dock it should open up
' speeding up drag and drop - It could also be that the writing of the data has been moved to the quit command in RD. To speed up the drag and drop I could set a flag and then
' move the saving of the data to a timer driven by that flag. Check other timers are operating.

' GDI
' ===

' Dock entering
' ==============
' entering the dock at the right hand side & leaving the dock from the left hand side
'   adding a blank icon to the existing dock works

    ' we will modify the dock arrays so that position 0 and the last positions are always populated with a blank
    ' but we can show them at different sizes
    ' this will mean we can use the existing code to animate the icons without changing the logic too much
    ' we will have to change the array handling to always take the first into account

' showing more than three icons in the current BUMP animation - it is possible

' animating and centring the three animated icons
' use of math.sin
    
    ' new bounce timer
    ' math.sin fed into the timer
    ' look at the values
    ' only the two outer icons are animated and they are +/- by a value
    ' that value can be replaced by the result of the math.sin calculation
    
' SD will not support Zoom: Flat as it is a rubbish animation - documentation updated

' Advanced animations
' ===================

'Rocketdock - icon sizing
'
'When you enter an icon in Rocketdock it is not full size, it does not grow to full size until the middle of the icon is reached.
'This is unlike SD that makes sure that the centre icon is full size so that when you traverse across it you have
'a fixed size icon to use to calculate the distance across the central icon.
'
'When you scroll further across the icon mid-point it then starts to decrease in size.
'
'This implies that RD is using a fixed width area to determine the icon sizes and not calculating across one icon's width as SD is currently doing.

' The current dock stores only the left hand position of each icon and as such advanced animation cannot take place due to
' that limitation. The icons to the left and right, their size is a proportion of how far the cursor is across the central icon.
'
' No advanced animation can be performed as the properties of each icon are not known so we cannot currently transform them.
' When we store the left hand position of the icon, we have also started to gather the icon's right hand location
' iconStoreRightPixels. So, we are already holding two x values, left and right, we now need only to store the Y values,
' top and bottom. Storage for those have been added but we still need to populate those values during the icon drawing
' routines:

    ' drawSmallStaticIcons - quick and simple routine using little cpu that keeps the dock on screen in small mode
    ' drawDockByCursorEntryPosition - calculates the initial dock position from the cursor entry point
    ' sequentialBubbleAnimation - performs the animation of the central three icons.

' animating the entry of the cursor into the dock
    ' the timer modifies a grow value by incrementing a value
    ' this value is subtracted from the maxbyte value in the bump animations
    ' until the value reaches the maximum maxbyte value when the timer is stopped
    ' this will cause the icon size of the current icon to grow and not just appear instantly
    ' the same growth value will be applied to the icons to two the left and right
    ' probably as a percentage of growth
    ' we should modify sizeModifierPxls for this to work using the current animateTimer
    ' the concept is that the animate timer is animating according to the horizontal diffference using sizeModifierPxls
    ' so it should be able to animate the vertical aspect too, all we need to do is to increase sizeModifierPxls
    ' rocketdock only grows the selected icon when when the small icons have been entered

' the bounce animation is much slower in Rocketdock


' Displaying a particular icon with a varied opacity
'======================================================

' should be possible using a matrix and an opacity setting
' you have to create a colour matrix, creat a structure to store the attributes, set them and then draw using those attributes
' it does what I want but it does not scale the image output like the first option
' the image size is shrunk but the image within is simply translated into that box without being scaled

 
' Improving graphical quality
' ============================
' GdipSetCompositingMode               used for alpha blending  compositing mode specifies how source colours are combined with background colors
' GdipSetCompositingQuality            Sets the compositing quality of this Graphics object. Speed vs quality


' Graphics Next Technology
' ========================
' Next step is to convert the dock to direct3D - we have the code already and a sample dock to use a crib. The dock however has a black background, is that normal?
    
' cairo and RC5, Cairo will provide an open source replacement for vector graphics, Cairo is still cpu-bound and will require translation
' of the graphics created using Cairo to open GL in order to use the GPU.

' consolidate the two small icon drawing routines
    ' drawSmallIconDockWithFadeEffects
    ' drawDockByCursorEntryPosition
    
' DirectX 2D Jacob Roman's training utilities to implement 2D graphics in place of GDI+
'       in addition there is the VB6 dock version from the same author as the original GDI+ dock used as inspiration here,
'       that uses DirectX 2D.
'
' Avant manager - test the animation routine for the dock, circledock might be worth looking
' re: the dock disappearance option on a particular icon. The dock must check that the application is no longer running before it automatically
' returns the dock to visibility.
    
    
' Running on a second monitor
' ============================

' much more difficult than expected.

'You 'll have to use Windows API to determine the virtual screen size for a multi-monitor setup:
'
'Private Const SM_CXVIRTUALSCREEN = 78
'Private Const SM_CYVIRTUALSCREEN = 79
'Private Const SM_CMONITORS = 80
'Private Const SM_SAMEDISPLAYFORMAT = 81
'
'Private Declare Function GetSystemMetrics Lib "user32" ( _
'   ByVal nIndex As Long) As Long
'
'Public Property Get VirtualScreenWidth() As Long
'   VirtualScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
'End Property
'Public Property Get VirtualScreenHeight() As Long
'   VirtualScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
'End Property
'Public Property Get DisplayMonitorCount() As Long
'   DisplayMonitorCount = GetSystemMetrics(SM_CMONITORS)
'End Property
'Public Property Get AllMonitorsSame() As Long
'   AllMonitorsSame = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
'End Property


    ' the form needs to fill the whole virtual screen area, currently it is only filling the default form...

    ' monitors have different twip per pixel ratios and that has to be taken into account, we have a tool for that
    ' screen.twipsperpixel X & Y have been modified
    ' the monitors run in a square virtual screen and you can position the monitors within that virtual space
    ' the current monitor is determined by where you are in that virtual space
    ' if the monitor number two is set then we use the left hand position of that monitor as the left start point for the dock
    ' we determine the bounds of monitor 0?
    ' then see if the monitor is set to 1
    ' it may affect the other two tools in the way they deal with positioning certain elements - need to test that

    ' tested placing the dock using absolute positioning and it will not display on the second monitor so GDI+ is not using the virtual screen for multiple monitors
    ' some research shows C++ code that tells me to enumerate the monitos and get the device context for each and then supply that to GDI+ initialisation
    ' routine that sets the device context. I think I can do that.
    
    ' GDI+ is still not placing the output on the second monitor, send a forum post after contacting Olaf.
    ' the dock positioning occurs during setWindowCharacteristics the setWindowsPos call puts it at 0,0 as well as layering it
    ' when it was moved the dock is cut off at the edge of the window. We need to see the dock on the next monitor
    ' when a change is made then the mouse positioning needs to be moved by the same amount as it is specifying the wrong icon.
    ' consider extending GDI to cover the whole virtual screen.
    
    ' it might be useful to make the dock slightly visible so we can see where it is, the method is on the net.
    

 
' Extracting embedded icons from DLLs and EXEs
' ==============================================

' Status - We area able to extract icons using privateExtractIcon and assign them to a picture box. This
' is what we do in iconSettings and it works.
'
' We need to interface between the extracted icon and GDI+
'
' See Cintanotes GDI+Icons

' GDIIcons

'
' Other
' ======
'
' code signing certificate
'
' multi threading using TwinBasic/RadBasic some of what RD achieves may be to do with multi-threading, being able to perform
' two tasks concurrently without any delay apparent to the user. In VB6 mode can we farm any task to another thread?
'
' Build the setup2go binary for SD and the sub-components


'
    
    
    
'========================================================================================================
' SteamyDock
'
' A VB6 GDI+ dock for Reactos, XP, Win7, 8 and 10.
' SteamyDock is a functional reproduction of the dock we all know and love - Rocketdock for Windows from Punklabs.
'
' Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
'
'           MZ-TOOLS https://www.mztools.com/
'           CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
'           Rubberduck http://rubberduckvba.com/
'           Rocketdock https://punklabs.com/
'           Registry code ALLAPI.COM
'           La Volpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'           PrivateExtractIcons code http://www.activevb.de/rubriken/
'           Persistent debug code http://www.vbforums.com/member.php?234143-Elroy
'           Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
'
'   Tested on :
'           ReactOS 0.4.14 32bit on virtualBox
'           Windows 7 Professional 32bit on Intel
'           Windows 7 Ultimate 64bit on Intel
'           Windows 7 Professional 64bit on Intel
'           Windows XP SP3 32bit on Intel
'           Windows 10 Home 64bit on Intel
'           Windows 10 Home 64bit on AMD
'
' Dependencies:
'           GDI+
'           A windows-alike o/s such as Windows or ReactOS
'
'========================================================================================================
'
' Credits
'
' I have really tried to maintain the credits as the project has progressed. If I have made a mistake and left someone out then
' do forgive me. I will make amends if anyone points out my mistake in leaving someone out.
'
' Peacemaker2000    Original idea for a GDI+ dock came from here:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=55352&lngWId=1&fbclid=IwAR2FeR12CdaxyOoY-muw-b6_oDW-_19oLrt8syEL6BQSX4PMEfHyWpfqpzM
'
' Olaf Schmidt    - used some of Olaf's code as examples of how to implement the handling of images using GDI+
'                   and specifically used two routines, CreateScaledImg & ReadBytesFromFile.
'
'                   Also critically, the idea of using the scripting dictionary as a repository for a collection of
'                   image bitmaps.
'
'                   In addition, the easeing functions to do the bounce animation, I initially used a .js
'                   implementation but Olaf's was better.
'
' Spider Harper     Is64bit() function.
'
' Wayne Phillips    Used a heavily modified version of his code to bring an external application window to the foreground
' https://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground
'
' www.thescarms.com Provided the code to enumerate through windows using a callback routine
'
' dee-u Candon City, Ilocos   Used a modified version of his code to obtain a window handle from a PID.
' https://www.vbforums.com/showthread.php?561413-getting-hwnd-from-process
'
' Shuja Ali @ codeguru for his settings.ini code.
'
' An unknown, untraceable source, possibly on MSN - for the KillApp code
'
' ALLAPI.COM        For the registry reading code.
'
' Elroy on VB forums for his Persistent debug window
' http://www.vbforums.com/member.php?234143-Elroy
'
' Rxbagain on codeguru for his Open File common dialog code without a dependent OCX
' http://forums.codeguru.com/member.php?92278-rxbagain
'
' si_the_geek       for his special folder code
'
' Aaron Young       for his code for registering a keypress system wide
'
'                   Lots of GDI+ examples gleaned from here:
' http://read.pudn.com/downloads29/sourcecode/windows/control/93919/Use_GDI+_(1627568102003/frmMain.frm__.htm
'
' La Volpe          Routine to check return value from any GDI++ function

' Jacques Lebrun    Function to Provide resolution of shortcuts
' https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information
'
' Dragokas systray code
'
'========================================================================================================
'
' The core of this program are the routines from Olaf Schmidt that open the image files as an ADO stream of bytes and feed
' those into GDI+. These images are then stored as bitmaps and fed into dictionary objects for storage.
'
' NOTE - Do not end this program within the IDE by RUN/END, do that a few times and GDI+ will consume all your memory until the IDE falls over. When this happens
' just close the IDE and re-open it. Instead, ALWAYS use the QUIT option on Steamydock's right click menu.
'
' NOTE - The keyboard capture for F11 key to hide the dock, is disabled during a debug run in the IDE.
'
' NOTE - The enumWindows callback function does not find certain minimised systray apps so we have a list, a kludge.
' You have to update it manually, simply add to the list those apps you find that 'can' be minimised to the systray
' if they are in the list then the program will identify them by their caption and then be able to maximise them.
'
' NOTE - Calls to subroutines are generally (not always) made using the obsolete CALL statement making them more obvious. I also work with
' other languages where the the use of brackets is required, it makes shifting from one language to another slightly less jarring.
' Functions are just referenced in the usual fashion, returning a value.
' Exception - Even though the GDI+ APIs are "Functions" they are run using the CALL statement. GDIP functions only return a zero or an error
' code whilst any returned pointers &c are provided as passed arguments and not as the function's return value. Having the call statement in
' place merely allows easy substitution for some error handling during debugging.

' Program Structure:
'
' There is a response timer and an animate timer.
' The responseTimer draws the small icons once and monitors the mouse position, the animateTimer runs at a high frequency and draws
' the whole dock multiple times per second providing the animation effect. The relationship of the timers is found in an Impress or Powerpoint type
' document in the documentation folder. There are several timers and they really control the operation of everything.
'
' Before those timers start, the program reads all the icon locations from the settings file and loads the icons into memory using a dictionary
' object to hold the data. The location of the objects is keyed. This occurs on startup. During runtime, the various images are
' recalled from memory and drawn to the screen using a for...loop.
'
' Only the central n(3) icons are resized. This way CPU usage is minimised. Memory usage is also minimal but
' all the icons must be stored in memory so there is a natural overhead. The right-click menu sits upon an invisible form
' as GDI+ does not like a menu on the same form as the GDI+ graphics. The associated icon data is stored in temporary arrays so that it
' can be processed quickly. The program keeps track of dock-initiated processes using these arrays.

' For the background image, we do NOT retain skin compatibility with Rocketdock. This is due to Punklabs overly-complex use of GDI+ in
' RD to stretch and manipulate the single small theme image into something wider that fits the whole dock.
' Instead, we have two small right/left image and one centre image that is sized in Photoshop -
' to 2000px, then we crop the image to size as required using GDI+. This cropping occurs when the image is loaded into the dictionary
' rather than when it is displayed. As SD is FOSS, a future developer can implement Rocketdock's themeing if it is really required.

' The data source has three locations. The first is the registry (obsolete), the settings.ini file in the program folder (obsolete) and the
' user data area. The first two are hangovers from Rocketdock.
'
' The registry and the original settings.ini that Rocketdock provides for variable storage are
' left-overs from XP days when the registry storage was trendy and encouraged by MS, the use of program files
' for the settings.ini, was a left-over from the days before the registry when settings were stored locally
' within the program files folder, before MS implemented folder security. Steamydock allows access to these obsolete locations to retain
' compatibility with Rocketdock but offers a third storage option in AppData compatible with modern windows requirements.

' BUILD: The program runs without any Microsoft OCX plugins. It is simply compile and go.

' Detail regarding data sources:

' origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file in program files.
' tmpSettingsFile = App.path & "\tmpSettings.ini" ' a temporary copy of the settings file that we work on.
' toolSettingsFile = SpecialFolder_AppData & <utilityName> "\settings.ini" the tool's own settings file.

' o The fourth settings file is the dockSettings.ini that sits in this location:
' C:\Users\<username>\AppData\Roaming\steamyDock\
'
' When the flag to write the 3rd settings file is set in the dock settings utility,
' we will write the rocketdock variable values to this file.
'
' docksettings.ini is partitioned as follows:
'
' [Software\SteamyDock\DockSettings] - the dockSettings tool writes here
' [Software\SteamyDock\IconSettings\Icons] - the iconSettings tool writes here
'
' re: toolSettingsFile - The utilities read their own config files for their own personal set up in their own folders in appdata
' Settings.ini, this is just for local settings that concern only the utility, look and feel, fonts &c
'
' eg.
' C:\Users\<username>\AppData\Roaming\dockSettings\settings.ini
'
' toolSettingsFile - Dock - the following items are currently inserted into the toolSettingsFile for the dockSettings utility
'
' [Software\SteamyDockSettings]
' defaultStrength = 400
' defaultStyle = False
' defaultFont=Centurion Light SF

' toolSettingsFile - Icons - the following items are currently inserted into the toolSettingsFile for the iconSettings utility

' [Software\SteamyDockSettings]
' defaultFolderNodeKey=C:\Program Files (x86)\SteamyDock\iconSettings\my collection ' this could be moved to the docksettings.ini later
' rdMapState = Visible ' as could this
' defaultSize = 8
' defaultStrength = False
' defaultStyle = False
' Quality = 1
' defaultFont=Centurion Light SF
'
'========================================================================================================
'
'    LICENSE AGREEMENTS:
'
'    Copyright © 2019 Dean Beedell
'
'    Using this program implies you have accepted the licence. The GPL licence applies to the code
'    this software Is provided 'as-is', without any express or implied warranty. In no event will the
'    author be held liable for any damages arising from the use of this software. Permission is granted to
'    anyone to use this software for any purpose, including commercial applications, and to alter it and
'    redistribute it freely, subject to the following restrictions:
'
'    1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation is required.
'    2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.
'    3. This notice may not be removed or altered from any source distribution.
'
'    This program is free software; you can redistribute it and/or modify it under the terms of the
'    GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'    even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    General Public Licence for more details.
'
'    You should have received a copy of the GNU General Public Licence along with this program; if not,
'    write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
'    USA
'
'    If you use this software in any way whatsoever then that implies acceptance of the licence. If you
'    do not wish to comply with the licence terms then please remove the download, binary and source code
'    from your systems immediately.
'
'--------------------------------------------------------------------------------------------------------------


Option Explicit

Private Declare Function OLE_CLSIDFromString Lib "ole32" Alias "CLSIDFromString" (ByVal lpszProgID As Long, ByVal pCLSID As Long) As Long


Private Declare Function Ole_CreatePic Lib "olepro32" _
                Alias "OleCreatePictureIndirect" ( _
                ByRef lpPictDesc As PictDesc, _
                ByVal riid As Long, _
                ByVal fPictureOwnsHandle As Long, _
                ByRef iPic As IPicture) As Long
                
                ' API to determine whether the program is running with administrator rights
Private Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Private Enum OLE_ERROR_CODES
    S_OK = 0
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
    E_FAIL = &H80004005
    E_UNEXPECTED = &H8000FFFF
    E_INVALIDARG = &H80070057
End Enum

' vars to obtain correct screen width (to correct VB6 bug) STARTS
Private Const HORZRES = 8
Private Const VERTRES = 10

Private lngHeight As Long
Private lngWidth As Long
Private lngCursor As Long
Private iconIndex As Single

Private sizeModifierPxls As Double
Private differenceFromLeftMostResizedIconPxls As Double
Private animateStep As Single
Private dockUpperMostPxls As Single
'Private dockTopPxls As Single '.nn
Private iconLeftmostPointPxls As Single
Private lngFont As Long
Private lngBrush As Long
Private lngFontFamily As Long
Private lngCurrentFont As Long
Private lngFormat As Long
Private iconHeightPxls As Single
'Private iconWidthPxls As Single
Private iconPosLeftPxls As Single
Private iconCurrentTopPxls As Single
Private iconCurrentBottomPxls As Single ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
Private screenBottomPxls As Single


Private bDrawn As Boolean
Private savApIMouseX As Long
Private savApIMouseY As Long
Private cHandle As Boolean

'general vars
Private fileNameArray() As String
Private normalDockWidthPxls As Long
Private expandedDockWidth As Long
Private leftIconSize As Long
Private dockJustEntered As Boolean
Private rdDefaultYPos As Integer
'Private saveStartLeftTwps As Long
Private saveStartLeftPxls As Long ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion



' bounce variables
Private sDBounceStep As Integer ' add to configuration later
Private sDBounceInterval As Integer
Private b1 As Double 'not all used yet
Private b2 As Double
Private b3 As Double
Private b4 As Double
Private b5 As Double
Private b6 As Double
Private b7 As Double
Private b8 As Double
Private b9 As Double
Private b0 As Double

' theme variables
Private rDThemeImage As String
Private rDThemeLeftMargin As Integer
Private rDThemeTopMargin  As Integer
Private rDThemeRightMargin  As Integer
Private rDThemeBottomMargin  As Integer
Private rDThemeOutsideLeftMargin  As Integer
Private rDThemeOutsideTopMargin  As Integer
Private rDThemeOutsideRightMargin  As Integer
Private rDThemeOutsideBottomMargin  As Integer

' Vars for

Private rDSeparatorImage As String
Private rDSeparatorTopMargin As Integer
Private rDSeparatorBottomMargin As Integer

Private xAxisModifier As Integer ' .57 DAEB 19/04/2021 frmMain.frm modifedAmountToSlide renamed to xAxisModifier for clarity's sake
Private yAxisModifier As Integer '.nn added for future Y axis animation
Private autoHideMode As String
Private autoSlideMode As String
Private slideOutFlag As Boolean
Private currentDockTopPxls As Integer
Private nMinuteExposeTimerCount As Integer

' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
' .23 DAEB frmMain.frm 08/02/2021 Changed from an array to a single var
Private lHotKey As Long
Public lPressed As Long '.nn


Private dockZorder As String '.nn
' .58 DAEB 21/04/2021 frmMain.frm added timer and vars to check to see if the system has just emerged from sleep
Dim strTimeThen As String



' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
Private hourglassimage As String
Private hourglassTimerCount As Integer

' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private mouseDownTime As Long

' .84 DAEB 20/07/2021 frmMain.frm Added prevention of the dock returning until the hiding application is no longer running.
Private autoHideProcessName As String


Private soundtoplay As String
Private delayRunTimerCount As Integer

   



' .nn DAEB 16/04/2022 frmMain.frm new timer to force reveal the dock when the hiding process has ended
'---------------------------------------------------------------------------------------
' Procedure : forceHideRevealTimer_Timer
' DateTime  : 16/04/2022 12:59
' Author    : beededea
' Purpose   : Reveals the dock 0 - 1.5 secs after the hiding process has ended
'---------------------------------------------------------------------------------------
'
Private Sub forceHideRevealTimer_Timer()
    Dim itIs As Boolean: itIs = False

   On Error GoTo forceHideRevealTimer_Timer_Error

        'if the dock has been manually revealed by the user and another app has been run in the meantime
        ' then the autoHideProcessName will be blank
        If autoHideProcessName = vbNullString Then
            forceHideRevealTimer.Enabled = False
            Exit Sub
        End If
        
        ' check to see if the process that hid the dock is still running
        ' the dock will not automatically appear until the process that hid it has finished (full screen games)
        itIs = IsRunning(autoHideProcessName, vbNull)
        If itIs = True Then
            ' the timer will continue to run
            Exit Sub
        Else
            autoHideProcessName = vbNullString
            forceHideRevealTimer.Enabled = False
            Call ShowDockNow
        End If

   On Error GoTo 0
   Exit Sub

forceHideRevealTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure forceHideRevealTimer_Timer of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 01/05/2021
' Purpose   : We handle the mouse events during mouseUp, we only set some states here
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Form_MouseDown_Error
    
    ' .65 DAEB 30/04/2021 frmMain.frm Added mouseDown event to capture the time of first press and code to simulate a drag and drop of an icon from the dock
    dragFromDockOperating = False
    mouseDownTime = GetTickCount 'we do not use TimeValue(Now) as it does not count milliseconds
    
    ' .75 DAEB 12/05/2021 frmMain.frm Changed Form_MouseMove to act as the correct event to a drag and drop operating from the dock
    selectedIconIndex = iconIndex ' this is the icon we will be bouncing
    dragImageToDisplay = selectedIconIndex & "ResizedImg" & LTrim$(Str$(iconSizeLargePxls))
    
'    dock.animateTimer.Enabled = False
'    dock.responseTimer.Enabled = False

    On Error GoTo 0
    Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form dock"
End Sub

' .75 DAEB 12/05/2021 frmMain.frm Changed Form_MouseMove to act as the correct event to measure a drag and drop operating from the dock
'---------------------------------------------------------------------------------------
' Procedure : Form_MouseMove
' Author    : beededea
' Date      : 12/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim timeDiff As Integer: timeDiff = 0
    Dim tickCount As Long: tickCount = 0
    
    On Error GoTo Form_MouseMove_Error

    If mouseDownTime = "0" Then Exit Sub

    ' calculates the time since the mouseDown and if no mouseup within 1/4 of a second assume it is a drag from the dock
    If mouseDownTime <> "0" Then ' time since the mouseDown event occurred
            tickCount = GetTickCount
            timeDiff = tickCount - mouseDownTime
            If Val(rDLockIcons) = 0 And timeDiff > 250 Then
                mouseDownTime = "0" ' reset
                dragFromDockOperating = True
                dragToDockOperating = True
                hourGlassTimer.Enabled = True
            End If
        End If

    On Error GoTo 0
    Exit Sub

Form_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Initialize
' Author    : beededea
' Date      : 28/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Initialize()
     fInitialise ' we can call this routine from elsewhere whereas we can't easily call Form_Initialize during our program
End Sub
    


    
'---------------------------------------------------------------------------------------
' Procedure : fInitialise
' Author    : beededea
' Date      : 15/04/2020
' Purpose   : All the form initialisation code moved to here so we can call this routine
'             from elsewhere whereas we can't call Form_Initialize directly
'---------------------------------------------------------------------------------------

Public Sub fInitialise()

   On Error GoTo fInitialise_Error

    Call initialiseGlobalVars
    
    ' local variables declared

    Dim thiskey As String: thiskey = vbNullString
    Dim a As Integer: a = 0
    Dim strKey As String: strKey = vbNullString

    ' other global variable assignments
    
    debugflg = 0
    animationFlg = False
    dockHidden = False
    dockOpacity = 100
    
    screenWidthTwips = 0
    screenHeightTwips = 0
    screenWidthPixels = 0
    screenHeightPixels = 0
    
    ' animation timers
    selectedIconIndex = 999 ' sets the icon to bounce index to something that will never occur
    bounceTimerRun = 1
    sDBounceStep = 4 ' we can add a slider for this in the dockSettings later
    sDBounceInterval = 5
    'bounceUpTimer.Interval = sDBounceInterval * 3
    'bounceDownTimer.Interval = sDBounceInterval * 3
    
    
    autoFadeOutTimerCount = 0
    autoFadeInTimerCount = 0 ' .01 DAEB 24/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
    autoSlideOutTimerCount = 0 ' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions
    autoSlideInTimerCount = 0 ' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions
    autoHideRevealTimerCount = 0
    
    'other vars
    iconCurrentTopPxls = 0
    iconCurrentBottomPxls = 0 ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
    
    dockUpperMostPxls = 0
    rdDefaultYPos = 6
    readEmbeddedIcons = False
    dockJustEntered = True

    sixtyFourBit = False
    
    ' useful global variable set
    sixtyFourBit = Is64bit()
    
    xAxisModifier = 0 ' .57 DAEB 19/04/2021 frmMain.frm modifedAmountToSlide renamed to xAxisModifier for clarity's sake
    yAxisModifier = 0 '.nn
    
    autoHideMode = "fadeout"
    'autoSlideMode = "slideout" ' .47 DAEB 01/04/2021 frmMain.frm autoSlideMode is now undefined at startup - this allowed the top position to operate as expected
    autoSlideMode = vbNullString
    slideOutFlag = False
    
    nMinuteExposeTimerCount = 0
    autoHideProcessName = vbNullString
    
    hourglassimage = vbNullString ' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
    hourglassTimerCount = 1
    strTimeThen = Now
    
    bounceZone = 75 ' .82 DAEB 12/07/2021 frmMain.frm Add the BounceZone as a configurable variable.


    msgBoxOut = True
    msgLogOut = True
    
    delayRunTimerCount = 0
    
    ' .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
    ' .05 DAEB frmMain.frm 10/02/2021 changes to handle invisible windows that exist in the known apps systray list
    'appSystrayTypes = "GPU-Z|XWidget|Lasso|Open Hardware Monitor|CintaNotes" ' systray apps list, add to the list those apps you find that can be minimised to the systray
    
    '=========================================
    ' program starts!
    '=========================================
    
    ' comment the following function back in only when debugging
    Call toggleDebugging
        
    ' write to the debuglog to log
    debugLog "*****************************"
    debugLog "% SteamyDock program started."
    debugLog "*****************************"
    
    ' extracts all the known drive names using Windows APIs to a useful global var
    Call getAllDriveNames(sAllDrives)
        
    'if the process already exists then kill it
    Call testDockRunning
    
    ' check the state of the licence
    Call checkLicenceState
    
    ' check the Windows version
    Call testWindowsVersion(classicThemeCapable)
    
    ' turn off the option to run as administrator
    Call disableAdmin  ' .17 DAEB frmMain.frm 27/01/2021 Moved disabling admin to a separate routine

    ' we check to see if rocketdock is installed in order to know the location of the settings.ini file used by Rocketdock
    'Call checkRocketdockInstallation ' also sets rdAppPath
    
    ' check where steamyDock is installed, seems obvious but someone could be running the binary somewhere remote from its default location
    Call checkSteamyDockInstallation ' in any case it sets the sdPathPath

    ' validate all the components are in place for this program to run.
    If fValidateComponents = False Then
        ' at the moment if components are missing we do nothing, just let SD attempt to start,
    End If
    ' get the location of the dock's new settings file
    Call locateDockSettingsFile

    ' read the dock settings from INI or from registry
    Call readDockConfiguration
    
    ' set the hotkey toggle to the user's chosen function key
    Call setUserHotKey ' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
    
    ' working here!
    ' no need to determine which monitor we should use, we know this from rdMonitor gleaned from readDockConfiguration above.
    ' monitor validation, despite the value set in config, we need to check again as a monitor may have been disconnected.
    If Val(rDMonitor) + 1 > GetMonitorCount Then
        rDMonitor = "0" 'validate
    End If
    
'    If Val(rDMonitor) > 0 Then
'        ' get screen bounds
'        ' position the dock onto the correct monitor using the current monitor left position plus 1
'        getDeviceHdc
'
'        ' set the device (screen) context default to primary monitor
'        If hdcScreen = 0 Then
            hdcScreen = Me.hdc
'        End If
'
'        'CenterFormOnMonitorTwo dock
'    End If
        
        
    ' check the selected monitor properties and determine the number of twips per pixel for this screen
    Call monitorProperties(dock, screenTwipsPerPixelX, screenTwipsPerPixelY)
    
    ' resolve VB6 sizing width bug
    Call resolveVB6SizeBug ' requires MonitorProperties to be in place above to assign a value to screenTwipsPerPixelY
    
    ' configuration private numeric vars that are easier to manipulate throughout the program than the string equivalents
    Call setLocalConfigurationVars
    
    ' get the location of the dock's theme settings file
    Call locateThemeSettingsFile
        
    ' read the background theme settings from INI
    Call readThemeConfiguration
    
    ' read the tool settings file and do some things for the first and only time
    'Call readToolSettings ' program specific settings do not apply to the dock, left here just in case we need it
    
    ' Initialises GDI Plus
    Call initialiseGDIStartup
    
    ' Create the VB collection object where the image bitmaps will be stored
    Call createDictionaryObjects

    ' Resize data arrays and load the icon images into the collections
    Call prepareArraysAndCollections
    
    ' sets bmpInfo object to create a bitmap of the whole screen size and get a handle to the Device Context
    Call createGDIPlusElements
           
    ' briefly display the product splash screen if set to do so
    Call showSplashScreen ' has to be at the end of the start up as we need to read the config file but also so as to not cause a clear outline to appear where the splash screen should be
    
    'creates a bitmap section in memory that applications can write to directly
    If debugflg = 1 Then debugLog "% sub readyGDIPlus" ' the debug needs to be here
    Call readyGDIPlus
        
    ' set autohide characteristics, needs to be exactly here
    Call setAutoHide
    
    ' update the window with the appropriately sized and qualified image
    Call setWindowCharacteristics ' This is the function that actually changes the display, called by animate timers, must also be here
        
    ' set up the timers that check to see if each process is running
    Call setUpProcessTimers
    
    ' Checks each target command for validity and sets a flag in an array to place a red X on the icon.
    Call checkTargetCommandValidity
    
    
    If rDSoundSelection = "0" Then
        soundtoplay = vbNullString
    ElseIf rDSoundSelection = "1" Then
        soundtoplay = sdAppPath & "\sounds\ting.wav"
    ElseIf rDSoundSelection = "2" Then
        soundtoplay = sdAppPath & "\sounds\click.wav"
    End If
    
    debugLog "******************************"
    debugLog "% SteamyDock startup complete."
    debugLog "******************************"
    
    
   On Error GoTo 0
   Exit Sub

fInitialise_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fInitialise of Form dock"

End Sub
' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
'---------------------------------------------------------------------------------------
'Procedure:   setUserHotKey
'Author:      beededea
' Date      : 26/01/2021
' Purpose   : using the user's choice, set the default keypress to work system wide
'---------------------------------------------------------------------------------------
'
Private Sub setUserHotKey()
   On Error GoTo setUserHotKey_Error
   
    If debugflg = 1 Then debugLog "% sub setUserHotKey"

    ' check to see whether the program is running within the IDE - if so disable system key hooks that capture F key by default
    ' if the program is run in the IDE (Debug mode) with the system wide key hook operative, the IDE will crash shortly afterward
    If Not InIDE Then
        ' .23 DAEB frmMain.frm 08/02/2021 Changed from an array to a single var
        If rDHotKeyToggle = "F1" Then lHotKey = SetHotKey(0, vbKeyF1)
        If rDHotKeyToggle = "F2" Then lHotKey = SetHotKey(0, vbKeyF2)
        If rDHotKeyToggle = "F3" Then lHotKey = SetHotKey(0, vbKeyF3)
        If rDHotKeyToggle = "F4" Then lHotKey = SetHotKey(0, vbKeyF4)
        If rDHotKeyToggle = "F5" Then lHotKey = SetHotKey(0, vbKeyF5)
        If rDHotKeyToggle = "F6" Then lHotKey = SetHotKey(0, vbKeyF6)
        If rDHotKeyToggle = "F7" Then lHotKey = SetHotKey(0, vbKeyF7)
        If rDHotKeyToggle = "F8" Then lHotKey = SetHotKey(0, vbKeyF8)
        If rDHotKeyToggle = "F9" Then lHotKey = SetHotKey(0, vbKeyF9)
        If rDHotKeyToggle = "F10" Then lHotKey = SetHotKey(0, vbKeyF10)
        If rDHotKeyToggle = "F11" Then lHotKey = SetHotKey(0, vbKeyF11)
        If rDHotKeyToggle = "F12" Then lHotKey = SetHotKey(0, vbKeyF12)
        If rDHotKeyToggle = "Disabled" Then lHotKey = 0
    End If
   On Error GoTo 0
   Exit Sub

setUserHotKey_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setUserHotKey of Form dock"
    End Sub
' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support

    

'---------------------------------------------------------------------------------------
' Procedure : showSplashScreen
' Author    : beededea
' Date      : 01/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub showSplashScreen()
   On Error GoTo showSplashScreen_Error
   
    If debugflg = 1 Then debugLog "% sub showSplashScreen"

    If sDSplashStatus = "1" Then
        splashForm.splashTimer.Enabled = True

        splashForm.Show
        splashForm.chkSplashDisable.Value = 0
    Else
        splashForm.Hide
    End If


   On Error GoTo 0
   Exit Sub

showSplashScreen_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showSplashScreen of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_MouseUp
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : this is the equivalent of an icon MouseUp event, a click anywhere on the form
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   On Error GoTo Form_MouseUp_Error

    Call fMouseUp(Button) ' occasionally we want to be able to trigger this manually and you can't call a Form_MouseUp


   On Error GoTo 0
   Exit Sub

Form_MouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseUp of Form dock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fMouseUp
' Author    : beededea
' Date      : 11/04/2020
' Purpose   : you cannot directly call a form mouseUp event from anywhere else so this is the equivalent that is called by the
'             Form_MouseUp event and we can also call fMouseUp as and when we require.
'---------------------------------------------------------------------------------------
' the mouse up event handles the left button click event and the right click menu activation. It also identifies a drag to or
' from the dock. Identifying a drag from the dock cannot be done in a traditional manner as we are not dropping it onto
' any traditional VB6 control. So a drag over or drop can never be captured. Instead, if we measure the time between mousedown
' and mouse up then we have an indication of a drag from the dock in progress. A workaround.


Public Sub fMouseUp(Button As Integer)
   On Error GoTo fMouseUp_Error
   
    Dim timeDiff As Integer:  timeDiff = 0
    Dim tickCount As Long: tickCount = 0
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim thisFilename As String: thisFilename = vbNullString
    
    Dim sourceIconIndex As Integer: sourceIconIndex = 0
    Dim targetIconIndex As Integer: targetIconIndex = 0
        
    mouseDownTime = "0"
    
    '.76 DAEB 12/05/2021 frmMain.frm Moved from the runtimer as some of the data is required before the run begins
    Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", selectedIconIndex, dockSettingsFile)
    
    If dragToDockOperating = True Then
        hourGlassTimer.Enabled = False
        dragToDockOperating = False
    End If
    
    If Button = 2 Then 'right click to display a menu
        dragFromDockOperating = False
        
        If dragToDockOperating = True Then
            hourGlassTimer.Enabled = False ' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
        Else
            animateTimer.Enabled = False ' stops the animation
            responseTimer.Enabled = False ' stops the assessment of the mouse position
        End If
        
        
        ' check the current process is running by looking into the array that contains a list of running processes using selectedIconIndex
        If processCheckArray(selectedIconIndex) = "False" Then
            forceRunNewAppFlag = False

            menuForm.mnuCloseApp.Visible = False
            menuForm.mnuRunNewApp.Visible = False
            menuForm.mnuRun.Visible = True
            menuForm.mnuFocusApp.Visible = False
            menuForm.mnuBackApp.Visible = False
        Else
            ' if the process is marked as running then enable the menu options
            menuForm.mnuCloseApp.Visible = True
            menuForm.mnuRunNewApp.Visible = True
            menuForm.mnuRun.Visible = False
            menuForm.mnuFocusApp.Visible = True
            menuForm.mnuBackApp.Visible = True
        End If
        
        PopupMenu menuForm.mnuMainMenu, vbPopupMenuRightButton
        'the popupmenu event returns here and re-enables the mouse response and animation timers
        
        If hideDockForNMinutes = False Then ' re-enable timers only when the dock is operating normally and not when instructed to hide
            animateTimer.Enabled = True
            responseTimer.Enabled = True
        End If
        
        usedMenuFlag = True ' essential
        
    Else  'any normal left button click
    
        ' .79 DAEB 21/05/2021 frmMain.frm Disable any currently active bounce up or down
        bounceCounter = 0
        bounceUpTimer.Enabled = False
        bounceDownTimer.Enabled = False

    
        ' identify drag from the dock cannot be done in a traditional manner as we are not dropping it onto anything, so a drag
        ' over or drop is not initiated. Instead if we measure the time between mousedown and mouse up then we have an indication of a drag from the dock
        
        ' .75 DAEB 12/05/2021 frmMain.frm Changed Form_MouseMove to act as the correct event to a drag and drop operating from the dock
        If dragFromDockOperating = True Then
            If insideDockFlg = False Then
                Call deleteThisIcon
                Exit Sub
            End If
               
            ' at this point we drop an icon from one part of the dock to another
            If insideDockFlg = True Then 'allow a MouseUp to capture a drag from one part of the dock to another
                dragFromDockOperating = False
                dragInsideDockOperating = True 'check for dragInsideDockOperating
                If selectedIconIndex <> iconIndex Then ' cannot drop onto itself
                    ' we read the source icon details
                    sourceIconIndex = selectedIconIndex
                    targetIconIndex = iconIndex
                    
                    selectedIconIndex = targetIconIndex ' reset the selectedIconIndex
                    thisFilename = sFilename
                    Call menuAddSummat(thisFilename, sTitle, sCommand, sArguments, sWorkingDirectory, sShowCmd, sOpenRunning, sIsSeparator, sDockletFile, sUseContext, sUseDialog, sUseDialogAfter, sQuickLaunch)
                    Call menuForm.postAddIConTasks(thisFilename, sTitle)
                    
                    'delete the old icon at its new location
                    If sourceIconIndex < targetIconIndex Then
                        selectedIconIndex = sourceIconIndex
                    Else
                        selectedIconIndex = sourceIconIndex + 1
                    End If
                    Call deleteThisIcon
                    
                    'MsgBox "Dragged icon " & dragImageToDisplay & " " & selectedIconIndex & " " & sCommand & " to position " & iconIndex
                
                Else
                    
                    If Val(rDHoverFX) = 1 Then Call selectBubbleAnimation(3) ' select drawSmallStaticIcons redraw the icons if dragged to the same position
                End If

                ' we use the existing "add an icon" or icon deletion code to move the icon collection to a new temporary dock and write the new details there and then back again to the icon collection
                ' inserting as we go, the icon in its new position and not in its old
                
                Exit Sub
            End If
        End If
        
        
        ' check the current process is running by looking into the array that contains a list of running processes using selectedIconIndex
        If processCheckArray(selectedIconIndex) = "False" Then
            ' it would be nice to lock the x axis during the bounce animation
            If userLevel <> "runas" Then userLevel = "open"
                        
            ' the runCommand is called from within the bounceDownTimer
            
            bounceUpTimer.Enabled = True
            'animateTimer.Enabled = True
            
        Else
            ' the runCommand is called directly when the app is already running to avoid delay, no bounce
            If userLevel <> "runas" Then userLevel = "open"
            Call runCommand("run", "") ' added new parameter to allow override .68
        End If
        
        If FExists(soundtoplay) And rDSoundSelection <> 0 Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        
    End If


   On Error GoTo 0
   Exit Sub

fMouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fMouseUp of Form dock"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_OLEDragDrop
' Author    : beededea
' Date      : 15/04/2020
' Purpose   : Handles drag and drop to the dock, only file types accepted. If an image, drops it straight onto the dock.
'             If it is a binary then we use code to try to extract the embededded icons using privateExtractIcons API
'             especially when the icon is a bigger one, if it is only a low resolution icon then we give it an icon based upon its suffix.
'             direct from the icon collection.
'             If it is a special binary, msc, cpl then it is given an icon from the collection
'             If it is a shortcut we have some code to investigate the shortcut for the link details
'
'             I have made the decision not to use the embedded icons by default as for the majority of
'             Win o/ses before 10 the embedded icons are low resolution and puny. Instead we use document types from the collection.
'             This IS STEAMYDOCK!
'---------------------------------------------------------------------------------------
    'The Format numbers used in the OLE DragDrop data structure, are:
    'Text = 1 (vbCFText)
    'Bitmap = 2 (vbCFBitmap)
    'Metafile = 3
    'Emetafile = 14
    'DIB = 8
    'Palette = 9
    'Files = 15 (vbCFFiles)
    'RTF = -16639
    '
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    Dim suffix As String: suffix = vbNullString
    Dim Filename As String: Filename = vbNullString
    Dim iconImage As String: iconImage = vbNullString
    Dim iconTitle As String: iconTitle = vbNullString
    Dim iconFileName As String: iconFileName = vbNullString
    Dim iconCommand As String: iconCommand = vbNullString
    Dim iconArguments As String: iconArguments = vbNullString
    Dim iconWorkingDirectory As String: iconWorkingDirectory = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim nname As String: nname = vbNullString
    Dim npath As String: npath = vbNullString
    Dim ndesc As String: ndesc = vbNullString
    Dim nwork As String: nwork = vbNullString
    Dim nargs As String: nargs = vbNullString
    Dim thisShortcut As Link

    On Error GoTo Form_OLEDragDrop_Error
    
    ' if the dock is not the bottom layer then pop up a message box
    ' ie. don't pop it up if layered underneath everything as no-one will see the msgbox
    If rDLockIcons = 1 And (rDzOrderMode = "0" Or rDzOrderMode = "1") Then
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, "Sorry, the dock is currently locked, so drag and drop is disabled!", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Sorry, the dock is currently locked, so drag and drop is disabled!"
        Exit Sub
    End If
    
    iconImage = vbNullString
    iconTitle = vbNullString
    iconArguments = vbNullString
    iconWorkingDirectory = vbNullString
        
    selectedIconIndex = iconIndex ' this is the icon we will be bouncing
    
    ' if there is more than one file dropped reject the drop
    ' if the dock is not the bottom layer then pop up a message box
    ' ie. don't pop it up if layered underneath everything as no-one will see the msgbox
    If Data.Files.Count > 1 And (rDzOrderMode = "0" Or rDzOrderMode = "1") Then
       ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, "Sorry, can only accept one icon drop at a time, you have dropped " & Data.Files.Count, "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        '        MsgBox "Sorry, can only accept one icon drop at a time, you have dropped " & Data.Files.count
        Exit Sub
    End If
    
    If Data.GetFormat(vbCFFiles) Then
        ' if it is of type 'file' then determine what to do, I have specific catch-alls and
        ' sections for each, just in case any specific tasks are required for each type.
        ' this could all be removed later as the actions seem more or less the same for each.
        
        ' Data.Files.Item
    
        iconTitle = Data.Files(1) ' set the title for all types
        iconCommand = Data.Files(1) ' set the command for all types
        
        ' is it a folder, does the folder exist
        If DirExists(iconTitle) Then
            iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-dir.png"
            If FExists(iconFileName) Then
                iconImage = iconFileName
            End If
        Else
    
              suffix = LCase(ExtractSuffixWithDot(Data.Files(1)))
              If InStr(".exe,.bat,.msc,.cpl,.lnk", suffix) <> 0 Then
                  
                  Effect = vbDropEffectCopy
                 
                  'if an exe is dragged and dropped onto RD it is given an id, that it appends to the binary name after an additional "?"
                  ' that ? signifies what? Well, possibly it is the handle of the embedded icon only added the one time, so that when the binary is read in the future the handle is already there
                  ' and that can be used to populate image array? Untested.
                  ' in this case we just need to note the ? and then query the binary for an embedded icon handle and compare it to the id that RD has given it.
                  ' if it is the same then we can perhaps simulate the same.
                  
                  If suffix = ".exe" Then
                    ' we should, if it is a EXE dig into it to determine the icon using privateExtractIcon
                                         
                    ' However, we do not extract the icon from the shortcut as it will be useless for steamydock
                    ' VB6 not being able to extract and handle a transparent PNG form
                    ' even if it was we have no current method of making a transparent PNG from a bitmap or ICO that
                    ' I can easily transfer to the GDI collection - but I am working on it...
                    ' the vast majority of default icons are far too small for steamydock in any case.
                    ' the result of the above is that there is currently no icon extracted, though that may change.
                    
                    ' instead we have a list of apps that we can match the shortcut name against, it exists in an external comma
                    ' delimited file. The list has two identification factors that are used to find a match and then we find an
                    ' associated icon to use with a relative path.
                       
                    iconFileName = identifyAppIcons(iconCommand) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                       
                    If FExists(iconFileName) Then
                      iconImage = iconFileName
                    Else
                      iconImage = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-EXE.png"
                    End If
                    
                  End If
                  
                  If suffix = ".msc" Then
                      ' if it is a MSC then  give it a SYSTEM type icon (EVENT)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-msc.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
                  If suffix = ".bat" Then
                      ' if it is a BAT then give it a BATCH type icon (NOTEPAD)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-bat.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
                  If suffix = ".cpl" Then
                      ' if it is a CPL then give it a SYSTEM type icon (CONSOLE)
                      
                      ' if there is no icon embedded found then use the default icon
                       ' check the icon exists
                      iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-cpl.png"
                      If FExists(iconFileName) Then
                          iconImage = iconFileName
                      End If
                  End If
                  
            '       If it is a shortcut we have some code to investigate the shortcut for the link details
                  If suffix = ".lnk" Then
                        ' if it is a short cut then you can use two methods, the first is currently limited to only
                        ' producing the path alone but it does avoid using the shell method that causes FPs to occur in av tools

                        Call GetShortcutInfo(iconCommand, thisShortcut) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                                       
                        iconTitle = GetFileNameFromPath(thisShortcut.Filename)
                        
                        If Not thisShortcut.Filename = vbNullString Then
                            iconCommand = LCase(thisShortcut.Filename)
                        End If
                        iconArguments = thisShortcut.Arguments
                        iconWorkingDirectory = thisShortcut.RelPath
                        
                        ' .55 DAEB 19/04/2021 frmMain.frm Added call to the older function to identify an icon using the shell object
                        'if the icontitle and command are blank then this is user-created link that only provides the relative path
                        If iconTitle = vbNullString And thisShortcut.Filename = vbNullString And Not iconWorkingDirectory = vbNullString Then
                            Call GetShellShortcutInfo(iconCommand, nname, npath, ndesc, nwork, nargs)
                    
                            iconTitle = nname
                            iconCommand = npath
                            iconArguments = nargs
                            iconWorkingDirectory = nwork
                        End If
                       
                      ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                      
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
                        iconImage = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-lnk.png"
                      End If
                  End If
            
              ElseIf InStr(".png, .bmp, .gif, .jpg, .jpeg, .ico, .tif, .tiff", suffix) <> 0 Then
                  ' See if this is a file name ending in bmp, gif, jpg, or jpeg or tiff.
                  ' if so use the filename and drop it onto the dock
                  
                  Effect = vbDropEffectCopy
                  
                  iconImage = iconCommand
                  If Not FExists(iconImage) Then
                      Exit Sub
                  End If
              
              ElseIf InStr(".zip, .7z, .arj, .deb, .pkg, .rar, .rpm, .tar, .gz, .z, .bck", suffix) <> 0 Then
                  
                '    .7z - 7-Zip compressed file
                '    .arj - ARJ compressed file
                '    .deb - Debian software package file
                '    .pkg - Package file
                '    .rar - RAR file
                '    .rpm - Red Hat Package Manager
                '    .tar.gz - Tarball compressed file
                '    .z - Z compressed file
                '    .zip - Zip compressed file
                
                ' See if this is a file name ending in the above
                ' if so use the filename and drop it onto the dock
                  
                Effect = vbDropEffectCopy
                  
                iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-zip.png"
                If FExists(iconFileName) Then
                    iconImage = iconFileName
                End If
    
                iconImage = iconCommand
                If Not FExists(iconImage) Then
                    Exit Sub
                End If
            
                      
              Else ' does not match any given type so see if we have an icon in the collection ready for it.
              
                  ' take the suffix and find a file in the collection that matches
                  ' if the file exists then add it to the menu
                  ' otherwise just do an empty default icon
                  
                  Effect = vbDropEffectCopy
                  
                  suffix = LCase(ExtractSuffix(Data.Files(1)))
                  iconImage = App.Path & "\iconSettings\my collection\steampunk icons MKVI\document-" & suffix & ".png"
                  iconCommand = Data.Files(1)
                  If Not FExists(iconImage) Then
                      iconImage = App.Path & "\nixietubelargeQ.png"
                  End If
                      
              End If
        End If
        
        ' if no specific image found
        If iconImage = vbNullString Then
            iconImage = App.Path & "\nixietubelargeQ.png"
        End If
        
        If FExists(iconImage) Then ' last check that the default ? image has not been deleted.
            Call menuAddSummat(iconImage, iconTitle, iconCommand, iconArguments, iconWorkingDirectory, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
            ' .51 DAEB 08/04/2021 frmMain.frm calls mnuIconSettings_Click to start up the icon settings tools and display the properties of the new icon.
            Call menuForm.postAddIConTasks(iconImage, iconTitle)
            ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
             'MessageBox Me.hwnd, iconTitle & " dropped successfully to the dock. ", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
             '            MsgBox iconTitle & " dropped successfully to the dock. ", vbSystemModal
        Else
            ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
             'MessageBox Me.hwnd, iconImage & " missing default image, " & App.Path & "\nixietubelargeQ.png" & " drop unsuccessful. ", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
             '            MsgBox iconImage & " missing default image, " & App.Path & "\nixietubelargeQ.png" & " drop unsuccessful. ", vbSystemModal
        End If
        
        
        'Call menuForm.mnuIconSettings_Click
        
    Else
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, " unknown file Object OLE dropped onto SteamyDock.", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        'MsgBox " unknown file Object OLE dropped onto SteamyDock."
    End If
    
    dragToDockOperating = False

    On Error GoTo 0
    Exit Sub

Form_OLEDragDrop_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_OLEDragDrop of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_OLEDragOver
' Author    : beededea
' Date      : 28/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
   On Error GoTo Form_OLEDragOver_Error

    If rDLockIcons = 0 Then
        dragToDockOperating = True
        hourGlassTimer.Enabled = True
    End If

   On Error GoTo 0
   Exit Sub

Form_OLEDragOver_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_OLEDragOver of Form dock"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

    Call dock.shutdwnGDI

    ' shutdown GDI
'    If lngImage Then
'        Call GdipReleaseDC(lngImage, dcMemory)
'        Call GdipDeleteGraphics(lngImage)
'    End If
'    If lngBitmap Then Call GdipDisposeImage(lngBitmap)
'    If lngGDI Then Call GdiplusShutdown(lngGDI)
    
    ' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
'    Dim lIndex As Long
'    For lIndex = 0 To 3
'        RemoveHotKey lHotKey(lIndex) ' removes the keys set when the app ends
'    Next

    ' .23 DAEB frmMain.frm 08/02/2021 Changed from an array to a single var
     RemoveHotKey lHotKey

    ' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form dock"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : initiatedProcessTimer
' Author    : beededea
' Date      : 10/07/2020
' Purpose   :
' Provides regular checking of ONLY processes initiated by the dock, removes the running indicator cog
' an array of the same size as the main icon arrays, each dock-initiated item resides in its own numbered location.
' Checking for just a few elements in an array, the empty elements can be bypassed, instead probing just these few processes
' for existence, this can be carried out much more frequently than the current once every 10-65 seconds full process check.
' If the result of the search is false then the program has completed and the cog can be removed.
' processCheckArray(useloop) - is the array that determines whether a cog is placed on an icon.
'---------------------------------------------------------------------------------------

Private Sub initiatedProcessTimer_Timer()

    Dim useloop As Long: useloop = 0
    Dim itIs As Boolean: itIs = False
     
    On Error GoTo initiatedProcessTimer_Error

        For useloop = 0 To rdIconMaximum
            ' instead of looping through all elements in the docksettings.ini file, we now store all the current commands in an array
            ' we loop through the array much quicker than looping through the temporary settings file.
            ' all we have to do is to remember to populate the array whenever an icon is added or deleted
            If Not initiatedProcessArray(useloop) = vbNullString Then
                itIs = IsRunning(initiatedProcessArray(useloop), vbNull)
                If itIs = False Then
                    processCheckArray(useloop) = False ' the cog array
                    initiatedProcessArray(useloop) = vbNullString ' removes the entry from the test array so it isn't caught again
                End If
                ' .81 DAEB 28/05/2021 frmMain.frm Refresh the running process with a cog when the process is running, this had been removed earlier
                bDrawn = False
                If smallDockBeenDrawn = True Then
                    If Val(rDHoverFX) = 1 Then Call selectBubbleAnimation(3) ' select drawSmallStaticIcons redraw the icons if dragged to the same position
                End If
            End If
        Next useloop

   On Error GoTo 0
   Exit Sub

initiatedProcessTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initiatedProcessTimer of Form dock"

End Sub








'---------------------------------------------------------------------------------------
' Procedure : positionZTimer_Timer
' Author    : beededea
' Date      : 02/03/2021
' Purpose   : .32 DAEB 03/03/2021 frmMain.frm Placing the dock back in the defined z-order
'---------------------------------------------------------------------------------------
'
Private Sub positionZTimer_Timer()
   On Error GoTo positionZTimer_Timer_Error

    If animateTimer.Enabled = True Or animatedIconsRaised = True Or menuForm.Visible = True Then
        Exit Sub
    End If
    
    If dockZorder = "high" Then
        If rDzOrderMode = "0" Then
            SetWindowPos dock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
        ElseIf rDzOrderMode = "1" Then
            SetWindowPos dock.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE
        ElseIf rDzOrderMode = "2" Then
            SetWindowPos dock.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE
        End If
        dockZorder = "low"
    End If

   On Error GoTo 0
   Exit Sub

positionZTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionZTimer_Timer of Form dock"
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : responseTimer
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : Determines whether to turn on the animate timer
'             It determines the position of the mouse and then if no animation required, draws the icons
'             in small size (only) when the mouse cursor leaves the dock area
'---------------------------------------------------------------------------------------
'
Private Sub responseTimer_Timer()
    On Error GoTo responseTimer_Error
    
    Dim dockHeightPxls As Long: dockHeightPxls = 0
    Dim outsideDock As Boolean: outsideDock = False
     
    lngReturn = GetCursorPos(apiMouse) ' return the mouse position every 200ms - sufficient
        
    dockHeightPxls = fSetDockUpperHeightLimit()
    Call tuneResponseTimerInterval
    Call defineDockEntranceLocation(dockHeightPxls)
    outsideDock = fTestCursorWithinDockYPosition(outsideDock)
    
    insideDockFlg = Not outsideDock '.nn Added as part of the drag and drop functionality
    
    ' the mouse has left the Max icon area
    If outsideDock = True And dragFromDockOperating = False Then
        Call stopAnimating
        Exit Sub ' leave the timer loop and do nothing else
    End If
    
     ' dragging from the dock for deletion
    If (outsideDock = True And dragFromDockOperating = True) Then
        If animateTimer = False Then Call startAnimating
        hourGlassTimer.Enabled = False
        dragToDockOperating = False
        Exit Sub
    End If
    
    ' the mouse is now within the icon area or being dragged so start animating and using cpu...
    If insideDockFlg = True Or dragFromDockOperating = True Or dragToDockOperating = True Then
        If animateTimer = False Then Call startAnimating
        If dragFromDockOperating = True Then
            hourGlassTimer.Enabled = True
            dragToDockOperating = True
        End If
    End If
   
   On Error GoTo 0
   Exit Sub

responseTimer_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure responseTimer of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fSetDockUpperHeightLimit
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function fSetDockUpperHeightLimit() As Long
    Dim dockHeightPxls As Long: dockHeightPxls = 0
     
    On Error GoTo fSetDockUpperHeightLimit_Error

    dockHeightPxls = 0
    ' 22/10/2020 .01 frmMain.frm responsetimer fix the incorrect check of the timer state to determine the dock upper limit when entering and triggering the main animation
    If animatedIconsRaised = True Then
        dockHeightPxls = iconSizeLargePxls + rDvOffset + rdDefaultYPos
    Else
        dockHeightPxls = iconSizeSmallPxls + rDvOffset + rdDefaultYPos
    End If
    
    fSetDockUpperHeightLimit = dockHeightPxls

    On Error GoTo 0
    Exit Function

fSetDockUpperHeightLimit_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fSetDockUpperHeightLimit of Form dock"
            Resume Next
          End If
    End With

End Function

'---------------------------------------------------------------------------------------
' Procedure : tuneResponseTimerInterval
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tuneResponseTimerInterval()
    
    ' .18 STARTS DAEB frmMain.frm 31/01/2021 reinstated checks of fade out and slide timers to set a more frequent response timer to improve animation
    On Error GoTo tuneResponseTimerInterval_Error

    If animatedIconsRaised = True Or autoFadeOutTimer.Enabled = True Or autoSlideOutTimer.Enabled = True Then ' logic to test on states needs to be refined
        responseTimer.Interval = 5 ' tests the mouse position more regularly, making dock much more responsive and fadeouts smoother
    Else
        responseTimer.Interval = 200 ' reduced to 5 times per second
    End If

    On Error GoTo 0
    Exit Sub

tuneResponseTimerInterval_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tuneResponseTimerInterval of Form dock"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : defineDockEntranceLocation
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub defineDockEntranceLocation(ByVal dockHeightPxls As Long)

    ' set the area of the screen that responds to the cursor entering to be only 10 pixels from the bottom of the screen
    ' it does this on a slide out and the instant options only, giving more room to display other apps without the dock interfering
    ' And Val(sDAutoHideType) <> 0
    
    ' .11 DAEB changed the setting of the dock top to a better place
    On Error GoTo defineDockEntranceLocation_Error

    If Not (rDAutoHide = "1" And dockHidden = True) Then
        currentDockTopPxls = (Me.Height / screenTwipsPerPixelY) - dockHeightPxls  ' sets the dock top to normal position
    End If

    On Error GoTo 0
    Exit Sub

defineDockEntranceLocation_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure defineDockEntranceLocation of Form dock"
            Resume Next
          End If
    End With
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fTestCursorWithinDockYPosition
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function fTestCursorWithinDockYPosition(ByVal dockHeightPxls As Long) As Boolean
    Dim outsideDock  As Boolean
    On Error GoTo fTestCursorWithinDockYPosition_Error

    outsideDock = False
    
    ' checks the mouse Y position - ie. is the mouse outside the vertical/horizontal dock area
    If dockPosition = vbbottom Then
        outsideDock = apiMouse.Y < currentDockTopPxls Or apiMouse.X < iconLeftmostPointPxls Or apiMouse.X > iconStoreLeftPixels(UBound(iconStoreLeftPixels))    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    End If
    If dockPosition = vbtop Then
        outsideDock = apiMouse.Y > dockHeightPxls Or apiMouse.X < iconLeftmostPointPxls Or apiMouse.X > iconStoreLeftPixels(UBound(iconStoreLeftPixels)) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    End If
    
    fTestCursorWithinDockYPosition = outsideDock ' return

    On Error GoTo 0
    Exit Function

fTestCursorWithinDockYPosition_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTestCursorWithinDockYPosition of Form dock"
            Resume Next
          End If
    End With
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : startHidingDockTimers
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startHidingDockTimers()

    On Error GoTo startHidingDockTimers_Error

        If rDAutoHide = "1" Then ' if hidden dynamically restore the opacity when hoverered upon
            If Val(sDAutoHideType) = 0 Then ' fade animation
                If funcBlend32bpp.SourceConstantAlpha < 255 Then ' fade back in
                    autoHideMode = "fadein"
                    autoFadeInTimer.Enabled = True
                    bDrawn = False
                    smallDockBeenDrawn = False ' allows the dock to redraw on the next response cycle
                End If
            ElseIf Val(sDAutoHideType) = 1 Then ' slide animation as per Rocketdock
                ' check whether the dock has been slid out already
                If iconCurrentTopPxls > (screenHeightPixels) Then
                    autoSlideMode = "slidein"
                    autoSlideInTimer.Enabled = True
                    bDrawn = False
                    smallDockBeenDrawn = False ' allows the dock to redraw on the next response cycle
                End If
            ElseIf Val(sDAutoHideType) = 2 Then 'instant invisible
                ' set the opacity of the whole dock, used to display solidly and for instant autohide
                funcBlend32bpp.SourceConstantAlpha = 255 * Val(dockOpacity) / 100
                bDrawn = False
                smallDockBeenDrawn = False ' allows the dock to redraw on the next response cycle
            End If

            dockHidden = False
        End If

    On Error GoTo 0
    Exit Sub

startHidingDockTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startHidingDockTimers of Form dock"
            Resume Next
          End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : startAnimating
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startAnimating()
    On Error GoTo startAnimating_Error

        animatedIconsRaised = True
        dockZorder = "high"
        dockOpacity = Val(rDIconOpacity) ' the default opacity for the icons
        smallDockBeenDrawn = False
        
        animateTimer.Enabled = True
       
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        
        '.nn Set the cursor to a pointer, if for some reason it has been set to anything other than a normal pointy cursor
        lngCursor = LoadCursor(0, 32512&)
        If (lngCursor > 0) Then SetCursor lngCursor

        Call startHidingDockTimers

    On Error GoTo 0
    Exit Sub

startAnimating_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startAnimating of Form dock"
            Resume Next
          End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : stopAnimating
' Author    : beededea
' Date      : 19/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopAnimating()
        ' only cancel the running of the animation timer when neither of the two bounce timers are running
        ' this allows the bouncetimers to complete even if the mouse moves out of the dock area.
        ' these timers control the initiation of the chosen application so it is important that they both complete.
    On Error GoTo stopAnimating_Error

        dockJustEntered = True
        dragToDockOperating = False ' stops the middle invisible icon on the sequentialBubbleAnimation routine

        If animatedIconsRaised = False Then
            If smallDockBeenDrawn = False Then ' only draws the dock when it has not yet been drawn
                If Val(rDHoverFX) = 1 Then Call selectBubbleAnimation(3) ' select drawSmallStaticIcons redraw the icons if dragged to the same position
            End If
            If animateTimer.Enabled = True Then
                
                If bounceUpTimer.Enabled = False Or bounceDownTimer.Enabled = False Then ' .80 DAEB 28/05/2021 frmMain.frm Keep the animateTimer and therefore the bounceTimers operating in order to run the chosen app.
                    animateTimer.Enabled = False ' stops the cpu costly animation timer
                End If
            End If
        Else ' if it was true
            animatedIconsRaised = False
            dockLoweredTime = TimeValue(Now)
        End If

    On Error GoTo 0
    Exit Sub

stopAnimating_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopAnimating of Form dock"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : animateTimer
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :   this is the X millisecond timer that does the animation for the dock icons
'               determines if and where exactly the mouse is in the < horizontal > icon hover area and determines the icon index
'               clears the whole previously drawn image section
'               calls the chosen animation method which animates
'---------------------------------------------------------------------------------------
'
Private Sub animateTimer_Timer()
    Dim showsmall As Boolean: showsmall = False
    Dim useloop As Integer: useloop = 0
    Dim thiskey As String: thiskey = vbNullString
    Dim textWidth As Integer: textWidth = 0
    Dim insideDock As Boolean: insideDock = False
   
    'Dim bumpFactor As Single' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
    'bumpFactor = 0' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
    
    On Error GoTo animateTimer_Error
    
    'lngReturn = GetCursorPos(apiMouse) ' not needed as it retruns a value in the response timer which is sufficient
    
    ' if the bounce or fade timere are running cause animation to continue even if the mouse is stationary.
    If bounceUpTimer.Enabled = True Or bounceDownTimer.Enabled = True Or hourGlassTimer.Enabled = True Or autoFadeOutTimer.Enabled = True Or autoFadeInTimer.Enabled = True Or autoSlideOutTimer.Enabled = True Or autoSlideInTimer.Enabled = True Then ' .nn Changed or added as part of the drag and drop functionality
        ' carry on as usual and animate when any animation timers are running
    Else ' we are only interested in analysing if there is any Y axis movement
        ' however, if the animate timers are not running and the cursor position is static then do no animation - just exit, saving CPU '
        If savApIMouseX = apiMouse.X And savApIMouseY = apiMouse.Y Then
            animateTimer.Enabled = False
            'animateTimer.Interval = 200
            'responseTimer.Interval = 200 ' nn
            Exit Sub             ' if the timer that does the bouncing is running then we need to animate even if the mouse is stationary...
        End If
        If savApIMouseX = apiMouse.X And savApIMouseY <> apiMouse.Y Then Exit Sub ' if moving in the x axis but not in the y axis we also exit
    End If

    savApIMouseY = apiMouse.Y
    savApIMouseX = apiMouse.X
    
    showsmall = True
    bDrawn = False
    expandedDockWidth = 0
        
    ' determines if and where exactly the mouse is in the < horizontal > icon hover area and if so, determine the icon index
    For useloop = 0 To iconArrayUpperBound
        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
        insideDock = apiMouse.X >= iconStoreLeftPixels(useloop) And apiMouse.X <= iconStoreRightPixels(useloop)
    
        If insideDock Then
                iconIndex = useloop ' this is the current icon number being hovered over
                Exit For ' as soon as we have the index we no longer have to stay in the loop
        End If
    Next useloop
    
    iconPosLeftPxls = iconLeftmostPointPxls ' put starting left position back again for the dock bg
    
' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
'    If usedMenuFlag = False Then ' only recalculate sizeModifierPxls for the bump animation when the menu has not recently been used
'        'sizeModifierPxls is the variance from one side of the 'main' icon to the cursor point that is applied to the icons either side in order to resize them
'         ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        sizeModifierPxls = ((apiMouse.x) - iconStoreLeftPixels(iconIndex)) / (bumpFactor)
'        'sizeModifierPxls = ((apiMouse.x * screenTwipsPerPixelX) - iconPosLeftTwips(iconIndex)) / (bumpFactor * screenTwipsPerPixelX)
'    Else
'        usedMenuFlag = False ' the menu causes the mouse to move far away from the icon centre and so icon sizing was massive
'    End If
    
        
    ' NOTE:
    ' if it is the first time the dock is entered then it is drawDockByCursorEntryPosition that draws all the icons into the correct location.
    ' when the icons have been ordered correctly then sequentialBubbleAnimation provides the animation from that point on.
    
    If dockJustEntered = True Then
        If Val(rDHoverFX) = 1 Then Call selectBubbleAnimation(2) ' select drawDockByCursorEntryPosition - finds horizontal start point for the dock and place icons accordingly
        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
        saveStartLeftPxls = iconStoreLeftPixels(0) ' we now have the dock start position for sequentialBubbleAnimation to do its stuff
        dockJustEntered = False
    Else
        'none
        'bubble
        'plateau
        'flat
        'bumpy
    
        ' 27/10/2020 .04 DAEB alternative animations to zoom: Bubble enabled as options STARTS
        If Val(rDHoverFX) = 0 Then
            ' the none choice, simply bounces the small icon without growing it at all
        ElseIf Val(rDHoverFX) = 1 Then
            Call selectBubbleAnimation(1) ' select sequentialBubbleAnimation
        ElseIf Val(rDHoverFX) = 2 Then
            'Call sequentialBubbleAnimation ' the current zoom: Bubble animation
            ' the zoom plateau animation, as per the current animation makes n number of central icons assume the full size
        ElseIf Val(rDHoverFX) = 3 Then
            ' the zoom flat animation all are shown in their large mode and the mouse scrolls from right to left according to mouse position
        ElseIf Val(rDHoverFX) = 4 Then
            'Call sequentialBubbleAnimation ' the current zoom: Bubble animation
        End If
        ' 27/10/2020 .04 DAEB alternative animations to zoom: Bubble enabled as options ENDS.
    End If
    
    'stores the current icon position
    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    iconStoreLeftPixels(UBound(iconStoreLeftPixels)) = iconPosLeftPxls
                

   On Error GoTo 0
   Exit Sub

animateTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure animateTimer of Form dock"

End Sub

Private Sub updateGDIPlus()

    Call GdipDeleteGraphics(lngImage)  'The graphics may now be deleted
    
    ' the third parameter is a pointer to a structure that specifies the new screen position of the layered window.
    ' If the current position is not changing, pptDst can be NULL. It is null. We can play with it to move the screen
    
    'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
    UpdateLayeredWindow Me.hWnd, hdcScreen, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA

End Sub
'
'Private Function Draw(Func As String) As Integer
'  Dim i As Integer: i = 0
'  Dim sinwave As Integer: sinwave = 0
'  Dim cen As Double: cen = 0
'
'  On Error Resume Next 'only for tan
'
'  Const pi = 3.14159265358979
'  sinwave = 0
'
'    Select Case Func
'        Case "sin"
'             sinwave = Sin(i * pi / 720)
'        Case "cos"
'        Case "tan"
'    End Select
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : sequentialBubbleAnimation
' Author    : beededea
' Date      : 01/05/2020
' Purpose   : sequentialBubbleAnimation is the main animator. It places the icons from left to right, storing the icon
'             positions in an array so the current chosen icon can be acted upon.
'             The previous drawDockByCursorEntryPosition places all the icons according to the position they find themselves in when the cursor enters the dock.
'             This routine simply takes those stored positions and then animates them sequentially from a to z
'---------------------------------------------------------------------------------------
'
Private Sub sequentialBubbleAnimation()
 
    Dim showsmall As Boolean: showsmall = False
    Dim useloop As Integer: useloop = 0
    Dim useloop2 As Integer: useloop2 = 0
    Dim thiskey As String: thiskey = ""
    Dim thiskey2 As String: thiskey2 = ""
    Dim textWidth As Integer: textWidth = 0
    Dim dockSkinStart As Long: dockSkinStart = 0
    Dim dockSkinWidth As Long: dockSkinWidth = 0
    Dim leftGrpMember As Integer: leftGrpMember = 0
    Dim leftmostResizedIcon As Integer: leftmostResizedIcon = 0
    Dim rightmostResizedIcon As Integer: rightmostResizedIcon = 0
    Dim bumpFactor As Single: bumpFactor = 0
    
    On Error GoTo sequentialBubbleAnimation_Error
    
    DeleteObject bmpMemory ' the bitmap deleted
    Call readyGDIPlus ' clears the whole previously drawn image section and the animation continues

    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecessary twip conversion
    iconPosLeftPxls = saveStartLeftPxls

    If rDtheme <> vbNullString And rDtheme <> "Blank" Then Call applyThemeSkinToDock(dockSkinStart, dockSkinWidth)
    
    Call determineIconRangeToAnimate(leftmostResizedIcon, rightmostResizedIcon)
    
    ' .61 DAEB 26/04/2021 frmMain.frm size modifier moved to the sequential bump animation
    bumpFactor = 1.2 ' this determines the bumpiness of the animation, change at your peril
    If usedMenuFlag = False Then ' only recalculate sizeModifierPxls for the bump animation when the menu has not recently been used
         ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
        sizeModifierPxls = ((apiMouse.X) - iconStoreLeftPixels(iconIndex)) / (bumpFactor)
    Else
        usedMenuFlag = False ' the menu causes the mouse to move far away from the icon centre and so icon sizing was massive
    End If

    For useloop = 0 To iconArrayUpperBound ' loop through all the icons one by one
        'size icon 0 to very small
        
        'Call sizePositionZero(useloop, showsmall)
        
        ' small icons to the left shown in small mode
        Call sizeEachSmallIconToLeft(useloop, leftmostResizedIcon, showsmall)

        ' the group of icons to the left of the main icon, resized dynamically
        Call sizeEachResizedIconToLeft(useloop, leftmostResizedIcon, showsmall)

         ' the main fullsize icon
        Call sizeFullSizeIcon(useloop, showsmall)

        ' the group of icons to the right of the main icon, resized dynamically
        Call sizeEachResizedIconToRight(useloop, rightmostResizedIcon, showsmall)

        ' small icons to the right shown in small mode
        Call sizeEachSmallIconToRight(useloop, rightmostResizedIcon, showsmall)
       
        ' display the icon in the dock
        If showsmall = True Then ' display the small size icon or the red X if icon missing
            Call showSmallIcon(useloop)
        Else
            Call showLargeIconTypes(useloop) ' display the larger size icon or the
        End If

        'now draw the icon text above the selected icon
        Call drawTextAboveIcon(useloop, textWidth)
        
        ' store the icon current position in the array
        Call storeCurrentIconPositions(useloop)

        iconPosLeftPxls = iconPosLeftPxls + iconWidthPxls
   
    Next useloop

    ' .nn Changed or added as part of the drag and drop functionality
    ' 12/05/2021 .nn DAEB Displays a smaller size icon at the cursor position when a drag from the dock is underway.
    If dragFromDockOperating = True Then
        updateDisplayFromDictionary collLargeIcons, vbNullString, dragImageToDisplay, (apiMouse.X - iconSizeLargePxls / 2), (apiMouse.Y - iconSizeLargePxls / 2), (iconSizeLargePxls * 0.75), (iconSizeLargePxls * 0.75)
    End If
    
    Call updateGDIPlus
    
'    If debugflg = 1 Then
'       DrawTheText "animateTimer.Enabled " & animateTimer.Enabled, 200, 50, 300, rDFontName, Val(Abs(rDFontSize))
'        DrawTheText "bounceHeight " & bounceHeight, 580, 50, 300, rDFontName, Val(Abs(rDFontSize))
'    End If
   On Error GoTo 0
   Exit Sub

sequentialBubbleAnimation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sequentialBubbleAnimation of Form dock"

End Sub


Private Sub determineIconRangeToAnimate(ByRef leftmostResizedIcon As Integer, ByRef rightmostResizedIcon As Integer)
    
    rDZoomWidth = 2 'override until the animation takes this into account
    If CBool(rDZoomWidth And 1) = False Then
        rDZoomWidth = rDZoomWidth + 1  ' must be 3,5,7,9 so convert to an odd number
    End If
     
    ' what is the group size? extract the index of the group and calculate the leftmost member
    leftmostResizedIcon = iconIndex - (rDZoomWidth - 1) / 2 '
    rightmostResizedIcon = iconIndex + (rDZoomWidth - 1) / 2
End Sub
Private Sub storeCurrentIconPositions(useloop)
        
        iconStoreLeftPixels(useloop) = Int(iconPosLeftPxls)
        iconStoreRightPixels(useloop) = Int(iconStoreLeftPixels(useloop) + iconWidthPxls) ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
        iconStoreTopPixels(useloop) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
        'iconStoreBottomPixels(useloop) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
End Sub

Private Sub sizePositionZero(ByVal useloop As Integer, ByRef showsmall As Boolean)
                
        If useloop = 0 Then 'small icons to the left shown in small mode
            iconHeightPxls = iconSizeSmallPxls
            iconWidthPxls = iconSizeSmallPxls

            If dockPosition = vbbottom Then
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) + xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) - xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                Else
                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
                    iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
                    iconCurrentBottomPxls = dockUpperMostPxls + iconSizeLargePxls ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                End If
            End If
            
            If dockPosition = vbtop Then
                
                ' NOTE: everything is inverted...
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier '.nn added the slidein/out
                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                Else
                    iconCurrentTopPxls = dockUpperMostPxls '.48 DAEB 01/04/2021 frmMain.frm  removed the vertical adjustment already applied to iconCurrentTopPxls
                End If
            End If

            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeSmallPxls
            showsmall = True
            expandedDockWidth = expandedDockWidth + iconWidthPxls
        End If
End Sub


Private Sub sizeEachSmallIconToLeft(ByVal useloop As Integer, ByVal leftmostResizedIcon As Integer, ByRef showsmall As Boolean)
                
        If useloop < leftmostResizedIcon Then  'small icons to the left shown in small mode
            iconHeightPxls = iconSizeSmallPxls
            iconWidthPxls = iconSizeSmallPxls

            If dockPosition = vbbottom Then
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) + xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) - xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                Else
                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
                    iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
                    iconCurrentBottomPxls = dockUpperMostPxls + iconSizeLargePxls ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                End If
            End If
            
            If dockPosition = vbtop Then
                
                ' NOTE: everything is inverted...
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier '.nn added the slidein/out
                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
                Else
                    iconCurrentTopPxls = dockUpperMostPxls '.48 DAEB 01/04/2021 frmMain.frm  removed the vertical adjustment already applied to iconCurrentTopPxls
                End If
            End If

            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeSmallPxls
            showsmall = True
            expandedDockWidth = expandedDockWidth + iconWidthPxls
        End If
End Sub

Private Sub sizeEachResizedIconToLeft(ByVal useloop As Integer, ByVal leftmostResizedIcon As Integer, ByRef showsmall As Boolean)

    Dim useloop2 As Integer: useloop2 = 0
    Dim resizeProportion As Double: resizeProportion = 0
        
    ' the group of icons to the left of the main icon, resized dynamically
    If useloop < iconIndex And useloop >= leftmostResizedIcon Then
       For useloop2 = leftmostResizedIcon To (iconIndex - 1)
            resizeProportion = 1 / ((rDZoomWidth - 1) / 2) ' 33, .50 &c
            resizeProportion = 1
            
'            If useloop = 0 Then
'                iconHeightPxls = iconSizeSmallPxls - (sizeModifierPxls * resizeProportion) 'sizeModifierPxls is the difference from the midpoint of the current icon in the x axis
'                iconWidthPxls = iconSizeSmallPxls - (sizeModifierPxls * resizeProportion)
'            Else
                iconHeightPxls = iconSizeLargePxls - (sizeModifierPxls * resizeProportion) 'sizeModifierPxls is the difference from the midpoint of the current icon in the x axis
                iconWidthPxls = iconSizeLargePxls - (sizeModifierPxls * resizeProportion)
'            End If
              
             If dockPosition = vbbottom Then
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) + xAxisModifier '.nn
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) - xAxisModifier '.nn
                Else
                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) '.nn
                End If
             End If
             
            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
                Else
                    iconCurrentTopPxls = dockUpperMostPxls
                End If
            End If
            
             'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)
            
             expandedDockWidth = expandedDockWidth + iconWidthPxls
             showsmall = False

        Next useloop2
    End If
End Sub


Private Sub sizeFullSizeIcon(ByVal useloop As Integer, ByRef showsmall As Boolean)
         ' the main fullsize icon
                     
        If useloop = iconIndex Then
'            If useloop = 0 Then
'                iconHeightPxls = iconSizeSmallPxls
'                iconWidthPxls = iconSizeSmallPxls
'            Else
                iconHeightPxls = iconSizeLargePxls
                iconWidthPxls = iconSizeLargePxls
'            End If
            
            If dockPosition = vbbottom Then
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = dockUpperMostPxls + xAxisModifier
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = dockUpperMostPxls - xAxisModifier
                Else
                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
                    iconCurrentTopPxls = dockUpperMostPxls
                End If
                
                If selectedIconIndex = iconIndex Then iconCurrentTopPxls = iconCurrentTopPxls - bounceHeight
            End If
            
            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
                
                '.nn added the slidein/out
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
                Else
                    iconCurrentTopPxls = dockUpperMostPxls
                End If
                
                If selectedIconIndex = iconIndex Then iconCurrentTopPxls = dockUpperMostPxls + bounceHeight
            End If
        
            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeLargePxls
'            If useloop = 0 Then
                showsmall = False
'            Else
'                showsmall = False
'            End If
            expandedDockWidth = expandedDockWidth + (iconWidthPxls)
    End If
End Sub
Private Sub sizeEachResizedIconToRight(ByVal useloop As Integer, ByVal rightmostResizedIcon As Integer, ByRef showsmall As Boolean)
    If useloop > iconIndex And useloop <= rightmostResizedIcon Then
    
        iconHeightPxls = iconSizeSmallPxls + sizeModifierPxls
        iconWidthPxls = iconSizeSmallPxls + sizeModifierPxls
    
        If dockPosition = vbbottom Then
            
            If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)) + xAxisModifier
            ElseIf autoSlideMode = "slidein" Then
                iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)) - xAxisModifier
            Else
                ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
                iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls))
            End If
            'If selectedIconIndex = iconIndex + 1 Then iconCurrentTopPxls = iconCurrentTopPxls - bounceHeight
        End If
        
        If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
            
            '.nn added the slidein/out
            If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
            ElseIf autoSlideMode = "slidein" Then
                iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
            Else
                iconCurrentTopPxls = dockUpperMostPxls
            End If
        End If
        
        'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)
        expandedDockWidth = expandedDockWidth + iconWidthPxls
        showsmall = False
    End If
End Sub

Private Sub sizeEachSmallIconToRight(ByVal useloop As Integer, ByVal rightmostResizedIcon As Integer, ByRef showsmall As Boolean)
            
        If useloop > rightmostResizedIcon Then 'small icons to the right
            iconHeightPxls = iconSizeSmallPxls
            iconWidthPxls = iconSizeSmallPxls

            If dockPosition = vbbottom Then
                
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
                Else
                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls))
                End If
            End If
            
            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
                
                '.nn added the slidein/out
                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
                ElseIf autoSlideMode = "slidein" Then
                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
                Else
                    iconCurrentTopPxls = dockUpperMostPxls
                End If
            End If

            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeSmallPxls
            expandedDockWidth = expandedDockWidth + iconWidthPxls
            showsmall = True
        End If
End Sub


Private Sub showSmallIcon(ByVal useloop As Integer)
    Dim thiskey As String: thiskey = ""
        
    thiskey = useloop & "ResizedImg" & LTrim$(Str$(iconSizeSmallPxls))
    updateDisplayFromDictionary collSmallIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
    If rDShowRunning = "1" Then
        If processCheckArray(useloop) = True Then
            'updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
            If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
            If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
         End If
    End If
    ' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
    If targetExistsArray(useloop) = 1 Then ' redxResizedImg64
        If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
        If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2)
    End If
End Sub

Private Sub showLargeIconTypes(ByVal useloop As Integer)
    Dim thiskey As String: thiskey = ""
         
    thiskey = useloop & "ResizedImg" & LTrim$(Str$(iconSizeLargePxls))
    ' add a 1% opaque background to the expanded image to catch click-throughs, blankresizedImg128 is the key name
    updateDisplayFromDictionary collLargeIcons, vbNullString, "blankresizedImg128", (iconPosLeftPxls), (iconCurrentTopPxls), (128), (128)

    ' .56 DAEB 19/04/2021 frmMain.frm Added a faded red background to the current image when the drag and drop is in operation.
    If dragToDockOperating = True And useloop = iconIndex Then
        updateDisplayFromDictionary collLargeIcons, vbNullString, "redresizedImg256", (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
    End If
    
    ' the current image itself always displays
    updateDisplayFromDictionary collLargeIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
                         
    ' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
    If dragToDockOperating = True And useloop = iconIndex Then
        If hourglassimage = vbNullString Then hourglassimage = "hourglass1resizedImg128"
        updateDisplayFromDictionary collLargeIcons, vbNullString, hourglassimage, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
    End If
    
    If rDShowRunning = "1" Then
        If processCheckArray(useloop) = True Then
            '                                                           thisCollection, strFilename,  key,                       Left,                                            Top,                                             Width,               Height
            If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls - (iconSizeLargePxls / 5)), (iconWidthPxls), (iconHeightPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
            If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls + (iconSizeLargePxls / 5)), (iconWidthPxls), (iconHeightPxls)
            
            'updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls - (iconSizeLargePxls / 5)), (iconSizeLargePxls), (iconSizeLargePxls)
        End If
    End If
    ' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
    If targetExistsArray(useloop) = 1 Then ' redxResizedImg64
            If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls - (iconSizeLargePxls / 5)), (iconWidthPxls / 2), (iconHeightPxls / 2) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
            If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls + (iconSizeLargePxls / 5)), (iconWidthPxls / 2), (iconHeightPxls / 2)
    End If
End Sub

Private Sub drawTextAboveIcon(ByVal useloop As Integer, ByVal textWidth As Integer)
        
        If useloop = iconIndex Then ' this section is located here to ensure the text is above the icon image
            'now draw the icon text above the selected icon
            If rDHideLabels = "0" Then
                
                If Not namesListArray(iconIndex) = "Separator" Then
                    textWidth = iconSizeLargePxls
                    If dockPosition = vbtop Then
                        DrawTheText namesListArray(iconIndex), iconCurrentTopPxls + iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
                    ElseIf dockPosition = vbbottom Then
                        ' puts the text 10% +10 px above the icon
                        DrawTheText namesListArray(iconIndex), (screenBottomPxls - ((iconSizeLargePxls / 10) + 40)) - iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
                        'DrawTheText textToDisplay, (screenBottomPxls - ((iconSizeLargePxls / 10) + 40)) - iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
                    End If
                End If
            End If
        End If


End Sub

'---------------------------------------------------------------------------------------
' Procedure : drawDockByCursorEntryPosition
' Author    : beededea
' Date      : 01/05/2020
' Purpose   : draws the icons once starting with the current MAIN icon and then positioning the others to the right or left of this first entry point icon.
'             This runs just ONCE before each animation period. The main function is to determine the leftmost position of the dock
'             in relation to the current icon highlighted. This is important as when one of the left or rightmost icons is selected
'             a sequential drawing of the icons might place the chosen icon off screen. We want to avoid that by focussing on the main icon.
'---------------------------------------------------------------------------------------
'
Private Sub drawDockByCursorEntryPosition()
    Dim showsmall As Boolean: showsmall = False
    Dim textWidth As Integer: textWidth = 0
    Dim leftmostResizedIcon As Integer: leftmostResizedIcon = 0
    Dim rightmostResizedIcon As Integer: rightmostResizedIcon = 0
    
    Dim useloop As Integer: useloop = 0
    Dim rightIconWidthPxls As Integer: rightIconWidthPxls = 0
    Dim mainIconWidthPxls  As Integer: mainIconWidthPxls = 0
    Dim thiskey As String: thiskey = vbNullString
    Dim dockSkinStart As Long: dockSkinStart = 0
    Dim dockSkinWidth As Long: dockSkinWidth = 0
    Dim offsetPxls As Integer: offsetPxls = 0
    Dim offsetProportion As Double: offsetProportion = 0
    
    On Error GoTo drawDockByCursorEntryPosition_Error
    'If debugflg = 1 Then debugLog "%drawDockByCursorEntryPosition"
    
    ' the small icon dock placement is inevitably incorrect at this point as the left most position of the dock, icon one,
    ' has not yet been calculated. However the code to theme the dock needs to be placed here as it is drawn first before the dock icons are drawn.
    ' this will be replaced by an animation timer that redraws the dock from the old to the current size.
    
    DeleteObject bmpMemory ' the bitmap deleted
    Call readyGDIPlus ' clears the whole previously drawn image section and the animation continues
    
    If rDtheme <> vbNullString And rDtheme <> "Blank" Then Call applyThemeSkinToDock(dockSkinStart, dockSkinWidth)
    
    Call determineIconRangeToAnimate(leftmostResizedIcon, rightmostResizedIcon)

    ' the main fullsize icon
    Call sizeAndShowFullSizeIconByCEP(iconIndex, showsmall)
    mainIconWidthPxls = iconWidthPxls
    
    ' what should be the group of icons to the left of the main icon, resized dynamically, currently caters only for one
    Call sizeAndShowSingleMainIconToLeftByCEP(iconIndex, leftmostResizedIcon, showsmall)

    ' what should be the group of icons to the right of the main icon, resized dynamically, currently caters only for one
    Call sizeAndShowSingleMainIconToRightByCEP(iconIndex, rightmostResizedIcon, mainIconWidthPxls, showsmall)
    rightIconWidthPxls = iconWidthPxls

    ' small icons to the left shown in small mode
    Call sizeAndShowSmallIconsToLeftByCEP(iconIndex, leftmostResizedIcon, showsmall)

    ' small icons to the right shown in small mode
    Call sizeAndShowSmallIconsToRightByCEP(iconIndex, rightmostResizedIcon, rightIconWidthPxls, showsmall)
   
    Call updateGDIPlus
   
   On Error GoTo 0
   Exit Sub

drawDockByCursorEntryPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawDockByCursorEntryPosition of Form dock"
End Sub
Private Sub sizeAndShowFullSizeIconByCEP(ByVal thisIconIndex As Integer, ByRef showsmall As Boolean)

    Dim mainIconWidthPxls  As Integer: mainIconWidthPxls = 0
    Dim textWidth As Integer: textWidth = 0

    '===================
    ' the main fullsize icon
    '==================
    iconHeightPxls = iconSizeLargePxls
    iconWidthPxls = iconSizeLargePxls
    mainIconWidthPxls = iconWidthPxls
    
    Call sizeFullSizeIcon(thisIconIndex, showsmall)

    ' the following two lines  position the main icon initially to the main icon's leftmost start point when small
    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    iconPosLeftPxls = iconStoreLeftPixels(iconIndex)
    
    Call storeCurrentIconPositions(thisIconIndex)
    
    ' display the icon in the dock
    Call showLargeIconTypes(thisIconIndex) ' display the larger size icon or the

    'now draw the icon text above the selected icon
    Call drawTextAboveIcon(thisIconIndex, textWidth)
    
End Sub
Private Sub sizeAndShowSingleMainIconToLeftByCEP(ByVal thisIconIndex As Integer, ByVal leftmostResizedIcon As Integer, ByRef showsmall As Boolean)

    '===================
    ' one icon to the left, resized dynamically
    '==================
    If thisIconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
        
        ' the icon to the left is currently sized full as the other on the right hand side is sized small.
        iconHeightPxls = iconSizeLargePxls
        iconWidthPxls = iconSizeLargePxls

        If dockPosition = vbbottom Then
            iconCurrentTopPxls = dockUpperMostPxls
        End If

        If dockPosition = vbtop Then
           iconCurrentTopPxls = dockUpperMostPxls
        End If

        iconPosLeftPxls = iconPosLeftPxls - iconWidthPxls
        
        Call storeCurrentIconPositions(thisIconIndex - 1)
        
        ' display the icon in the dock
        Call showLargeIconTypes(thisIconIndex - 1)
    End If

     ' iconLeftmostPointPxls = iconPosLeftPxls
End Sub

Private Sub sizeAndShowSingleMainIconToRightByCEP(ByVal thisIconIndex As Integer, ByVal rightmostResizedIcon As Integer, ByVal mainIconWidthPxls As Integer, ByRef showsmall As Boolean)
    Dim rightIconWidthPxls As Integer: rightIconWidthPxls = 0

    '===================
    ' one icon to the right, resized dynamically
    '==================
   If thisIconIndex <= rightmostResizedIcon And thisIconIndex < rdIconMaximum Then  '    If iconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
        
        ' the icon to the left is currently sized in small mode as the other on the left hand side is sized in full.
        iconHeightPxls = iconSizeSmallPxls
        iconWidthPxls = iconSizeSmallPxls

        rightIconWidthPxls = iconWidthPxls
         
        If dockPosition = vbbottom Then
            iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls) '.nn removal of sizeModifierPxls
        End If

        If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
            iconCurrentTopPxls = dockUpperMostPxls
        End If
        
        iconPosLeftPxls = (iconStoreLeftPixels(thisIconIndex)) + mainIconWidthPxls

        Call storeCurrentIconPositions(thisIconIndex + 1)

        Call showLargeIconTypes(thisIconIndex + 1)

    End If
End Sub

Private Sub sizeAndShowSmallIconsToLeftByCEP(ByVal thisIconIndex As Integer, ByRef leftmostResizedIcon As Integer, ByRef showsmall As Boolean)
    Dim leftLoop As Integer: leftLoop = 0
    Dim thiskey As String: thiskey = vbNullString

    ' all icons to the left
    '==================
    If thisIconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
        iconPosLeftPxls = iconStoreLeftPixels(thisIconIndex - 1)

        For leftLoop = thisIconIndex - 2 To 0 Step -1
'            iconHeightPxls = iconSizeSmallPxls
'            iconWidthPxls = iconSizeSmallPxls
'
'            If dockPosition = vbbottom Then
'                ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
'        ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'                iconCurrentTopPxls = dockUpperMostPxls
'            End If

            ' small icons to the left shown in small mode
            Call sizeEachSmallIconToLeft(leftLoop, leftmostResizedIcon, showsmall)

            iconPosLeftPxls = iconPosLeftPxls - iconWidthPxls
'            iconStoreLeftPixels(leftLoop) = iconPosLeftPxls ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'            iconStoreRightPixels(leftLoop) = iconStoreLeftPixels(leftLoop) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon             ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'            iconStoreTopPixels(leftLoop) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon

            Call storeCurrentIconPositions(leftLoop)

            Call showSmallIcon(leftLoop)
        Next leftLoop
    End If
End Sub
Private Sub sizeAndShowSmallIconsToRightByCEP(ByVal thisIconIndex As Integer, ByRef rightmostResizedIcon As Integer, ByRef rightIconWidthPxls As Integer, ByRef showsmall As Boolean)
    Dim rightLoop As Integer: rightLoop = 0
    Dim thiskey As String: thiskey = vbNullString
    '====================
    ' icons to the right
    '====================
    If thisIconIndex < rdIconMaximum Then   'check it isn't trying to animate a non-existent icon after the last icon

        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
       iconPosLeftPxls = (iconStoreLeftPixels(iconIndex + 1)) + rightIconWidthPxls
       
       For rightLoop = thisIconIndex + 2 To iconArrayUpperBound

            Call sizeEachSmallIconToRight(rightLoop, rightmostResizedIcon, showsmall)

            iconPosLeftPxls = iconPosLeftPxls + iconWidthPxls
        
            Call storeCurrentIconPositions(rightLoop)

'            thiskey = rightLoop & "ResizedImg" & LTrim$(Str$(iconSizeSmallPxls))
'            updateDisplayFromDictionary collSmallIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
'            If rDShowRunning = "1" Then
'                If processCheckArray(rightLoop) = True Then
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
'                End If
'            End If
'            ' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
'            If targetExistsArray(rightLoop) = 1 Then  ' redxResizedImg64
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2)
'            End If

            Call showSmallIcon(rightLoop)
            
        Next rightLoop
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadTheImageIntoGDIPlus
' Author    : beededea
' Date      : 18/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub loadTheImageIntoGDIPlus(ByVal iconIndexToShow As Single)
    Dim thiskey As String: thiskey = vbNullString

    On Error GoTo loadTheImageIntoGDIPlus_Error

    thiskey = iconIndexToShow & "ResizedImg" & LTrim$(Str$(iconSizeLargePxls))
    ' add a 1% opaque background to the expanded image to catch click-throughs
    updateDisplayFromDictionary collLargeIcons, vbNullString, "blankresizedImg128", (iconPosLeftPxls), (iconCurrentTopPxls), (128), (128)
    ' the current image itself
    updateDisplayFromDictionary collLargeIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
    If rDShowRunning = "1" Then
        If processCheckArray(iconIndexToShow) = True Then
            'updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeLargePxls / 2) - 3), (iconCurrentTopPxls - (iconSizeLargePxls / 5)), (iconSizeLargePxls), (iconSizeLargePxls)
            If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)  '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
            If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
        End If
        
    End If

   On Error GoTo 0
   Exit Sub

loadTheImageIntoGDIPlus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadTheImageIntoGDIPlus of Form dock"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : drawTheLabel
' Author    : beededea
' Date      : 18/06/2020
' Purpose   : now draw the icon text above the selected icon
'---------------------------------------------------------------------------------------
'
Private Sub drawTheLabel(ByVal iconIndexToShow As Single)
    Dim textWidth As Integer: textWidth = 0
    
   On Error GoTo drawTheLabel_Error

    If rDHideLabels = "0" Then
        Dim textToDisplay As String
        textToDisplay = iconCurrentTopPxls
        If Not namesListArray(iconIndexToShow) = "Separator" Then
            textWidth = iconSizeLargePxls
            If dockPosition = vbtop Then
                'DrawTheText textToDisplay, iconCurrentTopPxls + iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
                DrawTheText namesListArray(iconIndexToShow), iconCurrentTopPxls + iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
            ElseIf dockPosition = vbbottom Then
                ' puts the text 10% +10 px above the icon
                ' .73 DAEB 11/05/2021 frmMain.frm  sngBottom renamed to screenBottomPxls
                DrawTheText namesListArray(iconIndexToShow), (screenBottomPxls - ((iconSizeLargePxls / 10) + 40)) - iconSizeLargePxls, iconPosLeftPxls, textWidth, rDFontName, Val(Abs(rDFontSize))
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

drawTheLabel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawTheLabel of Form dock"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : runCommand
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : this routine passes the user-defined command to execute to an 'execute' routine
'             dependant upon the type of command. Note this routine is executed on a timer
'---------------------------------------------------------------------------------------
' .53 DAEB 11/04/2021 frmMain.frm changed all occurrences of sCommand to thisCommand to attain more compatibility with rdIconConfigFrm menuRun_click
' .68 DAEB 05/05/2021 frmMain.frm cause the docksettings utility to reopen if it has already been initiated

Public Sub runCommand(ByVal runAction As String, ByVal commandOverride As String) ' added new parameter to allow override .68
    
    Dim testURL As String: testURL = vbNullString
    Dim validURL As Boolean: validURL = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim folderPath As String: folderPath = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    Dim processID As Long:  processID = 0
    Dim windowHwnd As Long:  windowHwnd = 0
    Dim SetTopMostWindow As Long: SetTopMostWindow = 0
    Dim CurrentForegroundThreadID As Long: CurrentForegroundThreadID = 0
    Dim NewForegroundThreadID As Long: NewForegroundThreadID = 0
    Dim lngRetVal As Long: lngRetVal = 0
    Dim rmessage As String: rmessage = vbNullString ' .19 DAEB frmMain.frm 02/02/2021 added sArguments field to the confirmation dialog
    Dim hTray As Long: hTray = 0 ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    Dim hOverflow As Long: hOverflow = 0 ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    Dim userprof As String: userprof = vbNullString
    Dim intShowCmd As Integer: intShowCmd = 0

    On Error GoTo runCommand_Error
    'If debugflg = 1 Then debugLog "%runCommand"
    
    If userLevel = vbNullString Then userLevel = "open"

    'by default read the selected icon's data and set the command to execute
    If commandOverride = vbNullString Then
        'Call readIconData(selectedIconIndex) '.nn DAEB 12/05/2021 frmMain.frm Moved from the runtimer as some of the data is required before the run begins
        thisCommand = sCommand
    Else
        ' .68 DAEB 05/05/2021 frmMain.frm cause the docksettings utility to reopen if it has already been initiated
        
        thisCommand = commandOverride 'Not using the icon in the dock but the over-ridden command provided as a parameter
        ' therefore we must zero all the parameters, set from the last icon read, to empty values
        sFilename = vbNullString
        sFileName2 = vbNullString
        sTitle = vbNullString
        sCommand = vbNullString
        sArguments = vbNullString
        sWorkingDirectory = vbNullString
        sShowCmd = "0"
        sOpenRunning = "0"
        sIsSeparator = "0"
        sUseContext = "0"
        sDockletFile = "0"
        sUseDialog = "0"
        sUseDialogAfter = "0"
        sQuickLaunch = "0"
    End If
    
    If sIsSeparator = "1" Then
        Exit Sub
    End If
    
    intShowCmd = sShowCmd
    If sShowCmd = "0" Then
        intShowCmd = 1
    End If

    hTray = FindWindow_NotifyTray() ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    hOverflow = FindWindow_NotifyOverflow() ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas

    ' bring an already running process to the fore and then exit
    If rDOpenRunning = "1" And forceRunNewAppFlag = False Or sOpenRunning = "1" Then

    ' when the index is 999 this means that the cursor has left the area of the selected icon and is now 'browsing' the
    ' rest of the dock. Normally, this could not happen and would not matter - but for the additional second app that has
    ' a delayed start it is a normal condition. In this case we do not want to attempt to run an already-opened application so
    ' we bypass the process checking of the array and do not add this application to the list of running apps.
    
        If selectedIconIndex <> 999 Then
            If processCheckArray(selectedIconIndex) = True Or commandOverride <> vbNullString Then
                'is the quick way to check process is already running
                'but we need to run IsRunning again to get the process PID
                If IsRunning(thisCommand, processID) Then ' it checks again that the process is still running, as the check process timer that populates the processCheckArray is too infrequent to be relied upon
                    
                    ' .22 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
                    
                    'windowHwnd = getWindowHWndForPid(processID) ' old method of enumerating all windows and find the associated pid of each, returning the hWnd of the window associated with the PID
                    
                    'The EnumWindows function is more reliable than calling the GetWindow function in a loop as we used to do.
                    'ie. An application that calls GetWindow to perform this task risks being caught in an infinite
                    'loop or referencing a handle to a window that has been destroyed.
                    
                    ' enumerate all windows and find the associated pid of each, returning the hWnd of the window associated with the given PID
                    Call fEnumWindows(processID)
                    windowHwnd = storeWindowHwnd
                    
                    ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas STARTS
                    ' if the hwnd is zero then a matching process has not been found, in this case search the systray
                    If windowHwnd = 0 Then
    
                        Me.Print "Tray Handle: 0x" & Hex(hTray)
                        isSysTray hTray, processID, windowHwnd
    
                        Me.Print "Overflow Handle: 0x" & Hex(hOverflow)
                        isSysTray hOverflow, processID, windowHwnd
                    End If
                    ' .33 DAEB 03/03/2021 frmMain.frm New systray code from DragokasENDS
                    
                     'GetWindowRect windowHwnd, pRect unused
    
                    ' Get the thread for the current window that is to fore now (the dock)
                    CurrentForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
                    
                    ' Get the thread ID for the window we are trying to bring to the fore
                    NewForegroundThreadID = GetWindowThreadProcessId(windowHwnd, ByVal 0&)
            
                    'AttachThreadInput is used to ensure SetForegroundWindow will work
                    'even if our application isn't currently the foreground window
                    
                    '(e.g. a minimised application running in the background)
                    If CurrentForegroundThreadID <> NewForegroundThreadID Then
                        ' Attach shared keyboard input to the thread we are raising
                        Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, True)
                        ' Make the raised window the foreground window.
                        
                        
                        If runAction = "back" Then
                            ' .38 DAEB 18/03/2021 frmMain.frm utilised SetActiveWindow to give window focus without bringing it to fore
                            
                            '    The SetActiveWindow function activates a window, but not if the application is in the background.
                            '    The window will be brought into the foreground (top of Z order) if the application is in the foreground when it sets the activation.
                            lngRetVal = SetActiveWindow(windowHwnd)
                        Else
                            '     Brings the thread that created the specified window into the foreground and activates the window. Keyboard input is
                            '     directed to the window, and various visual cues are changed for the user. The system assigns a slightly higher
                            '     priority to the thread that created the foreground window than it does to other threads.
                            lngRetVal = SetForegroundWindow(windowHwnd)
                        End If
                        
                        ' break the thread's attachment to the newly raised window, breaking the association
                        ' effectively passing control to the raised window.
                        Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, False)
                    Else
                       lngRetVal = SetForegroundWindow(windowHwnd) ' bring window to the fore
                    End If
                    
                    ' .22 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function ENDS
                    If lngRetVal <> 0 Then
                                          
                        If IsIconic(windowHwnd) Then
                            Call ShowWindow(windowHwnd, SW_RESTORE) ' if a minimised window, bring to fore as a standard window
                        ElseIf IsZoomed(windowHwnd) Then
                            Call ShowWindow(windowHwnd, SW_MINIMIZE) ' if a full size window, minimise
                        ElseIf (Not IsIconic(windowHwnd) And Not IsZoomed(windowHwnd)) Then ' a normal window
                            
                            ' .42 DAEB 03/03/2021 frmMain.frm To support new receive focus menu option
                            If runAction = "focus" Then
                                BringWindowToTop windowHwnd ' .39 DAEB 18/03/2021 frmMain.frm utilised BringWindowToTop instead of SetWindowPos & HWND_TOP as that was used by a C program that worked perfectly.
                                'SetWindowPos windowHwnd, HWND_TOP, 0, 0, 0, 0, SWP_ACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                            End If
                            
                            ' .42 DAEB 03/03/2021 frmMain.frm To support new receive focus menu option
                            If runAction = "back" Then
                                ' .40 DAEB 18/03/2021 frmMain.frm Added SWP_NOOWNERZORDER as an additional flag as that was used by a C program that worked perfectly, fixing the z-order position problems
                                SetWindowPos windowHwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER
                            End If
                            
                            If prevIconIndex <> selectedIconIndex Then ' .27 DAEB frmMain.frm 11/02/2021 now operates like the standard Windows dock on a click, minimising then restoring
                                
                                ' .34 DAEB frmMain.frm 08/02/2021  - commented out the extra unwanted ShowWindow(windowHwnd, SW_RESTORE)
                                ' bringing this window to the fore, not needed, SetForegroundWindow does the job already - the showWindow causes a z-order problem if it is included.
                                ''''Call ShowWindow(windowHwnd, SW_RESTORE) ' < do not comment back in, leave all the commands in this if...else section commented out
                                
                                ' lngRetVal = SetForegroundWindow(windowHwnd) ' trial bring window to the fore
                                'SetWindowPos windowHwnd, HWND_TOPMOST, pRect.Left, pRect.Top, 0, 0, SWP_NOSIZE' trial bring window to the fore
                                
                            Else ' if the icon clicked is the same as the one before then
                                If runAction <> "focus" And runAction <> "back" Then
                                    'Call ShowWindow(windowHwnd, SW_MINIMIZE)   ' minimise the window
                                    Call ShowWindowAsync(windowHwnd, SW_MINIMIZE) ' .41 DAEB 18/03/2021 frmMain.frm utilised ShowWindowAsync instead of ShowWindow as the C program utilised it and it seemed to make sense to do so too
                                End If
                            End If
                            
                            ' I was not able to obtain a handle of a window with focus as it never matched
                            ' the selected window. It seems that you cannot check whether the chosen window already has focus as the
                            ' second you click an icon on the dock, the dock itself
                            ' seems to acquire focus.
    
                        ' .26 DAEB frmMain.frm 10/02/2021 added test to check window state and alter it accordingly
     
                        End If
    
                        prevIconIndex = selectedIconIndex ' .27 DAEB frmMain.frm 11/02/2021 now operates like the standard Windows dock on a click, minimising then restoring
    
                        Exit Sub ' if the app can be switched to successfully then do nothing else
                    End If
        
                End If ' IsRunning(thisCommand, processID)
            End If ' processCheckArray(selectedIconIndex)
        End If ' 999
    End If ' rDOpenRunning = "1"
    
    forceRunNewAppFlag = False
    
    ' run the selected program
    If sUseDialog = "1" Then
        ' .19 DAEB frmMain.frm 02/02/2021 added sArguments field to the confirmation dialog
        ' .21 DAEB frmMain.frm 07/02/2021 slight improvement to the confirmation dialog
        rmessage = "Are you sure you wish to run the following command - " & sTitle & "?" & vbCr & thisCommand
        If sArguments <> vbNullString Then rmessage = rmessage & " " & sArguments
        answer = MsgBox(rmessage, vbYesNo)
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    ' contains "shutdown"
    If InStr(thisCommand, "shutdown.exe") <> 0 Then
        Call shellExecuteWithDialog(userLevel, Environ$("windir") & "\SYSTEM32\shutdown", sArguments, 0&, intShowCmd)
        Exit Sub
    End If
    
    ' is the target a URL?
    testURL = Left$(thisCommand, 3)
    If testURL = "htt" Or testURL = "www" Then
        validURL = True
        Call shellExecuteWithDialog(userLevel, thisCommand, vbNullString, vbNullString, intShowCmd)
        Exit Sub
    End If
                
' .37 DAEB 03/03/2021 frmMain.frm removed the individual references to a Windows class ID
'     instead we check whether the prefix ::{ exists and then we run explorer passing the correct CLSID
    
'    If thisCommand = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
'        '  my computer
'        Call shellCommand("explorer.exe /e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", intShowCmd)
'        Exit Sub
'    End If
'
'    If thisCommand = "::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" Then
'        ' network
'        Call shellCommand("explorer.exe /e,::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", intShowCmd)
'        Exit Sub
'    End If
'
'    If thisCommand = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then
'        ' network
'        Call shellCommand("explorer.exe /e,::{208D2C60-3AEA-1069-A2D7-08002B30309D}", intShowCmd)
'        Exit Sub
'    End If

'    'printer
'    If thisCommand = "::{2227a280-3aea-1069-a2de-08002b30309d}" Then
'        Call shellCommand("explorer.exe /e,::{2227a280-3aea-1069-a2de-08002b30309d}", intShowCmd)
'        Exit Sub
'    End If
' .37 DAEB 03/03/2021 frmMain.frm removed the individual references to a Windows class ID ENDS

    ' control panel     ' .44 DAEB 01/04/2021 frmMain.frm put the control panel reference back
    If thisCommand = "control" Then
        Call shellCommand("rundll32.exe shell32.dll,Control_RunDLL", intShowCmd)
        Exit Sub
    End If

    ' Rocketdock quit compatibility
    If thisCommand = "[Quit]" Then
        MsgBox "I am sure you don't really want me to quit Steamydock... test cancelled."
        Exit Sub
    End If
    ' Rocketdock settings compatibility
    If thisCommand = "[Settings]" Then
        Call menuForm.mnuDockSettings_Click
        Exit Sub
    End If
    ' Rocketdock settings compatibility
    If thisCommand = "[Icons]" Then
        'Call menuForm.mnuIconSettings_Click
        Exit Sub
    End If
    
    ' .35 DAEB 03/03/2021 frmMain.frm check whether the prefix command required to access a Windows class ID is present
    If InStr(thisCommand, "explorer.exe /e,::{") Then
        Call shellCommand(thisCommand, intShowCmd)
        Exit Sub
    End If
    
    ' .36 DAEB 03/03/2021 frmMain.frm check whether the prefix is present that indicates a Windows class ID is present
    ' this allows SD to act like Rocketdock which only needs the CLSID to operate eg. ::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}
    If InStr(thisCommand, "::{") Then
        Call shellCommand("explorer.exe /e," & thisCommand, intShowCmd)
        Exit Sub
    End If
    
'    If InStr(inputData, "[defaultDockLocation]") Then
'        s = Replace(inputData, "[defaultDockLocation]", sdAppPath)
'    End If
    
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If

    ' .37 DAEB 03/03/2021 frmMain.frm removed the individual references to a Windows class ID
'    ' program files folder
'    If thisCommand = "::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}" Then
'        Call shellCommand("explorer.exe /e,::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}", intShowCmd)
'        Exit Sub
'    End If

     ' applications And features
    If thisCommand = "appwiz.cpl" Then
        'If debugflg = 1 Then debugLog "Shell " & "rundll32.exe shell32.dll,Control_RunDLL " & thisCommand
        Call shellCommand("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl", intShowCmd)
        Exit Sub
    End If
    
    ' recycle bin
    If thisCommand = "[RecycleBin]" Then
        Call shellCommand("explorer.exe /e,::{645ff040-5081-101b-9f08-00aa002f954e}", intShowCmd)
        Exit Sub
    End If
    
    ' admin tools
    If InStr(thisCommand, ".msc") <> 0 Then
        If FExists(thisCommand) Then ' if the file exists and is valid - run it
            Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, vbNullString, intShowCmd)
            Exit Sub ' .89 DAEB 08/12/2022 frmMain.frm Fixed duplicate run of .msc files.
        Else
            folderPath = GetDirectory(thisCommand)  ' extract the default folder from the full path
            
            ' .45 DAEB 01/04/2021 frmMain.frm Changed the logic to remove the code around a folder path existing...
            If Not DirExists(folderPath) Then
                 ' if there is no folder path provided then attempt it on its own hoping that the windows PATH will find it
                On Error GoTo tryMSCFullPAth ' if it is in the path then it will run
                Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
                Exit Sub
tryMSCFullPAth:
                On Error GoTo runCommand_Error
                Call shellExecuteWithDialog(userLevel, Environ$("windir") & "\SYSTEM32\" & GetFileNameFromPath(thisCommand), sArguments, sWorkingDirectory, intShowCmd)
                Exit Sub
            End If
            
        End If
        'Exit Sub
    End If
    

    ' task manager
    If thisCommand = "taskmgr" Then
        Call shellExecuteWithDialog(userLevel, Environ$("windir") & "\SYSTEM32\taskmgr", 0&, 0&, intShowCmd)
        Exit Sub
    End If
    
    ' RocketdockEnhancedSettings.exe (the .NET version of this program)
    If GetFileNameFromPath(thisCommand) = "RocketdockEnhancedSettings.exe" Then
        Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
         Exit Sub
    End If

    ' bat files
    If ExtractSuffixWithDot(UCase$(thisCommand)) = ".BAT" Then
        'If debugflg = 1 Then debugLog "ShellExecute " & thisCommand
        thisCommand = """" & sCommand & """" ' put the command in quotes so it handles spaces in the path
        folderPath = GetDirectory(thisCommand)  ' extract the default folder from the batch full path
        If FExists(sCommand) Then
            Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
        Else
            ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
            MessageBox Me.hWnd, thisCommand & " - this batch file does not exist", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
            ' MsgBox (thisCommand & " - this batch file does not exist")
        End If
        Exit Sub
    End If
    
    'anything else
    If FExists(thisCommand) Then
        'If debugflg = 1 Then debugLog "ShellExecute " & thisCommand
        If sWorkingDirectory <> vbNullString Then
            Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, sWorkingDirectory, intShowCmd)
            Exit Sub
        Else
            Call shellExecuteWithDialog(userLevel, thisCommand, sArguments, vbNullString, intShowCmd)
            Exit Sub
        End If
    ElseIf DirExists(thisCommand) Then
        'If debugflg = 1 Then debugLog "ShellExecute " & thisCommand
        Call shellExecuteWithDialog("open", thisCommand, sArguments, sWorkingDirectory, intShowCmd)
        Exit Sub
    ElseIf validURL = False Then
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, thisCommand & " - That isn't valid as a target file or a folder, or it doesn't exist - so SteamyDock can't run it.", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
    End If
    
    On Error GoTo 0
    Exit Sub

runCommand_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure runCommand of Form dock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : isSysTray
' Author    : beededea
' Date      : 20/02/2021
' Purpose   : .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
'---------------------------------------------------------------------------------------
'
Function isSysTray(hTray As Long, ByRef processID As Long, ByRef hWnd As Long)

    Dim Count As Long: Count = 0
    Dim hIcon() As Long: 'hIcon() = 0
    Dim i As Long: i = 0
    Dim pid As Long: pid = 0

    On Error GoTo isSysTray_Error

    Count = GetIconCount(hTray)

    If Count <> 0 Then
        Call GetIconHandles(hTray, Count, hIcon)
    End If

    For i = 0 To Count - 1
        pid = GetPidByWindow(hIcon(i))
        'if the extracted pid matches the supplied processID then we have the window handle
        If pid = processID Then
            hWnd = hIcon(i)
            Exit Function
        End If
    Next

   On Error GoTo 0
   Exit Function

isSysTray_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure isSysTray of Form dock"
End Function

'---------------------------------------------------------------------------------------
' Procedure : shellExecuteWithDialog
' Author    : beededea
' Date      : 31/01/2021
' Purpose   : handler for shellexecute allowing a subsequent dialog to be inititated
'---------------------------------------------------------------------------------------
'
Private Sub shellExecuteWithDialog(ByRef userLevel As String, ByVal sCommand As String, ByVal sArguments As String, ByVal sWorkingDirectory As String, ByVal windowState As Integer)

   On Error GoTo shellExecuteWithDialog_Error
   
   
   If windowState = 0 Then windowState = 1 ' .67 DAEB 01/05/2021 frmMain.frm Added creation of Windows in the states as provided by sShowCmd value in RD
   
    '.nn Added new check box to allow autohide of the dock prior to launch of the chosen app
    If sAutoHideDock = "1" Then
        'MessageBox Me.hwnd, sTitle & " Hiding the dock ", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        ' store the process name that caused the dock to auto hide
        autoHideProcessName = sCommand ' .84 DAEB 20/07/2021 frmMain.frm Added prevention of the dock returning until the hiding application is no longer running.
        Call HideDockNow
        
        '.85 Added new timer to allow auto-reveal of the dock once the chosen app has closed within 1.5 secs
        forceHideRevealTimer.Enabled = True
        
    Else
       autoHideProcessName = vbNullString
    End If
   
    ' run the selected program
    Call ShellExecute(hWnd, userLevel, sCommand, sArguments, sWorkingDirectory, windowState) ' .67 DAEB 01/05/2021 frmMain.frm Added creation of Windows in the states as provided by sShowCmd value in RD
            
    userLevel = "open" ' return to default
    
    ' add the process to a list of processes initiated by the dock (except for a secondary app - 999)
    If selectedIconIndex <> 999 Then initiatedProcessArray(selectedIconIndex) = sCommandArray(selectedIconIndex)
    Call dockProcessTimer ' trigger a test of all running processes

    ' call up a dialog box if required
    If sUseDialogAfter = "1" Then
        'MsgBox sTitle & " Command Issued - " & sCommand, vbSystemModal + vbExclamation, "SteamyDock Confirmation Message"
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, sTitle & " Command Issued - " & sCommand, "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
    End If
    
    
   On Error GoTo 0
   Exit Sub

shellExecuteWithDialog_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure shellExecuteWithDialog of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : shellCommand
' Author    : beededea
' Date      : 31/01/2021
' Purpose   : handler for shell command allowing a subsequent dialog to be initiated
'---------------------------------------------------------------------------------------
'
Private Sub shellCommand(ByVal shellparam1 As String, ByVal windowState As Integer)

   On Error GoTo shellCommand_Error
        
    '.nn Added new check box to allow autohide of the dock prior to launch of the chosen app
    If sAutoHideDock = "1" Then
        'MessageBox Me.hwnd, sTitle & " Hiding the dock ", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
        Call HideDockNow
    End If
    
    ' .67 DAEB 01/05/2021 frmMain.frm Added creation of Windows in the states as provided by sShowCmd value in RD
    ' run the selected program
    If windowState = 0 Then windowState = 1
    If windowState = 1 Then Call Shell(shellparam1, vbNormalFocus)
    If windowState = 2 Then Call Shell(shellparam1, vbMinimizedFocus)
    If windowState = 3 Then Call Shell(shellparam1, vbMaximizedFocus)
    
    userLevel = "open" ' return to default
    
    ' add the process to a list of processes initiated by the dock
    initiatedProcessArray(selectedIconIndex) = sCommandArray(selectedIconIndex)
    Call dockProcessTimer ' trigger a test of all running processes

    ' call up a dialog box if required
    If sUseDialogAfter = "1" Then
        ' .43 DAEB 01/04/2021 frmMain.frm Replaced the modal msgbox with the non-modal form
        MessageBox Me.hWnd, sTitle & " Command Issued - " & sCommand, "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
    End If
    


   On Error GoTo 0
   Exit Sub

shellCommand_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure shellCommand of Form dock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawTheText
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : draws any text onto the device with the characteristics required
'---------------------------------------------------------------------------------------
'
Private Sub DrawTheText(ByVal strText As String, ByVal yTop As Single, ByVal xLeft As Single, ByVal textWidth As Integer, Optional ByVal strFont As String = "Tahoma", Optional ByVal bytFontSize As Byte, Optional ByVal bytBorderSize As Byte = 1)
    Dim borderRGBColour As Long: borderRGBColour = 0
    Dim borderARGBColour As Long: borderARGBColour = 0
    Dim borderOpacity As Integer: borderOpacity = 0
    Dim strFontRGBColour As String: strFontRGBColour = vbNullString
    Dim strBorderRGBColour As String: strBorderRGBColour = vbNullString
    Dim strShadowRGBColour As String: strShadowRGBColour = vbNullString
    Dim fontRGBColour As Long: fontRGBColour = 0
    Dim fontARGBColour As Long: fontARGBColour = 0
    Dim shadowRGBColour As Long: shadowRGBColour = 0
    Dim shadowARGBColour As Long: shadowARGBColour = 0
    Dim shadowOpacity As Integer: shadowOpacity = 0
    Dim fontOpacity As Integer: fontOpacity = 0
    Dim rctTextBottom As Integer: rctTextBottom = 0
    
    On Error GoTo DrawTheText_Error
    
    rctTextBottom = 64
        
    Call GdipCreateFromHDC(dcMemory, lngFont)
    Call GdipCreateFontFamilyFromName(StrConv(strFont, vbUnicode), 0, lngFontFamily)
    
    ' if the font has bold then we can handle that here
    Call GdipCreateFont(lngFontFamily, bytFontSize, FontStyleRegular, UnitPoint, lngCurrentFont)
    
    Call GdipCreateStringFormat(0, 0, lngFormat)
    Call GdipSetStringFormatAlign(lngFormat, StringAlignmentCenter)
    Call GdipSetStringFormatLineAlign(lngFormat, StringAlignmentNear)
    
     
    'do the shadow first
    ' convert decimal colour to ARGB (opacity then RGB)
    shadowRGBColour = rDFontShadowColor
    shadowOpacity = 255 * Val(rDFontShadowOpacity) / 100
    shadowARGBColour = Color_RGBtoARGB(shadowRGBColour, shadowOpacity) 'shadowOpacity)

    Call GdipCreateSolidFill(shadowARGBColour, lngBrush)
    rctText.Left = xLeft + 3
    rctText.Top = yTop + 3
    rctText.Right = textWidth ' Me.ScaleWidth
    rctText.Bottom = rctTextBottom
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)


    ' Draw the border around the text

    ' set the colour for all the borders
    ' convert decimal colour to ARGB (opacity then RGB)
    borderRGBColour = rDFontOutlineColor ' an RGB long required by GDI conversion tools
    borderOpacity = 255 * Val(rDFontOutlineOpacity) / 100
    borderARGBColour = Color_RGBtoARGB(borderRGBColour, borderOpacity) ' borderOpacity)

    Call GdipCreateSolidFill(borderARGBColour, lngBrush)  ' This API requires ARGB format

    ' border to the left
    rctText.Left = xLeft - bytBorderSize
    rctText.Top = yTop
    rctText.Right = textWidth
    rctText.Bottom = rctTextBottom
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)

    ' border to the right
    rctText.Left = xLeft + bytBorderSize
    rctText.Top = yTop
    rctText.Right = textWidth
    rctText.Bottom = rctTextBottom
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)

    ' border to the top
    rctText.Left = xLeft
    rctText.Top = yTop - bytBorderSize
    rctText.Right = textWidth
    rctText.Bottom = rctTextBottom
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)

    ' border to the bottom
    rctText.Left = xLeft
    rctText.Top = yTop + bytBorderSize
    rctText.Right = textWidth
    rctText.Bottom = rctTextBottom
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)
    



    ' Now draw the text
    
    ' set the colour for the text itself
    ' convert RD decimal colour to ARGB (opacity followed by RGB)
                    
    fontRGBColour = rDFontColor ' an RGB long required by GDI conversion tools
    fontOpacity = 255 * Val(sDFontOpacity) / 100
    fontARGBColour = Color_RGBtoARGB(fontRGBColour, fontOpacity) ' wants a RGB long and gives a long.
        
    Call GdipCreateSolidFill(fontARGBColour, lngBrush)
    
    rctText.Left = xLeft
    rctText.Top = yTop
    rctText.Right = textWidth
    rctText.Bottom = rctTextBottom
    
    'legend =      graphic bitmap, StringToDraw, lengthOfTheStringToDraw, chosenFont, layoutRectangle, StringFormat As Long, brush As Long
    Call GdipDrawString(lngFont, StrConv(strText, vbUnicode), Len(strText), lngCurrentFont, rctText, lngFormat, lngBrush)
    
    Call GdipDeleteStringFormat(lngFormat)
    Call GdipDeleteFont(lngCurrentFont)
    Call GdipDeleteFontFamily(lngFontFamily)
    Call GdipDeleteBrush(lngBrush)
    Call GdipDeleteGraphics(lngFont)

   On Error GoTo 0
   Exit Sub

DrawTheText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DrawTheText of Form dock"

End Sub





        

        
'---------------------------------------------------------------------------------------
' Procedure : runTimer
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : calls the subroutine that runs the actual command from the selected icon
'---------------------------------------------------------------------------------------
'
Private Sub runTimer_Timer()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim rmessage As String: rmessage = vbNullString

   On Error GoTo runTimer_Error
    
    runTimer.Enabled = False
    Call runCommand("run", "") ' added new parameter to allow override ref: .68

    If sSecondApp <> vbNullString Then
    
        ' run the selected program
        If sUseDialog = "1" Then
            rmessage = "Are you sure you wish to run the associated second application? - " & sTitle & "?" & vbCr & sSecondApp
            answer = MsgBox(rmessage, vbYesNo)
            If answer = vbNo Then
                Exit Sub
            End If
        End If
        If FExists(sSecondApp) Then ' .78 DAEB 21/05/2021 frmMain.frm Added new field for second program to be run
            
            delayRunTimer.Enabled = True
           
        End If
    End If
    
    selectedIconIndex = 999 ' sets the icon to bounce index to something that will never occur

    
   On Error GoTo 0
   Exit Sub

runTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure runTimer of Form dock"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : delayRunTimer_Timer
' Author    : beededea
' Date      : 22/12/2022
' Purpose   : The timer is 3 seconds, starts the secondary program after the first run.
'---------------------------------------------------------------------------------------
'
Private Sub delayRunTimer_Timer()
    On Error GoTo delayRunTimer_Timer_Error

    delayRunTimerCount = delayRunTimerCount + 1
    If delayRunTimerCount >= 1 Then
        delayRunTimer.Enabled = False
        delayRunTimerCount = 0
        Call runCommand("run", sSecondApp)
    End If

    On Error GoTo 0
    Exit Sub

delayRunTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure delayRunTimer_Timer of Form dock"
            Resume Next
          End If
    End With
End Sub

' .72 DAEB 06/05/2021 frmMain.frm Created two timers that controls the bouncing when the icon is clicked, replacing the old timers
'---------------------------------------------------------------------------------------
' Procedure : bounceDownTimer_Timer
' Author    : beededea
' Date      : 11/05/2021
' Purpose   : 'timer that controls the bounce Downward when the icon is clicked
'---------------------------------------------------------------------------------------
'
Private Sub bounceDownTimer_Timer()
    Dim bvalue As Double: bvalue = 0

    On Error GoTo bounceDownTimer_Timer_Error
    
    ' first type of animation using a tall double bounce
    If rDIconActivationFX = "1" Then

        bounceCounter = bounceCounter - 4
    
        bvalue = BounceIn(bounceCounter / bounceZone) ' uses the same bounce IN type as the bounce IN
        bounceHeight = bounceZone * bvalue
    
        If bounceCounter <= 0 Then
            bounceHeight = 0
            bounceCounter = 0
            bounceDownTimer.Enabled = False
            If Val(sQuickLaunch) = 0 Then
                ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
                runTimer.Enabled = True
            End If
        End If
    End If

    ' second type of animation, a simple bounce up and down
    If rDIconActivationFX = "2" Then
        bounceDownTimer.Interval = 30
        bounceCounter = bounceCounter - sDBounceStep
        If bounceTimerRun = 2 Then bounceUpTimer.Interval = sDBounceInterval * 3
        If bounceTimerRun = 4 Then bounceUpTimer.Interval = sDBounceInterval * 4
        bounceHeight = bounceCounter
        If bounceCounter <= 0 And bounceTimerRun = 2 Then
            bounceTimerRun = bounceTimerRun + 1
            If Val(sQuickLaunch) = 0 Then
                ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
                 runTimer.Enabled = True  ' .77 DAEB 20/05/2021 frmMain.frm Added new check box to allow a quick launch of the chosen app
            End If
            bounceUpTimer.Enabled = True
            bounceDownTimer.Enabled = False
        End If
    
        If bounceCounter <= 0 And bounceTimerRun = 4 Then
            bounceCounter = 0
            bounceTimerRun = bounceTimerRun + 1
            bounceUpTimer.Enabled = True
            bounceDownTimer.Enabled = False
        End If
    End If



    On Error GoTo 0
    Exit Sub

bounceDownTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bounceDownTimer_Timer of Form dock"
End Sub
' .72 DAEB 06/05/2021 frmMain.frm Created two timers that controls the bouncing when the icon is clicked, replacing the old timers
'---------------------------------------------------------------------------------------
' Procedure : bounceUpTimer
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : timer that controls the bounceUpward when the icon is clicked
'---------------------------------------------------------------------------------------
'
Private Sub bounceUpTimer_Timer()
   On Error GoTo bounceUpTimer_Error
   
    Dim bvalue As Double: bvalue = 0
    
    If rDIconActivationFX = "0" Then ' no icon animation at all
        bounceUpTimer.Enabled = False
        runTimer.Enabled = True
        Exit Sub
    End If
    
    If Val(sQuickLaunch) = 1 Then
        ' .11 DAEB 21/05/2021 common.bas Added new field for second program to be run
         bounceUpTimer.Enabled = False
         runTimer.Enabled = True  ' .77 DAEB 20/05/2021 frmMain.frm Added new check box to allow a quick launch of the chosen app
         Exit Sub
    End If

    If rDIconActivationFX = "1" Then
        
        bounceCounter = bounceCounter + 4
    
        bvalue = BounceIn(bounceCounter / bounceZone)
        bounceHeight = bounceZone * bvalue
    
        If bounceCounter >= bounceZone Then
            bounceUpTimer.Enabled = False
            bounceDownTimer.Enabled = True
'            If Val(sQuickLaunch) = 1 Then
'                runTimer.Enabled = True  ' .77 DAEB 20/05/2021 frmMain.frm Added new check box to allow a quick launch of the chosen app
'            End If
        End If
    End If
    
    
    If rDIconActivationFX = "2" Then
        bounceUpTimer.Interval = 30
        bounceCounter = bounceCounter + sDBounceStep
        If bounceTimerRun = 3 Then bounceUpTimer.Interval = sDBounceInterval * 5
        bounceHeight = bounceCounter
        
        If bounceCounter >= 50 Then
            bounceTimerRun = bounceTimerRun + 1
            bounceUpTimer.Enabled = False
            bounceDownTimer.Enabled = True
        End If
    
        If bounceTimerRun = 5 Then
            bounceCounter = 0
            bounceTimerRun = 1
            bounceUpTimer.Enabled = False
            bounceDownTimer.Enabled = False
            bounceUpTimer.Interval = 10
            bounceDownTimer.Interval = 10
'            If Val(sQuickLaunch) = 1 Then
'                runTimer.Enabled = True  ' .77 DAEB 20/05/2021 frmMain.frm Added new check box to allow a quick launch of the chosen app
'            End If
        End If
    End If
    
    

   On Error GoTo 0
   Exit Sub

bounceUpTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bounceUpTimer of Form dock"
    
End Sub


''---------------------------------------------------------------------------------------
'' Procedure : bounceDownTimer
'' Author    : beededea
'' Date      : 19/04/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub bounceDownTimer_Timer()
'   On Error GoTo bounceDownTimer_Error
'   'If debugflg = 1 Then debugLog "%bounceDownTimer"
'
''    bounceCounter = bounceCounter - sDBounceStep
''    If bounceTimerRun = 2 Then bounceUpTimer.Interval = sDBounceInterval * 3
''    If bounceTimerRun = 4 Then bounceUpTimer.Interval = sDBounceInterval * 4
''
''    If bounceCounter <= 0 And bounceTimerRun = 2 Then
''        bounceTimerRun = bounceTimerRun + 1
''        runTimer.Enabled = True ' the runtimer start used to be here but occasionally an app will take time to start and a delay is introduced into the bounce animation
''        bounceUpTimer.Enabled = True
''        bounceDownTimer.Enabled = False
''    End If
''
''    If bounceCounter <= 0 And bounceTimerRun = 4 Then
''        bounceCounter = 0
''        bounceTimerRun = bounceTimerRun + 1
''        bounceUpTimer.Enabled = True
''        bounceDownTimer.Enabled = False
''    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'bounceDownTimer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bounceDownTimer of Form dock"
'
'End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setInitialStartPoint
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : used to determine the initial dock start position for the small icon display only
'---------------------------------------------------------------------------------------
'
Private Sub setInitialStartPoint()

    Dim proportionalOffset As Integer: proportionalOffset = 0
    Dim hOffsetPxls As Integer: hOffsetPxls = 0
    
    ' set the starting point for the icons to be drawn
    On Error GoTo setInitialStartPoint_Error

    'If debugflg = 1 Then debugLog "%" & "setInitialStartPoint"

    If dockPosition = vbbottom Then
        screenBottomPxls = Me.Height / screenTwipsPerPixelX ' .73 DAEB 11/05/2021 frmMain.frm  sngBottom renamed to screenBottomPxls
        
        If slideOutFlag = True Then
            dockUpperMostPxls = (screenHeightPixels - 10) ' 10 pixels above the bottom of the screen ' .nn
        Else
            ' the dock at the bottom of the screen taking into account the largest icons size
            dockUpperMostPxls = (Me.Height / screenTwipsPerPixelX) - iconSizeLargePxls
            ' the dock uppermost position now taking into account the dock vertical offset as defined by the user
            dockUpperMostPxls = dockUpperMostPxls - rDvOffset - rdDefaultYPos
        End If

    End If
    
    If dockPosition = vbtop Then ' .nn STARTS
        screenBottomPxls = 0 ' .nn 'the top of the screen, position 0
        
        If slideOutFlag = True Then ' if the dock has slid out then we need to expose just the first 10 pixels of the dock
            dockUpperMostPxls = 10
        Else
'           ' the dock uppermost position at the top of the screen taking into account the dock vertical offset as defined by the user
            dockUpperMostPxls = rDvOffset + rdDefaultYPos '.nn
        End If
         ' .nn ENDS
    End If
    

    normalDockWidthPxls = (rdIconMaximum * iconSizeSmallPxls)
    hOffsetPxls = ((screenWidthPixels - normalDockWidthPxls) / 2)
    proportionalOffset = hOffsetPxls + (hOffsetPxls * (Val(rDOffset) / 100))
    iconLeftmostPointPxls = proportionalOffset

    iconPosLeftPxls = iconLeftmostPointPxls ' rDOffset


   On Error GoTo 0
   Exit Sub

setInitialStartPoint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setInitialStartPoint of Form dock"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : shutdwnGDI
' Author    : beededea
' Date      : 08/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub shutdwnGDI()
   On Error GoTo shutdwnGDI_Error

    If lngImage Then
        Call GdipReleaseDC(lngImage, dcMemory)
        Call GdipDeleteGraphics(lngImage)
    End If
    If lngBitmap Then Call GdipDisposeImage(lngBitmap)
    If lngGDI Then Call GdiplusShutdown(lngGDI)

   On Error GoTo 0
   Exit Sub

shutdwnGDI_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure shutdwnGDI of Form dock"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : processTimer
' Author    : beededea
' Date      : 11/04/2020
' Purpose   : checks whether the listed processes are currently running every 5-65 secs (10 by default)
'---------------------------------------------------------------------------------------
'
Private Sub processTimer_Timer()
   On Error GoTo processTimer_Error
   
   Call dockProcessTimer

   On Error GoTo 0
   Exit Sub

processTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure processTimer of Form dock"
End Sub






'---------------------------------------------------------------------------------------
' Procedure : checkQuestionMark
' Author    : beededea
' Date      : 16/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub checkQuestionMark(ByVal key As String, ByRef Filename As String, ByVal Size As Byte)
    Dim filestring As String: filestring = vbNullString
    Dim suffix As String: suffix = vbNullString
    Dim qPos As Integer: qPos = 0

    ' does the string contain a ? if so it probably has an embedded .ICO
   On Error GoTo checkQuestionMark_Error

    qPos = InStr(1, Filename, "?")
    If qPos <> 0 Then
        ' extract the string before the ? (qPos)
        filestring = Mid$(Filename, 1, qPos - 1)
    End If
    
    ' test the resulting filestring exists
    If FExists(filestring) Then
        ' extract the suffix
        suffix = ExtractSuffixWithDot(filestring)
        ' test as to whether it is an .EXE or a .DLL
        If InStr(".exe,.dll", LCase(suffix)) <> 0 Then
            Call displayEmbeddedIcons(key, filestring, hiddenForm.hiddenPicbox, Size)
        Else
            ' the file may have a ? in the string but does not match otherwise in any useful way
            Filename = sdAppPath & "\icons\" & "help.png" ' .12 25/01/2021 DAEB Change to sdAppPath
        End If
    Else
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

checkQuestionMark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkQuestionMark of Form dock"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : GetShortcutInfoNET
' Author    : beededea
' Date      : 17/04/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetShortcutInfoNET(ByVal ShortcutPath As String) As String

    Dim Begin As Long: Begin = 0
    Dim EndV As Long: EndV = 0
    Dim FileInfoStartsAt As Long: FileInfoStartsAt = 0
    Dim FileOffset As Long: FileOffset = 0
    Dim FirstPart As String: FirstPart = vbNullString
    Dim flags As Long: flags = 0
    Dim fileH As Long: fileH = 0
    Dim Offset As Integer: Offset = 0
    Dim Link As String: Link = vbNullString
    Dim LinkTarget As String: LinkTarget = vbNullString
    Dim PathLength As Long: PathLength = 0
    Dim SecondPart As String: SecondPart = vbNullString
    Dim TotalStructLength As Long: TotalStructLength = 0

   On Error GoTo GetShortcutInfoNET_Error

   fileH = FreeFile()
   If Dir$(ShortcutPath, vbNormal) = vbNullString Then Error 53
   
   Open ShortcutPath For Binary Lock Read Write As fileH
      Seek #fileH, &H15
      Get #fileH, , flags
      If (flags And &H1) = &H1 Then
         Seek #fileH, &H4D
         Get #fileH, , Offset
         Seek #fileH, Seek(fileH) + Offset
      End If

      FileInfoStartsAt = Seek(fileH) - 1
      Get #fileH, , TotalStructLength
      Seek #fileH, Seek(fileH) + &HC
      Get #fileH, , FileOffset
      Seek #fileH, FileInfoStartsAt + FileOffset + 1
      
      PathLength = (TotalStructLength + FileInfoStartsAt) - Seek(fileH) - 1
      LinkTarget = Input$(PathLength, fileH)
      Link = LinkTarget
      
      Begin = InStr(Link, vbNullChar & vbNullChar)
      If Begin > 0 Then
         EndV = InStr(Begin + 2, Link, "\\")
         EndV = InStr(EndV, Link, vbNullChar) + 1
       
         FirstPart = Mid$(Link, 1, Begin - 1)
         SecondPart = Mid$(Link, EndV)
 
         GetShortcutInfoNET = FirstPart & SecondPart
         Exit Function
      End If

      GetShortcutInfoNET = Link
      Exit Function
   Close fileH

GetShortcutInfoNET = vbNullString

   On Error GoTo 0
   Exit Function

GetShortcutInfoNET_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetShortcutInfoNET of Form dock"
End Function


'---------------------------------------------------------------------------------------
' Procedure : displayEmbeddedAllIcons
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : The program extracts icons embedded within a DLL or an executable
'             you pass the name of the picbox you require and the image is displayed there
'             it should return all and not only the 16 and 32 bit icons
'
'             I may not have coded this particularly well - but it works.
'---------------------------------------------------------------------------------------
'
Public Sub displayEmbeddedIcons(ByVal key As String, ByVal Filename As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer)
    
    Dim sExeName As String: sExeName = vbNullString
    Dim lIconIndex As Long: lIconIndex = 0
    Dim xSize As Long: xSize = 0
    Dim ySize As Long: ySize = 0
    Dim hIcon() As Long: 'hIcon() = 0  cannot initialise
    Dim hIconID() As Long: 'hIconID() = 0  cannot initialise
    Dim nIcons As Long: nIcons = 0
    Dim Result As Long: Result = 0
    Dim flags As Long: flags = 0
    Dim i As Long: i = 0
    Dim pic As IPicture: 'pic cannot initialise
    Dim thiskey As String: thiskey = vbNullString
    Dim bytesFromFile() As Byte
    Dim Strm As stdole.IUnknown '  cannot initialise
    Dim img As Long: img = 0
    Dim dx As Long: dx = 0
    Dim dy As Long: dy = 0
    Dim strFilename As String: strFilename = vbNullString
    Dim opacity As String: opacity = vbNullString

    
    On Error Resume Next

    sExeName = Filename
    lIconIndex = 0
    strFilename = App.Path & "\tmp.bmp"
    
    i = 2 ' need some experimentation here
    
    'the boundaries of the icons you wish to extract, can be set to somethink like 256, 256 but that is all
    ' you will extract, just the one icon
    xSize = make32BitLong(CInt("256"), CInt("16"))
    ySize = make32BitLong(CInt("256"), CInt("16"))
    
    flags = LR_LOADFROMFILE

    ' Get the total number of Icons in the file.
    Result = PrivateExtractIcons(sExeName, lIconIndex, xSize, ySize, ByVal 0&, ByVal 0&, 0&, 0&)
    
    ' The sExeName is the resource string/filepath.
    ' lIconIndex Index is the index.
    ' xSize and ySize are the desired sizes.
    ' phicon is the returned array of icon handles.
    ' So you could call it with phicon set to nothing and it will return the number of icons in the file.
    
    ' piconid ?
    
    ' nicons is just the number of icons you wish to extract.
    ' Then you call it again with nicon set to this number and niconindex=0. Then it will extract ALL icons in one go.
    ' flags
    '
    '    LR_DEFAULTCOLOR
    '    LR_CREATEDIBSECTION
    '    LR_DEFAULTSIZE
    '    LR_LOADFROMFILE
    '    LR_LfsOADMAP3DCOLORS
    '    LR_LOADTRANSPARENT
    '    LR_MONOCHROME
    '    LR_SHARED
    '    LR_VGACOLOR
    '
    ' eg. PrivateExtractIcons ('C:\Users\Public\Documents\RAD Studio\Projects\2010\Aero Colorizer\AeroColorizer.exe', 0, 128, 128, @hIcon, @nIconId, 1, LR_LOADFROMFILE)
    ' PrivateExtractIcons(sExeName, nIcon, cxIcon, cyIcon, phicon, piconid, nicons, 0)

    nIcons = 2 ' Result
    
    ' Dimension the arrays to the number of icons.
    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
    ReDim hIconID(lIconIndex To lIconIndex + nIcons * 2 - 1)

'  Rocketdock always uses the same ID for the same exe

'   C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE?5063424
'   E:\games\World_of_Tanks_NA\WorldOfTanks.exe?184608432

' if an exe is dragged and dropped onto RD it is given an id, it is appended to the binary name after an additional "?"
' that question mark signifies what? Possibly the handle of the embedded icon only added the first time,
' so that when the binary is read in the future the handle is already there?
' and that can be used to populate image array? Untested.
' in this case we just need to note the ? and then query the binary for an embedded icon and compare it to the id that RD has given it.
        
    ' use the undocumented PrivateExtractIcons to extract the icons we require
    Result = PrivateExtractIcons(sExeName, lIconIndex, xSize, _
                            ySize, hIcon(LBound(hIcon)), _
                            hIconID(LBound(hIconID)), _
                            nIcons * 2, flags)
        
    ' create an icon with a handle so we can test the result
    'pic = CreateIcon(hIcon(LBound(hIcon)))
    
    ' Creates a GDI+ Image object based on the stream, Olaf Schmidt
'    GdipLoadImageFromStream ObjPtr(Strm), img
'    If img = 0 Then MsgBox "Could not load image with GDIPlus"
'
'    'GDI+ API to determine image dimensions, Olaf Schmidt
'    GdipGetImageWidth pic, dx
'    GdipGetImageHeight pic, dy
'
'    ' uses a function extracted from Olaf Schmidt's code in gdiPlusCacheCls to create and resize the image
'    lngBitmap = CreateScaledImg(pic, dx, dy, IconSize, IconSize, opacity)
'
'    ' create a unique key string
'    thiskey = key & "ResizedImg" & LTrim$(str$(IconSize))
'
'    ' add the bitmap to the dictionary collection
'    collLargeIcons.Add thiskey, lngBitmap
'
'   ' get rid of the icon we created
'    Call DestroyIcon(hIcon(i + lIconIndex - 1))
            
    
    
    
    
    'MsgBox hIcon(LBound(hIcon))
    
    ' Draw the icon to a hidden picturebox control.
    ' this is a bit of a temporary kludge just seeing how to extract the embedded icon from the exe to a GDI+ image
    If Not (pic Is Nothing) Then
        With targetPicBox
        
            .Width = IconSize * screenTwipsPerPixelX
            .Height = IconSize * screenTwipsPerPixelX

            'ensure the picbox is empty first
            Set .Picture = LoadPicture(vbNullString)
            .AutoRedraw = True

            Call DrawIconEx(.hdc, 0, 0, hIcon(LBound(hIcon)), IconSize, IconSize, 0, 0, DI_NORMAL)
            .Refresh
            
            SavePicture .image, strFilename
        
            'hiddenForm.Show ' uses a hidden form to host the picbox so we can see the icon if needs be.
        
            ' uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes
            bytesFromFile = ReadBytesFromFile(strFilename)
        
            ' creates a stream object stored in global memory using the location address of the variable where the data resides, Olaf Schmidt
            CreateStreamOnHGlobal VarPtr(bytesFromFile(0)), 0, Strm
        
            ' Creates a GDI+ Image object based on the stream, Olaf Schmidt
            Call GdipLoadImageFromStream(ObjPtr(Strm), img)
            If img = 0 Then MsgBox "Could not load image with GDIPlus"
        
            'GDI+ API to determine image dimensions, Olaf Schmidt
            Call GdipGetImageWidth(img, dx)
            Call GdipGetImageHeight(img, dy)
        
            ' uses a function extracted from Olaf Schmidt's code in gdiPlusCacheCls to create and resize the image
            lngBitmap = CreateScaledImg(img, dx, dy, IconSize, IconSize, opacity)
        
            ' create a unique key string
            thiskey = key & "ResizedImg" & LTrim$(Str$(IconSize))
        
            ' add the bitmap to the dictionary collection
            collLargeIcons.Add thiskey, lngBitmap
        
           ' get rid of the icon we created
            Call DestroyIcon(hIcon(i + lIconIndex - 1))
            
            Kill strFilename
        
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateIcon
' Author    : beededea
' Date      : 14/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CreateIcon(ByVal hImage As Long) As IPicture
    ' This method creates an icon based on a handle
    Dim pic As IPicture
    Dim dsc As PictDesc
    Dim iid(0 To 15) As Byte
    Dim Result As Long: Result = 0
    
   On Error GoTo CreateIcon_Error

    Set CreateIcon = Nothing
    If hImage <> 0 Then
        With dsc
           .cbSizeofStruct = Len(dsc)
           .hImage = hImage
           .PicType = VBRUN.PictureTypeConstants.vbPicTypeIcon
        End With
        
        Result = OLE_CLSIDFromString(StrPtr(IID_IPicture), _
                                                        VarPtr(iid(0)))
                                                    
        If (Result = OLE_ERROR_CODES.S_OK) Then
            Result = Ole_CreatePic(dsc, VarPtr(iid(0)), True, pic)
            
            If (Result = OLE_ERROR_CODES.S_OK) Then
                Set CreateIcon = pic
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

CreateIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateIcon of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : make32BitLong
' Author    : beededea
' Date      : 20/11/2019
' Purpose   : packing variables into a 32bit LONG for an API call
'---------------------------------------------------------------------------------------
'
Private Function make32BitLong(ByVal LoWord As Integer, Optional ByVal HiWord As Integer = 0) As Long
   On Error GoTo make32BitLong_Error
   'If debugflg = 1 Then debugLog "%make32BitLong"

    make32BitLong = CLng(HiWord) * CLng(&H10000) + CLng(LoWord)

   On Error GoTo 0
   Exit Function

make32BitLong_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure make32BitLong of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : drawSmallStaticIcons
' Author    : beededea
' Date      : 28/07/2020
' Purpose   : Displays small icon images from the small collection.
'---------------------------------------------------------------------------------------
'
Public Sub drawSmallStaticIcons()
'    Dim a As Integer: a = 0
    Dim useloop As Integer: useloop = 0
    Dim dockHeight As Long: dockHeight = 0
    Dim thiskey As String: thiskey = vbNullString
    Dim dockSkinStart As Long: dockSkinStart = 0
    Dim dockSkinWidth As Long: dockSkinWidth = 0
    
    On Error GoTo drawSmallStaticIcons_Error

    Call setInitialStartPoint ' return the dock start point when small
    
    ' Check bDrawn so the program doesn't redraw the whole icon picture more than once
    If bDrawn = False Then
        iconPosLeftPxls = iconLeftmostPointPxls
        normalDockWidthPxls = 0
        iconHeightPxls = iconSizeSmallPxls
        iconWidthPxls = iconSizeSmallPxls
                    
        'Call drawSmallIconDockWithFadeEffects
                                            
        DeleteObject bmpMemory ' the bitmap deleted
        Call readyGDIPlus ' clears the whole previously drawn image section and the animation continues
    
        If rDtheme <> vbNullString And rDtheme <> "Blank" Then Call applyThemeSkinToDock(dockSkinStart, dockSkinWidth)
                
        ' this loop redraws all the icons at the same small size after the mouse has left the icon area
        For useloop = 0 To rdIconMaximum  'File1.ListCount - 1
            
            'Call sizePositionZero(useloop, showsmall)

            Call sizeEachSmallIconToLeft(useloop, rdIconMaximum, True)
            
            ' display the small size icons
            Call showSmallIcon(useloop)
                            
            ' store the icon current position in the array
            Call storeCurrentIconPositions(useloop)
                    
            iconPosLeftPxls = iconPosLeftPxls + iconWidthPxls
            
        Next useloop
                                            
         ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    '            iconStoreLeftPixels(UBound(iconStoreLeftPixels)) = iconPosLeftPxls
    '            iconStoreRightPixels(UBound(iconStoreRightPixels)) = iconStoreLeftPixels(UBound(iconStoreLeftPixels)) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
    '            iconStoreTopPixels(UBound(iconStoreRightPixels)) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
                    
        Call storeCurrentIconPositions(UBound(iconStoreLeftPixels))
        
        Call updateGDIPlus
            
        smallDockBeenDrawn = True
        bDrawn = True
    
    End If

   On Error GoTo 0
   Exit Sub

drawSmallStaticIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawSmallStaticIcons of Form dock"

End Sub


''---------------------------------------------------------------------------------------
'' Procedure : drawSmallIconDockWithFadeEffects
'' Author    : beededea
'' Date      : 04/05/2020
'' Purpose   : Starting at a set LEFT hand side, it loops through each element in the small dictionary and adds each icon to the
''             combined image for display - no animation performed. This runs in conjunction with the responseTimer that operates
''             at a much reduced rate to avoid overuse of the CPU.
''             It only displays small icon images from the small collection.
''---------------------------------------------------------------------------------------
''
'Public Sub drawSmallIconDockWithFadeEffects()
'    Dim useloop As Integer
'    Dim thiskey As String: thiskey = vbNullString
'    Dim dockSkinStart As Long: dockSkinStart = 0
'    Dim dockSkinWidth As Long: dockSkinWidth = 0
'
'    iconWidthPxls = iconSizeSmallPxls
'
'    On Error GoTo drawSmallIconDockWithFadeEffects_Error
'   'If debugflg = 1 Then debugLog "%drawSmallIconDockWithFadeEffects"
'
'            DeleteObject bmpMemory ' Now the bitmap may be deleted
'            Call readyGDIPlus
'
'            If rDtheme <> vbNullString And rDtheme <> "Blank" Then Call applyThemeSkinToDock(dockSkinStart, dockSkinWidth)
'
'            ' this loop redraws all the icons at the same small size after the mouse has left the icon area
'            For useloop = 0 To rdIconMaximum  'File1.ListCount - 1
'
''                If dockPosition = vbbottom Then
''                    If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
''                        iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
''                    ElseIf autoSlideMode = "slidein" Then
''                        iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
''                    Else
''                        iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) '.nn
''                    End If
''                End If
''                If dockPosition = vbtop Then
''                    ' .nn
''                    If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
''                        'iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
''                    ElseIf autoSlideMode = "slidein" Then
''                        iconCurrentTopPxls = dockUpperMostPxls - xAxisModifier
''                        'iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
''                    Else
''                        iconCurrentTopPxls = dockUpperMostPxls - xAxisModifier
''                    End If
''
''                End If
'
'                ' NOTE: re-using the subroutine that is normally used to put small icons to the left shown in small mode
'                ' used here instead to resize all icons
'
'                Call sizeEachSmallIconToLeft(useloop, rdIconMaximum, True)
'
'                ' display the small size icons
'                Call showSmallIcon(useloop)
'
'                ' store the icon current position in the array
'                Call storeCurrentIconPositions(useloop)
'
'                iconPosLeftPxls = iconPosLeftPxls + iconWidthPxls
'            Next useloop
'
''       DrawTheText "responseTimer.Enabled " & responseTimer.Enabled, 440, 50, 300, rDFontName, Val(Abs(rDFontSize))
''       DrawTheText "responseTimer.interval " & responseTimer.Interval, 460, 50, 300, rDFontName, Val(Abs(rDFontSize))
'
'
'   On Error GoTo 0
'   Exit Sub
'
'drawSmallIconDockWithFadeEffects_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure drawSmallIconDockWithFadeEffects of Form dock"
'End Sub


'            'section width in pixels
'            animSectPixelWidth = (rDZoomWidth * (rdIconMax / 2)) / 2 ' the max icon pixel width /2 multiplied by the number of animated icons
'            animSectTwipWidth = animSectPixelWidth * screenTwipsPerPixelX  '
'
'            ' distance of the current icon from the centre of the section in twips
'            h = (apiMouse.X * screenTwipsPerPixelX) - iconPosLeftTwips(startAnimSec)
'
'            'proportion of the current icon from the centre of the section
'            animateStep = h / animSectTwipWidth
'            If animateStep >= 1 Then animateStep = 1
'
'            'the closer to the centre of the section the larger the icon until reaches maxbytesize
'            'no smaller than minbytesize
'
'            'animateStep = ( / (2 * screenTwipsPerPixelX)
'
'            iconHeightPxls = iconSizeLargePxls * animateStep 'animateStep is the difference from the midpoint of the current icon in the x axis
'            iconWidthPxls = iconSizeLargePxls * animateStep
'
'            If dockPosition = vbBottom Then
'                iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls * animateStep)
'            End If
'
'            If selectedIconIndex = iconIndex Then
'                iconCurrentTopPxls = iconCurrentTopPxls - bounceCounter
'            End If





'---------------------------------------------------------------------------------------
' Procedure : prepareArraysAndCollections
' Author    : beededea
' Date      : 02/05/2020
' Purpose   : resize arrays and load the images into the collections
'---------------------------------------------------------------------------------------
'
Public Sub prepareArraysAndCollections()
    Dim a As Integer: a = 0
    Dim strKey As String: strKey = vbNullString
    'sDSkinSize = 30
    
    ' redimension the arrays to cater for the number of icons in the dock
    On Error GoTo prepareArraysAndCollections_Error
    If debugflg = 1 Then debugLog "% sub prepareArraysAndCollections"

    ReDim fileNameArray(rdIconMaximum) As String ' the file location of the original icons
    ReDim namesListArray(rdIconMaximum) As String ' the name assigned to each icon
    ReDim sCommandArray(rdIconMaximum) As String ' the Windows command or exe assigned to each icon
    ReDim targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
    ReDim processCheckArray(rdIconMaximum) As String ' the array that contains true/false according to the running state of the associated process
    ReDim initiatedProcessArray(rdIconMaximum) As String ' the array containing the binary name of any process initiated by the dock

    Call loadAdditionalImagestoDictionary ' the additional images need to be added to the dictionary
    
    ' extract filenames from Rocketdock registry or settings.ini
    For a = 0 To rdIconMaximum
        'readIconData (a)
        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", a, dockSettingsFile

        strKey = LTrim$(Str$(a))
        ' read the two main icon variables into arrays, one for each
        fileNameArray(a) = sFilename
        namesListArray(a) = sTitle
        sCommandArray(a) = sCommand
        
        ' here is the code to cache the images to the collection at a small size
        If FExists(sFilename) Then
            resizeAndLoadImgToDict collSmallIcons, strKey, fileNameArray(a), namesListArray(a), (0), (0), (iconSizeSmallPxls), (iconSizeSmallPxls)
        ElseIf InStr(sFilename, "?") And readEmbeddedIcons = True Then ' Note: the question mark is an illegal character and test for a valid file will fail in VB.NET despite working in VB6 so we test it as a string instead
            checkQuestionMark strKey, fileNameArray(a), iconSizeSmallPxls ' if the question mark appears in the icon string - test it for validity and an embedded icon
        Else ' if the image is not found display an 'x'
            resizeAndLoadImgToDict collSmallIcons, strKey, App.Path & "\red-X.png", "buggered", (0), (0), (iconSizeSmallPxls), (iconSizeSmallPxls)
        End If
        
        ' now cache all the images to the collection at the larger size
        If FExists(sFilename) Then
            resizeAndLoadImgToDict collLargeIcons, strKey, fileNameArray(a), namesListArray(a), (0), (0), (iconSizeLargePxls), (iconSizeLargePxls)
        ElseIf InStr(sFilename, "?") And readEmbeddedIcons = True Then  ' Note: the question mark is an illegal character and test for a valid file will fail in VB.NET despite working in VB6 so we test it as a string instead
            checkQuestionMark strKey, fileNameArray(a), iconSizeLargePxls ' if the question mark appears in the icon string - test it for validity and an embedded icon
        Else
            resizeAndLoadImgToDict collLargeIcons, strKey, App.Path & "\red-X.png", "buggered", (0), (0), (iconSizeLargePxls), (iconSizeLargePxls)
        End If
        
        ' check to see if each process is running and store the result away - this is also run on a 10s timer
        'processCheckArray(a) = isProcessInTaskList(sCommand)
        processCheckArray(a) = IsRunning(sCommand, vbNull)
    Next a
    
    'redimension the array that is used to store all of the icon current positions
    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    ReDim Preserve iconStoreLeftPixels(theCount)
    ReDim Preserve iconStoreRightPixels(theCount) ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
    ReDim Preserve iconStoreTopPixels(theCount) ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
    ReDim Preserve iconStoreBottomPixels(theCount) ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon

    
    iconArrayUpperBound = rdIconMaximum

   On Error GoTo 0
   Exit Sub

prepareArraysAndCollections_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prepareArraysAndCollections of Form dock"

End Sub






'---------------------------------------------------------------------------------------
' Procedure : readToolSettings
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read this utilties' own settings.ini file and do some things using the data - unused
'---------------------------------------------------------------------------------------
'
Private Sub readToolSettings()
    Dim sfirst As String: sfirst = vbNullString

    On Error GoTo readToolSettings_Error
    'If debugflg = 1 Then debugLog "%" & "readToolSettings"
   
    If Not FExists(toolSettingsFile) Then Exit Sub ' does the tool's own settings.ini exist?
    
    'test to see if the tool has ever been run before
    sfirst = GetINISetting("Software\SteamyDockSettings", "First", toolSettingsFile)
    
    If sfirst = "True" Then
    
        sfirst = "False"
        
        'write the updated test of first run to false
        PutINISetting "Software\SteamyDockSettings", "First", sfirst, toolSettingsFile
        
    End If

    If IsUserAnAdmin() = 0 And requiresAdmin = True Then
        MsgBox "This tool requires to be run as administrator on Windows 8 and above in order to function. Admin access is NOT required on Win7 and below. If you aren't entirely happy with that then you'll need to remove the software now. This is a limitation imposed by Windows itself. To enable administrator access find this tool's exe and right-click properties, compatibility - run as administrator. YOU have to do this manually, I can't do it for you."
    End If
    
   On Error GoTo 0
   Exit Sub

readToolSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readToolSettings of Form rDIconConfigForm"
    
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : locateDockSettingsFile
'' Author    : beededea
'' Date      : 17/10/2019
'' Purpose   : get this tool's settings file
''---------------------------------------------------------------------------------------
''
'Private Sub locateDockSettingsFile()
'    Dim dockSettingsDir As String
'
'    On Error GoTo locateDockSettingsFile_Error
'    'If debugflg = 1 Then debugLog "%locateDockSettingsFile"
'
'    dockSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\steamyDock" ' just for this user alone
'    dockSettingsFile = dockSettingsDir & "\docksettings.ini" ' the third config option for steamydock alone
'
'    'if the folder does not exist then create the folder
'    If Not DirExists(dockSettingsDir) Then
'        MkDir dockSettingsDir
'    End If
'
'    'if the settings.ini does not exist then create the file by copying
'    If Not FExists(dockSettingsFile) Then
'        If FExists(App.Path & "\defaultDockSettings.ini") Then
'            FileCopy App.Path & "\defaultDockSettings.ini", dockSettingsFile
'            MsgBox ("Creating default sample dock, feel free to Edit/Delete items as you require.")
'        End If
'    End If
'
'    'confirm the settings file exists, if not use the version in the app itself
'    If Not FExists(dockSettingsFile) Then
'            dockSettingsFile = App.Path & "\settings.ini"
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'locateDockSettingsFile_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure locateDockSettingsFile of Form dockSettings"
'
'End Sub





    
'---------------------------------------------------------------------------------------
' Procedure : readThemeConfiguration
' Author    : beededea
' Date      : 09/07/2020
' Purpose   : ' read the background theme settings from INI
'---------------------------------------------------------------------------------------
'
Private Sub readThemeConfiguration()

    On Error GoTo readThemeConfiguration_Error
    'If debugflg = 1 Then debugLog "%readThemeConfiguration"
    
    'background.ini
    
'    [Background]
'    Image = Milk2.png
'    LeftMargin = 8
'    TopMargin = 8
'    RightMargin = 8
'    BottomMargin = 8
'    Outside-LeftMargin = 8
'    Outside-TopMargin = 8
'    Outside-RightMargin = 8
'    Outside-BottomMargin = 8

    If validTheme = False Then Exit Sub

    rDThemeImage = GetINISetting("Background", "Image", rdThemeSkinFile)
    rDThemeLeftMargin = Val(GetINISetting("Background", "LeftMargin", rdThemeSkinFile))
    rDThemeTopMargin = Val(GetINISetting("Background", "TopMargin", rdThemeSkinFile))
    rDThemeRightMargin = Val(GetINISetting("Background", "RightMargin", rdThemeSkinFile))
    rDThemeBottomMargin = Val(GetINISetting("Background", "BottomMargin", rdThemeSkinFile))
    rDThemeOutsideLeftMargin = Val(GetINISetting("Background", "Outside-LeftMargin", rdThemeSkinFile))
    rDThemeOutsideTopMargin = Val(GetINISetting("Background", "Outside-TopMargin", rdThemeSkinFile))
    rDThemeOutsideRightMargin = Val(GetINISetting("Background", "Outside-RightMargin", rdThemeSkinFile))
    rDThemeOutsideBottomMargin = Val(GetINISetting("Background", "Outside-BottomMargin", rdThemeSkinFile))
    
    'validate the inputs
    
'    rDThemeImage ' must not be empty, set to a default
     'If rDThemeImage = "" Then
'    rDThemeLeftMargin ' must be a n ineteger value less than 20
'    rDThemeTopMargin ' must be an integer value less than 20
'    rDThemeRightMargin  ' must be an integer value less than 20
'    rDThemeBottomMargin  ' must be an integer value less than 20
'    rDThemeOutsideLeftMargin  ' must be an integer value less than 20
'    rDThemeOutsideTopMargin  ' must be an integer value less than 20
'    rDThemeOutsideRightMargin  ' must be an integer value less than 20
'    rDThemeOutsideBottomMargin ' must be an integer value less than 20

    ' separator.ini
    
'    [Separator]
'    Image = Separator.png
'    TopMargin = 3
'    BottomMargin = 3
    
    rDSeparatorImage = GetINISetting("Separator", "Image", rdThemeSeparatorFile)
    rDSeparatorTopMargin = Val(GetINISetting("Separator", "TopMargin", rdThemeSeparatorFile))
    rDSeparatorBottomMargin = Val(GetINISetting("Separator", "BottomMargin", rdThemeSeparatorFile))

'    rDSeparatorImage  '  must not be empty
'    rDSeparatorTopMargin  ' must be an integer value less than 20
'    rDSeparatorBottomMargin  ' must be an integer value less than 20

    ' the skin size is validated here as it is a skin variable, however, it is stored in the main configuration file and currently not the theme file.
    ' I am unsure whether we will continue to support the RD theme methods.
    
    If Val(rDSkinSize) <= 0 Or Val(rDSkinSize) > 177 Then
        sDSkinSize = 1
    End If

On Error GoTo 0
   Exit Sub

readThemeConfiguration_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readThemeConfiguration of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : locateThemeSettingsFile
' Author    : beededea
' Date      : 09/07/2020
' Purpose   : get the location of the dock's theme settings file
'---------------------------------------------------------------------------------------
'
Private Sub locateThemeSettingsFile()

    validTheme = False

    On Error GoTo locateThemeSettingsFile_Error
    'If debugflg = 1 Then debugLog "%readThemeConfiguration"
    
    ' read the default theme name from the setting file
    If rDtheme = vbNullString Then
        MsgBox ("Theme not set")
        Exit Sub
    End If
    
    ' if it exists set the theme file to the settings file found
    rdThemeSkinFile = App.Path & "\Skins\" & rDtheme & "\background.ini"
    rdThemeSeparatorFile = App.Path & "\Skins\" & rDtheme & "\separator.ini"
    ' test existence of the set theme file
    If Not FExists(rdThemeSkinFile) Then
        validTheme = False
        Exit Sub
    End If
    If Not FExists(rdThemeSeparatorFile) Then
        validTheme = False
        Exit Sub
    End If
 
    validTheme = True ' if we arrived this far the theme exists
    If validTheme = False Then
        MsgBox ("Selected Theme " & rDtheme & " does not exist within Steamydock")
    End If
    On Error GoTo 0
   Exit Sub

locateThemeSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure locateThemeSettingsFile of Form dock"
End Sub






' test on Win10
'
'    Dim ab As String
'    Dim b As Boolean
'
'    ab = "true"
'    b = CBool(ab)
'
'    ab = "True"
'    b = CBool(ab)
'
'    ab = "True"
'    b = CBool(LCase(ab))


'---------------------------------------------------------------------------------------
' Procedure : applyThemeSkinToDock
' Author    : beededea
' Date      : 13/08/2020
' Purpose   : draw the theme or skin behind the icons this method is not compatible with Rocketdock skins
'---------------------------------------------------------------------------------------
'
' Rocket Dock themeing method described for future implementation and improvement.
'
' In Rocketdock, the bg is used thusly;
' Starting from the original 118x 118 image it extracts a left hand crop of approx. 37 pixels and uses that as the left hand image
' then it takes a small sliver of three or so pixels from that same crop and scales it (stretches it rightward) for 150 pixels or so
' it appears that this image is blended or a gradient fade out is applied to the right hand portion
' it appears as if this is place on top of image 2 and the left is blended...
' from the original image the central section is taken, approx. another 18-20 pixels from the left hand side to the middle of the image
' this is then stretched to the centre of the dock.
' Either all of these GDI+ functions are carried out or these stetching, blending operations are carried out using a 3rd party graphics library
' the same is then performed for the right hand side of the dock.

' Needless overkill, it has been replaced with a left hand image, a right hand image and a centre image, rectangular and 2000px wide that is cropped to fit.

' There are three issues to resolve:
' i.  the bottom few pixels that trigger the dock at the bottom need to be transposed to the top
' ii. the dock theme needs to be accounted for at the top position in dothedocktheme
' iii.the busy cog needs to appear on the bottom of the icons

Private Sub applyThemeSkinToDock(ByVal dockSkinStart As Long, ByVal dockSkinWidth As Long)
    
    Dim thiskey As String: thiskey = vbNullString
    Dim bgThemeTopPxls As Long: bgThemeTopPxls = 0
    
    On Error GoTo applyThemeSkinToDock_Error
    
    dockSkinStart = iconPosLeftPxls - (sDSkinSize)
    dockSkinWidth = (rdIconMaximum * iconSizeSmallPxls) + iconSizeLargePxls * 2
    
    ' .49 DAEB 01/04/2021 frmMain.frm added the vertical adjustment for sliding in and out STARTS
    If autoSlideMode = "slideout" Then
        If dockPosition = vbtop Then
            ' set the skin to a position above the icons and modified in the Y axis by the slideTimer
            bgThemeTopPxls = (dockUpperMostPxls) - xAxisModifier '.nn
        Else ' dockPosition = vbBottom
            ' set the skin to a position above the small icons and modified in the Y axis by the slideTimer if the slider timer is running
            bgThemeTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
        End If
    ElseIf autoSlideMode = "slidein" Then
        If dockPosition = vbtop Then
            ' set the skin to a position above the icons and modified in the Y axis by the slideTimer
            bgThemeTopPxls = (dockUpperMostPxls) + xAxisModifier '.nn
        Else ' dockPosition = vbBottom
            ' set the skin to a position above the small icons and modified in the Y axis by the slideTimer if the slider timer is running
            bgThemeTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
        End If
    Else
        If dockPosition = vbtop Then
            ' set the skin to a position above the icons
            bgThemeTopPxls = (dockUpperMostPxls)  '.nn
        Else ' dockPosition = vbBottom
            ' set the skin to a position above the small icons
            bgThemeTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls))
        End If
    End If
    ' .49 DAEB 01/04/2021 frmMain.frm added the vertical adjustment for sliding in and out ENDS
    
    
    ' display the start theme left hand
    thiskey = "sDSkinLeft" & "ResizedImg" & LTrim$(Str$(sDSkinSize))
    updateDisplayFromDictionary collLargeIcons, vbNullString, thiskey, (dockSkinStart), ((bgThemeTopPxls + iconSizeSmallPxls / 2) - sDSkinSize / 2), (sDSkinSize), (sDSkinSize)

    ' display the middle section in one 2000px length already cropped to the calculated dock size
    thiskey = "sDSkinMid" & "ResizedImg" & LTrim$(Str$(sDSkinSize))
    updateDisplayFromDictionary collLargeIcons, vbNullString, thiskey, (dockSkinStart + sDSkinSize), ((bgThemeTopPxls + iconSizeSmallPxls / 2) - sDSkinSize / 2), (dockSkinWidth), (sDSkinSize)

   ' display the end theme background
    thiskey = "sDSkinRight" & "ResizedImg" & LTrim$(Str$(sDSkinSize))
    updateDisplayFromDictionary collLargeIcons, vbNullString, thiskey, ((dockSkinStart + dockSkinWidth + sDSkinSize)), ((bgThemeTopPxls + iconSizeSmallPxls / 2) - sDSkinSize / 2), (sDSkinSize), (sDSkinSize)

   On Error GoTo 0
   Exit Sub

applyThemeSkinToDock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure applyThemeSkinToDock of Form dock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BounceIn
' Author    : Olaf Schmidt
' Date      : 13/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function BounceIn(ByVal t As Double)
   On Error GoTo BounceIn_Error

  BounceIn = 1 - BounceOut(1 - t) ' return

   On Error GoTo 0
   Exit Function

BounceIn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BounceIn of Form dock"
End Function
'---------------------------------------------------------------------------------------
' Procedure : BounceOut
' Author    : Olaf Schmidt
' Date      : 13/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function BounceOut(ByVal t As Double)
   On Error GoTo BounceOut_Error

  If (t < (1 / 2.75)) Then BounceOut = 7.5625 * t ^ 2: Exit Function
  If (t < (2 / 2.75)) Then t = t - 1.5 / 2.75: BounceOut = 7.5625 * t ^ 2 + 0.75: Exit Function
  If (t < (2.5 / 2.75)) Then t = t - 2.25 / 2.75: BounceOut = 7.5625 * t ^ 2 + 0.9375: Exit Function
  t = t - 2.625 / 2.75: BounceOut = 7.5625 * t ^ 2 + 0.984375 ' return

   On Error GoTo 0
   Exit Function

BounceOut_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BounceOut of Form dock"
End Function

'Function BounceOut2(t)

' The above function runs faster than this one...

'    If (t < 1 / 2.75) Then BounceOut2 = 7.5625 * t * t: Exit Function
'    If (t < 2 / 2.75) Then BounceOut2 = 7.5625 * (t = t - 1.5 / 2.75) * t + 0.75: Exit Function
'    If (t < 2.5 / 2.75) Then BounceOut2 = 7.5625 * (t = t - 2.25 / 2.75) * t + 0.9375: Exit Function
'    BounceOut2 = 7.5625 * (t = t - 2.625 / 2.75) * t + 0.984375: Exit Function


'End Function




''---------------------------------------------------------------------------------------
'' Procedure : BounceInOut
'' Author    : Olaf Schmidt
'' Date      : 13/09/2020
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Function BounceInOut(t)
'   On Error GoTo BounceInOut_Error
'
'  If t < 0.5 Then BounceInOut = BounceIn(t * 2) / 2 Else BounceInOut = (BounceOut(t * 2 - 1) + 1) / 2
'
'   On Error GoTo 0
'   Exit Function
'
'BounceInOut_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BounceInOut of Form dock"
'End Function






'---------------------------------------------------------------------------------------
' Procedure : resolveVB6SizeBug
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : VB6 has a bug - it should return 28800 on my screen but often returns 16200 when a game runs full screen, changing the resolution
'             the screen width determination is incorrect, the API call below resolves this.
'             NOTE: the dock program is the size of the whole screen
'---------------------------------------------------------------------------------------
'
Private Sub resolveVB6SizeBug()

   On Error GoTo resolveVB6SizeBug_Error
   
    If debugflg = 1 Then debugLog "% sub resolveVB6SizeBug"


'    screenWidthTwips = 0 ' private wide vars
'    screenHeightTwips = 0
'    screenWidthPixels = 0
'    screenHeightPixels = 0
    
'    Me.Height = Screen.Height '16200 correct
'    Me.Width = Screen.Width ' 16200 < VB6 bug here


    screenHeightTwips = GetDeviceCaps(dock.hdc, VERTRES) * screenTwipsPerPixelY
    screenWidthTwips = GetDeviceCaps(dock.hdc, HORZRES) * screenTwipsPerPixelX
    
    screenHeightPixels = GetDeviceCaps(dock.hdc, VERTRES)
    screenWidthPixels = GetDeviceCaps(dock.hdc, HORZRES)
        
    'set the form to the size of the whole monitor, has to be done in twips
    Me.Height = screenHeightTwips
    Me.Width = screenWidthTwips

    'Me.Left = 1000
    
   On Error GoTo 0
   Exit Sub

resolveVB6SizeBug_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure resolveVB6SizeBug of Form dock"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setLocalConfigurationVars
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : configuration private numeric vars that are easier to manipulate throughout the program than the string equivalents
'---------------------------------------------------------------------------------------
'
Private Sub setLocalConfigurationVars()
   On Error GoTo setLocalConfigurationVars_Error
   
    If debugflg = 1 Then debugLog "% sub setLocalConfigurationVars"

    iconSizeSmallPxls = Val(rDIconMin) ' in dock icon size to display
    iconSizeLargePxls = Val(rdIconMax)  ' maximum dock icon size to display

   On Error GoTo 0
   Exit Sub

setLocalConfigurationVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setLocalConfigurationVars of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : initialiseGDIStartup
' Author    : beededea
' Date      : 18/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGDIStartup()
    ' Initialises GDI Plus
   On Error GoTo initialiseGDIStartup_Error
   
    If debugflg = 1 Then debugLog "% sub initialiseGDIStartup"

    gdipInit.GDIPlusVersion = 1
    If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
        MsgBox "Error loading GDI+", vbCritical
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

initialiseGDIStartup_Error:

    If debugflg = 1 Then debugLog "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGDIStartup of Form dock"

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGDIStartup of Form dock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : createDictionaryObjects
' Author    : beededea
' Date      : 18/09/2020
' Purpose   :  This initialises each VB collection object where the image bitmaps will be stored
'              This method of using the scripting dictionary as an object collection was suggested by Olaf Schmidt.
'---------------------------------------------------------------------------------------
'

Private Sub createDictionaryObjects()
    
   On Error GoTo createDictionaryObjects_Error
   
    If debugflg = 1 Then debugLog "% sub createDictionaryObjects"
   
    ' dictionary for the larger icons
    Set collLargeIcons = CreateObject("Scripting.Dictionary")
    collLargeIcons.CompareMode = 1 'case-insenitive Key-Comparisons
    
    'dictionary for the smaller icons
    Set collSmallIcons = CreateObject("Scripting.Dictionary")
    collSmallIcons.CompareMode = 1 'case-insenitive Key-Comparisons
    
    ' .64 DAEB 30/04/2021 frmMain.frm Deleted the temporary collection, now unused.
    'third temporary dictionary that is used for temporary storage whilst resizing the collection
'    Set collTemporaryIcons = CreateObject("Scripting.Dictionary")
'    collTemporaryIcons.CompareMode = 1 'case-insenitive Key-Comparisons

   On Error GoTo 0
   Exit Sub

createDictionaryObjects_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createDictionaryObjects of Form dock"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : createGDIPlusElements
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : sets bmpInfo object to create a bitmap the whole screen size and get a handle to the Device Context
'---------------------------------------------------------------------------------------
'
Private Sub createGDIPlusElements()
    ' sets the bmpInfo object containing data to create a bitmap the whole screen size
    ' used later when creating DIB section of the correct size, width &c
    On Error GoTo createGDIPlusElements_Error
   
    If debugflg = 1 Then debugLog "% sub createGDIPlusElements"

    bmpInfo.bmpHeader.Size = Len(bmpInfo.bmpHeader)
    bmpInfo.bmpHeader.BitCount = 32
    bmpInfo.bmpHeader.Height = Me.ScaleHeight
    
    bmpInfo.bmpHeader.Width = screenWidthPixels  ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    
    bmpInfo.bmpHeader.Planes = 1
    bmpInfo.bmpHeader.SizeImage = bmpInfo.bmpHeader.Width * bmpInfo.bmpHeader.Height * (bmpInfo.bmpHeader.BitCount / 8)
    
    ' A handle to the Device Context (HDC) is obtained before output is written and then released after elements have been written.
    ' Get a device context compatible with the screen
    dcMemory = CreateCompatibleDC(hdcScreen)

    ' A device context is a generalized rendering abstraction. It serves as a proxy between your rendering code and the output device.
    ' It allows you to use the same rendering code, regardless of the destination; the low-level details are handled for you,
    ' depending on the output device, including clipping, scaling, and viewport translation.
    
   On Error GoTo 0
   Exit Sub

createGDIPlusElements_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createGDIPlusElements of Form dock"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setAutoHide
' Author    : beededea
' Date      : 18/09/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setAutoHide()
    ' allows the autohide check timer to lower the dock after a short delay during startup
   On Error GoTo setAutoHide_Error
   
    If debugflg = 1 Then debugLog "% sub setAutoHide"

    If rDAutoHide = "1" Then
        autoHideChecker.Interval = 1
        dockLoweredTime = TimeValue(Now)
    End If

   On Error GoTo 0
   Exit Sub

setAutoHide_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAutoHide of Form dock"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setUpProcessTimers
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : set up the timers that check to see if each process is running
'---------------------------------------------------------------------------------------
'
Private Sub setUpProcessTimers()
    
    ' start the 10s timer that checks to see if each process is running
   On Error GoTo setUpProcessTimers_Error
   
    If debugflg = 1 Then debugLog "% sub setUpProcessTimers"

    processTimer.Interval = Val(rDRunAppInterval) * 1000
    If rDShowRunning = "1" Then
        processTimer.Enabled = True
    Else
        processTimer.Enabled = False
    End If
    
    initiatedProcessTimer.Enabled = True ' this was enabled by default on a 5 second timer but is now here with a reduced interval, this manual start giving time to the whole
                                         ' tool to get its stuff done before it runs.
    targetExistsTimer.Enabled = True
    
   On Error GoTo 0
   Exit Sub

setUpProcessTimers_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setUpProcessTimers of Form dock"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : GetMonitorCount
' Author    : beededea
' Date      : 02/03/2020
' Purpose   : The number of monitors is known by RD
'---------------------------------------------------------------------------------------
'
Private Function GetMonitorCount() As Integer

    ' variables declared
   Dim NumberOfMonitors As Integer: NumberOfMonitors = 0


   NumberOfMonitors = 1

   On Error GoTo GetMonitorCount_Error
   'If debugflg = 1 Then debugLog "%GetMonitorCount"

   NumberOfMonitors = GetSystemMetrics(SM_CMONITORS)

   GetMonitorCount = NumberOfMonitors

   On Error GoTo 0
   Exit Function

GetMonitorCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetMonitorCount of Form dockSettings"

End Function



'---------------------------------------------------------------------------------------
' Procedure : autoHideChecker_Timer
' Author    : beededea
' Date      : 01/05/2020
' Purpose   : checks to see if the dock needs to be hidden, if so, initiates one of the hider timers
'             runs from the outset on a half second interval controls when the dock is lowered
'---------------------------------------------------------------------------------------
'
Private Sub autoHideChecker_Timer()
   Dim secondDiff As Integer: secondDiff = 0
   'On Error GoTo autoHideChecker_Error
   ''If debugflg = 1 Then debugLog "%autoHideChecker"

        If rDAutoHide = "1" And animatedIconsRaised = False And dockHidden = False Then
            autoHideChecker.Interval = 500
            If dockLoweredTime <> "00:00:00" Then
                secondDiff = DateDiff("s", dockLoweredTime, TimeValue(Now))
            End If
            ' time since the dock was lowered
            If secondDiff > (Val(rDAutoHideDelay) / 1000) Then
                If Val(sDAutoHideType) = 0 Then ' fade animation
                    autoHideMode = "fadeout"
                    autoFadeOutTimer.Enabled = True ' .nn
                ElseIf Val(sDAutoHideType) = 1 Then ' slide animation as per Rocketdock
                    'xAxisModifier = 0 ' .nn not needed and commented out to prevent slider oscillating
                    autoSlideMode = "slideout"
                    autoSlideOutTimer.Enabled = True
                ElseIf Val(sDAutoHideType) = 2 Then 'instant invisible
                    ' set the opacity of the whole dock, used to display solidly and for instant autohide
                    funcBlend32bpp.SourceConstantAlpha = 0
                    bDrawn = False
                    smallDockBeenDrawn = False ' allows the dock to redraw on the next response cycle
                    Exit Sub
                End If
            End If
        End If

   On Error GoTo 0
   Exit Sub

autoHideChecker_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoHideChecker of Form dock"
End Sub
' 24/01/2021 .09 DAEB created a separate fade in timer and function
'---------------------------------------------------------------------------------------
' Procedure : autoFadeInTimer
' Author    : beededea
' Date      : 18/05/2020
' Purpose   : the timer's interval is set as a slider within dock settings
'             this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha
'---------------------------------------------------------------------------------------
'
Private Sub autoFadeInTimer_Timer()
    Dim newDockOpacity As Integer: newDockOpacity = 0
    Dim autoHideGranularity  As Double: autoHideGranularity = 0
    
    On Error GoTo autoFadeInTimer_Error
    
    newDockOpacity = 0
    dockOpacity = 100
    
    autoFadeOutTimer.Enabled = False
    
    responseTimer.Interval = 5  ' this frequency is also maintained within the responseTimer event. This event does the animation that actually
                                ' accomplishes the fade
                                ' it stays at this frequency until the fadeTimer is done when it reverts to 200
                                ' it is important as this maintains the smoothness of the fadeout.
    autoFadeInTimerCount = autoFadeInTimerCount + 10  ' .10 DAEB 25/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer

    If rDPopupDelay = 0 Then rDPopupDelay = 1 ' .24 DAEB frmMain.frm 09/02/2021 handling any potential divide by zero

    autoHideGranularity = dockOpacity / rDPopupDelay ' set the factor by which the timer should decrease the opacity
    newDockOpacity = 1 + (autoFadeInTimerCount * autoHideGranularity) ' .10 DAEB 25/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
    
    If newDockOpacity > 100 Then newDockOpacity = 100 ' funcBlend32bpp.SourceConstantAlpha does not like values less than 0
    
    ' set the increasingly increased opacity of the whole dock
    funcBlend32bpp.SourceConstantAlpha = 255 * newDockOpacity / 100
    
    If autoFadeInTimerCount >= Val(rDPopupDelay) Then ' .10 DAEB 25/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
        ' ensure the opacity of the whole dock is solid
        funcBlend32bpp.SourceConstantAlpha = 255
        dockHidden = False
    
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        autoFadeInTimer.Enabled = False
        autoFadeInTimerCount = 0 ' .10 DAEB 25/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
    End If
    
    bDrawn = False
    smallDockBeenDrawn = False ' set a flag to allow the animation to redraw
            
   On Error GoTo 0
   Exit Sub

autoFadeInTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoFadeInTimer of Form dock"
    
End Sub

' .01 24/01/2021 DAEB modified to handle the new timer name
'---------------------------------------------------------------------------------------
' Procedure : autoFadeOutTimer
' Author    : beededea
' Date      : 18/05/2020
' Purpose   : the timer's interval is set as a slider within dock settings
'             this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha
'---------------------------------------------------------------------------------------
'
Private Sub autoFadeOutTimer_Timer()

    Dim newDockOpacity As Integer: newDockOpacity = 0
    Dim autoHideGranularity  As Double: autoHideGranularity = 0
    
    On Error GoTo autoFadeOutTimer_Error
    
    newDockOpacity = 0
    dockOpacity = 100
    
    If animatedIconsRaised = True Then
        autoFadeOutTimer.Enabled = False
        Exit Sub
    End If
    
    If autoFadeInTimer.Enabled = True Then
        autoFadeOutTimer.Enabled = False
        Exit Sub
    End If
        
    responseTimer.Interval = 5  ' this frequency is also maintained within the responseTimer event. This event does the animation that actually
                                ' accomplishes the fade
                                ' it stays at this frequency until the fadeTimer is done when it reverts to 200
                                ' it is important as this maintains the smoothness of the fadeout.
    autoFadeOutTimerCount = autoFadeOutTimerCount + 10
    If rDAutoHideTicks = 0 Then rDAutoHideTicks = 1 ' .24 DAEB frmMain.frm 09/02/2021 handling any potential divide by zero
    autoHideGranularity = dockOpacity / rDAutoHideTicks ' set the factor by which the timer should decrease the opacity
    
    ' 24/01/2021 .08 DAEB removed the fade in functions from the fade out subroutine

    newDockOpacity = 100 - (autoFadeOutTimerCount * autoHideGranularity)
    If newDockOpacity < 0 Then newDockOpacity = 0 ' funcBlend32bpp.SourceConstantAlpha does not like values less than 0
    
    ' set the increasingly reduced/increased opacity of the whole dock
    funcBlend32bpp.SourceConstantAlpha = 255 * newDockOpacity / 100
    
    If autoFadeOutTimerCount >= Val(rDAutoHideTicks) Then
        ' ensure the opacity of the whole dock is zero
        funcBlend32bpp.SourceConstantAlpha = 0
        dockHidden = True
    
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        autoFadeOutTimer.Enabled = False
        autoFadeOutTimerCount = 0
        
        currentDockTopPxls = screenHeightPixels - 10
    End If
    
    bDrawn = False
    smallDockBeenDrawn = False ' set a flag to allow the animation to redraw
            
   On Error GoTo 0
   Exit Sub

autoFadeOutTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoFadeOutTimer of Form dock"
    
End Sub
''---------------------------------------------------------------------------------------
'' Procedure : autoFadeOutTimer
'' Author    : beededea
'' Date      : 18/05/2020
'' Purpose   : the timer's interval is set as a slider within dock settings
''             this routine simply gradually sets the opacity of the dock when triggered using funcBlend32bpp.SourceConstantAlpha
''---------------------------------------------------------------------------------------
''
'Private Sub autoFadeOutTimer_Timer()
'
'    Dim newDockOpacity As Integer : = 0
'    Dim autoHideGranularity  As double: = 0
'
'    On Error GoTo autoFadeOutTimer_Error
'
'
'    newDockOpacity = 0
'    dockOpacity = 100
'
'    If autoHideMode = "fadeout" And animatedIconsRaised = True Then
'        autoHideMode = "fadein" 'if the cursor enters the dock during a fade out this will turn it into a fade in
'    End If
'
'    responseTimer.Interval = 5  ' this frequency is also maintained within the responseTimer event. This event does the animation that actually
'                                ' accomplishes the fade
'                                ' it stays at this frequency until the fadeTimer is done when it reverts to 200
'                                ' it is important as this maintains the smoothness of the fadeout.
'    autoFadeOutTimerCount = autoFadeOutTimerCount + 10
'    autoHideGranularity = dockOpacity / rDAutoHideTicks ' set the factor by which the timer should decrease the opacity
'
'    If autoHideMode = "fadeout" Then
'        newDockOpacity = 100 - (autoFadeOutTimerCount * autoHideGranularity)
'    Else
'        newDockOpacity = 1 + (autoFadeOutTimerCount * autoHideGranularity)
'    End If
'
'    If newDockOpacity < 0 Then newDockOpacity = 0 ' funcBlend32bpp.SourceConstantAlpha does not like values less than 0
'    If newDockOpacity > 100 Then newDockOpacity = 100 ' funcBlend32bpp.SourceConstantAlpha does not like values less than 0
'
'    ' set the increasingly reduced/increased opacity of the whole dock
'    funcBlend32bpp.SourceConstantAlpha = 255 * newDockOpacity / 100
'
'    If autoFadeOutTimerCount >= Val(rDAutoHideTicks) Then
'        If autoHideMode = "fadeout" Then
'            ' ensure the opacity of the whole dock is zero
'            funcBlend32bpp.SourceConstantAlpha = 0
'            dockHidden = True
'        Else
'            ' ensure the opacity of the whole dock is solid
'            funcBlend32bpp.SourceConstantAlpha = 255
'            dockHidden = False
'        End If
'
'        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
'        autoFadeOutTimer.Enabled = False
'        autoFadeOutTimerCount = 0
'    End If
'
'    bDrawn = False
'    smallDockBeenDrawn = False ' set a flag to allow the animation to redraw
'
'   On Error GoTo 0
'   Exit Sub
'
'autoFadeOutTimer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoFadeOutTimer of Form dock"
'
'End Sub

' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions STARTS
'---------------------------------------------------------------------------------------
' Procedure : autoSlideOutTimer
' Author    : beededea
' Date      : 25/09/2020
' Purpose   : slide the dock in the Y axis
'---------------------------------------------------------------------------------------
'
Private Sub autoSlideOutTimer_Timer()
    Dim autoSlideGranularity  As Double: autoSlideGranularity = 0
    Dim amountToSlidePxls As Integer: amountToSlidePxls = 0
    
    amountToSlidePxls = 25

    On Error GoTo autoSlideOutTimer_Error

    If animatedIconsRaised = True Then
        autoSlideOutTimer.Enabled = False
        Exit Sub
    End If
    
    If autoSlideInTimer.Enabled = True Then
        autoSlideOutTimer.Enabled = False
        Exit Sub
    End If
        
    amountToSlidePxls = 25

    'If animatedIconsRaised = True Then autoSlideMode = "slidein" 'if the cursor enters the dock during a fade out this will turn it into a fade in

    responseTimer.Interval = 5  ' this frequency is also maintained within the responseTimer event. This event does the animation that actually
                                ' accomplishes the fade
                                ' it stays at this frequency until the fadeTimer is done when it reverts to 200
                                ' it is important as this maintains the smoothness of the slideout.
    autoSlideOutTimerCount = autoSlideOutTimerCount + 5  '10ms
    If rDAutoHideTicks = 0 Then rDAutoHideTicks = 1 ' .24 DAEB frmMain.frm 09/02/2021 handling any potential divide by zero
    autoSlideGranularity = amountToSlidePxls / rDAutoHideTicks ' set the factor by which the timer should slide out the dock
    
    ' modify the whole dock in the Y axis here using
    xAxisModifier = xAxisModifier + (autoSlideOutTimerCount * autoSlideGranularity)
    
    ' check whether the sliding dock is below the level of the screen
    If iconCurrentTopPxls - 10 > (screenHeightPixels) Then ' the extra 10 pixels is to ensure the theme is off screen too
        autoSlideOutTimer.Enabled = False
        autoSlideOutTimerCount = 0
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        slideOutFlag = True ' we need a flag to state that the dock has 'slidden' to determine the position just the first 10 pixels of the dock
        dockHidden = True
    End If
    
    If autoSlideOutTimerCount >= Val(rDAutoHideTicks) Then
        ' ensure the opacity of the whole dock is zero
        'funcBlend32bpp.SourceConstantAlpha = 0
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        autoSlideOutTimer.Enabled = False
        autoSlideOutTimerCount = 0
        slideOutFlag = True ' we need a flag to state that the dock has 'slidden'
        dockHidden = True
    End If

    bDrawn = False
    smallDockBeenDrawn = False ' set a flag to allow the animation to redraw

    On Error GoTo 0
    Exit Sub

autoSlideOutTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoSlideOutTimer_ of Form dock"

End Sub
' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions ENDS


' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions STARTS
'---------------------------------------------------------------------------------------
' Procedure : autoSlideInTimer_Timer
' Author    : beededea
' Date      : 25/09/2020
' Purpose   : slide the dock in the Y axis
'---------------------------------------------------------------------------------------
'
Private Sub autoSlideInTimer_Timer()
    Dim autoSlideGranularity  As Double: autoSlideGranularity = 0
    Dim amountToSlidePxls As Integer: amountToSlidePxls = 0
    
    On Error GoTo autoSlideInTimer_Error
    
    amountToSlidePxls = 25
    autoSlideOutTimer.Enabled = False
    slideOutFlag = False
 
    'animateTimer.Enabled = True
 
    responseTimer.Interval = 5  ' this frequency is also maintained within the responseTimer event. This event does the animation that actually
                                ' accomplishes the fade
                                ' it stays at this frequency until the fadeTimer is done when it reverts to 200
                                ' it is important as this maintains the smoothness of the slideout.
    autoSlideInTimerCount = autoSlideInTimerCount + 5  '10ms
    If rDAutoHideTicks = 0 Then rDAutoHideTicks = 1 ' .24 DAEB frmMain.frm 09/02/2021 handling any potential divide by zero

    autoSlideGranularity = amountToSlidePxls / rDAutoHideTicks ' set the factor by which the timer should slide out the dock
    
    If iconCurrentTopPxls < 860 Then ' .nn this is the bug just here
        iconCurrentTopPxls = 860 '.nn
        autoSlideInTimer.Enabled = False
        autoSlideInTimerCount = 0
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        dockHidden = False
        autoSlideMode = vbNullString 'nn Set to nothing to ensure that the modifiedslide position is not taken into account when redrawing the static loop
    Else
        ' modify the whole dock in the Y axis here using .nn
        xAxisModifier = xAxisModifier + (autoSlideInTimerCount * autoSlideGranularity)
    End If
    
    If autoSlideInTimerCount >= Val(rDAutoHideTicks) Then
        ' ensure the opacity of the whole dock is zero
        'funcBlend32bpp.SourceConstantAlpha = 0
        responseTimer.Interval = 200 ' return the responseTimer interval to normal, may not be necessary here
        autoSlideInTimer.Enabled = False
        autoSlideInTimerCount = 0
        slideOutFlag = True ' we need a flag to state that the dock
        dockHidden = True
    End If

    bDrawn = False
    smallDockBeenDrawn = False ' set a flag to allow the animation to redraw

    On Error GoTo 0
    Exit Sub

autoSlideInTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure autoSlideInTimer of Form dock"

End Sub
' .28 DAEB frmMain.frm 16/02/2021 Seperated the autoSlide Timers to in and out versions ENDS


' .14 DAEB frmMain.frm 27/01/2021 Add new subroutine to make the dock transparent and shutdown timers
'---------------------------------------------------------------------------------------
' Procedure : HideDockNow
' Author    : beededea
' Date      : 25/01/2021
' Purpose   : hides the dock when the user presses F11 or when the menu option is selected to hide, sets the alpha and
'             stops all timers
'---------------------------------------------------------------------------------------
'
Public Sub HideDockNow()
   On Error GoTo HideDockNow_Error
    
    dock.nMinuteExposeTimer.Enabled = True ' timers are associated with forms, stupid VB6
    hideDockForNMinutes = True
    
    funcBlend32bpp.SourceConstantAlpha = 0
    
    Call GdipDeleteGraphics(lngImage)  'The graphics may now be deleted

    'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
    UpdateLayeredWindow Me.hWnd, hdcScreen, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA
    
    responseTimer.Enabled = False
    animateTimer.Enabled = False
    autoFadeOutTimer.Enabled = False
    autoFadeInTimer.Enabled = False
    autoSlideOutTimer.Enabled = False
    autoSlideInTimer.Enabled = False

   On Error GoTo 0
   Exit Sub

HideDockNow_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HideDockNow of Form dock"
    
End Sub
' .14 DAEB frmMain.frm 27/01/2021 Add new subroutine to make the dock transparent and shutdown timers

' .15 frmMain.frm STARTS DAEB 27/01/2021 Add new subroutine to show the dock after it has been manually hidden by the user
'---------------------------------------------------------------------------------------
' Procedure : ShowDockNow
' Author    : beededea
' Date      : 26/01/2021
' Purpose   : Shows the dock after it has been manually hidden by the user
'---------------------------------------------------------------------------------------
'
Public Sub ShowDockNow()
   On Error GoTo ShowDockNow_Error

        nMinuteExposeTimer.Enabled = False ' timers are associated with forms, stupid VB6
        nMinuteExposeTimerCount = 0
        hideDockForNMinutes = False
        
        funcBlend32bpp.SourceConstantAlpha = 255
        
        Call GdipDeleteGraphics(lngImage)  'The graphics may now be deleted

        'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
        UpdateLayeredWindow Me.hWnd, hdcScreen, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA
        
        responseTimer.Enabled = True

   On Error GoTo 0
   Exit Sub

ShowDockNow_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowDockNow of Form dock"
End Sub
' .15 frmMain.frm ENDS DAEB 27/01/2021 Add new subroutine to show the dock after it has been manually hidden by the user


'---------------------------------------------------------------------------------------
' Procedure : nMinuteExposeTimer
' Author    : beededea
' Date      : 25/01/2021
' Purpose   : Causes the dock to re-appear in its default state after N mins
'             Shows the dock after it has been manually hidden by the user
'---------------------------------------------------------------------------------------
'
Private Sub nMinuteExposeTimer_Timer()
    Dim itIs As Boolean: itIs = False         ' .84 DAEB 20/07/2021 frmMain.frm Added prevention of the dock returning until the hiding application is no longer running.

    On Error GoTo nMinuteExposeTimer_Error

    ' when a timer is initiated it runs immediately, we want it to do nothing until the 10 mins is up
    
    ' the default timer interval is 60000 milliseconds or 60 seconds,
    ' every 60 seconds it increments the nMinuteExposeTimerCount by one
    
    ' reason for this is that a VB6 timer can only extend up to 65 secs/65000 millisecs,
    
    ' .52 DAEB 09/04/2021 frmMain.frm add code to increase the timer to 120 minutes
        
    ' if both the timer set value is greater than 65 and the current count is at the max then
    ' stop and restart the timer
    If Val(sDContinuousHide) >= 65 And nMinuteExposeTimerCount = 65 Then
        nMinuteExposeTimer.Enabled = False
        nMinuteExposeTimerCount = 0
        nMinuteExposeTimer.Enabled = True
        Exit Sub ' exit as the timer will start immediately and the count will be incremented on that very same run
    End If

    If nMinuteExposeTimerCount <= Val(sDContinuousHide) - 1 Then  ' .16 DAEB frmMain.frm 27/01/2021 Added the user set parameter sDContinuousHide
        If Not nMinuteExposeTimerCount = 65 Then     ' .52 DAEB 09/04/2021 frmMain.frm add code to increase the timer to 120 minutes
            nMinuteExposeTimerCount = nMinuteExposeTimerCount + 1
        End If
        Exit Sub
    Else
        ' .84 DAEB 20/07/2021 frmMain.frm Added prevention of the dock returning until the hiding application is no longer running.
        
        ' .nn DAEB 16/04/2022 frmMain.frm was the dock hidden by the running of a utility with the hide dock flag set?
        If autoHideProcessName <> vbNullString Then
            ' check to see if the process that hid the dock is still running
            ' the dock will not automatically appear until the process that hid it has finished (full screen games)
            itIs = IsRunning(autoHideProcessName, vbNull)
            If itIs = True Then
                ' the timer will continue to run
                Exit Sub
            Else
                autoHideProcessName = vbNullString
                Call ShowDockNow
            End If
        Else
            Call ShowDockNow ' normal timed run, just show the dock
        End If
    End If
    

   On Error GoTo 0
   Exit Sub

nMinuteExposeTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure nMinuteExposeTimer of Form dock"
    
End Sub




'---------------------------------------------------------------------------------------
' Procedure : disableAdmin
' Author    : beededea
' Date      : 28/01/2021
' Purpose   : turn off the run as administrator option for XP
'---------------------------------------------------------------------------------------
'
Private Sub disableAdmin()
   On Error GoTo disableAdmin_Error
   
    If debugflg = 1 Then debugLog "% sub disableAdmin"

    If InStr(WindowsVer, "Windows XP") <> 0 Then
        menuForm.mnuAdmin.Enabled = False
    End If

   On Error GoTo 0
   Exit Sub

disableAdmin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure disableAdmin of Form dock"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : Retn
' Author    : La Volpe
' Date      : 22/02/2021
' Purpose   : ' .31 DAEB 03/03/2021 frmMain.frm Check return value from any GDI++ function
'---------------------------------------------------------------------------------------
'
'Private Sub Retn(GdipReturn As Long)
'    ' Just to check for any errors.
'    '
'   On Error GoTo Retn_Error
'
'    If GdipReturn = OK Then Exit Sub
'                                        Debug.Print "GDI+ Error:  ";
'    Select Case GdipReturn
'    Case GenericError:                  Debug.Print "Generic Error"
'    Case InvalidParameter:              Debug.Print "Invalid Parameter/Argument"
'    Case OutOfMemory:                   Debug.Print "Out Of Memory"
'    Case ObjectBusy:                    Debug.Print "Object Busy, already in use in another thread"
'    Case InsufficientBuffer:            Debug.Print "Insufficient Buffer, buffer specified as an argument in the API call is not large enough"
'    Case NotImplemented:                Debug.Print "Method Not Implemented"
'    Case Win32Error:                    Debug.Print "Win32 Error"
'    Case WrongState:                    Debug.Print "Wrong State"
'    Case Aborted:                       Debug.Print "Method Aborted"
'    Case FileNotFound:                  Debug.Print "File Not Found"
'    Case ValueOverflow:                 Debug.Print "Value Overflow, arithmetic operation that produced a numeric overflow"
'    Case AccessDenied:                  Debug.Print "Access Denied"
'    Case UnknownImageFormat:            Debug.Print "Unknown Image Format"
'    Case FontFamilyNotFound:            Debug.Print "Font Family Not Found"
'    Case FontStyleNotFound:             Debug.Print "Font Style Not Found"
'    Case NotTrueTypeFont:               Debug.Print "Not TrueType Font"
'    Case UnsupportedGdiplusVersion:     Debug.Print "Unsupported Gdiplus Version"
'    Case GdiplusNotInitialized:         Debug.Print "Gdiplus Not Initialized"
'    Case PropertyNotFound:              Debug.Print "Property Not Found, does not exist in the image"
'    Case PropertyNotSupported:          Debug.Print "Property Not Supported, not supported by the format of the image"
'    Case ProfileNotFound:               Debug.Print "Profile Not Found, color profile required to save an image in CMYK format was not found"
'    Case Else:                          Debug.Print "Error Not Specified"
'    End Select
'    '
'    Stop
'
'   On Error GoTo 0
'   Exit Sub
'
'Retn_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Retn of Form dock"
'End Sub




'
'
'
'
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3">
'    <assemblyIdentity
'        Version = "2002.10.0.25"
'        processorArchitecture = "X86"
'        name = "vb6.exe"
'        type="win32" />
'    <description>WindowsExecutable</description>
'    <dependency>
'        <dependentAssembly>
'            <assemblyIdentity
'                type="win32"
'                name = "Microsoft.Windows.Common-Controls"
'                Version = "6.0.0.0"
'                processorArchitecture = "X86"
'                publicKeyToken = "6595b64144ccf1df"
'                language="*" />
'        </dependentAssembly>
'    </dependency>
'    <asmv3:application>
'        <asmv3:windowsSettings xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">
'            <dpiAware>true</dpiAware>
'        </asmv3:windowsSettings>
'    </asmv3:application>
'    <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
'        <application>
'            <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}" />
'            <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}" />
'            <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}" />
'            <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}" />
'            <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}" />
'        </application>
'    </compatibility>
'</assembly>

'Private Sub sleepTimer_Timer()
'        ' .nn
'    ' The device went to sleep with the use of the sleep command, therefore
'    ' the dock has been raised but when the system restarted the mouse ended up outside
'    ' of the dock area. So, lower dock by redrawing the small icons.
'    If outsideDock = True And animatedIconsRaised = True Then
'        If msgCnt = 0 Then MsgBox "1. just woke from sleep"
'        msgCnt = 1
'    End If
'End Sub




'Private Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'   ' // capture messages //
'
'   Select Case wMsg
'
'         Case WM_POWERBROADCAST ' Sent to an application to notify it of power-management events.
'
'            ' show messages in listbox for testing
'            'FMsgSink.List1.AddItem "WM_POWERBROADCAST, wParam = " & wParam & " lParam = " & lParam
'
'            ' coming out of sleep mode would be...?
'            If wParam = enPowerBroadcastType.PBT_APMRESUMESUSPEND Then
'                ' do something here
'                'MsgBox "2. just woke from sleep"
'            End If
'
'            ' going in to sleep mode would be...?
'            If wParam = enPowerBroadcastType.PBT_APMSUSPEND Then
'
'
'                ' do something here
'            End If
'
'            ' .Do Not Remove!
'            WindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
'
'      Case Else
'         ' Default processing...Do Not Remove!
'         WindowProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
'
'   End Select
'
'End Function








' .58 DAEB 21/04/2021 frmMain.frm added timer and vars to check to see if the system has just emerged from sleep
'---------------------------------------------------------------------------------------
' Procedure : sleepTimer_Timer
' Author    : beededea
' Date      : 21/04/2021
' Purpose   : timer that stores the last time
' if the current time is greater than the last time stored by more than 30 seconds we can assume the system
' has been sent to sleep, if the two are significantly different then we reorganise the dock
'---------------------------------------------------------------------------------------
'
Private Sub sleepTimer_Timer()
    Dim strTimeNow As String: strTimeNow = vbNullString 'set a variable to compare for the NOW time
    Dim lngSecondsGap As Long: lngSecondsGap = 0  ' set a variable for the difference in time
    
    On Error GoTo sleepTimer_Timer_Error

    strTimeNow = Now()
    
    lngSecondsGap = DateDiff("s", strTimeThen, strTimeNow)

    If lngSecondsGap > 30 Then
        'MsgBox "System has just woken up from a sleep"
        MessageBox Me.hWnd, "System has just woken up from a sleep - animatedIconsRaised =" & animatedIconsRaised, "SteamyDock Information Message", vbOKOnly

        ' at this point we should lower the dock and redraw the small icons.
        'If animatedIconsRaised = True Then
        
        ' the dock thinks the animatedIconsRaised is false!
        'Call sequentialBubbleAnimation 'here deanie
        Call animateTimer_Timer
        strTimeThen = Now
    Else
        strTimeThen = Now
    End If
    

    On Error GoTo 0
    Exit Sub

sleepTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sleepTimer_Timer of Form dock"

End Sub


Public Sub clearCollections()
' it is NOT possible to reclaim the memory taken by a collection that has been emptied - so don't use this
' This is retained to remind me to NOT to research this again.

' clear down the collections
    collLargeIcons.RemoveAll
    collSmallIcons.RemoveAll
    'collTemporaryIcons.RemoveAll' .64 DAEB 30/04/2021 frmMain.frm Deleted the temporary collection, now unused.


    ' dictionary for the larger icons
    Set collLargeIcons = Nothing
    Set collLargeIcons = New Scripting.Dictionary

    'dictionary for the smaller icons
    Set collSmallIcons = Nothing
    Set collSmallIcons = New Scripting.Dictionary

    'third temporary dictionary that is used for temporary storage whilst resizing the collection
'    Set collTemporaryIcons = Nothing
'    Set collTemporaryIcons = New Scripting.Dictionary' .64 DAEB 30/04/2021 frmMain.frm Deleted the temporary collection, now unused.
    
    'collTemporaryIcons = New Scripting.Dictionary ' to do the SET NEW here, support for MS scripting must be enabled in project - references
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : transit
' Author    : beededea
' Date      : 17/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub transit(fromX, fromY, toX, toY)

    On Error GoTo transit_Error

    

    On Error GoTo 0
    Exit Sub

transit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure transit of Form dock"
End Sub

' .86 DAEB 08/12/2022 frmMain.frm Added new timer to inspect each target command in turn and highlight if missing.
'---------------------------------------------------------------------------------------
' Procedure : targetExistsTimer_Timer
' Author    : beededea
' Date      : 08/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub targetExistsTimer_Timer()
    On Error GoTo targetExistsTimer_Timer_Error

    Call checkTargetCommandValidity

    On Error GoTo 0
    Exit Sub

targetExistsTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure targetExistsTimer_Timer of Form dock"
            Resume Next
          End If
    End With
End Sub
' .86 DAEB 08/12/2022 frmMain.frm Added new timer to inspect each target command in turn and highlight if missing.
'---------------------------------------------------------------------------------------
' Procedure : checkTargetCommandValidity
' Author    : beededea
' Date      : 08/12/2022
' Purpose   : Checks each target command for validity and sets a flag in an array to place a red X on the icon.
'             Written using GOTOs as VB6 does not have a CONTINUE command, I will rewrite this.
'---------------------------------------------------------------------------------------
'
Private Sub checkTargetCommandValidity()
    Dim useloop As Integer: useloop = 0
    Dim thisCommand As String: thisCommand = vbNullString
    Dim userprof As String: userprof = vbNullString
    Dim folderPath As String: folderPath = vbNullString
    
    Dim fileSuffixArray() As String: ' fileSuffixArray() = vbNullString
    Dim executableFileString As String: executableFileString = vbNullString
    Dim suffixElement As Variant
    Dim testURL As String: testURL = vbNullString
    
    Dim pathString As String: pathString = vbNullString
    Dim pathArray() As String: ' pathArray() = vbNullString
    Dim pathElement As Variant
    Dim currentCommand As String: currentCommand = vbNullString
    
    On Error GoTo checkTargetCommandValidity_Error

    executableFileString = "com cmd msc cpl bat exe"
    pathString = Environ$("path")
    
    For useloop = 0 To rdIconMaximum
        targetExistsArray(useloop) = 0

        ' instead of looping through all the command stored in the docksettings.ini file, we now store all the current commands in an array
        ' we loop through the array much quicker than looping through the temporary settings file and extracting the commands from each

        ' if the array location is empty then use goto to jump to the next iteration, ' sorry! VB6 has no continue.
        If sCommandArray(useloop) = vbNullString Then GoTo l_next_iteration
        thisCommand = sCommandArray(useloop)

        If FExists(thisCommand) Then
            GoTo l_next_iteration ' when we match a condition we loop over the subsequent conditions to iterate over the next item.
        End If
                    
        If DirExists(thisCommand) Then         ' is it a folder?
            GoTo l_next_iteration
        End If

        If InStr(thisCommand, "::{") Then
            GoTo l_next_iteration
        End If
                        
        If InStr(thisCommand, "%userprofile%") Then
            userprof = Environ$("USERPROFILE")
            thisCommand = Replace(thisCommand, "%userprofile%", userprof)
            If FExists(thisCommand) Then
                GoTo l_next_iteration
            End If
        End If
        
        If InStr(thisCommand, "%systemroot%") Then
            userprof = Environ$("SYSTEMROOT")
            thisCommand = Replace(thisCommand, "%systemroot%", userprof)
            If FExists(thisCommand) Then
                GoTo l_next_iteration
            End If
        End If
        
        ' Rocketdock commands compatibility
        If thisCommand = "[Quit]" Then
            GoTo l_next_iteration
        End If

        If thisCommand = "[Settings]" Then
            GoTo l_next_iteration
        End If

        If thisCommand = "[Icons]" Then
            GoTo l_next_iteration
        End If

        If thisCommand = "[RecycleBin]" Then
            GoTo l_next_iteration
        End If
        
        ' is the target a URL?
        testURL = Left(thisCommand, 3)
        If testURL = "htt" Or testURL = "www" Then
            GoTo l_next_iteration
        End If

        ' check in the windows folder, this is also done in the PATH check below but this one is quicker.
        If DirExists(Environ$("windir") & thisCommand) Then
            GoTo l_next_iteration
        End If
        
        ' these next two splits are meant to be at this location, to minimise them occurring
        
        'Use Split function to divide up the individual parts of the environment PATH string
        ' we do not want to do this every time, only when necessary and only once.
        If Not IsArrayInitialized(pathArray) Then pathArray = Split(pathString, ";")
        
        ' Use Split function to divide up the component parts of the suffix string
        ' we do not want to do this every time, only when necessary and only once.
        If Not IsArrayInitialized(fileSuffixArray) Then fileSuffixArray = Split(executableFileString)

        'iterate through the array created to work on each value, admin tools ends with .msc, .cpl, bat or exe
        For Each suffixElement In fileSuffixArray
            ' extract the suffix
            ' if the suffix is valid
            
            If InStr(thisCommand, "." & suffixElement) = 0 Then
                currentCommand = thisCommand & "." & suffixElement
            Else
                currentCommand = thisCommand
            End If
            
            If FExists(currentCommand) Then ' if the file exists and is valid - run it
                    GoTo l_next_iteration
            Else
                For Each pathElement In pathArray
                    If FExists(pathElement & "\" & currentCommand) Then
                        GoTo l_next_iteration
                    End If
                Next pathElement
            End If
        Next suffixElement

l_set_flag:
        ' set a flag to enable a small 'x' on this icon
        targetExistsArray(useloop) = 1

l_next_iteration:
    Next useloop

    On Error GoTo 0
    Exit Sub

checkTargetCommandValidity_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkTargetCommandValidity of Form dock"
            Resume Next
          End If
    End With

End Sub
' .90 DAEB 08/12/2022 frmMain.frm Added routine to check for an array that has already been initialised
'---------------------------------------------------------------------------------------
' Procedure : IsArrayInitialized
' Author    : beededea
' Date      : 09/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function IsArrayInitialized(arr As Variant) As Boolean
    On Error GoTo IsArrayInitialized_Error

    If Not IsArray(arr) Then Err.Raise 13
    On Error Resume Next
    IsArrayInitialized = (LBound(arr) <= UBound(arr))

    On Error GoTo 0
    Exit Function

IsArrayInitialized_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsArrayInitialized of Form dock"
            Resume Next
          End If
    End With
End Function



'---------------------------------------------------------------------------------------
' Procedure : transitTimer_Timer
' Author    : beededea
' Date      : 17/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub transitTimer_Timer()

    On Error GoTo transitTimer_Timer_Error


    On Error GoTo 0
    Exit Sub

transitTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure transitTimer_Timer of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : hourglassTimer_Timer
' Author    : beededea
' Date      : 30/04/2021
' Purpose   : ' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
'---------------------------------------------------------------------------------------
'
Private Sub hourglassTimer_Timer()
' load a small rotating hourglass image into the collection, used to signify running actions
    On Error GoTo hourglassTimer_Timer_Error

    hourglassTimerCount = hourglassTimerCount + 1
    If hourglassTimerCount > 5 Then hourglassTimerCount = 1
    
    hourglassimage = "hourglass" & hourglassTimerCount & "resizedImg128"

    On Error GoTo 0
    Exit Sub

hourglassTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure hourglassTimer_Timer of Form dock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fValidateComponents
' Author    : beededea
' Date      : 21/12/2022
' Purpose   : exits immediately if a component is missing
'---------------------------------------------------------------------------------------
'
Private Function fValidateComponents() As Boolean
    On Error GoTo fValidateComponents_Error

    ' folder checks
    fValidateComponents = reportMissingDir(sdAppPath & "\sounds")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingDir(sdAppPath & "\icons")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingDir(sdAppPath & "\skins")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingDir(sdAppPath & "\dockSettings")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingDir(sdAppPath & "\help")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingDir(sdAppPath & "\iconSettings")
    If fValidateComponents = False Then Exit Function
'    fValidateComponents = reportMissingDir(sdAppPath & "\arse")
'    If fValidateComponents = False Then Exit Function
'
    
    fValidateComponents = reportMissingFile(sdAppPath & "\appIdent.csv")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F1-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F2-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F3-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F4-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F5-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\busy-F6-32x32x24.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\red.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\SteamyDock-splash.jpg")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\red-X.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\blank.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\separator.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\tinyCircle.png")
    If fValidateComponents = False Then Exit Function
    fValidateComponents = reportMissingFile(sdAppPath & "\nixietubelargeQ.png")
    If fValidateComponents = False Then Exit Function
    
    On Error GoTo 0
    Exit Function

fValidateComponents_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fValidateComponents of Form dock"
            Resume Next
          End If
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : reportMissingDir
' Author    : beededea
' Date      : 09/01/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function reportMissingDir(ByVal folderToCheck As String) As Boolean
    On Error GoTo reportMissingDir_Error

    reportMissingDir = True
    If Not DirExists(folderToCheck) Then
        MsgBox "Essential component missing. Unable to find this folder: " & vbCr & vbCr & folderToCheck & vbCr & _
             vbCr & "Please re-install as some functions may not work properly." & _
             vbCr & "SteamyDock will now attempt to run but you may have to kill the steamydock process manually."
        reportMissingDir = False
    End If

    On Error GoTo 0
    Exit Function

reportMissingDir_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reportMissingDir of Form dock"
            Resume Next
          End If
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : reportMissingFile
' Author    : beededea
' Date      : 09/01/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function reportMissingFile(ByVal fileToCheck As String) As Boolean
    On Error GoTo reportMissingFile_Error

    reportMissingFile = True
    If Not FExists(fileToCheck) Then
        MsgBox "Essential component missing. Unable to find this folder: " & vbCr & vbCr & fileToCheck & vbCr & _
             vbCr & "Please re-install as some functions may not work properly." & _
             vbCr & "SteamyDock will now attempt to run but you may have to kill the steamydock process manually."
        reportMissingFile = False
    End If

    On Error GoTo 0
    Exit Function

reportMissingFile_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reportMissingFile of Form dock"
            Resume Next
          End If
    End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : selectBubbleAnimation
' Author    : beededea
' Date      : 09/01/2023
' Purpose   : there are three animation subroutines for the bubble animation
'---------------------------------------------------------------------------------------
'
Private Sub selectBubbleAnimation(ByVal animationType As Integer)

    On Error GoTo selectBubbleAnimation_Error

    Select Case animationType
        Case 1
            Call sequentialBubbleAnimation
        Case 2
            Call drawDockByCursorEntryPosition
        Case 3
            Call drawSmallStaticIcons
    End Select

    On Error GoTo 0
    Exit Sub

selectBubbleAnimation_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure selectBubbleAnimation of Form dock"
            Resume Next
          End If
    End With
    
End Sub






' change to shellExecuteWithDialog to allow apps to run unelevated: CREDIT - fafalone
' Requires oleexp.tlb (only for the IDE, compiled apps don't need it) and mIID.bas that's included with oleexp.

'Private Sub LaunchUnelevated(sPath As String)
'    Dim pShWin As ShellWindows
'    Set pShWin = New ShellWindows
'
'    Dim pDispView As oleexp.IDispatch 'Can't use the built-in VB6 version, need to specify our unrestricted implementation
'    Dim pServ As IServiceProvider
'    Dim pSB As IShellBrowser
'    Dim pDual As IShellFolderViewDual
'    Dim pView As IShellView
'
'    Dim vrEmpty As Variant
'    Dim hwnd As Long
'
'    Set pServ = pShWin.FindWindowSW(CVar(CSIDL_DESKTOP), vrEmpty, SWC_DESKTOP, hwnd, SWFO_NEEDDISPATCH)
'
'    pServ.QueryService SID_STopLevelBrowser, IID_IShellBrowser, pSB
'
'    pSB.QueryActiveShellView pView
'
'    pView.GetItemObject SVGIO_BACKGROUND, IID_IDispatch, pDispView
'    Set pDual = pDispView
'
'    Dim pDispShell As IShellDispatch2
'    Set pDispShell = pDual.Application
'
'    pDispShell.ShellExecute sPath
'End Sub




'        If useloop < leftmostResizedIcon Then  'small icons to the left shown in small mode
'            iconHeightPxls = iconSizeSmallPxls
'            iconWidthPxls = iconSizeSmallPxls
'
'            If dockPosition = vbbottom Then
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
'                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) + xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
'                    iconCurrentBottomPxls = ((dockUpperMostPxls + iconSizeLargePxls)) - xAxisModifier ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'                Else
'                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                    iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
'                    iconCurrentBottomPxls = dockUpperMostPxls + iconSizeLargePxls ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'                End If
'            End If
'
'            If dockPosition = vbtop Then
'
'                ' NOTE: everything is inverted...
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier '.nn added the slidein/out
'                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
'                    iconCurrentBottomPxls = ((dockUpperMostPxls)) + xAxisModifier  ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'                Else
'                    iconCurrentTopPxls = dockUpperMostPxls '.48 DAEB 01/04/2021 frmMain.frm  removed the vertical adjustment already applied to iconCurrentTopPxls
'                End If
'            End If
'
'            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeSmallPxls
'            showsmall = True
'            expandedDockWidth = expandedDockWidth + iconWidthPxls
'        End If
'
'
'
'
'
'    ' the group of icons to the left of the main icon, resized dynamically
'    If useloop < iconIndex And useloop >= leftmostResizedIcon Then
'        ' this is the area that we are currently changing
'        ' loop through all resized icons to the left
'
'       For useloop2 = leftmostResizedIcon To (iconIndex - 1)
'              ' if the icon number shown is 5
'             ' for the shrinking icon next to the main icon there will be a minimum size that it does not shrink below = 50% of the maximum it can grow, it grows to the max
'             ' if rDZoomWidth = 5 then
'             '
'             ' endif
'             ' for the next shrinking icon it will grow from minimum size to 50% of the maximum it can grow
'
'            resizeProportion = 1 / ((rDZoomWidth - 1) / 2) ' 33, .50 &c
'
'
'            ' for five icons that means two to the left
'
'
'            'leftmostResizedIcon.height and width = maximum iconsize *resizeProportion ie.50%
'            ' sizeModifierPxls = offsetFromLeftPxls
'            ' useloop * resizeProportion
'
''                iconHeightPxls = iconSizeLargePxls - (sizeModifierPxls * (useloop * resizeProportion)) 'sizeModifierPxls is the difference from the midpoint of the current icon in the x axis
''                iconWidthPxls = iconSizeLargePxls - (sizeModifierPxls * (useloop * resizeProportion))
'
'
'            'next icon height and width =up to maximum iconsize
'            'middle icon maximum icon size of course
'
'            resizeProportion = 1
'
'
'
'             iconHeightPxls = iconSizeLargePxls - (sizeModifierPxls * resizeProportion) 'sizeModifierPxls is the difference from the midpoint of the current icon in the x axis
'             iconWidthPxls = iconSizeLargePxls - (sizeModifierPxls * resizeProportion)
'
'
'             If dockPosition = vbbottom Then
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    'iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)) + xAxisModifier
'                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) + xAxisModifier '.nn
'                ElseIf autoSlideMode = "slidein" Then
'                    'iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)) - xAxisModifier
'                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) - xAxisModifier '.nn
'                Else
'                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                    'iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls))
'                    iconCurrentTopPxls = (dockUpperMostPxls + sizeModifierPxls) '.nn
'                End If
'
'                'If selectedIconIndex = iconIndex - 1 Then iconCurrentTopPxls = iconCurrentTopPxls - bounceCounter
'             End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'
'                '.nn added the slidein/out
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
'                Else
'                    iconCurrentTopPxls = dockUpperMostPxls
'                End If
'            End If
'
'             'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)
'            showsmall = False
'             expandedDockWidth = expandedDockWidth + iconWidthPxls
'
'        Next useloop2
'    End If
'
'
'        ' the main fullsize icon
'        If useloop = iconIndex Then
'            iconHeightPxls = iconSizeLargePxls
'            iconWidthPxls = iconSizeLargePxls
'
'            If dockPosition = vbbottom Then
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = dockUpperMostPxls + xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = dockUpperMostPxls - xAxisModifier
'                Else
'                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                    iconCurrentTopPxls = dockUpperMostPxls
'                End If
'
'                If selectedIconIndex = iconIndex Then iconCurrentTopPxls = iconCurrentTopPxls - bounceHeight
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'
'                '.nn added the slidein/out
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
'                Else
'                    iconCurrentTopPxls = dockUpperMostPxls
'                End If
'
'                If selectedIconIndex = iconIndex Then iconCurrentTopPxls = dockUpperMostPxls + bounceHeight
'            End If
'
'            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeLargePxls
'            showsmall = False
'            expandedDockWidth = expandedDockWidth + (iconWidthPxls)
'    End If
'
'
'        If useloop > iconIndex And useloop <= rightMostResizedIcon Then
'            iconHeightPxls = iconSizeSmallPxls + sizeModifierPxls
'            iconWidthPxls = iconSizeSmallPxls + sizeModifierPxls
'
'            If dockPosition = vbbottom Then
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)) + xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)) - xAxisModifier
'                Else
'                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                    iconCurrentTopPxls = (dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls))
'                End If
'                'If selectedIconIndex = iconIndex + 1 Then iconCurrentTopPxls = iconCurrentTopPxls - bounceHeight
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'
'                '.nn added the slidein/out
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
'                Else
'                    iconCurrentTopPxls = dockUpperMostPxls
'                End If
'            End If
'
'            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)
'            expandedDockWidth = expandedDockWidth + iconWidthPxls
'            showsmall = False
'        End If
'
'        If useloop > rightMostResizedIcon Then 'small icons to the right
'            iconHeightPxls = iconSizeSmallPxls
'            iconWidthPxls = iconSizeSmallPxls
'
'            If dockPosition = vbbottom Then
'
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) + xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls)) - xAxisModifier
'                Else
'                    ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls))
'                End If
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'
'                '.nn added the slidein/out
'                If autoSlideMode = "slideout" Then 'slideout is the default but if the slider timer is not running then xAxisModifier = 0
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) - xAxisModifier
'                ElseIf autoSlideMode = "slidein" Then
'                    iconCurrentTopPxls = ((dockUpperMostPxls + iconSizeSmallPxls)) + xAxisModifier
'                Else
'                    iconCurrentTopPxls = dockUpperMostPxls
'                End If
'            End If
'
'            'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - iconSizeSmallPxls
'            expandedDockWidth = expandedDockWidth + iconWidthPxls
'            showsmall = True
'        End If



























'
'    '===================
'    ' the main fullsize icon
'    '==================
'    iconHeightPxls = iconSizeLargePxls
'    iconWidthPxls = iconSizeLargePxls
'    mainIconWidthPxls = iconWidthPxls
'
'    If dockPosition = vbbottom Then
'        ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'        iconCurrentTopPxls = dockUpperMostPxls
'        ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'    End If
'
'    If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'        iconCurrentTopPxls = dockUpperMostPxls
'    End If
'
'    ' the following two lines  position the main icon initially to the main icon's leftmost start point when small
'    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    iconPosLeftPxls = (iconStoreLeftPixels(iconIndex)) '
'
'    ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    iconStoreRightPixels(iconIndex) = iconStoreLeftPixels(iconIndex) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'    iconStoreTopPixels(iconIndex) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'
'    'iconStoreBottomPixels(thisIconIndex) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'    ' any alteration to the above two lines to offset the icon start position causes a cascade in the subsequent animation routine moving it drastically to the left/right
'
'    Call loadTheImageIntoGDIPlus(iconIndex)
'    Call drawTheLabel(iconIndex)
'
'
''     all other icons are positioned in relation to the large main icon
''
''    ===================
'    ' one icon to the left, resized dynamically
'    '==================
'    If iconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
''        iconHeightPxls = iconSizeLargePxls - sizeModifierPxls 'sizeModifierPxls is the difference from the midpoint of the current icon in the x axis
''        iconWidthPxls = iconSizeLargePxls - sizeModifierPxls
'
'        iconHeightPxls = iconSizeLargePxls '.nn removal of sizeModifierPxls
'        iconWidthPxls = iconSizeLargePxls
'
'        If dockPosition = vbbottom Then
'            'iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)
'            iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeLargePxls) '.nn removal of sizeModifierPxls
'        ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'        End If
'
'        If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'           iconCurrentTopPxls = dockUpperMostPxls
'        End If
'
'        'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeLargePxls - sizeModifierPxls)
'
'        iconPosLeftPxls = iconPosLeftPxls - iconWidthPxls
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconStoreLeftPixels(iconIndex - 1) = iconPosLeftPxls
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconStoreRightPixels(iconIndex - 1) = iconStoreLeftPixels(iconIndex - 1) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'        iconStoreTopPixels(iconIndex - 1) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'
'        'iconStoreBottomPixels(iconIndex - 1) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'        Call loadTheImageIntoGDIPlus(iconIndex - 1)
'    End If
'
'
'
'    '===================
'    ' one icon to the right, resized dynamically
'    '==================
'    If iconIndex < rdIconMaximum Then  '    If iconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
'
''        iconHeightPxls = iconSizeSmallPxls + sizeModifierPxls '.nn removal of sizeModifierPxls
''        iconWidthPxls = iconSizeSmallPxls + sizeModifierPxls
'
'        iconHeightPxls = iconSizeSmallPxls  '.nn removal of sizeModifierPxls
'        iconWidthPxls = iconSizeSmallPxls
'
'        rightIconWidthPxls = iconWidthPxls
'
'        If dockPosition = vbbottom Then
''                ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                'iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls) '.nn removal of sizeModifierPxls
'
'                iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - (iconSizeSmallPxls) '.nn removal of sizeModifierPxls
'
'        ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'        End If
'
'        If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'            iconCurrentTopPxls = dockUpperMostPxls
'        End If
'
'        'If dockPosition = vbRight Then iconPosLeftPxls = iconLeftmostPointPxls + iconSizeLargePxls - (iconSizeSmallPxls + sizeModifierPxls)
'
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconPosLeftPxls = (iconStoreLeftPixels(iconIndex)) + mainIconWidthPxls
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconStoreLeftPixels(iconIndex + 1) = iconPosLeftPxls
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconStoreRightPixels(iconIndex + 1) = iconStoreLeftPixels(iconIndex + 1) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'        iconStoreTopPixels(iconIndex + 1) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'
'        'iconStoreBottomPixels(iconIndex + 1) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'        Call loadTheImageIntoGDIPlus(iconIndex + 1)
'    End If
'
'
''    ===================
'    ' all icons to the left
'    '==================
'    If iconIndex > 0 Then 'check it isn't trying to animate a non-existent icon before the first icon
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'        iconPosLeftPxls = iconStoreLeftPixels(iconIndex - 1)
'
'        For useloop = iconIndex - 2 To 0 Step -1
'            iconHeightPxls = iconSizeSmallPxls
'            iconWidthPxls = iconSizeSmallPxls
'
'            If dockPosition = vbbottom Then
'                ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
'        ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'                iconCurrentTopPxls = dockUpperMostPxls
'            End If
'
'            iconPosLeftPxls = iconPosLeftPxls - iconWidthPxls
'            iconStoreLeftPixels(useloop) = iconPosLeftPxls ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'            iconStoreRightPixels(useloop) = iconStoreLeftPixels(useloop) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon             ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'            iconStoreTopPixels(useloop) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'
'            'iconStoreBottomPixels(useloop) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'            thiskey = useloop & "ResizedImg" & LTrim$(Str$(iconSizeSmallPxls))
'            updateDisplayFromDictionary collSmallIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
'            If rDShowRunning = "1" Then
'                If processCheckArray(useloop) = True Then
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
'                End If
'                ' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
'                If targetExistsArray(useloop) = 1 Then
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
'                End If
'
'            End If
'        Next useloop
'    End If
'
'    '====================
'    ' icons to the right
'    '====================
'    If iconIndex < rdIconMaximum Then   'check it isn't trying to animate a non-existent icon after the last icon
'
'        ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'       iconPosLeftPxls = (iconStoreLeftPixels(iconIndex + 1)) + rightIconWidthPxls
'       For useloop = iconIndex + 2 To iconArrayUpperBound
'
'            iconHeightPxls = iconSizeSmallPxls
'            iconWidthPxls = iconSizeSmallPxls
'
'            If dockPosition = vbbottom Then
'                ' .46 DAEB 01/04/2021 frmMain.frm Ensured that there is a line to calculate iconCurrentTopPxls now that autoSlideMode is now undefined at startup
'                iconCurrentTopPxls = dockUpperMostPxls + iconSizeLargePxls - iconSizeSmallPxls
'                ' .50 DAEB 01/04/2021 frmMain.frm Pruned all the redundant code for positioniong according to the slideIn/Out state, not done here
'            End If
'
'            If dockPosition = vbtop Then ' .48 DAEB 01/04/2021 frmMain.frm removed the vertical adjustment already applied to iconCurrentTopPxls
'                iconCurrentTopPxls = dockUpperMostPxls
'            End If
'
'
'            iconPosLeftPxls = iconPosLeftPxls + iconWidthPxls
'            ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'            iconStoreLeftPixels(useloop) = iconPosLeftPxls
'            iconStoreRightPixels(useloop) = iconStoreLeftPixels(useloop) + iconWidthPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'            iconStoreTopPixels(useloop) = iconCurrentTopPxls ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'            'iconStoreBottomPixels(useloop) =' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'            thiskey = useloop & "ResizedImg" & LTrim$(Str$(iconSizeSmallPxls))
'            updateDisplayFromDictionary collSmallIcons, vbNullString, thiskey, (iconPosLeftPxls), (iconCurrentTopPxls), (iconWidthPxls), (iconHeightPxls)
'            If rDShowRunning = "1" Then
'                If processCheckArray(useloop) = True Then
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "tinycircleResizedImg128", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls), (iconSizeSmallPxls)
'                End If
'            End If
'            ' .87 DAEB 08/12/2022 frmMain.frm Target command validity flag places a red X on the icon
'            If targetExistsArray(useloop) = 1 Then  ' redxResizedImg64
'                    If dockPosition = vbbottom Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconCurrentTopPxls - (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2) '.69 DAEB 06/05/2021 frmMain.frm Draw the small cog in the right place for the vbtop position
'                    If dockPosition = vbtop Then updateDisplayFromDictionary collLargeIcons, vbNullString, "redxResizedImg64", (iconPosLeftPxls + (iconSizeSmallPxls / 2) - 3), (iconSizeSmallPxls + (iconSizeSmallPxls / 5)), (iconSizeSmallPxls / 2), (iconSizeSmallPxls / 2)
'            End If
'        Next useloop
'    End If
'
