Attribute VB_Name = "Module1"
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context

Option Explicit
'
Dim specifiedMonitor As Long
Dim glWindowToCheckOnMonitor As Long
Dim glMonitorWidth As Long
Dim glMonitorHeight As Long
Dim gbMonitorIsPrimary As Boolean
Dim gbWindowIsOnMonitor As Boolean
Dim glSecondaryMonitorNumber As Long
Dim thisHdc As Long
'
Public Type RECT
  Left As Long
  Top As Long
  Right As Long ' This is +1 (right - left = width)
  Bottom As Long ' This is +1 (bottom - top = height)
End Type
Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Private Const MONITORINFOF_PRIMARY = &H1
Private Const MONITOR_DEFAULTTONEAREST = &H2
Private Const MONITOR_DEFAULTTONULL = &H0
Private Const MONITOR_DEFAULTTOPRIMARY = &H1
'
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long

Public hdcScreen As Long


Public Sub CenterFormOnMonitorTwo(ByRef frm As Form) ' This is actually just the non-primary monitor.
    ' This assumes that the second monitor is to the right of the first.
    ' Other work will be needed if this isn't the case.
    Dim MonitorTwoHeight As Long
    Dim MonitorTwoWidth As Long
    Dim iSecondaryMonitorNumber As Long
    Dim MonitorCount As Long
    '
    MonitorCount = countMonitors() ' function to enumerate all screens
    If MonitorCount = 1 Then Exit Sub
    iSecondaryMonitorNumber = SecondaryMonitorNumber()
    '
    MonitorTwoWidth = MonitorPixelWidth(iSecondaryMonitorNumber) * Screen.twipsPerPixelX
    MonitorTwoHeight = MonitorPixelHeight(iSecondaryMonitorNumber) * Screen.twipsPerPixelY
    '
    frm.Left = ((MonitorTwoWidth - frm.Width) \ 2) + Screen.Width
    frm.Top = (MonitorTwoHeight - frm.Height) \ 3
    If Not WindowIsOnMonitor(iSecondaryMonitorNumber, frm.hWnd) Then
        ' Primary must be on right if not visible.
        frm.Left = ((MonitorTwoWidth - frm.Width) \ 2) - Screen.Width
        frm.Top = (MonitorTwoHeight - frm.Height) \ 3
    End If
    If Not WindowIsOnMonitor(iSecondaryMonitorNumber, frm.hWnd) Then
        ' Couldn't do it for some reason, so forget it.
        ' Maybe monitors are on top of one another.
        frm.Left = (ScreenWidth - frm.Width) \ 2
        frm.Top = (ScreenHeight - frm.Height) \ 3
    End If
End Sub

Public Function SecondaryMonitorNumber() As Long
    Dim iCount As Long
    '
    glSecondaryMonitorNumber = 0
    specifiedMonitor = -1 ' Don't specify one to just execute.
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, iCount 'callback that tests all monitors attached
    SecondaryMonitorNumber = glSecondaryMonitorNumber
End Function

Public Function WindowIsOnMonitor(iMonitor As Long, hWnd As Long) As Boolean
    Dim iCount As Long
    '
    gbWindowIsOnMonitor = False
    specifiedMonitor = iMonitor
    glWindowToCheckOnMonitor = hWnd
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, iCount 'callback that tests all monitors attached
    WindowIsOnMonitor = gbWindowIsOnMonitor
End Function

Public Function countMonitors() As Long
    Dim iCount As Long
    '
    specifiedMonitor = -1 ' This time do not specify a monitor number in order to just count them, -1 counts them all
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, iCount 'callback that tests all monitors attached
    countMonitors = iCount
End Function

Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hDCMonitor As Long, uRect As RECT, dwData As Long) As Long
    Dim MI As MONITORINFO
    Dim hdc As Long
    
    ' This is just used for the above functions.
    dwData = dwData + 1 ' Increase monitor count.
    If specifiedMonitor = dwData Then
        Call SetMonitorGlobals(hMonitor, uRect)
    
'                hdc = GetDC(0) 'get an hDC directly to the screen
        'hdcScreen = hDCMonitor
'                If Not specifiedMonitor = dwData Then
'                    ReleaseDC 0, hdc '1795237634 ' 939599720
'                End If
    End If
    
    'hdc = CreateDC(("DISPLAY"), vbNullString, vbNullString, ByVal 0&) ' GetDC(0)
    'thisHdc = hdc
    'ReleaseDC 0, hdc

    
    ' We must still set the glSecondaryMonitorNumber.
    MI.cbSize = Len(MI)
    Call GetMonitorInfo(hMonitor, MI)
    If Not CBool(MI.dwFlags = MONITORINFOF_PRIMARY) Then ' This is the "primary monitor test".
        glSecondaryMonitorNumber = dwData
    End If
    

    
    MonitorEnumProc = 1
End Function

Public Function MonitorPixelWidth(iMonitor As Long) As Long
    ' Returns the width of the specified monitor.
    ' Handles systems with multiple monitors.
    Dim iCount As Long
    '
    specifiedMonitor = iMonitor
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, iCount 'callback that tests all monitors attached
    MonitorPixelWidth = glMonitorWidth
End Function

Public Function MonitorPixelHeight(iMonitor As Long) As Long
    ' Returns the height of the specified monitor.
    ' Handles systems with multiple monitors.
    Dim iCount As Long
    '
    specifiedMonitor = iMonitor
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, iCount 'callback that tests a specifc monitors
    MonitorPixelHeight = glMonitorHeight
End Function

Public Function ScreenWidth()
    ' This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
    Dim Pixels As Long
    Const SM_CXSCREEN = 0
    '
    Pixels = GetSystemMetrics(SM_CXSCREEN)
    ScreenWidth = Pixels * twipsPerPixelX
End Function

Public Function ScreenHeight(Optional bSubtractTaskbar As Boolean = False)
    ' This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
    Dim Pixels As Long
    Const SM_CYSCREEN = 1
    '
    Pixels = GetSystemMetrics(SM_CYSCREEN)
    If bSubtractTaskbar Then
        ' The taskbar is typically 30 pixels or 450 twips, or, at least, this is the assumption made here.
        ' It can actually be multiples of this, or possibly moved to the side or top.
        ' This procedure does not account for these possibilities.
        ScreenHeight = (Pixels - 30) * twipsPerPixelY
    Else
        ScreenHeight = Pixels * twipsPerPixelY
    End If
End Function

Private Sub SetMonitorGlobals(hMonitor As Long, uRect As RECT)
    Dim MI As MONITORINFO
    Dim hdc As Long

    '
    glMonitorWidth = uRect.Right - uRect.Left
    glMonitorHeight = uRect.Bottom - uRect.Top
    MI.cbSize = Len(MI)
    GetMonitorInfo hMonitor, MI
    gbMonitorIsPrimary = CBool(MI.dwFlags = MONITORINFOF_PRIMARY)
    gbWindowIsOnMonitor = (MonitorFromWindow(glWindowToCheckOnMonitor, MONITOR_DEFAULTTONEAREST) = hMonitor)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : twipsPerPixelX
' Author    : beededea
' Date      : 23/01/2021
' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
'---------------------------------------------------------------------------------------
'
Public Function twipsPerPixelX() As Single
    Dim hdc As Long
    Dim lPixelsPerInch As Long
    Const LOGPIXELSX = 88        '  Logical pixels/inch in X
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    '
    On Error GoTo twipsPerPixelX_Error
    
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hdc = GetDC(0)
    If hdc <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
        ReleaseDC 0, hdc
        twipsPerPixelX = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        twipsPerPixelX = Screen.twipsPerPixelX
    End If

   On Error GoTo 0
   Exit Function

twipsPerPixelX_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure twipsPerPixelX of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : twipsPerPixelY
' Author    : beededea
' Date      : 23/01/2021
' Purpose   : This works even on Tablet PC.  The problem is: when the tablet screen is rotated, the "Screen" object of VB doesn't pick it up.
'---------------------------------------------------------------------------------------
'
Public Function twipsPerPixelY() As Single
    Dim hdc As Long
    Dim lPixelsPerInch As Long
    Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
    Const POINTS_PER_INCH As Long = 72 ' A point is defined as 1/72 inches.
    Const TWIPS_PER_POINT As Long = 20 ' Also, by definition.
    
   On Error GoTo twipsPerPixelY_Error
   
    ' 23/01/2021 .01 monitorModule.bas DAEB added if then else if you can't get device context
    hdc = GetDC(0)
    If hdc <> 0 Then
        lPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
        ReleaseDC 0, hdc
        twipsPerPixelY = TWIPS_PER_POINT * (POINTS_PER_INCH / lPixelsPerInch) ' Cancel units to see it.
    Else
        twipsPerPixelY = Screen.twipsPerPixelY
    End If

   On Error GoTo 0
   Exit Function

twipsPerPixelY_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure twipsPerPixelY of Module Module1"

End Function

'---------------------------------------------------------------------------------------
' Procedure : getDeviceHdc
' Author    : beededea
' Date      : 23/09/2020
' Purpose   : Handles systems with multiple monitors.
'---------------------------------------------------------------------------------------
'
Public Sub getDeviceHdc()
    Dim iCount As Long

    On Error GoTo getDeviceHdc_Error

        specifiedMonitor = Val(rDMonitor) + 1
        EnumDisplayMonitors 0, ByVal 0&, AddressOf hdcEnumProc, iCount 'callback that calls all monitors

    On Error GoTo 0
    Exit Sub

getDeviceHdc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDeviceHdc of Module Module1"
End Sub

Private Function hdcEnumProc(ByVal hMonitor As Long, ByVal hDCMonitor As Long, uRect As RECT, dwData As Long) As Long
    'Dim hdc As Long
    Dim nDC As Long
   
    dwData = dwData + 1 ' Increment monitor number as we loop through any attached monitors
    
    'get the data from the specified monitor only
    If specifiedMonitor = dwData Then
'            hdc = GetDC(0) 'get an hDC directly to the screen
'            hdcScreen = hdc
        
        'Create Device Context Compatible With Screen
        nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
        hdcScreen = nDC
            
    End If
    
    hdcEnumProc = 1
End Function

