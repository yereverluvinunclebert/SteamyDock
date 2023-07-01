Attribute VB_Name = "modPersistentDebugPrint"
Option Explicit
'
Private Const WM_DESTROY                As Long = &H2&  ' All other needed constants are declared within the procedures.
Dim bSetWhenSubclassing_UsedByIdeStop   As Boolean      ' Never goes false once set by first subclassing, unless IDE Stop button is clicked.
'
Private Declare Function SetWindowSubclass Lib "comctl32.dll" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, Optional ByVal dwRefData As Long) As Long
Private Declare Function GetWindowSubclass Lib "comctl32.dll" Alias "#411" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, pdwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function NextSubclassProcOnChain Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function DefSubclassProc Lib "comctl32.dll" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
'
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
'
Private Type POINTAPI
    x As Long
    Y As Long
End Type
'
Private Type ChooseColorType
    lStructSize         As Long
    hWndOwner           As Long
    hInstance           As Long
    rgbResult           As Long
    lpCustColors        As Long
    Flags               As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type
Private Enum ChooseColorFlagsEnum
    CC_RGBINIT = &H1                  ' Make the color specified by rgbResult be the initially selected color.
    CC_FULLOPEN = &H2                 ' Automatically display the Define Custom Colors half of the dialog box.
    CC_PREVENTFULLOPEN = &H4          ' Disable the button that displays the Define Custom Colors half of the dialog box.
    CC_SHOWHELP = &H8                 ' Display the Help button.
    CC_ENABLEHOOK = &H10              ' Use the hook function specified by lpfnHook to process the Choose Color box's messages.
    CC_ENABLETEMPLATE = &H20          ' Use the dialog box template identified by hInstance and lpTemplateName.
    CC_ENABLETEMPLATEHANDLE = &H40    ' Use the preloaded dialog box template identified by hInstance, ignoring lpTemplateName.
    CC_SOLIDCOLOR = &H80              ' Only allow the user to select solid colors. If the user attempts to select a non-solid color, convert it to the closest solid color.
    CC_ANYCOLOR = &H100               ' Allow the user to select any color.
End Enum
#If False Then ' Intellisense fix.
    Public CC_RGBINIT, CC_FULLOPEN, CC_PREVENTFULLOPEN, CC_SHOWHELP, CC_ENABLEHOOK, CC_ENABLETEMPLATE, CC_ENABLETEMPLATEHANDLE, CC_SOLIDCOLOR, CC_ANYCOLOR
#End If
Private Type KeyboardInput        '
    dwType As Long                ' Set to INPUT_KEYBOARD.
    wVK As Integer                ' shift, ctrl, menukey, or the key itself.
    wScan As Integer              ' Not being used.
    dwFlags As Long               '            HARDWAREINPUT hi;
    dwTime As Long                ' Not being used.
    dwExtraInfo As Long           ' Not being used.
    dwPadding As Currency         ' Not being used.
End Type
'
Dim msColorTitle                As String
Dim pt32                        As POINTAPI
Dim muEvents(1)                 As KeyboardInput    ' Just used to emulate "Enter" key.
'
Private Declare Function ChooseColorAPI Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColorType) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal uFlags As Long) As Long
Private Declare Function SetFocusTo Lib "user32" Alias "SetFocus" (Optional ByVal hWnd As Long) As Long
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'
Public ColorDialogSuccessful    As Boolean
Public ColorDialogColor         As Long
'

Public Function ShowColorDialog(hWndOwner As Long, lDefaultColor As Long, Optional NewColor As Long, Optional Title As String = "Select Color", Optional CustomColorsHex As String) As Boolean
    ' You can optionally use ColorDialogSuccessful & ColorDialogColor or the return of ShowColorDialog and NewColor.  They will be the same.
    '
    ' CustomColorHex is a comma separated hex string of 16 custom colors.  It's best to just let the user specify these, starting out with all black.
    ' If this CustomColorHex string doesn't separate into precisely 16 values, it's ignored, resulting with all black custom colors.
    ' The string is returned, and it's up to you to save it if you wish to save your user-specified custom colors.
    ' These will be specific to this program, because this is your CustomColorsHex string.
    '
    Dim uChooseColor As ChooseColorType
    Dim CustomColors(15) As Long
    Dim sArray() As String
    Dim i As Long
    '
    msColorTitle = Title
    '
    ' Setup custom colors.
    sArray = Split(CustomColorsHex, ",")
    If UBound(sArray) = 15 Then
        For i = 0 To 15
            CustomColors(i) = Val("&h" & sArray(i))
        Next
    End If
    '
    uChooseColor.hWndOwner = hWndOwner
    uChooseColor.lpCustColors = VarPtr(CustomColors(0))
    uChooseColor.Flags = CC_ENABLEHOOK Or CC_FULLOPEN Or CC_RGBINIT
    uChooseColor.hInstance = App.hInstance
    uChooseColor.lStructSize = LenB(uChooseColor)
    uChooseColor.lpfnHook = ProcedureAddress(AddressOf ColorHookProc)
    uChooseColor.rgbResult = lDefaultColor
    '
    ColorDialogSuccessful = False
    If ChooseColorAPI(uChooseColor) = 0 Then
        Exit Function
    End If
    If uChooseColor.rgbResult > vbWhite Then Exit Function
    '
    ColorDialogColor = uChooseColor.rgbResult
    NewColor = uChooseColor.rgbResult
    ColorDialogSuccessful = True
    ShowColorDialog = True
    '
    ' Return custom colors.
    ReDim sArray(15)
    For i = 0 To 15
        sArray(i) = Hex$(CustomColors(i))
    Next
    CustomColorsHex = Join(sArray, ",")
End Function

Private Function ColorHookProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_SHOWWINDOW     As Long = 24&
    Const WM_LBUTTONDBLCLK  As Long = 515&
    Const KEYEVENTF_KEYUP   As Long = 2&
    Const KEYEVENTF_KEYDOWN As Long = 0&
    Const INPUT_KEYBOARD    As Long = 1&
    '
    If uMsg = WM_SHOWWINDOW Then
        SetWindowText hWnd, msColorTitle
        ColorHookProc = 1&
    End If
    '
    If uMsg = WM_LBUTTONDBLCLK Then
        '
        ' If we're on a hWnd with text, we probably should ignore the double-click.
        GetCursorPos pt32
        ScreenToClient hWnd, pt32
        '
        If WindowText(ChildWindowFromPointEx(hWnd, pt32.x, pt32.Y, 0&)) = vbNullString Then
            ' For some reason, this SetFocus is necessary for the dialog to receive keyboard input under certain circumstances.
            SetFocusTo hWnd
            ' Build EnterKeyDown & EnterKeyDown events.
            muEvents(0).wVK = vbKeyReturn: muEvents(0).dwFlags = KEYEVENTF_KEYDOWN: muEvents(0).dwType = INPUT_KEYBOARD
            muEvents(1).wVK = vbKeyReturn: muEvents(1).dwFlags = KEYEVENTF_KEYUP:   muEvents(1).dwType = INPUT_KEYBOARD
            ' Put it on buffer.
            SendInput 2&, muEvents(0), Len(muEvents(0))
            ColorHookProc = 1&
        End If
    End If
End Function

Private Function ProcedureAddress(AddressOf_TheProc As Long) As Long
    ProcedureAddress = AddressOf_TheProc
End Function

Private Sub SetWindowText(hWnd As Long, sText As String)
    Const WM_SETTEXT        As Long = &HC&
    SendMessageLongW hWnd, WM_SETTEXT, 0&, StrPtr(sText)
End Sub

Public Function WindowText(hWnd As Long) As String
    ' Form or control.
    WindowText = String$(GetWindowTextLength(hWnd) + 1&, vbNullChar)
    Call GetWindowText(hWnd, WindowText, Len(WindowText))
    WindowText = Left$(WindowText, InStr(WindowText, vbNullChar) - 1&)
End Function





Private Sub SubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, Optional dwRefData As Long, Optional uIdSubclass As Long)
    If uIdSubclass = 0& Then uIdSubclass = hWnd
    bSetWhenSubclassing_UsedByIdeStop = True
    Call SetWindowSubclass(hWnd, AddressOf_ProcToSubclass, hWnd, dwRefData)
End Sub

Private Sub UnSubclassSomeWindow(hWnd As Long, AddressOf_ProcToSubclass As Long, Optional uIdSubclass As Long)
    If uIdSubclass = 0& Then uIdSubclass = hWnd
    Call RemoveWindowSubclass(hWnd, AddressOf_ProcToSubclass, uIdSubclass)
End Sub

Private Function IdeStopButtonClicked() As Boolean
    IdeStopButtonClicked = Not bSetWhenSubclassing_UsedByIdeStop
End Function

Public Sub SubclassFormToReceiveStringMsg(frm As VB.Form)
    SubclassSomeWindow frm.hWnd, AddressOf StringMessage_Proc, ObjPtr(frm)
End Sub

Private Function StringMessage_Proc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    If uMsg = WM_DESTROY Then
        UnSubclassSomeWindow hWnd, AddressOf_StringMessage_Proc, uIdSubclass
        StringMessage_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    If IdeStopButtonClicked Then ' Protect the IDE.  Don't execute any specific stuff if we're stopping.  We may run into COM objects or other variables that no longer exist.
        StringMessage_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
        Exit Function
    End If
    '
    Dim cds As COPYDATASTRUCT
    Dim Buf() As Byte
    Dim sMsg As String
    Const WM_COPYDATA As Long = &H4A&
    '
    If uMsg = WM_COPYDATA Then
        Call CopyMemory(cds, ByVal lParam, Len(cds))
        ReDim Buf(1& To cds.cbData)
        Call CopyMemory(Buf(1&), ByVal cds.lpData, cds.cbData)
        sMsg = StrConv(Buf, vbUnicode)
        sMsg = Left$(sMsg, InStr(sMsg, vbNullChar) - 1&)
        '
        frmDebugPrint.Out sMsg  ' Out MUST be public, or we can't find it this way.
    End If
    '
    ' Give control to other procs, if they exist.
    StringMessage_Proc = NextSubclassProcOnChain(hWnd, uMsg, wParam, lParam)
End Function

Private Function AddressOf_StringMessage_Proc() As Long
    AddressOf_StringMessage_Proc = ProcedureAddress(AddressOf StringMessage_Proc)
End Function

