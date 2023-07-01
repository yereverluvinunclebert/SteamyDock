Attribute VB_Name = "modTooltips"

Option Explicit
'
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
Private Type TOOLINFO
    lSize       As Long
    lFlags      As Long
    hWnd        As Long
    lId         As Long
    '
    'lpRect      As RECT
    Left        As Long
    Top         As Long
    Right       As Long ' This is +1 (right - left = width)
    Bottom      As Long ' This is +1 (bottom - top = height)
    '
    hInstance   As Long
    lpStr       As String
    lParam      As Long
End Type
'
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'
Private Const WM_USER               As Long = &H400&
Private Const CW_USEDEFAULT         As Long = &H80000000
'
Private Const TTM_ACTIVATE          As Long = WM_USER + 1&
'Private Const TTM_ADDTOOLA          As Long = WM_USER + 4&
Private Const TTM_ADDTOOLW          As Long = WM_USER + 50&
Private Const TTM_SETDELAYTIME      As Long = WM_USER + 3&
'Private Const TTM_UPDATETIPTEXTA    As Long = WM_USER + 12&
Private Const TTM_UPDATETIPTEXTW    As Long = WM_USER + 57&
Private Const TTM_SETTIPBKCOLOR     As Long = WM_USER + 19&
Private Const TTM_SETTIPTEXTCOLOR   As Long = WM_USER + 20&
Private Const TTM_SETMAXTIPWIDTH    As Long = WM_USER + 24&
'Private Const TTM_SETTITLEA         As Long = WM_USER + 32&
Private Const TTM_SETTITLEW         As Long = WM_USER + 33&
'
Private Const TTS_NOPREFIX          As Long = &H2&
Private Const TTS_BALLOON           As Long = &H40&
Private Const TTS_ALWAYSTIP         As Long = &H1&
'
Private Const TTF_CENTERTIP         As Long = &H2&
Private Const TTF_IDISHWND          As Long = &H1&
Private Const TTF_SUBCLASS          As Long = &H10&
Private Const TTF_TRANSPARENT       As Long = &H100&
'
Private Const TTDT_AUTOPOP          As Long = 2&
Private Const TTDT_INITIAL          As Long = 3&
'
Private Const TOOLTIPS_CLASS        As String = "tooltips_class32"
'
Private Const GWL_EXSTYLE           As Long = &HFFFFFFEC
Private Const WS_EX_TOOLWINDOW      As Long = &H80&
Private Const WS_EX_TOPMOST         As Long = &H8&
'
Public Enum ttIconType
    TTNoIcon
    TTIconInfo
    TTIconWarning
    TTIconError
End Enum
#If False Then ' Intellisense fix.
    Public TTNoIcon, TTIconInfo, TTIconWarning, TTIconError
#End If
'
Private hwndTT As Long ' hwnd of the tooltip
'

Public Sub CreateToolTip(ByVal ParentHwnd As Long, _
                         ByVal TipText As String, _
                         Optional ByVal uIcon As ttIconType = TTNoIcon, _
                         Optional ByVal sTitle As String, _
                         Optional ByVal lForeColor As Long = -1&, _
                         Optional ByVal lBackColor As Long = -1&, _
                         Optional ByVal bCentered As Boolean, _
                         Optional ByVal bBalloon As Boolean, _
                         Optional ByVal lWrapTextLength As Long = 50&, _
                         Optional ByVal lDelayTime As Long = 600&, _
                         Optional ByVal lVisibleTime As Long = 7500&)
    '
    ' If lWrapTextLength = 0 then there will be no wrap.
    ' Also, lWrapTextLength = 40 is a minimum value.
    ' The max for lVisibleTime is 32767.
    '
    Static bCommonControlsInitialized   As Boolean
    Dim lWinStyle                       As Long
    Dim ti                              As TOOLINFO
    Static PrevParentHwnd               As Long
    Static PrevTipText                  As String
    Static PrevTitle                    As String
    '
    ' Don't do anything unless we need to.
    If hwndTT <> 0& And ParentHwnd = PrevParentHwnd And TipText = PrevTipText And sTitle = PrevTitle Then Exit Sub
    PrevParentHwnd = ParentHwnd
    PrevTipText = TipText
    PrevTitle = sTitle
    '
    If Not bCommonControlsInitialized Then
        InitCommonControls
        bCommonControlsInitialized = True
    End If
    '
    ' Destroy any previous tooltip.
    If hwndTT <> 0& Then DestroyWindow hwndTT
    '
    ' Format the text.
    FormatTooltipText TipText, lWrapTextLength
    '
    ' Initial style settings.
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    If bBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON ' Create baloon style if desired.
    ' Set the style.
    hwndTT = CreateWindowExW(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, StrPtr(TOOLTIPS_CLASS), 0&, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
    '
    ' Setup our tooltip info structure.
    ti.lFlags = TTF_SUBCLASS Or TTF_IDISHWND
    If bCentered Then ti.lFlags = ti.lFlags Or TTF_CENTERTIP
    ' Set the hwnd prop to our parent control's hwnd.
    ti.hWnd = ParentHwnd
    ti.lId = ParentHwnd
    ti.hInstance = App.hInstance
    ti.lpStr = TipText
    ti.lSize = LenB(ti)
    ' Set the tooltip structure
    SendMessageLongA hwndTT, TTM_ADDTOOLW, 0&, VarPtr(ti)
    SendMessageLongA hwndTT, TTM_UPDATETIPTEXTW, 0&, VarPtr(ti)
    '
    ' Colors.
    If lForeColor <> -1& Then SendMessageA hwndTT, TTM_SETTIPTEXTCOLOR, lForeColor, 0&
    If lBackColor <> -1& Then SendMessageA hwndTT, TTM_SETTIPBKCOLOR, lBackColor, 0&
    '
    ' Title or icon.
    If uIcon <> TTNoIcon Or sTitle <> vbNullString Then SendMessageLongA hwndTT, TTM_SETTITLEW, CLng(uIcon), StrPtr(sTitle)
    '
    SendMessageLongA hwndTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, lVisibleTime
    SendMessageLongA hwndTT, TTM_SETDELAYTIME, TTDT_INITIAL, lDelayTime
End Sub

Public Sub DestroyToolTip()
    ' It's not a bad idea to put this in the Form_Unload event just to make sure.
    If hwndTT <> 0& Then DestroyWindow hwndTT
    hwndTT = 0&
End Sub

Private Sub FormatTooltipText(TipText As String, lLen As Long)
    Dim s       As String
    Dim ss()    As String
    Dim i       As Long
    Dim j       As Long
    '
    ' Make sure we need to do anything.
    If lLen = 0& Then Exit Sub
    If lLen < 40& Then lLen = 40&
    If Len(TipText) <= lLen Then Exit Sub
    '
    ss = Split(TipText, vbCrLf)                     ' We split each line separately.
    For j = LBound(ss) To UBound(ss)
        If Len(ss(j)) > lLen Then
            s = vbNullString
            Do
                i = InStrRev(ss(j), " ", lLen + 1&)
                If i = 0& Then
                    s = s & Left$(ss(j), lLen) & vbCrLf ' Build "s" and trim from TipText.
                    ss(j) = Mid$(ss(j), lLen + 1&)
                Else
                    s = s & Left$(ss(j), i - 1&) & vbCrLf ' Build "s" and trim from TipText.
                    ss(j) = Mid$(ss(j), i + 1&)
                End If
                If Len(ss(j)) <= lLen Then
                    ss(j) = s & ss(j) ' Place "s" back into ss(j) and get out.
                    Exit Do
                End If
            Loop
        End If
    Next
    TipText = Join(ss, vbCrLf)
End Sub





