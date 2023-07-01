Attribute VB_Name = "modMsgBoxX"
    '*************************************************************
    '* MsgBoxEx() - Written by Aaron Young, February 7th 2000
    '*            - Edited by Philip Manavopoulos, May 19th 2005
    '*************************************************************
     
    Option Explicit
     
    Private Type CWPSTRUCT
        lParam As Long
        wParam As Long
        message As Long
        hwnd As Long
    End Type
     
    'Added by manavo11
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
     
    Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
    End Type
    'Added by manavo11
     
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
     
    'Added by manavo11
    Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
     
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
     
    Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
     
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    'Added by manavo11
     
    Private Const WH_CALLWNDPROC = 4
    Private Const GWL_WNDPROC = (-4)
    Private Const WM_CTLCOLORBTN = &H135
    Private Const WM_DESTROY = &H2
    Private Const WM_SETTEXT = &HC
    Private Const WM_CREATE = &H1
     
    'Added by manavo11
    ' System Color Constants
    Private Const COLOR_BTNFACE = 15
    Private Const COLOR_BTNTEXT = 18
     
    ' Windows Messages
    Private Const WM_CTLCOLORSTATIC = &H138
    Private Const WM_CTLCOLORDLG = &H136
     
    Private Const WM_SHOWWINDOW As Long = &H18
    'Added by manavo11
     
    Private lHook As Long
    Private lPrevWnd As Long
     
    Private bCustom As Boolean
    Private sButtons() As String
    Private lButton As Long
    Private sHwnd As String
     
    'Added by manavo11
    Private lForecolor As Long
    Private lBackcolor As Long
     
    Private sDefaultButton As String
     
    Private iX As String
    Private iY As String
    Private iWidth As String
    Private iHeight As String
     
    Private iButtonCount As Integer
    Private iButtonWidth As Integer
    'Added by manavo11
     
    Public Function SubMsgBox(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim sText As String
        
        Select Case Msg
        
        'Added by manavo11
        Case WM_SHOWWINDOW
            Dim MsgBoxRect As RECT
            
            GetWindowRect hwnd, MsgBoxRect
            
            If StrPtr(iX) = 0 Then
                iX = MsgBoxRect.Left
            End If
            
            If StrPtr(iY) = 0 Then
                iY = MsgBoxRect.Top
            End If
            
            If StrPtr(iWidth) = 0 Then
                iWidth = MsgBoxRect.Right - MsgBoxRect.Left
            Else
                Dim i As Integer
                Dim h As Long
                
                Dim ButtonRECT As RECT
                
                For i = 0 To iButtonCount
                    h = FindWindowEx(hwnd, h, "Button", vbNullString)
                    
                    GetWindowRect h, ButtonRECT
                    
                    MoveWindow h, 14 + (iButtonWidth * i) + (6 * i), iHeight - (ButtonRECT.Bottom - ButtonRECT.Top) - 40, iButtonWidth, ButtonRECT.Bottom - ButtonRECT.Top, 1
                Next
            End If
            
            If StrPtr(iHeight) = 0 Then
                iHeight = MsgBoxRect.Bottom - MsgBoxRect.Top
            End If
            
            MoveWindow hwnd, iX, iY, iWidth, iHeight, 1
        Case WM_CTLCOLORDLG, WM_CTLCOLORSTATIC
            Dim tLB As LOGBRUSH
            'Debug.Print wParam
            
            Call SetTextColor(wParam, lForecolor)
            Call SetBkColor(wParam, lBackcolor)
            
            tLB.lbColor = lBackcolor
            
            SubMsgBox = CreateBrushIndirect(tLB)
            Exit Function
        'Added by manavo11
        
        Case WM_CTLCOLORBTN
            'Customize the MessageBox Buttons if neccessary..
            'First Process the Default Action of the Message (Draw the Button)
            SubMsgBox = CallWindowProc(lPrevWnd, hwnd, Msg, wParam, ByVal lParam)
            'Now Change the Button Text if Required
            If Not bCustom Then Exit Function
            If lButton = 0 Then sHwnd = ""
            'If this Button has Been Modified Already then Exit
            If InStr(sHwnd, " " & Trim(Str(lParam)) & " ") Then Exit Function
            sText = sButtons(lButton)
            sHwnd = sHwnd & " " & Trim(Str(lParam)) & " "
            lButton = lButton + 1
            'Modify the Button Text
            SendMessage lParam, WM_SETTEXT, Len(sText), ByVal sText
            
            'Added by manavo11
            If sText = sDefaultButton Then
                SetFocus lParam
            End If
            'Added by manavo11
            
            Exit Function
            
        Case WM_DESTROY
            'Remove the MsgBox Subclassing
            Call SetWindowLong(hwnd, GWL_WNDPROC, lPrevWnd)
        End Select
        SubMsgBox = CallWindowProc(lPrevWnd, hwnd, Msg, wParam, ByVal lParam)
    End Function
     
    Private Function HookWindow(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim tCWP As CWPSTRUCT
        Dim sClass As String
        'This is where you need to Hook the Messagebox
        CopyMemory tCWP, ByVal lParam, Len(tCWP)
        If tCWP.message = WM_CREATE Then
            sClass = Space(255)
            sClass = Left(sClass, GetClassName(tCWP.hwnd, ByVal sClass, 255))
            If sClass = "#32770" Then
                'Subclass the Messagebox as it's created
                lPrevWnd = SetWindowLong(tCWP.hwnd, GWL_WNDPROC, AddressOf SubMsgBox)
            End If
        End If
        HookWindow = CallNextHookEx(lHook, nCode, wParam, ByVal lParam)
    End Function
     
    Public Function MsgBoxEx(ByVal Prompt As String, Optional ByVal Buttons As Long = vbOKOnly, Optional ByVal Title As String, Optional ByVal HelpFile As String, Optional ByVal Context As Long, Optional ByRef CustomButtons As Variant, Optional DefaultButton As String, Optional X As String, Optional Y As String, Optional Width As String, Optional Height As String, Optional ByVal ForeColor As ColorConstants = -1, Optional ByVal BackColor As ColorConstants = -1) As Long
        Dim lReturn As Long
        
        bCustom = (Buttons = vbCustom)
        If bCustom And IsMissing(CustomButtons) Then
            MsgBox "When using the Custom option you need to supply some Buttons in the ""CustomButtons"" Argument.", vbExclamation + vbOKOnly, "Error"
            Exit Function
        End If
        lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookWindow, App.hInstance, App.ThreadID)
        'Set the Defaults
        If Len(Title) = 0 Then Title = App.Title
        If bCustom Then
            'User wants to use own Button Titles..
            If TypeName(CustomButtons) = "String" Then
                ReDim sButtons(0)
                sButtons(0) = CustomButtons
                Buttons = 0
            Else
                sButtons = CustomButtons
                Buttons = UBound(sButtons)
            End If
        End If
        
        'Added by manavo11
        lForecolor = GetSysColor(COLOR_BTNTEXT)
        lBackcolor = GetSysColor(COLOR_BTNFACE)
        
        If ForeColor >= 0 Then lForecolor = ForeColor
        If BackColor >= 0 Then lBackcolor = BackColor
        
        sDefaultButton = DefaultButton
        
        iX = X
        iY = Y
        iWidth = Width
        iHeight = Height
        
        iButtonCount = UBound(sButtons)
        iButtonWidth = (iWidth - (2 * 14) - (6 * (Buttons + 1))) / (Buttons + 1)
        'Added by manavo11
        
        lButton = 0
        
        'Show the Modified MsgBox
        lReturn = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
        Call UnhookWindowsHookEx(lHook)
        'If it's a Custom Button MsgBox, Alter the Return Value
        If bCustom Then lReturn = lReturn - (UBound(CustomButtons) + 1)
        bCustom = False
        MsgBoxEx = lReturn
    End Function
