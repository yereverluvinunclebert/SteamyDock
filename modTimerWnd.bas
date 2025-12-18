Attribute VB_Name = "modTimerWnd"
'---------------------------------------------------------------------------------------
' Module    : modTimerWnd
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : creates a hidden window with zero dimensions intercepting timer events
'             used to host/create in-code timers that do not require VB6 timer controls to exist on a VBform
'             These timers are created using the form_load event of the frmTimer form. All timers together.
'---------------------------------------------------------------------------------------

Option Explicit

Public TimerManager As New clsTimerManager

' API used to create a hidden Window
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    ByVal lpParam As Long) As Long

' API used to register that Window
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" ( _
    ByRef wc As WNDCLASSEX) As Long

' API for catching messages
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" ( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

' Start Timer API
Public Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

' Kill Timer API
Public Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

' API to return a handle for the created window
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
    ByVal lpModuleName As String) As Long

Private Const WM_TIMER As Long = &H113

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As String
    hIconSm As Long
End Type

Private m_hWnd As Long


'---------------------------------------------------------------------------------------
' Procedure : CreateTimerWindow
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : Create a hidden timer window that captures the WM_TIMER messages
'---------------------------------------------------------------------------------------
'
Public Function CreateTimerWindow() As Long
    Static Registered As Boolean

    On Error GoTo CreateTimerWindow_Error

    If Not Registered Then
        Dim wc As WNDCLASSEX
        wc.cbSize = Len(wc)
        wc.lpfnWndProc = pDefFarWndProc(AddressOf WindowProc) ' Long Pointer to the Windows Procedure function that will be called
        wc.hInstance = GetModuleHandle(vbNullString)
        wc.lpszClassName = "VB6_TimerMgr"
        RegisterClassEx wc
        Registered = True
    End If

    If m_hWnd = 0 Then
        ' get the handle of the window we are now creating with these characteristics, title, co-ords and size (nil) and the API to return the handle
        m_hWnd = CreateWindowEx(0, "VB6_TimerMgr", vbNullString, _
                                0, 0, 0, 0, 0, 0, 0, _
                                GetModuleHandle(vbNullString), 0)
    End If

    CreateTimerWindow = m_hWnd

    On Error GoTo 0
    Exit Function

CreateTimerWindow_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateTimerWindow of Module modTimerWnd"
End Function


'---------------------------------------------------------------------------------------
' Procedure : pDefFarWndProc
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : used above within wc.lpfnWndProc to return a long integer pointer
'             to the Windows Procedure function that will be called
'---------------------------------------------------------------------------------------
'
Private Function pDefFarWndProc(ByVal CBFunc As Long) As Long
    On Error GoTo pDefFarWndProc_Error

    pDefFarWndProc = CBFunc

    On Error GoTo 0
    Exit Function

pDefFarWndProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pDefFarWndProc of Module modTimerWnd"
End Function


'---------------------------------------------------------------------------------------
' Procedure : WindowProc
' Author    : beededea & chatGPT (with corrections/documentation by me)
' Date      : 17/12/2025
' Purpose   : called on create window, manages messages and causes the timer events to trigger
'---------------------------------------------------------------------------------------
'
Private Function WindowProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    On Error GoTo WindowProc_Error

    ' handle each timer message
    If uMsg = WM_TIMER Then
        ' cause the dispatch function in the timeManager class to kick off
        TimerManager.Dispatch wParam
        Exit Function
    End If

    ' returning value
    WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)

    On Error GoTo 0
    Exit Function

WindowProc_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WindowProc of Module modTimerWnd"
End Function

