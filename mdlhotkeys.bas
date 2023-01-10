Attribute VB_Name = "mdlhotkeys"
' .01 mdlhotkeys.bas DAEB 27/01/2021 Added the hotkeys module to support system wide keypresses

'------------------------------------------------------------
' mdlhotkeys.bas
'
' Author: Aaron Young
' Origin: Written
' Purpose: Register system wide hotkeys
'
'------------------------------------------------------------

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WH_GETMESSAGE = 3
Private Const WM_HOTKEY = &H312

Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
    
Private lHookID As Long
Private lHotKeys As Long


' .01 mdlhotkeys.bas DAEB 27/01/2021 Added the hotkeys module to support system wide keypresses
'---------------------------------------------------------------------------------------
' Procedure : CallBackHook
' Author    : beededea
' Date      : 28/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CallBackHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tMSG As Msg
        
    On Error GoTo CallBackHook_Error

    CopyMemory tMSG, ByVal lParam, Len(tMSG)
    If tMSG.message = WM_HOTKEY And wParam Then
        'Execute whatever for the HotKey here
        dock.lPressed = tMSG.wParam
        'dock.Command1_Click
        'MsgBox "You pressed the Hotkey with the ID of: " & lPressed, vbSystemModal
        If hideDockForNMinutes = True Then
            hideDockForNMinutes = False
            Call dock.ShowDockNow
        Else
            Call dock.HideDockNow
        End If
    End If
    If nCode < 0 Then CallBackHook = CallNextHookEx(lHookID, nCode, wParam, ByVal lParam)

   On Error GoTo 0
   Exit Function

CallBackHook_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CallBackHook of Module mdlhotkeys"
End Function

' .01 mdlhotkeys.bas DAEB 27/01/2021 Added the hotkeys module to support system wide keypresses
'---------------------------------------------------------------------------------------
' Procedure : SetHotKey
' Author    : beededea
' Date      : 28/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SetHotKey(ByVal lSpecial As Long, ByVal lKey As Long) As Long
   Static lHotKeyID As Long
    
   On Error GoTo SetHotKey_Error

    lHotKeyID = lHotKeyID + 1
    If RegisterHotKey(0&, lHotKeyID, lSpecial, lKey) <> 0 Then
        lHotKeys = lHotKeys + 1
        SetHotKey = lHotKeyID
        If lHookID = 0 Then lHookID = SetWindowsHookEx(WH_GETMESSAGE, AddressOf CallBackHook, App.hInstance, App.ThreadID)
    End If

   On Error GoTo 0
   Exit Function

SetHotKey_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetHotKey of Module mdlhotkeys"
    
End Function
' .01 mdlhotkeys.bas DAEB 27/01/2021 Added the hotkeys module to support system wide keypresses
'---------------------------------------------------------------------------------------
' Procedure : RemoveHotKey
' Author    : beededea
' Date      : 28/01/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub RemoveHotKey(ByVal lHotKeyID As Long)
   On Error GoTo RemoveHotKey_Error

    If UnregisterHotKey(0&, lHotKeyID) Then
        If lHotKeys > 0 Then lHotKeys = lHotKeys - 1
    End If
    If lHotKeys = 0 And lHookID <> 0 Then
        Call UnhookWindowsHookEx(lHookID)
        lHookID = 0
    End If

   On Error GoTo 0
   Exit Sub

RemoveHotKey_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveHotKey of Module mdlhotkeys"
End Sub







