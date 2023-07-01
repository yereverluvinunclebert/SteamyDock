Attribute VB_Name = "modDebugPrint"
'
' This is a Stand-Alone module that can be thrown into any project.
' It works in conjunction with the PersistentDebugPrint program, and that program must be running to use this module.
' The only procedure you should worry about is the DebugPrint procedure.
' Basically, it does what it says, provides a "Debug" window that is persistent across your development IDE exits and starts (even IDE crashes).
'
Option Explicit
'
Private Type COPYDATASTRUCT
    dwData  As Long
    cbData  As Long
    lpData  As Long
End Type
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'
Dim mhWndTarget As Long
'
Const DoDebugPrint As Boolean = True
'

Public Sub DebugPrint(ParamArray vArgs() As Variant)
    ' Commas are allowed, but not semicolons.
    '
    If Not DoDebugPrint Then Exit Sub
    '
    Static bErrorMessageShown As Boolean
    ValidateTargetHwnd
    If mhWndTarget = 0& Then
        If Not bErrorMessageShown Then
            MsgBox "The Persistent Debug Print Window couldn't be found.  Be sure the PersistentDebugPrint program is running.", vbCritical, "Persistent Debug Message"
            bErrorMessageShown = True
            Exit Sub
        End If
    End If
    '
    Dim v       As Variant
    Dim sMsg    As String
    Dim bNext   As Boolean
    For Each v In vArgs
        If bNext Then
            sMsg = sMsg & Space$(8&)
            sMsg = Left$(sMsg, (Len(sMsg) \ 8&) * 8&)
        End If
        bNext = True
        sMsg = sMsg & CStr(v)
    Next
    '
    SendStringToAnotherWindow sMsg
End Sub

Private Sub ValidateTargetHwnd()
    If IsWindow(mhWndTarget) Then
        Select Case WindowClass(mhWndTarget)
        Case "ThunderForm", "ThunderRT6Form"
            If WindowText(mhWndTarget) = "Persistent Debug Print Window" Then
                Exit Sub
            End If
        End Select
    End If
    EnumWindows AddressOf EnumToFindTargetHwnd, 0&
End Sub

Private Function EnumToFindTargetHwnd(ByVal hwnd As Long, ByVal lParam As Long) As Long
    mhWndTarget = 0&                        ' We just set it every time to keep from needing to think about it before this is called.
    Select Case WindowClass(hwnd)
    Case "ThunderForm", "ThunderRT6Form"
        If WindowText(hwnd) = "Persistent Debug Print Window" Then
            mhWndTarget = hwnd
            Exit Function
        End If
    End Select
    EnumToFindTargetHwnd = 1&               ' Keep looking.
End Function

Private Function WindowClass(hwnd As Long) As String
    WindowClass = String$(1024&, vbNullChar)
    WindowClass = Left$(WindowClass, GetClassName(hwnd, WindowClass, 1024&))
End Function

Private Function WindowText(hwnd As Long) As String
    ' Form or control.
    WindowText = String$(GetWindowTextLength(hwnd) + 1&, vbNullChar)
    Call GetWindowText(hwnd, WindowText, Len(WindowText))
    WindowText = Left$(WindowText, InStr(WindowText, vbNullChar) - 1&)
End Function

Private Sub SendStringToAnotherWindow(sMsg As String)
    Dim cds             As COPYDATASTRUCT
    Dim lpdwResult      As Long
    Dim Buf()           As Byte
    Const WM_COPYDATA   As Long = &H4A&
    '
    ReDim Buf(1 To Len(sMsg) + 1&)
    Call CopyMemory(Buf(1&), ByVal sMsg, Len(sMsg)) ' Copy the string into a byte array, converting it to ASCII.
    cds.dwData = 3&
    cds.cbData = Len(sMsg) + 1&
    cds.lpData = VarPtr(Buf(1&))
    'Call SendMessage(hWndTarget, WM_COPYDATA, Me.hwnd, cds)
    SendMessageTimeout mhWndTarget, WM_COPYDATA, 0&, cds, 0&, 1000&, lpdwResult ' Return after a second even if receiver didn't acknowledge.
End Sub

