Attribute VB_Name = "messagebox"
Option Explicit

Private Type OSVERSIONINFO
' used by API call GetVersionExW
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion(1 To 256) As Byte
End Type
   
'#If VBA7 Then
'Private Declare PtrSafe Function GetVersionExW Lib "kernel32" (lpOSVersinoInfo As OSVERSIONINFO) As Long
'' http://msdn.microsoft.com/en-us/library/ms724451%28VS.85%29.aspx
'
'Private Declare PtrSafe Function MessageBoxW Lib "user32.dll" ( _
'   ByVal hwnd As LongPtr, _
'   ByVal PromptPtr As LongPtr, _
'   ByVal TitlePtr As LongPtr, _
'   ByVal UType As VbMsgBoxStyle) _
'      As VbMsgBoxResult
'' http://msdn.microsoft.com/en-us/library/ms645505(VS.85).aspx
'
'Private Declare PtrSafe Function MessageBoxTimeoutW Lib "user32.dll" ( _
'      ByVal WindowHandle As LongPtr, _
'      ByVal PromptPtr As LongPtr, _
'      ByVal TitlePtr As LongPtr, _
'      ByVal UType As VbMsgBoxStyle, _
'      ByVal Language As Integer, _
'      ByVal Miliseconds As Long _
'      ) As VbMsgBoxResult
'' http://msdn.microsoft.com/en-us/library/windows/desktop/ms645507(v=vs.85).aspx (XP+, undocumented)
'
'#Else
' for Office before 2010 and also VB6
Private Declare Function GetVersionExW Lib "kernel32" (lpOSVersinoInfo As OSVERSIONINFO) As Long
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal PromptPtr As Long, _
   ByVal TitlePtr As Long, ByVal UType As VbMsgBoxStyle) As VbMsgBoxResult
Private Declare Function MessageBoxTimeoutW Lib "user32.dll" (ByVal HandlePtr As Long, _
   ByVal PromptPtr As Long, ByVal TitlePtr As Long, ByVal UType As VbMsgBoxStyle, _
   ByVal Language As Integer, ByVal Miliseconds As Long) As VbMsgBoxResult
'#End If

Public Const vbTimedOut As Long = 32000 ' return if MsgBoxW times out


Public OSVersion As Long
Public OSBuild As Long
Public OSBits As Long

' NumBits will be 32 if the VB/VBA system running this code is 32-bit. VB6 is always 32-bit
'  and all versions of MS Office up until Office 2010 are 32-bit. Office 2010+ can be installed
'  as either 32 or 64-bit
#If Win64 Then
Public Const NumBits As Byte = 64
#Else
Public Const NumBits As Byte = 32
#End If



Sub Init()

' Sets the operating system major version * 100 plus the Minor version in a long
' Ex- Windows Xp has major version = 5 and the minor version equal to 01 so the return is 501
Dim version_info As OSVERSIONINFO
OSBuild = 0
version_info.dwOSVersionInfoSize = LenB(version_info)  '276
If GetVersionExW(version_info) = 0 Then
   OSVersion = -1 ' error of some sort. Shouldn't happen.
Else
   OSVersion = (version_info.dwMajorVersion * 100) + version_info.dwMinorVersion
   If version_info.dwPlatformId = 0 Then
      OSVersion = 301 ' Win 3.1
   Else
      OSBuild = version_info.dwBuildNumber
      End If
   End If

' Sets OSBits=64 if running on a 64-bit OS, 32 if on a 32-bit OS. NOTE- This is not the
'  # bits of the program executing the program. 32-bit  OFFice or VBA6 would return
'  OSBits = 64 if the code is running on a machine that has is running 64-bit Windows.
If Len(Environ$("PROGRAMFILES(X86)")) > 0 Then OSBits = 64 Else OSBits = 32 ' can't be 16

End Sub


'#If VBA7 Then
'Public Function MsgBoxW( _
' Optional Prompt As String = "", _
' Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
' Optional Title As String = "", _
' Optional ByVal TimeOutMSec As Long = 0, _
' Optional flags As Long = 0, _
' Optional ByVal hwnd As LongPtr = 0) As VbMsgBoxResult
'#Else
Public Function MsgBoxW( _
 Optional Prompt As String = "", _
 Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
 Optional Title As String = "", _
 Optional ByVal TimeOutMSec As Long = 0, _
 Optional flags As Long = 0, _
 Optional ByVal hwnd As Long = 0) As VbMsgBoxResult
'#End If
' A UniCode replacement for MsgBox with optional Timeout
' Returns are the same as for VB/VBA's MsgBox call except
'  If there is an error (unlikely) the error code is returned as a negative value
'  If you specify a timeout number of milliseconds and the time elapses without
'   the user clicking a button or pressing Enter, the return is "vbTimedOut" (numeric value = 32000)
' Inuts are the same as for the VB/VBA version except for the added in;ut variable
'  TimeOutMSec which defaults to 0 (infinite time) but specifies a time that if the
'  message box is displayed for that long it will automatically close and return "vbTimedOut"
' NOTE- The time out feature was added in Windows XP so it is ignored if you run this
'  code on Windows 2000 or earlier.
' NOTE- The time out feature uses an undocumented feature of Windows and is not guaranteed
'  to be in future versions of Windows although it has been in all since XP.

If OSVersion < 600 Then ' WindowsVersion less then Vista
   Init
   If OSVersion < 600 Then ' earlier than Vista
      If (Buttons And 15) = vbAbortRetryIgnore Then Buttons = (Buttons And 2147483632) Or 6 ' (7FFFFFFF xor 15) or 6
      End If
   End If
If (OSVersion >= 501) And (TimeOutMSec > 0) Then ' XP and later only
   MsgBoxW = MessageBoxTimeoutW(hwnd, StrPtr(Prompt), StrPtr(Title), Buttons Or flags, 0, TimeOutMSec)
Else ' earlier than XP does not have timeout capability for MessageBox
   MsgBoxW = MessageBoxW(hwnd, StrPtr(Prompt), StrPtr(Title), Buttons Or flags)
   End If
If MsgBoxW = 0 Then MsgBoxW = Err.LastDllError ' this should never happen
End Function

