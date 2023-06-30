With CreateObject("Scripting.FileSystemObject")
   Dim RC5Path: RC5Path = .GetParentFolderName(WScript.ScriptFullName) & "\vbWidgets.dll"
   Dim RegSvr:  RegSvr  = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe" 
   If Not .FileExists(RC5Path) Then MsgBox "Couldn't find vbWidgets.dll in:" & vbLF & .GetParentFolderName(WScript.ScriptFullName): WScript.Quit
   If Not .FileExists(RegSvr)  Then RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe" 
 
   CreateObject("Shell.Application").ShellExecute """" & RegSvr & """", """" & RC5Path & """", "", "runas", 1
End With