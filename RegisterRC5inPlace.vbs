With CreateObject("Scripting.FileSystemObject")
   Dim RC5Path: RC5Path = .GetParentFolderName(WScript.ScriptFullName) & "\vbRichClient5.dll"
   Dim RegSvr:  RegSvr  = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe" 
   If Not .FileExists(RC5Path) Then MsgBox "Couldn't find vbRichClient5.dll in:" & vbLF & .GetParentFolderName(WScript.ScriptFullName): WScript.Quit
   If Not .FileExists(RegSvr)  Then RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe" 
 
   CreateObject("Shell.Application").ShellExecute """" & RegSvr & """", """" & RC5Path & """", "", "runas", 1
End With