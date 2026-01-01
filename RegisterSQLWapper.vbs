With CreateObject("Scripting.FileSystemObject")
   Dim SQLPath: SQLPath = .GetParentFolderName(WScript.ScriptFullName) & "\VBSQLite12.DLL"
   Dim RegSvr:  RegSvr  = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe" 
   If Not .FileExists(SQLPath) Then MsgBox "Couldn't find VBSQLite12.dll in:" & vbLF & .GetParentFolderName(WScript.ScriptFullName): WScript.Quit
   If Not .FileExists(RegSvr)  Then RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe" 
 
   CreateObject("Shell.Application").ShellExecute """" & RegSvr & """", """" & SQLPath & """", "", "runas", 1
End With