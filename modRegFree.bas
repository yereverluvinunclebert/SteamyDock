Attribute VB_Name = "modRegFree"
'A readymade "PlugIn-Module" you can include into your Projects, to be able to load Classes from COM-Dlls
'without registering them priorily - and just to be clear - there is no "dynamic re-registering" involved -
'DirectCOM.dll will load the appropriate Class-Instances without touching the Systems Registry in any way...
'
'There's 3 Functions exposed from this Module:
'- GetInstanceFromBinFolder ... loads Classes from Dlls located in a SubFolder, relative to your Exe-Path
'- GetInstance              ... same as above - but allowing absolute Dll-Paths (anywhere in the FileSystem)
'- GetExePath               ... just a small helper, to give the correct Exe-Path, even when called from within Dlls
'
'the approach is *very* reliable (in use for nearly two decades now, it works from Win98 to Windows-10)
'So, happy regfree COM-Dll-loading... :-) (Olaf Schmidt, in Dec. 2014)

Option Explicit

Public New_c As Object

'we need only two exports from the small DirectCOM.dll Helper here
Private Declare Function GetInstanceEx Lib "DirectCOM" (spFName As Long, spClassName As Long, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
Private Declare Function GETINSTANCELASTERROR Lib "DirectCOM" () As String

Private Declare Function LoadLibraryW& Lib "kernel32" (ByVal lpLibFileName&)
Private Declare Function GetModuleFileNameW& Lib "kernel32" (ByVal hMod&, ByVal lpFileName&, ByVal nSize&)
 
'a convenience-function which loads Classes from Dlls (residing in a SubFolder below your Exe-Path)
'just adjust the Optional RelBinFolderName-Param to your liking (currently specified as "Bin")
Public Function GetInstanceFromBinFolder(ByVal ShortDllFileName As String, ClassName As String, _
                                         Optional RelBinFolderName$ = "Bin") As Object
  Select Case LCase$(Right$(ShortDllFileName, 4))
    Case ".dll", ".ocx" 'all fine, nothing to do
    Case Else: ShortDllFileName = ShortDllFileName & ".dll" 'expand the ShortFileName about the proper file-ending when it was left out
  End Select

  Set GetInstanceFromBinFolder = GetInstance(GetExePath & RelBinFolderName & "\" & ShortDllFileName, ClassName)
End Function

'the generic Variant, which needs a full (user-provided), absolute Dll-PathFileName in the first Param
Public Function GetInstance(FullDllPathFileName As String, ClassName As String) As Object
  If Len(FullDllPathFileName) = 0 Or Len(ClassName) = 0 Then Err.Raise vbObjectError, , "Empty-Param(s) were passed to GetInstance"
  
  EnsureDirectCOMDllPreLoading FullDllPathFileName 'will raise an Error, when DirectCOM.dll was not found in "relative Folders"
  
  On Error Resume Next
    Set GetInstance = GetInstanceEx(StrPtr(FullDllPathFileName), StrPtr(ClassName), True)
  If Err Then
    On Error GoTo 0: Err.Raise vbObjectError, Err.Source & ".GetInstance", Err.Description
  ElseIf GetInstance Is Nothing Then
    On Error GoTo 0: Err.Raise vbObjectError, Err.Source & ".GetInstance", GETINSTANCELASTERROR()
  End If
End Function

'always returns the Path to the Executable (even when called from within COM-Dlls, which resolve App.Path to their own location)
Public Function GetExePath(Optional ExeName As String) As String
Dim S As String, Pos As Long: Const MaxPath& = 260
Static stExePath As String, stExeName As String
  If Len(stExePath) = 0 Then 'resolve it once
    S = Space$(MaxPath)
    S = Left$(S, GetModuleFileNameW(0, StrPtr(S), Len(S)))
    Pos = InStrRev(S, "\")
    
    stExeName = Mid$(S, Pos + 1)
    stExePath = Left$(S, Pos) 'preserve the BackSlash at the end
    Select Case UCase$(stExeName) 'when we run in the VB-IDE, ...
      Case "VB6.EXE", "VB5.EXE": stExePath = App.Path & "\" 'we resolve to the App.Path instead
    End Select
  End If
  
  ExeName = stExeName
  GetExePath = stExePath
End Function

Private Sub EnsureDirectCOMDllPreLoading(FullDllPathFileName As String)
Static hDirCOM As Long
  If hDirCOM Then Exit Sub  'nothing to do, DirectCOM.dll was already found and pre-loaded
  If hDirCOM = 0 Then hDirCOM = LoadLibraryW(StrPtr(GetExePath & "Bin\vbRichClient5.dll"))
  If hDirCOM = 0 Then Err.Raise vbObjectError, Err.Source & ".GetInstance", "Couldn't pre-load vbRichClient5.dll"
End Sub


