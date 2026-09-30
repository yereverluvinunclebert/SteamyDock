Attribute VB_Name = "modCopyAPI"
'---------------------------------------------------------------------------------------
' Module    : modCopyAPI
' Author    : Isladogs
' Date      : 30/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

'Type declarations
'============================
Private Type SHFILEOPSTRUCT

#If VBA7 Then

      hWnd As LongPtr
#Else
      hWnd As Long
#End If

wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long

#If VBA7 Then
      hNameMappings As LongPtr
#Else
      hNameMappings As Long
#End If

lpszProgressTitle As String

End Type
'============================

'API declarations
'============================
#If VBA7 Then       'A2010 or later (32/64-bit)
      Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" _
            Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
#Else       'A2007 or earlier
      Private Declare Function SHFileOperation Lib "shell32.dll" _
            Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
#End If
'============================

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FO_COPY = &H2
'============================

'---------------------------------------------------------------------------------------
' Procedure : apiFileCopy
' Author    : Isladogs
' Date      : 30/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function apiFileCopy(src As String, Dest As String, _
      Optional NoConfirm As Boolean = False) As Boolean

      'PARAMETERS: src: Source File (FullPath)
      'dest: Destination File (FullPath or directory)
      'NoConfirm (Optional): If set to true, no confirmation box
      'is displayed when overwriting existing files, and no
      'copy progress dialog box is displayed
      'Returns (True if Successful, false otherwise)

      Dim WinType_SFO As SHFILEOPSTRUCT
      Dim lRet As Long
      Dim lflags As Long

    On Error GoTo apiFileCopy_Error

      lflags = FOF_ALLOWUNDO
      If NoConfirm Then lflags = lflags & FOF_NOCONFIRMATION

      With WinType_SFO
            .wFunc = FO_COPY
            .pFrom = src
            .pTo = Dest
            .fFlags = lflags
      End With

      lRet = SHFileOperation(WinType_SFO)
      apiFileCopy = (lRet = 0)

    On Error GoTo 0
    Exit Function

apiFileCopy_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure apiFileCopy of Module modCopyAPI"

End Function
