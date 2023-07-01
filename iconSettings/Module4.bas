Attribute VB_Name = "Module4"
'
'' APIs and structures for opening a common dialog box to select files without OCX dependencies
'
'Public Type OPENFILENAME
'    lStructSize As Long    'The size of this struct (Use the Len function)
'    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
'    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
'    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
'    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
'    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
'    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
'    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
'    nMaxFile As Long             'The length of lpstrFile + 1
'    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
'    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
'    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
'    lpstrTitle As String         'The caption of the dialog.
'    flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
'    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
'    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
'    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
'    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
'    lpfnHook As Long             'Pointer to the hook procedure.
'    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
'End Type
'
'Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" ( _
'    lpofn As OPENFILENAME) As Long
'
'Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" ( _
'    lpofn As OPENFILENAME) As Long
'
'Public OF As OPENFILENAME
'Public x_OpenFilename As OPENFILENAME
'
'
'
'
'
'Public Type BROWSEINFO
'    hwndOwner As Long
'    pidlRoot As Long 'LPCITEMIDLIST
'    pszDisplayName As String
'    lpszTitle As String
'    ulFlags As Long
'    lpfn As Long  'BFFCALLBACK
'    lParam As Long
'    iImage As Long
'End Type
'Public Declare Function SHBrowseForFolderA Lib "Shell32.dll" (binfo As BROWSEINFO) As Long
'Public Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl&, ByVal szPath$) As Long
'Public Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
''these functions need to be in a BAS module and not a form or the AddressOf does not work.
'
'
'Private Function BrowseCallbackProc(ByVal hWnd&, ByVal msg&, ByVal lp&, ByVal InitDir$) As Long
'   Const BFFM_INITIALIZED As Long = 1
'   Const BFFM_SETSELECTION As Long = &H466
'   If (msg = BFFM_INITIALIZED) And (InitDir <> "") Then
'      Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal InitDir$)
'   End If
'   BrowseCallbackProc = 0
'End Function
'
'Private Function GetAddress(ByVal Addr As Long) As Long
'   GetAddress = Addr
'End Function
'
'Public Function BrowseFolder(ByVal hwndOwner&, DefFolder$) As String
'   Dim bi As BROWSEINFO, pidl&, newPath$
'   bi.hwndOwner = hwndOwner
'   bi.lpfn = GetAddress(AddressOf BrowseCallbackProc)
'   bi.lParam = StrPtr(DefFolder)
'   pidl = SHBrowseForFolderA(bi)
'   If (pidl) Then
'      newPath = String(260, 0)
'      If SHGetPathFromIDListA(pidl, newPath) Then
'         newPath = Left(newPath, InStr(1, newPath, Chr(0)) - 1)
'         BrowseFolder = newPath
'      End If
'      Call CoTaskMemFree(ByVal pidl&)
'   End If
'End Function

