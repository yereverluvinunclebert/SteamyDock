Attribute VB_Name = "mdlExplorerPaths"
'---------------------------------------------------------------------------------------
' Module    : mdlExplorerPaths
' Author    : fafalone
' Date      : 11/04/2023
' Purpose   : Lists all explorer window details, lovely useful code, don't know how Fafalone
'             figured out all this but it is all so very useful, thanks Faf.!
'---------------------------------------------------------------------------------------

Option Explicit

Public Declare Function PSGetNameFromPropertyKey Lib "propsys.dll" (PropKey As PROPERTYKEY, ppszCanonicalName As Long) As Long
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long) ' Frees memory allocated by the shell
Public Declare Function IUnknown_QueryService Lib "shlwapi" (ByVal pUnk As Long, guidService As UUID, riid As UUID, ppvOut As Any) As Long

Public Declare Function vbaObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
    SysReAllocString VarPtr(LPWSTRtoStr), lPtr
    If fFree Then
        Call CoTaskMemFree(lPtr)
    End If
End Function

'-----------------------------------------------------
'FOLLOWING CODE NOT NEEDED IF YOU USE mIID.bas v4 or higher!

Public Function IID_IShellBrowser() As UUID
'{000214E2-0000-0000-C000-000000000046}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H214E2, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IShellBrowser = iid
End Function
Public Function SID_STopLevelBrowser() As UUID
'{4C96BE40-915C-11CF-99D3-00AA004AE837}

Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4C96BE40, CInt(&H915C), CInt(&H11CF), &H99, &HD3, &H0, &HAA, &H0, &H4A, &HE8, &H37)
 SID_STopLevelBrowser = iid
End Function
Public Function IID_IShellItem() As UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Public Function IID_IFolderView() As UUID
'{cde725b0-ccc9-4519-917e-325d72fab4ce}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HCDE725B0, CInt(&HCCC9), CInt(&H4519), &H91, &H7E, &H32, &H5D, &H72, &HFA, &HB4, &HCE)
 IID_IFolderView = iid
End Function
Public Function IID_IFolderView2() As UUID
'{1af3a467-214f-4298-908e-06b03e0b39f9}
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1AF3A467, CInt(&H214F), CInt(&H4298), &H90, &H8E, &H6, &HB0, &H3E, &HB, &H39, &HF9)
 IID_IFolderView2 = iid
End Function

Public Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : isExplorerRunning
' Author    : beededea
' Date      : 10/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function isExplorerRunning(ByRef NameProcess As String) As Boolean
    Dim a As String
    Dim windowCount As Integer
    Dim openExplorerPathArray() As String
    Dim useloop As Integer
    
    On Error GoTo isExplorerRunning_Error
   
    Call enumerateExplorerWindows(openExplorerPathArray(), windowCount)
    
    For useloop = 0 To windowCount - 1
        If NameProcess <> "" And LCase$(NameProcess) = LCase$(openExplorerPathArray(useloop)) Then
            isExplorerRunning = True
            Exit Function
        End If
    Next useloop
    
    isExplorerRunning = False

    On Error GoTo 0
    Exit Function

isExplorerRunning_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure isExplorerRunning of Module common"
            Resume Next
          End If
    End With
End Function
'---------------------------------------------------------------------------------------
' Procedure : enumerateExplorerWindows
' Author    : fafalone
' Date      : 10/04/2023
' Purpose   : Obtains the path for each explorer window
'---------------------------------------------------------------------------------------
'
Public Sub enumerateExplorerWindows(ByRef openExplorerPaths() As String, ByRef windowCount As Integer)
    Dim punkitem As oleexp.IUnknown
    Dim spsb As IShellBrowser
    Dim spsv As IShellView
    Dim spfv As IFolderView2
    Dim spsi As IShellItem
    Dim lpPath As Long: lpPath = 0
    Dim strPath As String: strPath = vbNullString
    Dim lsiptr As Long: lsiptr = 0
    Dim openShellWindow As ShellWindows
    Dim pdp As oleexp.IDispatch
    Dim useloop As Integer: useloop = 0
    
    'On Error GoTo 0 ' l_start ' essential
    
    On Error Resume Next ' handles automation error

l_start:
    Set openShellWindow = New ShellWindows
    windowCount = openShellWindow.Count
    If windowCount < 1 Then Exit Sub
    ReDim openExplorerPaths(windowCount - 1)
    
    For useloop = 0 To windowCount - 1
        Set pdp = openShellWindow.Item(CVar(useloop))
        Set punkitem = pdp
    
        If True Then
            If (pdp Is Nothing) = False Then
    
                IUnknown_QueryService ObjPtr(punkitem), SID_STopLevelBrowser, IID_IShellBrowser, spsb
                If (spsb Is Nothing) = False Then
    
                    spsb.QueryActiveShellView spsv
                    If (spsv Is Nothing) = False Then
    
                        Dim pUnk As oleexp.IUnknown
                        Set pUnk = spsv
                        pUnk.QueryInterface IID_IFolderView2, spfv
                        If (spfv Is Nothing) = False Then
    
                            spfv.getFolder IID_IShellItem, lsiptr
                            If lsiptr Then vbaObjSetAddRef spsi, lsiptr
                            If (spsi Is Nothing) = False Then
                                
                                spsi.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpPath
                                strPath = LPWSTRtoStr(lpPath)
                                openExplorerPaths(useloop) = strPath
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Set spsi = Nothing
        Set spsv = Nothing
        Set spsb = Nothing
        lsiptr = 0
    
    Next useloop

    On Error GoTo 0
    Exit Sub

enumerateExplorerWindows_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure enumerateExplorerWindows of Module mdlExplorerPaths"
            Resume Next
          End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : findExplorerHwndByPath
' Author    : beededea based upon Fafalone's code
' Date      : 20/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function findExplorerHwndByPath(sPath As String) As Long

    On Error GoTo e0
    
    Dim pWindows As ShellWindows
    Set pWindows = New ShellWindows
    Dim pWB2 As IWebBrowser2
    #If TWINBASIC Then
        Dim pDisp As IDispatch
    #Else
        Dim pDisp As oleexp.IDispatch
    #End If
    Dim pSP As IServiceProvider
    Dim pSB As IShellBrowser
    Dim pSView As IShellView
    Dim pFView As IFolderView2
    Dim pFolder As IShellItem
    Dim lpPath As LongPtr, sCurPath As String
    Dim nCount As Long
    Dim i As Long
    Dim hr As Long
    Dim lThisHwnd As Long: lThisHwnd = 0
    Dim lProcessID As Long: lProcessID = 0
    Dim lProcessThread As Long: lProcessThread = 0
    
    findExplorerHwndByPath = 0
    
    nCount = pWindows.Count
    If nCount < 1 Then
        Debug.Print "No open Explorer windows found."
        Exit Function
    End If
    For i = 0 To nCount - 1
        Set pDisp = pWindows.Item(i)
        If (pDisp Is Nothing) = False Then
            Set pSP = pDisp
            If (pSP Is Nothing) = False Then
                pSP.QueryService SID_STopLevelBrowser, IID_IShellBrowser, pSB
                If (pSB Is Nothing) = False Then
                    pSB.QueryActiveShellView pSView
                    If (pSView Is Nothing) = False Then
                        Set pFView = pSView
                        If (pFView Is Nothing) = False Then
                            pFView.getFolder IID_IShellItem, pFolder
                            pFolder.GetDisplayName SIGDN_FILESYSPATH, lpPath
                            sCurPath = LPWSTRtoStr(lpPath)
                            Debug.Print "CompPath " & sCurPath & "||" & sPath
                            If LCase$(sCurPath) = LCase$(sPath) Then
                                Set pWB2 = pDisp
                                If (pWB2 Is Nothing) = False Then
                                    
                                    lThisHwnd = pWB2.hWnd
                                    
                                    findExplorerHwndByPath = lThisHwnd ' return
                                    Exit Function
                                Else
                                    Debug.Print "Couldn't get IWebWebrowser2"
                                End If
                            End If
                        Else
                            Debug.Print "Couldn't get IFolderView"
                        End If
                    Else
                        Debug.Print "Couldn't get IShellView"
                    End If
                Else
                    Debug.Print "Couldn't get IShellBrowser"
                End If
            Else
                Debug.Print "Couldn't get IServiceProvider"
            End If
        Else
            Debug.Print "Couldn't get IDispatch"
        End If
    Next
    Debug.Print "Couldn't find path."
Exit Function
e0:
    Debug.Print "CloseExplorerPathByWindow.Error->0x" & Hex$(Err.Number) & ", " & Err.Description

   On Error GoTo 0
   Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseExplorerWindowByPath
' Author    : fafalone
' Date      : 20/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CloseExplorerWindowByPath(sPath As String)

    On Error GoTo e0
    
    Dim pWindows As ShellWindows
    Set pWindows = New ShellWindows
    Dim pWB2 As IWebBrowser2
    #If TWINBASIC Then
        Dim pDisp As IDispatch
    #Else
        Dim pDisp As oleexp.IDispatch
    #End If
    Dim pSP As IServiceProvider
    Dim pSB As IShellBrowser
    Dim pSView As IShellView
    Dim pFView As IFolderView2
    Dim pFolder As IShellItem
    Dim lpPath As LongPtr, sCurPath As String
    Dim nCount As Long
    Dim i As Long
    Dim hr As Long
    
    nCount = pWindows.Count
    If nCount < 1 Then
        Debug.Print "No open Explorer windows found."
        Exit Sub
    End If
    For i = 0 To nCount - 1
        Set pDisp = pWindows.Item(i)
        If (pDisp Is Nothing) = False Then
            Set pSP = pDisp
            If (pSP Is Nothing) = False Then
                pSP.QueryService SID_STopLevelBrowser, IID_IShellBrowser, pSB
                If (pSB Is Nothing) = False Then
                    pSB.QueryActiveShellView pSView
                    If (pSView Is Nothing) = False Then
                        Set pFView = pSView
                        If (pFView Is Nothing) = False Then
                            pFView.getFolder IID_IShellItem, pFolder
                            pFolder.GetDisplayName SIGDN_FILESYSPATH, lpPath
                            sCurPath = LPWSTRtoStr(lpPath)
                            Debug.Print "CompPath " & sCurPath & "||" & sPath
                            If LCase$(sCurPath) = LCase$(sPath) Then
                                Set pWB2 = pDisp
                                If (pWB2 Is Nothing) = False Then
                                    pWB2.Quit
                                    'pWB2.
                                    Exit Sub
                                Else
                                    Debug.Print "Couldn't get IWebWebrowser2"
                                End If
                            End If
                        Else
                            Debug.Print "Couldn't get IFolderView"
                        End If
                    Else
                        Debug.Print "Couldn't get IShellView"
                    End If
                Else
                    Debug.Print "Couldn't get IShellBrowser"
                End If
            Else
                Debug.Print "Couldn't get IServiceProvider"
            End If
        Else
            Debug.Print "Couldn't get IDispatch"
        End If
    Next
    Debug.Print "Couldn't find path."
Exit Sub
e0:
    Debug.Print "CloseExplorerPathByWindow.Error->0x" & Hex$(Err.Number) & ", " & Err.Description

   On Error GoTo 0
   Exit Sub

End Sub
