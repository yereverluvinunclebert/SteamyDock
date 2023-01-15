Attribute VB_Name = "mdlExplorerPaths"
Option Explicit
Public openExplorerPaths() As String
Public Declare Function PSGetNameFromPropertyKey Lib "propsys.dll" (PropKey As PROPERTYKEY, ppszCanonicalName As Long) As Long
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long) ' Frees memory allocated by the shell
Public Declare Function IUnknown_QueryService Lib "shlwapi" (ByVal pUnk As Long, guidService As UUID, riid As UUID, ppvOut As Any) As Long

Public Declare Function vbaObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long

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
Public Sub EnumWindows()
    'Dim li As ListItem
    Dim i As Long, j As Long
    'Dim s1 As String, s2 As String, s3 As String
    Dim siaSel As IShellItemArray
    Dim lpText As Long
    Dim sText As String
    Dim sItems() As String
    Dim punkitem As oleexp.IUnknown
    Dim lPtr As Long
    Dim pclt As Long
    Dim spsb As IShellBrowser
    Dim spsv As IShellView
    Dim spfv As IFolderView2
    Dim spsi As IShellItem
    Dim lpPath As Long
    Dim sPath As String
    Dim lsiptr As Long
    Dim openShellWindow As ShellWindows
    Dim spev As oleexp.IEnumVARIANT
    Dim spunkenum As oleexp.IUnknown
    Dim pVar As Variant
    Dim pdp As oleexp.IDispatch
    Dim useloop As Integer
    
    Set openShellWindow = New ShellWindows
    If openShellWindow.Count < 1 Then Exit Sub
    ReDim openExplorerPaths(openShellWindow.Count - 1)
    
    For useloop = 0 To openShellWindow.Count - 1
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
                        
'                                Set li = ListView1.ListItems.Add(, , vbNullString)
'                                With li
'                                    spsi.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpPath
'                                    sPath = LPWSTRtoStr(lpPath)
'                                    .SubItems(5) = sPath
'                                End With
                                
                                spsi.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpPath
                                sPath = LPWSTRtoStr(lpPath)
                                openExplorerPaths(useloop) = sPath
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

End Sub
