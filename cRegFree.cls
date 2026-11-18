VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegFree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ConstantsEnum
    MEMBERID_NIL = -1
    S_OK
    S_FALSE
    REGKIND_NONE
    CC_STDCALL = 4
    #If WIN64 Then
        PTR_SIZE = 8
    #Else
        PTR_SIZE = 4
    #End If
End Enum

Private Enum vtbInterfaceOffsets
    ITypeLib_FindName = 11 * PTR_SIZE
    ITypeInfo_GetTypeAttr = 3 * PTR_SIZE
    ITypeInfo_ReleaseTypeAttr = 19 * PTR_SIZE
    IClassFactory_CreateInstance = 3 * PTR_SIZE
End Enum

Private Const sDllGetClassObject As String = "DllGetClassObject"

#If WIN64 Then
    Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As LongPtr) As LongPtr
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" Alias "#146" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long
    Private Declare PtrSafe Function LoadTypeLibEx Lib "oleaut32" Alias "#183" (ByVal lpszFile As LongPtr, ByVal RegKind As Long, pptLib As IUnknown) As Long
#Else
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
    Private Declare Function DispCallFunc Lib "oleaut32" Alias "#146" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long
    Private Declare Function LoadTypeLibEx Lib "oleaut32" Alias "#183" (ByVal lpszFile As Long, ByVal RegKind As Long, pptLib As IUnknown) As Long
#End If

Private ParamTypes(0 To 10) As Integer, ParamValues(0 To 10) As LongPtr, lParamCount As Long, lpInterface As LongPtr, vParams As Variant, _
        IID_IClassFactory(0 To 1) As Currency, IID_IUnknown(0 To 1) As Currency, lpDllGetClassObject As LongPtr, colClassFactory As Collection, ITypeLib As IUnknown

Friend Function LoadTypeLibFromDLL(ByVal sLibName As String, Optional lDllResourceNumber As Long) As Boolean
    If ITypeLib Is Nothing Then
        lpDllGetClassObject = GetModuleHandleW(StrPtr(sLibName)) ' Check if the library had already been loaded
        If lpDllGetClassObject = 0 Then lpDllGetClassObject = LoadLibraryW(StrPtr(sLibName)) ' If not then we load it
        lpDllGetClassObject = GetProcAddress(lpDllGetClassObject, sDllGetClassObject) ' Get the pointer to the DllGetClassObject function
        If lpDllGetClassObject Then
            If lDllResourceNumber Then sLibName = sLibName & ChrW$(92) & lDllResourceNumber ' Append a backslash and the resource number (if any) to the name of the library
            LoadTypeLibFromDLL = LoadTypeLibEx(StrPtr(sLibName), REGKIND_NONE, ITypeLib) = S_OK ' REGKIND_NONE calls LoadTypeLibEx without the registration process enabled
        End If
    Else
        LoadTypeLibFromDLL = True
    End If
End Function

Friend Function CreateObj(sClassName As String, Optional sLibName As String, Optional lDllResourceNumber As Long) As Object
 Dim IClassFactory As IUnknown, RegFreeIUnknown As IUnknown, ITypeInfo As IUnknown, rgMemId As Long, pcFound As Long
 
 #If WIN64 Then
    Dim lpTypeAttr As LongPtr
    If InvokeObj(ClassFactory(sClassName), IClassFactory_CreateInstance, vbNullPtr, VarPtr(IID_IUnknown(0)), VarPtr(RegFreeIUnknown)) = S_OK Then ' Create an instance of this class
        Set CreateObj = RegFreeIUnknown ' Get the IDispatch implementation of this class
    ElseIf LoadTypeLibFromDLL(sLibName, lDllResourceNumber) Then
        pcFound = 1 ' We want to find only one instance of this class name (there shouldn't be duplicates anyway)
        InvokeObj ITypeLib, ITypeLib_FindName, StrPtr(sClassName), 0&, VarPtr(ITypeInfo), VarPtr(rgMemId), VarPtr(pcFound) ' Search the TypeLib for our class name
        If rgMemId = MEMBERID_NIL Then ' If the class name is found then "rgMemId" will return MEMBERID_NIL
            InvokeObj ITypeInfo, ITypeInfo_GetTypeAttr, VarPtr(lpTypeAttr) ' The first member of the "TypeAttr" structure is the class GUID so we don't need to CopyMemory its contents
            If lpTypeAttr Then
                If InvokeObj(Nothing, lpDllGetClassObject, lpTypeAttr, VarPtr(IID_IClassFactory(0)), VarPtr(IClassFactory)) = S_OK Then ' Call DllGetClassObject to retrieve the class object from the DLL object handler
                    colClassFactory.Add IClassFactory, sClassName: Set CreateObj = CreateObj(sClassName) ' Add this ClassFactory to the collection and create an instance
                End If
                InvokeObj ITypeInfo, ITypeInfo_ReleaseTypeAttr, lpTypeAttr ' Release the previously allocated "TypeAttr" structure
            End If
        End If
    End If
 #Else
    Dim lpTypeAttr As Long
    If InvokeObj(ClassFactory(sClassName), IClassFactory_CreateInstance, 0&, VarPtr(IID_IUnknown(0)), VarPtr(RegFreeIUnknown)) = S_OK Then ' Create an instance of this class
        Set CreateObj = RegFreeIUnknown ' Get the IDispatch implementation of this class
    ElseIf LoadTypeLibFromDLL(sLibName, lDllResourceNumber) Then
        pcFound = 1 ' We want to find only one instance of this class name (there shouldn't be duplicates anyway)
        InvokeObj ITypeLib, ITypeLib_FindName, StrPtr(sClassName), 0&, VarPtr(ITypeInfo), VarPtr(rgMemId), VarPtr(pcFound) ' Search the TypeLib for our class name
        If rgMemId = MEMBERID_NIL Then ' If the class name is found then "rgMemId" will return MEMBERID_NIL
            InvokeObj ITypeInfo, ITypeInfo_GetTypeAttr, VarPtr(lpTypeAttr) ' The first member of the "TypeAttr" structure is the class GUID so we don't need to CopyMemory its contents
            If lpTypeAttr Then
                If InvokeObj(Nothing, lpDllGetClassObject, lpTypeAttr, VarPtr(IID_IClassFactory(0)), VarPtr(IClassFactory)) = S_OK Then ' Call DllGetClassObject to retrieve the class object from the DLL object handler
                    colClassFactory.Add IClassFactory, sClassName: Set CreateObj = CreateObj(sClassName) ' Add this ClassFactory to the collection and create an instance
                End If
                InvokeObj ITypeInfo, ITypeInfo_ReleaseTypeAttr, lpTypeAttr ' Release the previously allocated "TypeAttr" structure
            End If
        End If
    End If
#End If
 
End Function

Private Function ClassFactory(sClassName As String) As IUnknown
On Error Resume Next
    Set ClassFactory = colClassFactory(sClassName) ' Check whether this ClassFactory already exists in the collection
    If Err Then Err.Clear
End Function

Private Function InvokeObj(Interface As IUnknown, vtbOffset As LongPtr, ParamArray ParamsArray() As Variant) As Variant
Dim lRet As Long
    
    InvokeObj = S_FALSE: lpInterface = ObjPtr(Interface): vParams = ParamsArray ' Make a copy of the array of parameters to get rid of any VT_BYREF members
    For lParamCount = 0 To UBound(vParams): ParamTypes(lParamCount) = VarType(vParams(lParamCount)): ParamValues(lParamCount) = VarPtr(vParams(lParamCount)): Next lParamCount
    #If WIN64 Then

        If lpInterface <> vbNullPtr OrElse vtbOffset > &HFFFF& Then ' Call the object's method found at "vtbOffset" in its VTable or else a function pointer instead
            lRet = DispCallFunc(lpInterface, vtbOffset, CC_STDCALL, vbLong, lParamCount, ParamTypes(0), ParamValues(0), InvokeObj)
        End If
    #Else
        
        If lpInterface Then ' Call the object's method found at "vtbOffset" in its VTable
            lRet = DispCallFunc(lpInterface, vtbOffset, CC_STDCALL, vbLong, lParamCount, ParamTypes(0), ParamValues(0), InvokeObj)
        ElseIf vtbOffset > 1024 Then ' The object is "Nothing" so here we call a function pointer instead
            lRet = DispCallFunc(lpInterface, vtbOffset, CC_STDCALL, vbLong, lParamCount, ParamTypes(0), ParamValues(0), InvokeObj)
        End If
    #End If
    
    If lRet Then Debug.Print Hex$(lRet) ' Display a helpful error code if DispCallFunc was called with an incorrect number or type of parameters (and it didn't crash right away!)
End Function

Private Sub Class_Initialize()
    Set colClassFactory = New Collection
    IID_IClassFactory(0) = 0.0001@: IID_IClassFactory(1) = 504403158265495.5712@: IID_IUnknown(1) = IID_IClassFactory(1) ' These IIDs are very similar so we hold them as "Currency" constants
End Sub


