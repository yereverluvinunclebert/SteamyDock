VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMainIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHOffset As Long
Private mVOffset As Long
Private mIconWidth As Long
Private mIconIndex As Integer
Private mIconOpacity As Integer
Private mIconImage As String

Private Sub Class_Initialize()


    'mainIconxo = 0: mainIconyo = 0

End Sub
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get iconIndex() As Integer
   On Error GoTo iconIndexGet_Error

   iconIndex = mIconIndex

   On Error GoTo 0
   Exit Property

iconIndexGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let iconIndex(ByVal newValue As Integer)
   On Error GoTo iconIndexLet_Error

    If mIconIndex <> newValue Then mIconIndex = newValue Else Exit Property
    
    If mIconIndex = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

iconIndexLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'
'---------------------------------------------------------------------------------------
' Property  : hOffset
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get hOffset() As Long
   On Error GoTo hOffsetGet_Error

   hOffset = mHOffset

   On Error GoTo 0
   Exit Property

hOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property hOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : hOffset
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let hOffset(ByVal newValue As Long)
   On Error GoTo hOffsetLet_Error

    If mHOffset <> newValue Then mHOffset = newValue Else Exit Property
    
    If mHOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

hOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property hOffset of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : vOffset
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get vOffset() As Long
   On Error GoTo vOffsetGet_Error

   vOffset = mVOffset

   On Error GoTo 0
   Exit Property

vOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property vOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : vOffset
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let vOffset(ByVal newValue As Long)
   On Error GoTo vOffsetLet_Error

    If mVOffset <> newValue Then mVOffset = newValue Else Exit Property
    
    If mVOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

vOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property vOffset of Class Module cwIcon"
End Property
