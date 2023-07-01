Attribute VB_Name = "mdlMain"
Option Explicit


Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF

Public Const LOGPIXELSY = 90

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE = 32
Private Const FW_BOLD = 700
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_OEMTEXT = 7
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS

Public Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hWnd As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public debugflg As Integer
Public modFlag As Boolean
Public tmpSettingsFile As String


Public classicTheme As Boolean
Public storeThemeColour As Long

Public startupFlg As Boolean
'Public rDRunAppInterval As String
'Public rDAlwaysAsk As String
'Public rDGeneralReadConfig As String
'Public rDGeneralWriteConfig As String
'Public rDSkinTheme As String
'Public rDDefaultDock As String

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
(pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor _
    As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public rDEnableBalloonTooltips As Boolean
Public sdChkToggleDialogs As String

Public dockSettingsXPos As String
Public dockSettingsYPos As String




'---------------------------------------------------------------------------------------
' Procedure : MulDiv
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
    
    ' variables declared
    Dim lngTemp As Long
   On Error GoTo MulDiv_Error

  On Error GoTo MulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
MulDiv_end:
  MulDiv = lngTemp
  Exit Function
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err

   On Error GoTo 0
   Exit Function

MulDiv_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MulDiv of Module mdlMain"
End Function
'---------------------------------------------------------------------------------------
' Procedure : ByteToString
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ByteToString(aBytes() As Byte) As String
      
    ' variables declared
    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
   On Error GoTo ByteToString_Error

  dwBytePoint = LBound(aBytes)
  While dwBytePoint <= UBound(aBytes)
    dwByteVal = aBytes(dwBytePoint)
    If dwByteVal = 0 Then
      ByteToString = szOut
      Exit Function
    Else
      szOut = szOut & Chr$(dwByteVal)
    End If
    dwBytePoint = dwBytePoint + 1
  Wend
  ByteToString = szOut

   On Error GoTo 0
   Exit Function

ByteToString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ByteToString of Module mdlMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(InString As String, ByteArray() As Byte)
    
    ' variables declared
    Dim intLbound As Integer
  Dim intUbound As Integer
  Dim intLen As Integer
  Dim intX As Integer
   On Error GoTo StringToByte_Error

  intLbound = LBound(ByteArray)
  intUbound = UBound(ByteArray)
  intLen = Len(InString)
  If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
For intX = 1 To intLen
ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StringToByte of Module mdlMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : test_DialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function test_DialogFont(ctl As Control) As Boolean
        
    ' variables declared
    Dim f As FormFontInfo
    ' set the defaults
   On Error GoTo test_DialogFont_Error

    With f
      .Color = 0
      .Height = 12
      .Weight = 700
      .Italic = False
      .UnderLine = False
      .Name = "Arial"
    End With
    
    Call DialogFont(f)
    
    With f
'        DebugPrint "Font Name: "; .Name
'        DebugPrint "Font Size: "; .Height
'        DebugPrint "Font Weight: "; .Weight
'        DebugPrint "Font Italics: "; .Italic
'        DebugPrint "Font Underline: "; .UnderLine
'        DebugPrint "Font Color: "; .Color
        
        ctl.FontName = .Name
        ctl.FontSize = .Height
        ctl.FontWeight = .Weight
        ctl.FontItalic = .Italic
        ctl.FontUnderline = .UnderLine
        ctl.ForeColor = .Color
        ctl = .Name & " - Size:" & .Height
    End With
    test_DialogFont = True

   On Error GoTo 0
   Exit Function

test_DialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure test_DialogFont of Module mdlMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : DialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DialogFont(ByRef f As FormFontInfo) As Boolean
      
    ' variables declared
    Dim LF As LOGFONT, FS As FONTSTRUC
  Dim lLogFontAddress As Long, lMemHandle As Long, hWndAccessApp As Long

   On Error GoTo DialogFont_Error

  LF.lfWeight = f.Weight
  LF.lfItalic = f.Italic * -1
  LF.lfUnderline = f.UnderLine * -1
  LF.lfHeight = -MulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
  Call StringToByte(f.Name, LF.lfFaceName())
  FS.rgbColors = f.Color
  FS.lStructSize = Len(FS)

  lMemHandle = GlobalAlloc(GHND, Len(LF))
  If lMemHandle = 0 Then
    DialogFont = False
    Exit Function
  End If

  lLogFontAddress = GlobalLock(lMemHandle)
  If lLogFontAddress = 0 Then
    DialogFont = False
    Exit Function
  End If

  CopyMemory ByVal lLogFontAddress, LF, Len(LF)
  FS.lpLogFont = lLogFontAddress
  FS.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
  If ChooseFont(FS) = 1 Then
    CopyMemory LF, ByVal lLogFontAddress, Len(LF)
    f.Weight = LF.lfWeight
    f.Italic = CBool(LF.lfItalic)
    f.UnderLine = CBool(LF.lfUnderline)
    f.Name = ByteToString(LF.lfFaceName())
    f.Height = CLng(FS.iPointSize / 10)
    f.Color = FS.rgbColors
    DialogFont = True
  Else
    DialogFont = False
  End If

   On Error GoTo 0
   Exit Function

DialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DialogFont of Module mdlMain"
End Function




' Show the common dialog for choosing a color.
' Return the chosen color, or -1 if the dialog is canceled
'
' hParent is the handle of the parent form
' bFullOpen specifies whether the dialog will be open with the Full style
' (allows to choose many more colors)
' InitColor is the color initially selected when the dialog is open

' Example:
'    Dim oleNewColor As OLE_COLOR
'    oleNewColor = ShowColorsDialog(Me.hwnd, True, vbRed)
'    If oleNewColor <> -1 Then Me.BackColor = oleNewColor

'---------------------------------------------------------------------------------------
' Procedure : ShowColorDialog
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function ShowColorDialog(Optional ByVal hParent As Long, _
    Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR) _
    As Long
        
    ' variables declared
    Dim CC As ChooseColorStruct
    Dim aColorRef(15) As Long
    Dim lInitColor As Long

    ' translate the initial OLE color to a long value
   On Error GoTo ShowColorDialog_Error

    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    'fill the ChooseColorStruct struct
    With CC
        .lStructSize = Len(CC)
        .hWndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .rgbResult = lInitColor
        .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
            CC_FULLOPEN, 0)
    End With

    ' Show the dialog
    If ChooseColor(CC) Then
        'if not cancelled, return the color
        ShowColorDialog = CC.rgbResult
    Else
        'else return -1
        ShowColorDialog = -1
    End If

   On Error GoTo 0
   Exit Function

ShowColorDialog_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowColorDialog of Module mdlMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Convert_Dec2RGB
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Convert_Dec2RGB(ByVal myDECIMAL As Long) As String
      
    ' variables declared
    Dim myRED As Long
  Dim myGREEN As Long
  Dim myBLUE As Long

   On Error GoTo Convert_Dec2RGB_Error

  myRED = myDECIMAL And &HFF
  myGREEN = (myDECIMAL And &HFF00&) \ 256
  myBLUE = myDECIMAL \ 65536

  Convert_Dec2RGB = CStr(myRED) & "," & CStr(myGREEN) & "," & CStr(myBLUE)

   On Error GoTo 0
   Exit Function

Convert_Dec2RGB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Convert_Dec2RGB of Module mdlMain"

End Function

''# preparation (in a separate module)

'---------------------------------------------------------------------------------------
' Procedure : FindWindowHandle
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function FindWindowHandle(Caption As String) As Long
   On Error GoTo FindWindowHandle_Error

  FindWindowHandle = FindWindow(vbNullString, Caption)

   On Error GoTo 0
   Exit Function

FindWindowHandle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindWindowHandle of Module mdlMain"
End Function

