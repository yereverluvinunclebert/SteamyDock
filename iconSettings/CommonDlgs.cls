VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDlgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Common Dialog Choose font without the CommonDialog OCX.
'
'No dependencies.
'

Private Const LOGPIXELSY As Long = 90
Private Const LF_FACESIZE As Long = 32

Public Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum
'Case-preserving hack:
#If False Then
    Dim cdlOFNAllowMultiselect, cdlOFNCreatePrompt, cdlOFNExplorer, cdlOFNExtensionDifferent
    Dim cdlOFNFileMustExist, cdlOFNHideReadOnly, cdlOFNLongNames, cdlOFNNoChangeDir
    Dim cdlOFNNoDereferenceLinks, cdlOFNNoLongNames, cdlOFNNoReadOnlyReturn
    Dim cdlOFNNoValidate, cdlOFNOverwritePrompt, cdlOFNPathMustExist, cdlOFNReadOnly
    Dim cdlOFNShareAware
#End If

Public Enum FontsConstants
    'ShowFont constants.
    cdlCFANSIOnly = &H400&
    cdlCFApply = &H200&
    cdlCFBoth = &H3&
    cdlCFEffects = &H100&
    cdlCFFixedPitchOnly = &H4000&
    cdlCFForceFontExist = &H10000
    cdlCFLimitSize = &H2000&
    cdlCFInitFont = &H40& 'Loads our Font property values into the dialog as defaults.
    cdlCFNoScriptSel = &H800000
    cdlCFNoFaceSel = &H80000
    cdlCFNoSimulations = &H1000&
    cdlCFNoSizeSel = &H200000
    cdlCFNoStyleSel = &H100000
    cdlCFNoVectorFonts = &H800&
    cdlCFPrinterFonts = &H2&
    cdlCFScalableOnly = &H20000
    cdlCFScreenFonts = &H1&
    cdlCFTTOnly = &H40000
    cdlCFWYSIWYG = &H8000&
End Enum
'Case-preserving hack:
#If False Then
    Dim cdlCFANSIOnly, cdlCFApply, cdlCFBoth, cdlCFEffects, cdlCFFixedPitchOnly
    Dim cdlCFForceFontExist, cdlCFLimitSize, cdlCFInitFont, cdlCFNoScriptSel
    Dim cdlCFNoFaceSel, cdlCFNoSimulations, cdlCFNoSizeSel, cdlCFNoStyleSel
    Dim cdlCFNoVectorFonts, cdlCFPrinterFonts, cdlCFScalableOnly, cdlCFScreenFonts
    Dim cdlCFTTOnly, cdlCFWYSIWYG
#End If

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
   lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Private Type CHOOSEFONTType
   lStructSize As Long
   hwndOwner As Long
   hDC As Long
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


Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontA" ( _
    pChoosefont As CHOOSEFONTType) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nIndex As Long) As Long


Private Declare Function MulDiv Lib "kernel32" ( _
    ByVal nNumber As Long, _
    ByVal nNumerator As Long, _
    ByVal nDenominator As Long) As Long

'Shared properties, used with multiple dialog types.
Public flags As Long

'ShowOpen, ShowSave properties.
Public DefaultExt As String 'Value excludes period.
Public DialogTitle As String
Public FileName As String
Public FileTitle As String
Public Filter As String
Public FilterIndex As Long
Public InitDir As String
Public MaxFileSize As Long

'ShowFont properties.
Public Charset As Integer
Public FontColor As ColorConstants
Public FontBold As Boolean
Public FontItalic As Boolean
Public FontName As String
Public FontSize As Single
Public FontStrikeThru As Boolean
Public FontUnderline As Boolean
Public FontWeight As Integer
Public Max As Single
Public Min As Single

Private CF As CHOOSEFONTType
Private LF As LOGFONT


Private Sub InitChooseFont(ByVal hWnd As Long, ByVal hDC As Long)
    Dim bytFaceName() As Byte
    Dim intByte As Integer
    
    With LF
        If FontSize = 0 Then
            .lfHeight = 0
        Else
            .lfHeight = -MulDiv(FontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72)
        End If
        .lfWidth = 0
        .lfWeight = IIf(FontBold, 700, 400)
        .lfItalic = IIf(FontItalic, 1, 0)
        .lfUnderline = IIf(FontUnderline, 1, 0)
        .lfStrikeOut = IIf(FontStrikeThru, 1, 0)
        .lfCharSet = Charset
        bytFaceName = StrConv(Left$(FontName & String$(LF_FACESIZE, 0), LF_FACESIZE), vbFromUnicode)
        For intByte = 0 To LF_FACESIZE - 1
            .lfFaceName(intByte) = bytFaceName(intByte)
        Next
    End With
    With CF
        .hDC = hDC
        .hwndOwner = hWnd
        .nSizeMax = Max
        .nSizeMin = Min
        .rgbColors = FontColor
        If (flags And cdlCFBoth) = 0 Then
            .flags = flags Or cdlCFScreenFonts
        Else
            .flags = flags
        End If
    End With
End Sub

Private Sub ExtractChooseFont(ByVal hDC As Long)
    With LF
        FontSize = -MulDiv(.lfHeight, 72, GetDeviceCaps(hDC, LOGPIXELSY))
        FontBold = .lfWeight >= 600
        FontItalic = CBool(.lfItalic)
        FontStrikeThru = CBool(.lfStrikeOut)
        FontUnderline = CBool(.lfUnderline)
        FontWeight = .lfWeight
        Charset = .lfCharSet
        FontName = StrConv(.lfFaceName, vbUnicode)
        FontName = Left$(FontName, InStr(FontName, vbNullChar) - 1)
    End With
    With CF
        FontColor = .rgbColors
        flags = .flags
    End With
End Sub

Public Sub ClearFont()
    Dim NF As New StdFont
    
    With NF
        Charset = .Charset
        FontColor = vbBlack
        FontName = .Name
        FontSize = .Size
        FontBold = False
        FontItalic = False
        FontStrikeThru = False
        FontUnderline = False
        FontWeight = .Weight
        Max = 0
        Min = 0
    End With
End Sub

Public Function ShowFont(ByVal hWnd As Long, ByVal hDC As Long) As Boolean
    'Returns False on Cancel or error.
    InitChooseFont hWnd, hDC
    ShowFont = ChooseFont(CF) <> 0
    If ShowFont Then
        ExtractChooseFont hDC
    End If
End Function

Private Sub Class_Initialize()
    With CF
        .lStructSize = Len(CF)
        .hInstance = App.hInstance
        .lpLogFont = VarPtr(LF)
    End With
    ClearFont
    With OF
        .lStructSize = Len(OF)
        .hInstance = App.hInstance
    End With
    Filter = "All files (*.*)|*.*"
    FilterIndex = 1
    MaxFileSize = 256
End Sub
