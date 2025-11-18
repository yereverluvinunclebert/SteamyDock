Attribute VB_Name = "modEmbeddedIcons"
'---------------------------------------------------------------------------------------
' Module    : modEmbeddedIcons
' Author    : beededea
' Date      : 13/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
    
    
'Public Type ImageCodecInfo
'   ClassID As IID
'   FormatID As IID
'   CodecName As Long      ' String Pointer; const WCHAR*
'   DllName As Long        ' String Pointer; const WCHAR*
'   FormatDescription As Long ' String Pointer; const WCHAR*
'   FilenameExtension As Long ' String Pointer; const WCHAR*
'   MimeType As Long       ' String Pointer; const WCHAR*
'   flags As ImageCodecFlags   ' Should be a Long equivalent
'   Version As Long
'   SigCount As Long
'   SigSize As Long
'   SigPattern As Long      ' Byte Array Pointer; BYTE*
'   SigMask As Long         ' Byte Array Pointer; BYTE*
'End Type
'
'' Information flags about image codecs
'Public Enum ImageCodecFlags
'   ImageCodecFlagsEncoder = &H1
'   ImageCodecFlagsDecoder = &H2
'   ImageCodecFlagsSupportBitmap = &H4
'   ImageCodecFlagsSupportVector = &H8
'   ImageCodecFlagsSeekableEncode = &H10
'   ImageCodecFlagsBlockingDecode = &H20
'
'   ImageCodecFlagsBuiltin = &H10000
'   ImageCodecFlagsSystem = &H20000
'   ImageCodecFlagsUser = &H40000
'End Enum

Private Declare Function PrivateExtractIcons Lib "user32" _
                Alias "PrivateExtractIconsA" ( _
                ByVal lpszFile As String, _
                ByVal nIconIndex As Long, _
                ByVal cxIcon As Long, _
                ByVal cyIcon As Long, _
                ByRef phIcon As Long, _
                ByRef pIconId As Long, _
                ByVal nIcons As Long, _
                ByVal flags As Long _
) As Long

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal hbm As Long, ByRef pBitMap As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal CallbackData As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, ByRef clsidEncoder As Any, ByRef encoderParams As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus

Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private Const DI_NORMAL = 3
Private Const LR_LOADFROMFILE As Long = &H10

Private Type PictDesc
    cbSizeofStruct  As Long
    PicType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

' APIs for drawing icons START
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long

Private Declare Function Ole_CreatePic Lib "olepro32" _
                Alias "OleCreatePictureIndirect" ( _
                ByRef lpPictDesc As PictDesc, _
                ByVal riid As Long, _
                ByVal fPictureOwnsHandle As Long, _
                ByRef iPic As IPicture _
) As Long

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef clsid As IID) As Long

Private Declare Function OLE_CLSIDFromString Lib "ole32" Alias "CLSIDFromString" (ByVal lpszProgID As Long, ByVal pclsid As Long) As Long

Private Enum OLE_ERROR_CODES
    S_OK = 0
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
    E_FAIL = &H80004005
    E_UNEXPECTED = &H8000FFFF
    E_INVALIDARG = &H80070057
End Enum


' NOTE: Enums evaluate to a Long
Private Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
   ProfileNotFound = 21
End Enum



'---------------------------------------------------------------------------------------
' Procedure : fExtractEmbeddedPNGFromEXe
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : The program extracts icons embedded within a DLL or an executable
'             you pass the name of the picbox you require and the image is displayed there
'             it should return all and not only the 16 and 32 bit icons as does extractIconEx
'             Also, on request, writes a PNG to a file on disc in the special folder area.
'             Returns success only when the PNG is extracted
'---------------------------------------------------------------------------------------
'
Public Function fExtractEmbeddedPNGFromEXE(ByVal FileName As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer, ByVal writePNGToFile As Boolean) As String
    
    Dim lIconIndex As Long: lIconIndex = 0
    Dim lxSize As Long: lxSize = 0
    Dim lySize As Long: lySize = 0
    Dim lhIcon() As Long
    Dim lhIconID() As Long
    Dim nIcons As Integer: nIcons = 0
    Dim lResult As Long: lResult = 0
    Dim lFlags As Long: lFlags = 0
    Dim i As Long: i = 0
    Dim pic As StdPicture ' interface for a Picture object
    Dim GSI As GdiplusStartupInput
    Dim lhToken As Long: lhToken = 0
    Dim lhGraphics As Long: lhGraphics = 0
    Dim lhImage As Long: lhImage = 0
    Dim ImageFormatPNG As IID
    Dim sOutputFilename As String: sOutputFilename = vbNullString
    Dim sJustTheFilename As String: sJustTheFilename = vbNullString
    Dim bSuccessSaveToPNG As GpStatus
    Dim encoderCLSID As IID

    On Error GoTo fExtractEmbeddedPNGFromEXe_Error
    
    GSI.GdiplusVersion = 1
    GdiplusStartup lhToken, GSI

    On Error Resume Next ' debug
    
    If FileName = "" Then MsgBox "filename is missing"
    If targetPicBox = Null Then MsgBox "targetPicBox is missing"
    If IconSize = 0 Then MsgBox "IconSize is missing"

    lIconIndex = 0
    i = 2 ' need some experimentation here
    
    'the boundaries of the icons you wish to extract packed into a 32bit LONG for an API call
    lxSize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
    lySize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
    
    ' lFlags
    '
    '    LR_DEFAULTCOLOR
    '    LR_CREATEDIBSECTION
    '    LR_DEFAULTSIZE
    '    LR_LOADFROMFILE
    '    LR_LfsOADMAP3DCOLORS
    '    LR_LOADTRANSPARENT
    '    LR_MONOCHROME
    '    LR_SHARED
    '    LR_VGACOLOR
    '
    lFlags = LR_LOADFROMFILE '16

    ' Call PrivateExtractIcons with the 5th param set to nothing, solely to obtain the total number of Icons in the file.
    lResult = PrivateExtractIcons(FileName, lIconIndex, lxSize, lySize, ByVal 0&, ByVal 0&, 0&, 0&)
    
    If lResult = 0 Then
        'MsgBox "Failed to extract icon."
        GoTo CleanUp
    End If
    
    ' The Filename is the resource string/filepath.
    ' lIconIndex is the index.
    ' lxSize and lySize are the desired sizes.
    ' 5th parameter is a pointer to the returned array of icon handles.
    ' piconid is an ID of each icon that best fits the current display device. The returned identifier is 0 if not obtained.
    ' nicons is the number of icons you wish to extract.
    
    ' If you call it with nicon set to this number and niconindex=0 it will extract ALL your icons in one go.
    ' eg. PrivateExtractIcons(sExeName, lIconIndex, lxSize, lySize,  lhIcon(LBound(lhIcon)), lhIconID(LBound(lhIconID)), nIcons * 2, LR_LOADFROMFILE)

    nIcons = lResult
    
    ' Dimension the arrays to the number of icons.
    ReDim lhIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
    ReDim lhIconID(lIconIndex To lIconIndex + nIcons * 2 - 1)

    ' use the undocumented PrivateExtractIcons to extract the icons we require where the 5th param is a pointer to the returned array of handles to extracted icons
    lResult = PrivateExtractIcons(FileName, lIconIndex, lxSize, _
                            lySize, lhIcon(LBound(lhIcon)), _
                            lhIconID(LBound(lhIconID)), _
                            nIcons * 2, lFlags)
        
    ' create an Ipicture icon with a handle, no specific size - to check as to a valid pic before we write directly to the targetPicBox
    Set pic = CreateIcon(lhIcon(i + lIconIndex - 1))
        
    ' resize and place the target picbox according to the size of the icon
    ' (rather than placing the icon in the middle of the picbox as I should, I can code that later)
    Call centrePreviewImage(targetPicBox, IconSize, 1)
            
    ' Draw the icon directly onto the respective picturebox control and save as a PNG
    If Not (pic Is Nothing) Then
        With targetPicBox
        
            'ensure the picbox is empty first
            .Picture = LoadPicture(vbNullString)
            .Cls
            .AutoRedraw = True

            'creates a GDI+ image bitmap (lhImage) using the icon handle from the icon handle array populated by PrivateExtractIcons
            lResult = GdipCreateBitmapFromHICON(lhIcon(LBound(lhIcon)), lhImage)
            If lResult <> 0 Or lhImage = 0 Then
                ' MsgBox "Failed to create bitmap from icon."
                GoTo CleanUp
            Else
                ' Creates a GDIP Graphics object (lhGraphics) that is associated with the current device context, that being the target picbox
                GdipCreateFromHDC .hDC, lhGraphics ' wrap target DC in GDI+
                
                ' Draws an image at a specified location (targetPicBox) using the image bitmap and graphics object, in effect writing the image to the picbox
                '                      lhGraphics, lhImage, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&
                GdipDrawImageRectRectI lhGraphics, lhImage, 0, 0, IconSize, IconSize, 0, 0, 256, 256, 2&, 0, 0, 0
                
'               centre image using a better method
'                        ScaleX(x, ScaleMode, vbPixels) - WidthPx \ 2, _
'                        ScaleY(y, ScaleMode, vbPixels) - HeightPx \ 2, _
'                        IconSize, _
'                        IconSize, _

                ' Now the code to extract the embedded PNG to a file in the special folder location.
                If writePNGToFile = True Then
                
                    ' take the filename, extract just the filename body minus the suffix, then point it to the special folder with a PNG suffix.
                    sJustTheFilename = Mid(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
                    sJustTheFilename = ExtractFilenameWithoutSuffix(sJustTheFilename)
                    sOutputFilename = SpecialFolder(SpecialFolder_AppData) & "\steamyDock\images\" & sJustTheFilename & ".png"
    
                    ' set the encoder class identifier to handle the image bitmap as a PNG
                    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatPNG
                        
                    ' extract a PNG of the image bitmap and save to file
                    bSuccessSaveToPNG = GdipSaveImageToFile(lhImage, StrPtr(sOutputFilename), ImageFormatPNG, ByVal 0&) = 0&
                    If bSuccessSaveToPNG = False Then
                        fExtractEmbeddedPNGFromEXE = ""
                        ' MsgBox "Failed to save PNG."
                    Else
                        fExtractEmbeddedPNGFromEXE = sOutputFilename
                    End If
                    
                End If
                
            End If
            .Refresh
        End With
    End If
    
CleanUp:

    ' get rid of the icons we created
    Call DestroyIcon(lhIcon(i + lIconIndex - 1))
    Call GdipDeleteGraphics(lhGraphics)
    Call GdipDisposeImage(lhImage): lhImage = 0&
    Call GdiplusShutdown(lhToken)

    On Error GoTo 0
    Exit Function

fExtractEmbeddedPNGFromEXe_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fExtractEmbeddedPNGFromEXe of Module mdlMain"
    
End Function

''---------------------------------------------------------------------------------------
'' Procedure : displayEmbeddedIcons
'' Author    : beededea
'' Date      : 05/07/2019
'' Purpose   : The program extracts icons embedded within a DLL or an executable
''             you pass the name of the picbox you require and the image is displayed there
''             it should return all and not only the 16 and 32 bit icons as does extractIconEx
''
''             I may not have coded this particularly well - but it works.
''---------------------------------------------------------------------------------------
''
''
'Public Sub displayEmbeddedIcons(ByVal FileName As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer, ByVal writePNGToFile As Boolean)
'
'    Dim lIconIndex As Long: lIconIndex = 0
'    Dim xSize As Long: xSize = 0
'    Dim ySize As Long: ySize = 0
'    Dim hIcon() As Long
'    Dim hIconID() As Long
'    Dim nIcons As Long: nIcons = 0
'    Dim Result As Long: Result = 0
'    Dim flags As Long: flags = 0
'    Dim i As Long: i = 0
'    Dim pic As StdPicture ' interface for a Picture object
'    Dim outputFilename As String: outputFilename = vbNullString
'    Dim GSI As GdiplusStartupInput
'    Dim hToken As Long: hToken = 0
'    Dim hGraphics As Long: hGraphics = 0
'    Dim hImage As Long: hImage = 0
'    Dim ImageFormatPNG As IID
'    Dim sOutputFilename As String: sOutputFilename = vbNullString
'    Dim sJustTheFilename As String: sJustTheFilename = vbNullString
'    Dim successSaveToPNG As Boolean: successSaveToPNG = False
'
'    On Error GoTo displayEmbeddedIcons_Error
'
'    GSI.GdiplusVersion = 1
'    GdiplusStartup hToken, GSI
'
'    On Error Resume Next ' debug
'
'    lIconIndex = 0
'    i = 2 ' need some experimentation here
'
'    'the boundaries of the icons you wish to extract packed into a 32bit LONG for an API call
'    xSize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
'    ySize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
'
'    ' flags
'    '
'    '    LR_DEFAULTCOLOR
'    '    LR_CREATEDIBSECTION
'    '    LR_DEFAULTSIZE
'    '    LR_LOADFROMFILE
'    '    LR_LfsOADMAP3DCOLORS
'    '    LR_LOADTRANSPARENT
'    '    LR_MONOCHROME
'    '    LR_SHARED
'    '    LR_VGACOLOR
'    '
'    flags = LR_LOADFROMFILE '16
'
'    ' Call PrivateExtractIcons with the 5th param set to nothing, solely to obtain the total number of Icons in the file.
'    Result = PrivateExtractIcons(FileName, lIconIndex, xSize, ySize, ByVal 0&, ByVal 0&, 0&, 0&)
'
'    If Result = 0 Then
'        MsgBox "Failed to extract icon."
'        GoTo CleanUp
'    End If
'
'    ' The Filename is the resource string/filepath.
'    ' lIconIndex is the index.
'    ' xSize and ySize are the desired sizes.
'    ' 5th parameter is a pointer to the returned array of icon handles.
'    ' piconid is an ID of each icon that best fits the current display device. The returned identifier is 0 if not obtained.
'    ' nicons is the number of icons you wish to extract.
'
'    ' If you call it with nicon set to this number and niconindex=0 it will extract ALL your icons in one go.
'    ' eg. PrivateExtractIcons(sExeName, lIconIndex, xSize, ySize,  hIcon(LBound(hIcon)), hIconID(LBound(hIconID)), nIcons * 2, LR_LOADFROMFILE)
'
'    nIcons = Result
'
'    ' Dimension the arrays to the number of icons.
'    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
'    ReDim hIconID(lIconIndex To lIconIndex + nIcons * 2 - 1)
'
'    ' use the undocumented PrivateExtractIcons to extract the icons we require where the 5th param is a pointer to the returned array of handles to extracted icons
'    Result = PrivateExtractIcons(FileName, lIconIndex, xSize, _
'                            ySize, hIcon(LBound(hIcon)), _
'                            hIconID(LBound(hIconID)), _
'                            nIcons * 2, flags)
'
'    ' create an Ipicture icon with a handle, no specific size - to check as to a valid pic before we write directly to the targetPicBox
'    Set pic = CreateIcon(hIcon(i + lIconIndex - 1))
'
'    ' resize and place the target picbox according to the size of the icon
'    ' (rather than placing the icon in the middle of the picbox as I should, I can code that later)
'
'    Call centrePreviewImage(targetPicBox, IconSize, 1)
'
'    ' Draw the icon directly onto the respective picturebox control and save as a PNG
'    If Not (pic Is Nothing) Then
'        With targetPicBox
'
'            'ensure the picbox is empty first
'            .Picture = LoadPicture(vbNullString)
'            .Cls
'            .AutoRedraw = True
'
'            'creates a GDI+ image bitmap (hImage) using the icon handle from the icon handle array populated by PrivateExtractIcons
'            Result = GdipCreateBitmapFromHICON(hIcon(LBound(hIcon)), hImage)
'            If Result <> 0 Or hImage = 0 Then
'                MsgBox "Failed to create bitmap from icon."
'                GoTo CleanUp
'            Else
'                ' Creates a GDIP Graphics object (hGraphics) that is associated with the current device context, that being the target picbox
'                GdipCreateFromHDC .hDC, hGraphics
'
'                ' Draws an image at a specified location using the image bitmap and graphics object, in effect writing the image to the picbox
'                '                      hGraphics, hImage, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&
'                GdipDrawImageRectRectI hGraphics, hImage, 0, 0, IconSize, IconSize, 0, 0, 256, 256, 2&, 0, 0, 0
'
''               centre image using a better method
''                        ScaleX(x, ScaleMode, vbPixels) - WidthPx \ 2, _
''                        ScaleY(y, ScaleMode, vbPixels) - HeightPx \ 2, _
''                        IconSize, _
''                        IconSize, _
'
'                ' In iconSettings we prove that it is possible to extract the PNG from extract the PNG from the DLL and write that to a file
'                ' this is of little use here as we write to a picbox and display our PNG image there
'                ' In SD, we will take this routine and use it to write a PNG to the local profile area and then insert the PNG into the dictionary at runtime startup.
'
'                If writePNGToFile = True Then
'
'                    ' take the filename, extract just the filename body minus the suffix, then point it to the special folder with a PNG suffix.
'                    sJustTheFilename = Mid(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
'                    sJustTheFilename = ExtractFilenameWithoutSuffix(sJustTheFilename)
'                    sOutputFilename = SpecialFolder(SpecialFolder_AppData) & "\steamyDock\images\" & sJustTheFilename & ".png"
'
'                    ' set the encoder class identifier to handle the image bitmap as a PNG
'                    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatPNG
'    '
'                    ' extract a PNG of the image bitmap and save to file
'                    successSaveToPNG = GdipSaveImageToFile(hImage, StrPtr(sOutputFilename), ImageFormatPNG, ByVal 0&) = 0&
'                    If successSaveToPNG = False Then
'                        MsgBox "Failed to save PNG."
'                    End If
'                End If
'
'            End If
'            .Refresh
'        End With
'    End If
'
'CleanUp:
'
'    ' get rid of the icons we created
'    Call DestroyIcon(hIcon(i + lIconIndex - 1))
'    Call GdipDeleteGraphics(hGraphics)
'    Call GdipDisposeImage(hImage): hImage = 0&
'    Call GdiplusShutdown(hToken)
'
'   On Error GoTo 0
'   Exit Sub
'
'displayEmbeddedIcons_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayEmbeddedIcons of Module mdlMain"
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : make32BitLong
' Author    : beededea
' Date      : 20/11/2019
' Purpose   : packing variables into a 32bit LONG for an API call
'---------------------------------------------------------------------------------------
'
Private Function make32BitLong(ByVal LoWord As Integer, Optional ByVal HiWord As Integer = 0) As Long

   On Error GoTo make32BitLong_Error
   
   If debugflg = 1 Then debugLog "%make32BitLong"

    make32BitLong = CLng(HiWord) * CLng(&H10000) + CLng(LoWord)

   On Error GoTo 0
   Exit Function

make32BitLong_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure make32BitLong of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateIcon
' Author    : beededea
' Date      : 14/07/2019
' Purpose   : This method creates an icon based on a handle
'---------------------------------------------------------------------------------------
'
Private Function CreateIcon(ByVal hImage As Long) As IPicture
    
    Dim pic As IPicture
    Dim dsc As PictDesc
    Dim IID(0 To 15) As Byte
    Dim Result As Long: Result = 0
    
    On Error GoTo CreateIcon_Error

    Set CreateIcon = Nothing
    If hImage <> 0 Then
        With dsc
           .cbSizeofStruct = Len(dsc)
           .hImage = hImage
           .PicType = VBRUN.PictureTypeConstants.vbPicTypeBitmap
        End With
        
        Result = OLE_CLSIDFromString(StrPtr(IID_IPicture), VarPtr(IID(0)))
                                                    
        If (Result = OLE_ERROR_CODES.S_OK) Then
            Result = Ole_CreatePic(dsc, VarPtr(IID(0)), True, pic)
            
            ' Creates a new picture object initialized according to a PICTDESC structure.
            If (Result = OLE_ERROR_CODES.S_OK) Then
                Set CreateIcon = pic
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

CreateIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateIcon of Module Module1"
End Function





'---------------------------------------------------------------------------------------
' Procedure : centrePreviewImage
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : place the image correctly within the preview pane
'---------------------------------------------------------------------------------------
' because the icon images are drawn from the top left of the
' preview pictureBox we have to manually set the picbox to size and position for each icon size
' this could be done with padding but it matches the VB6 method (no padding there)
Public Sub centrePreviewImage(ByRef targetPicBox As PictureBox, ByVal IconSize As Integer, ByVal ResizeRatio As Double)

    If targetPicBox.Name = "picPreview" Then
        If IconSize = 16 Then
            targetPicBox.Left = (1900 * ResizeRatio)
            targetPicBox.Top = (1900 * ResizeRatio)
            targetPicBox.Width = (200 * ResizeRatio)
            targetPicBox.Height = (200 * ResizeRatio)
        ElseIf IconSize = 32 Then
            targetPicBox.Left = (1800 * ResizeRatio)
            targetPicBox.Top = (1800 * ResizeRatio)
            targetPicBox.Width = (2000 * ResizeRatio)
            targetPicBox.Height = (2000 * ResizeRatio)
        ElseIf IconSize = 64 Then
            targetPicBox.Left = (1450 * ResizeRatio)
            targetPicBox.Top = (1450 * ResizeRatio)
            targetPicBox.Width = (2000 * ResizeRatio)
            targetPicBox.Height = (2000 * ResizeRatio)
        ElseIf IconSize = 128 Then
            targetPicBox.Left = (1000 * ResizeRatio)
            targetPicBox.Top = (1000 * ResizeRatio)
            targetPicBox.Width = (2000 * ResizeRatio)
            targetPicBox.Height = (2000 * ResizeRatio)
        ElseIf IconSize = 256 Then
            targetPicBox.Left = (100 * ResizeRatio)
            targetPicBox.Top = (100 * ResizeRatio)
            targetPicBox.Width = (4000 * ResizeRatio)
            targetPicBox.Height = (4000 * ResizeRatio)
        End If
    End If
End Sub






'---------------------------------------------------------------------------------------
' Procedure : GetEncoderClsid
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
' Built-in encoders for saving: (You can *try* to get other types also)
'   image/bmp
'   image/jpeg
'   image/gif
'   image/tiff
'   image/png
'
' Notes When Saving:
'The JPEG encoder supports the Transformation, Quality, LuminanceTable, and ChrominanceTable parameter categories.
'The TIFF encoder supports the Compression, ColorDepth, and SaveFlag parameter categories.
'The BMP, PNG, and GIF encoders no do not support additional parameters.
'
' Purpose:
'The function calls GetImageEncoders to get an array of ImageCodecInfo objects. If one of the
'ImageCodecInfo objects in that array represents the requested encoder, the function returns
'the index of the ImageCodecInfo object and copies the CLSID into the variable pointed to by
'pClsid. If the function fails, it returns –1.

'Private Function GetEncoderClsid(strMimeType As String, ClassID As IID) As Long
'   Dim num As Long
'   Dim Size As Long
'   Dim i As Long
'
'   Dim ICI() As ImageCodecInfo
'   Dim Buffer() As Byte
'
'   On Error GoTo GetEncoderClsid_Error
'
'   GetEncoderClsid = -1 'Failure flag
'
'   ' Get the encoder array size
'   Call GdipGetImageEncodersSize(num, Size)
'   If Size = 0 Then Exit Function ' Failed!
'
'   ' Allocate room for the arrays dynamically
'   ReDim ICI(1 To num) As ImageCodecInfo
'   ReDim Buffer(1 To Size) As Byte
'
'   ' Get the array and string data
'   Call GdipGetImageEncoders(num, Size, Buffer(1))
'   ' Copy the class headers
'   Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * num))
'
'   ' Loop through all the codecs
'   For i = 1 To num
'      ' Must convert the pointer into a usable string
'      If StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0 Then
'         ClassID = ICI(i).ClassID   ' Save the class id
'         GetEncoderClsid = i        ' return the index number for success
'         Exit For
'      End If
'   Next
'   ' Free the memory
'   Erase ICI
'   Erase Buffer
'
'   On Error GoTo 0
'   Exit Function
'
'GetEncoderClsid_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetEncoderClsid of Module mdlMain"
'End Function
'
'
'
''---------------------------------------------------------------------------------------
'' Procedure : PtrToStrW
'' Author    : From www.mvps.org/vbnet...i think
'' Date      : 21/08/2020
'' Purpose   : '   Dereferences an ANSI or Unicode string pointer
''   and returns a normal VB BSTR
''---------------------------------------------------------------------------------------
''
'Private Function PtrToStrW(ByVal lpsz As Long) As String
'    Dim sOut As String
'    Dim lLen As Long
'
'   On Error GoTo PtrToStrW_Error
'
'    lLen = lstrlenW(lpsz)
'
'    If (lLen > 0) Then
'        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
'        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
'        PtrToStrW = StrConv(sOut, vbFromUnicode)
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'PtrToStrW_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PtrToStrW of Module mdlMain"
'End Function
'
