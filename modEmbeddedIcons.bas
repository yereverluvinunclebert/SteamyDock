Attribute VB_Name = "modEmbeddedIcons"
Option Explicit

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
    
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

'---------------------------------------------------------------------------------------
' Procedure : displayEmbeddedIcons
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : The program extracts icons embedded within a DLL or an executable
'             you pass the name of the picbox you require and the image is displayed there
'             it should return all and not only the 16 and 32 bit icons as does extractIconEx
'             Also, on request, writes a PNG to a file on disc in the special folder area.
'---------------------------------------------------------------------------------------
'
Public Sub displayEmbeddedIcons(ByVal FileName As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer, ByVal writePNGToFile As Boolean)
    
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
    Dim successSaveToPNG As Boolean: successSaveToPNG = False

    On Error GoTo displayEmbeddedIcons_Error
    
    GSI.GdiplusVersion = 1
    GdiplusStartup lhToken, GSI

    On Error Resume Next ' debug

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
        MsgBox "Failed to extract icon."
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
                MsgBox "Failed to create bitmap from icon."
                GoTo CleanUp
            Else
                ' Creates a GDIP Graphics object (lhGraphics) that is associated with the current device context, that being the target picbox
                GdipCreateFromHDC .hDC, lhGraphics
                
                ' Draws an image at a specified location using the image bitmap and graphics object, in effect writing the image to the picbox
                '                      lhGraphics, lhImage, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&
                GdipDrawImageRectRectI lhGraphics, lhImage, 0, 0, IconSize, IconSize, 0, 0, 256, 256, 2&, 0, 0, 0
                
'               centre image using a better method
'                        ScaleX(x, ScaleMode, vbPixels) - WidthPx \ 2, _
'                        ScaleY(y, ScaleMode, vbPixels) - HeightPx \ 2, _
'                        IconSize, _
'                        IconSize, _

                ' In iconSettings we prove that it is possible to extract the PNG from extract the PNG from the DLL and write that to a file
                ' this is of little use here as we write to a picbox and display our PNG image there
                ' In SD, we will take this routine and use it to write a PNG to the local profile area and then insert the PNG into the dictionary at runtime startup.
                
                If writePNGToFile = True Then
                
                    ' take the filename, extract just the filename body minus the suffix, then point it to the special folder with a PNG suffix.
                    sJustTheFilename = Mid(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
                    sJustTheFilename = ExtractFilenameWithoutSuffix(sJustTheFilename)
                    sOutputFilename = SpecialFolder(SpecialFolder_AppData) & "\steamyDock\images\" & sJustTheFilename & ".png"
    
                    ' set the encoder class identifier to handle the image bitmap as a PNG
                    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatPNG
    '
                    ' extract a PNG of the image bitmap and save to file
                    successSaveToPNG = GdipSaveImageToFile(lhImage, StrPtr(sOutputFilename), ImageFormatPNG, ByVal 0&) = 0&
                    If successSaveToPNG = False Then
                        MsgBox "Failed to save PNG."
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
   Exit Sub

displayEmbeddedIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayEmbeddedIcons of Module mdlMain"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : make32BitLong
' Author    : beededea
' Date      : 20/11/2019
' Purpose   : packing variables into a 32bit LONG for an API call
'---------------------------------------------------------------------------------------
'
Private Function make32BitLong(ByVal LoWord As Integer, _
                 Optional ByVal HiWord As Integer = 0) As Long
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

