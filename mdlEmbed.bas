Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function OLE_CLSIDFromString Lib "ole32" Alias "CLSIDFromString" (ByVal lpszProgID As Long, ByVal pclsid As Long) As Long

Private Enum OLE_ERROR_CODES
    S_OK = 0
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
    E_FAIL = &H80004005
    E_UNEXPECTED = &H8000FFFF
    E_INVALIDARG = &H80070057
End Enum

Private Declare Function Ole_CreatePic Lib "olepro32" _
                Alias "OleCreatePictureIndirect" ( _
                ByRef lpPictDesc As PictDesc, _
                ByVal riid As Long, _
                ByVal fPictureOwnsHandle As Long, _
                ByRef iPic As IPicture) As Long
                
                
                
                
                

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CreateDIBSection2 Lib "gdi32.dll" Alias "CreateDIBSection" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Const DIB_RGB_COLORS As Long = 0
Private Const BI_RGB         As Long = 0&

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader    As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_ICON      As Long = 1
Private Const IMAGE_CURSOR    As Long = 2
Private Const LR_LOADFROMFILE As Long = &H10

Private Declare Function GdipCreateImageAttributes Lib "GdiPlus.dll" (ByRef mImageattr As Long) As GpStatus
Private Declare Function GdipSetImageAttributesColorMatrix Lib "GdiPlus.dll" (ByVal mImageattr As Long, ByVal mType As ColorAdjustType, ByVal mEnableFlag As Long, ByRef mGpColorMatrix As ColorMatrix, ByRef mGrayMatrix As ColorMatrix, ByVal mFlags As ColorMatrixFlags) As GpStatus
Private Declare Function GdipDisposeImageAttributes Lib "GdiPlus.dll" (ByVal mImageattr As Long) As GpStatus
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = &H0
    ColorAdjustTypeBitmap = &H1
    ColorAdjustTypeBrush = &H2
    ColorAdjustTypePen = &H3
    ColorAdjustTypeText = &H4
    ColorAdjustTypeCount = &H5
    ColorAdjustTypeAny = &H6
End Enum

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = &H0
    ColorMatrixFlagsSkipGrays = &H1
    ColorMatrixFlagsAltGray = &H2
End Enum

Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As GpStatus
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mGraphics As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal mImage As Long) As GpStatus
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mDstx As Long, ByVal mDsty As Long, ByVal mDstwidth As Long, ByVal mDstheight As Long, ByVal mSrcx As Long, ByVal mSrcy As Long, ByVal mSrcwidth As Long, ByVal mSrcheight As Long, ByVal mSrcUnit As GpUnit, ByVal mImageAttributes As Long, ByRef mcallback As Long, ByRef mcallbackData As Long) As GpStatus
Private Declare Function GdipGetImageDimension Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Single, ByRef mHeight As Single) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As GpStatus
Private Declare Function GdipCloneBitmapAreaI Lib "GdiPlus.dll" (ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mPixelFormat As Long, ByVal mSrcBitmap As Long, ByRef mDstBitmap As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As GpStatus
Private Const PixelFormat32bppARGB As Long = &H26200A
Private Enum GpStatus
    Ok = &H0
    GenericError = &H1
    InvalidParameter = &H2
    OutOfMemory = &H3
    ObjectBusy = &H4
    InsufficientBuffer = &H5
    NotImplemented = &H6
    Win32Error = &H7
    WrongState = &H8
    Aborted = &H9
    FileNotFound = &HA
    ValueOverflow = &HB
    AccessDenied = &HC
    UnknownImageFormat = &HD
    FontFamilyNotFound = &HE
    FontStyleNotFound = &HF
    NotTrueTypeFont = &H10
    UnsupportedlusVersion = &H11
    lusNotInitialized = &H12
    PropertyNotFound = &H13
    PropertyNotSupported = &H14
End Enum

Private Enum GpUnit
    UnitWorld = &H0
    UnitDisplay = &H1
    UnitPixel = &H2
    UnitPoint = &H3
    UnitInch = &H4
    UnitDocument = &H5
    UnitMillimeter = &H6
End Enum

Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As SmoothingMode) As GpStatus
Private Enum SmoothingMode
    SmoothingModeInvalid = &HFFFFFFFF
    SmoothingModeDefault = &H0
    SmoothingModeHighSpeed = &H1
    SmoothingModeHighQuality = &H2
    SmoothingModeNone = &H3
    SmoothingModeAntiAlias = &H4
End Enum

Private Declare Function GdipImageRotateFlip Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mRfType As RotateFlipType) As GpStatus
Private Enum RotateFlipType
    RotateNoneFlipNone = &H0
    Rotate90FlipNone = &H1
    Rotate180FlipNone = &H2
    Rotate270FlipNone = &H3
    RotateNoneFlipX = &H4
    Rotate90FlipX = &H5
    Rotate180FlipX = &H6
    Rotate270FlipX = &H7
    RotateNoneFlipY = &H6
    Rotate90FlipY = &H7
    Rotate180FlipY = &H4
    Rotate270FlipY = &H5
    RotateNoneFlipXY = &H2
    Rotate90FlipXY = &H3
    Rotate180FlipXY = &H0
    Rotate270FlipXY = &H1
End Enum

Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal mtoken As Long)
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (ByRef mtoken As Long, ByRef mInput As GdiplusStartupInput, ByRef mOutput As GdiplusStartupOutput) As GpStatus
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

'ICON FILE FORMAT (src: wikipedia)

'STRUCTURE-------------------------------------------------------------------------------------
'Icon Header      Stores general information about the ICO file.
'Directory[1..n]  Stores general information about every image in the file.
'Icon #1          The actual "data" for the first image in old AND/XOR DIB format or newer PNG
'...
'Icon #n          Data for the last icon image
'----------------------------------------------------------------------------------------------

'HEADER----------------------------------------------------------------------------------------
'Offset# Size  Purpose
'0       2     reserved. should always be 0
'2       2     type. 1 for icon (.ICO), 2 for cursor (.CUR) file
'4       2     count; number of images in the file
'----------------------------------------------------------------------------------------------

'DIRECTORY-------------------------------------------------------------------------------------
'Offset# Size  Purpose
'0       1     width, should be 0 if 256 pixels
'1       1     height, should be 0 if 256 pixels
'2       1     colour count, should be 0 if more than 256 colours
'3       1     reserved, should be 0[1]
'4       2     colour planes when in .ICO format, should be 0 or 1[2], or the X hotspot when in .CUR format
'6       2     bits per pixel when in .ICO format[3], or the Y hotspot when in .CUR format
'8       4     size in bytes of the bitmap data
'12      4     offset, bitmap data address in the file
'
' 1) A BITMAPINFOHEADER structure
' 2) An array of RGBQUAD structures (missing if the colour depth of the bitmap is > 8bpp)
' 3) A colour DIB containing the AND bitmap bits
' 4) A mono DIB containing the XOR bitmap bits
'-----------------------------------------------------------------------------------------------
'
'[1] Although Microsoft's technical documentation states that this value must be zero, the icon encoder built into .NET
'    (System.Drawing.Icon.Save) sets this value to 255. It appears that the operating system ignores this value altogether.
'[2] Setting the colour planes to 0 or 1 is treated equivalently by the operating system, but if the colour planes are
'    set higher than 1, this value should be multiplied by the bits per pixel to determine the final colour depth of the image.
'    It is unknown if the various Windows operating system versions are resilient to different colour plane values.
'[3] The bits per pixel might be set to zero, but can be inferred from the other data; specifically, if the bitmap is
'    not PNG compressed, then the bits per pixel can be calculated based on the length of the bitmap data relative to the
'    size of the image. If the bitmap is PNG compressed, the bits per pixel are stored within the PNG data. It is unknown if the various Windows operating system versions contain logic to infer the bit depth for all possibilities if this value is set to zero.
'

Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Long)

Private Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal mStream As Long, ByRef mImage As Long) As GpStatus
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const GMEM_MOVEABLE As Long = &H2

Private Type IconHeader
    nReserved As Integer
    nType     As Integer
    nImgCnt   As Integer
End Type
    
Private Type IconDirectory
    bWidth        As Byte
    bHeight       As Byte
    bColorCnt     As Byte
    bReserved     As Byte
    nColPlanes    As Integer
    nBPP          As Integer
    lBmpDataSize  As Long
    lBmpAddOffset As Long
End Type

Private m_udtIconHeader      As IconHeader
Private m_udtIconDirArr()    As IconDirectory
Private m_szIconDetailsArr() As String
Private m_szIconFile         As String
Private m_hTokenGDIP         As Long

Public Function getIconAsGDIPImage(ByVal nIndex As Integer, ByVal key As String, ByVal FileName As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer) As Long
    
    On Error GoTo EH
    Dim udtBIH As BITMAPINFOHEADER, lpDIBBits As Long, lWid As Long, lHei As Long, hDIB As Long, hDCMem As Long, bDIBBitsArr() As Byte
    Dim hFile As Integer, lBmpDataSize As Long, lFileOffset As Long, eRet As GpStatus, hImage As Long, hBmpPrev As Long
    Dim lRet As Long, hImageFlipped As Long, lCnt As Long, szBuff As String, udtIH As IconHeader, udtID As IconDirectory
    
    
    Dim sExeName As String: sExeName = vbNullString
    Dim lIconIndex As Long: lIconIndex = 0
    Dim xSize As Long: xSize = 0
    Dim ySize As Long: ySize = 0
    Dim hIcon() As Long: 'hIcon() = 0  cannot initialise
    Dim hIconID() As Long: 'hIconID() = 0  cannot initialise
    Dim nIcons As Long: nIcons = 0
    Dim Result As Long: Result = 0
    Dim flags As Long: flags = 0
    Dim i As Long: i = 0
    Dim pic As IPicture: 'pic cannot initialise
    Dim thiskey As String: thiskey = vbNullString
    Dim bytesFromFile() As Byte
    Dim Strm As stdole.IUnknown '  cannot initialise
    Dim img As Long: img = 0
    Dim dx As Long: dx = 0
    Dim dy As Long: dy = 0
    Dim strFilename As String: strFilename = vbNullString
    Dim opacity As String: opacity = vbNullString

    
    On Error Resume Next

    sExeName = FileName
    lIconIndex = 0
    
    ' Init ret value
    getIconAsGDIPImage = 0
    
    
    i = 2 ' need some experimentation here
    
    'the boundaries of the icons you wish to extract, can be set to somethink like 256, 256 but that is all
    ' you will extract, just the one icon
    xSize = make32BitLong(CInt("256"), CInt("16"))
    ySize = make32BitLong(CInt("256"), CInt("16"))
    
    flags = LR_LOADFROMFILE

    ' Get the total number of Icons in the file.
    Result = PrivateExtractIcons(sExeName, lIconIndex, xSize, ySize, ByVal 0&, ByVal 0&, 0&, 0&)
    
    ' The sExeName is the resource string/filepath.
    ' lIconIndex Index is the index.
    ' xSize and ySize are the desired sizes.
    ' phicon is the returned array of icon handles.
    ' So you could call it with phicon set to nothing and it will return the number of icons in the file.
    
    ' piconid ?
    
    ' nicons is just the number of icons you wish to extract.
    ' Then you call it again with nicon set to this number and niconindex=0. Then it will extract ALL icons in one go.
    ' flags
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
    ' eg. PrivateExtractIcons ('C:\Users\Public\Documents\RAD Studio\Projects\2010\Aero Colorizer\AeroColorizer.exe', 0, 128, 128, @hIcon, @nIconId, 1, LR_LOADFROMFILE)
    ' PrivateExtractIcons(sExeName, nIcon, cxIcon, cyIcon, phicon, piconid, nicons, 0)

    nIcons = 2 ' Result
    
    ' Dimension the arrays to the number of icons.
    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
    ReDim hIconID(lIconIndex To lIconIndex + nIcons * 2 - 1)

'  Rocketdock always uses the same ID for the same exe

'   C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE?5063424
'   E:\games\World_of_Tanks_NA\WorldOfTanks.exe?184608432

' if an exe is dragged and dropped onto RD it is given an id, it is appended to the binary name after an additional "?"
' that question mark signifies what? Possibly the handle of the embedded icon only added the first time,
' so that when the binary is read in the future the handle is already there?
' and that can be used to populate image array? Untested.
' in this case we just need to note the ? and then query the binary for an embedded icon and compare it to the id that RD has given it.
        
    ' use the undocumented PrivateExtractIcons to extract the icons we require
    Result = PrivateExtractIcons(sExeName, lIconIndex, xSize, _
                            ySize, hIcon(LBound(hIcon)), _
                            hIconID(LBound(hIconID)), _
                            nIcons * 2, flags)
        
    ' create an icon with a handle so we can test the result
    Set pic = CreateIcon(hIcon(i + lIconIndex - 1))
    

    
'    ' Make sure an Icon file has been parsed
'    If (m_szIconFile = "") Then
'        Call Err.Raise(vbObjectError + 1, "getIconAsGDIPImage()", "No icon file parsed yet")
'    End If
'
'    ' Validate Icon index
'    If ((nIndex < 0) Or (nIndex >= m_udtIconHeader.nImgCnt)) Then
'        Call Err.Raise(vbObjectError + 1, "getIconAsGDIPImage()", "Invalid index Must be >= 0 and < icon count")
'    End If
'
'    ' Acquire Icon props from previously parsed icon file
'    With m_udtIconDirArr(nIndex)
'        lWid = IIf((.bWidth = 0), 256, .bWidth)
'        lHei = IIf((.bHeight = 0), 256, .bHeight)
'        lBmpDataSize = .lBmpDataSize       ' Total bytes that make up the AND/XOR DIB
'        lFileOffset = (.lBmpAddOffset + 1) ' Location of the first byte of the bitmap data
'    End With
'
'    With udtBIH
'        .biBitCount = 32                   ' For Alpha channel support (0xAARRGGBB)
'        .biClrImportant = 0                ' All colors are important
'        .biClrUsed = 0                     ' Use the maximum number of colors corresponding to the biBitCount
'        .biCompression = BI_RGB            ' Uncompressed, raw RGB Pixels
'        .biHeight = lHei                   ' DIB height
'        .biPlanes = 1                      ' Always 1
'        .biSize = Len(udtBIH)              ' Size of the BITMAPINFOHEADER UDT
'        .biSizeImage = (lWid * lHei * 4)   ' Amount of bytes that make up the Bitmap
'        .biWidth = lWid                    ' DIB width
'        .biXPelsPerMeter = 0               ' n/a
'        .biYPelsPerMeter = 0               ' n/a
'    End With
'
'    ' Create a buffer the size of the bitmap data
'    ReDim bDIBBitsArr(lBmpDataSize - 1)
'    hFile = FreeFile()
'    Open m_szIconFile For Binary Access Read Lock Write As hFile
'        ' Read out the bitmap data into the bDIBBitsArr array
'        Get hFile, lFileOffset, bDIBBitsArr
'    Close hFile
'
'    ' Same file so header is the same except for nImgCnt, which is gonna
'    ' be 1 since we're extracting this 1 Icon as index 'nIndex'
'    udtIH = m_udtIconHeader
'    udtIH.nImgCnt = 1
'    udtID = m_udtIconDirArr(nIndex)
'    ' Set the bitmapDataOffset right after the IconDirectory structure, remember just 1 Icon
'    udtID.lBmpAddOffset = Len(udtIH) + Len(udtID)
'
'    ' Get temp file path
'    szBuff = String(260, Chr(0))
'    lCnt = GetTempPath(260, szBuff)
'    szBuff = Left(szBuff, lCnt)
'    szBuff = szBuff + IIf((Right(szBuff, 1) = "\"), "", "\") & CStr(Timer()) & ".tmp"
'
'    ' Write out this Icon's data into the temp file
'    hFile = FreeFile()
'    Open szBuff For Binary Access Read Write Lock Write As hFile
'        Put hFile, , udtIH       ' Header
'        Put hFile, , udtID       ' Directory
'        Put hFile, , bDIBBitsArr ' Bitmap data
'    Close hFile
'
'    ' Load the extracted icon
'    hIcon = LoadImage(App.hInstance, szBuff, IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
'    ' Remove Temp file
'    Call Kill(szBuff)
'    ' Check the handle to make sure its valid
'    If (hIcon = 0) Then
'        Call Err.Raise(vbObjectError + 1, "getIconAsGDIPImage()", "Unable to load requested icon")
'    End If
'
'    ' Check a memory DC compatible with the screen
    hDCMem = CreateCompatibleDC(0)
    ' Create a DIB compatible with the screen (hDCMem)
    hDIB = CreateDIBSection(hDCMem, udtBIH, DIB_RGB_COLORS, lpDIBBits, 0, 0)
    ' Sanity checks...
    If ((hDIB = 0) Or (lpDIBBits = 0)) Then
        'Call DestroyIcon(hIcon)
        Call DeleteDC(hDCMem)
        Call Err.Raise(vbObjectError + 1, "getIconAsGDIPImage()", "Unable to create DIB")
    End If

    ' Remove old Bitmap and insert the DIB into the memory DC (hDCMem) so
    ' we can draw the icon on it, and then access the data via the pointer to the DIB's bits
    hBmpPrev = SelectObject(hDCMem, hDIB)
    ' Draw the icon as normal
    'lRet = DrawIconEx(hDCMem, 0, 0, hIcon, lWid, lHei, 0, 0, DI_NORMAL)
    
    ' Draw the icon as normal
    lRet = DrawIconEx(hDCMem, 0, 0, hIcon(LBound(hIcon)), IconSize, IconSize, 0, 0, DI_NORMAL)  '
    
    ' get rid of the icons we created
    Call DestroyIcon(hIcon(i + lIconIndex - 1))
    ' Call DestroyIcon(hIcon(LBound(hIcon))
    
    ' Cleanup...
    'Call DestroyIcon(hIcon)
    Call SelectObject(hDCMem, hBmpPrev)
    Call DeleteDC(hDCMem)
    
    ' Create a 32-bit bitmap with pixelformat ARGB to "house" the 32-bit DIB
    ' (lWid * 4) is the offset of the beginning of one scan line with the next, usually with the formula:
    ' number of bytes in the pixel format (for example, 4 for 32-bits per pixel(Ex ARGB)) multiplied by the width of the bitmap
    eRet = GdipCreateBitmapFromScan0(lWid, lHei, (lWid * 4), PixelFormat32bppARGB, lpDIBBits, hImage)
    ' Sanity check...
    If (eRet <> Ok) Then
        Call DeleteObject(hDIB)
        Call Err.Raise(vbObjectError + 1, "getIconAsGDIPImage()", "Unable to create GDIP bitmap from DIB Bits pointer")
    End If
    
    ' Create another GDIP bitmap and clone the GDIP bitmap based on the DIB so we dont have to keep the
    ' DIB alive to have the bitmap remain valid. If you dont clone the bitmap, you will have to
    ' Dispose the Bitmap via GdipDisposeImage() then delete the DIB via DeleteObject()
    eRet = GdipCreateBitmapFromScan0(lWid, lHei, (lWid * 4), PixelFormat32bppARGB, 0, hImageFlipped)
    eRet = GdipCloneBitmapAreaI(0, 0, lWid, lHei, PixelFormat32bppARGB, hImage, hImageFlipped)
    ' DIB is stored bottom-to-top so filp it vertically (same with Render function of the StdPicture)
    eRet = GdipImageRotateFlip(hImageFlipped, RotateNoneFlipY)
    
    ' Cleanup....
    Call GdipDisposeImage(hImage)
    Call DeleteObject(hDIB)
    
    ' Done!
    getIconAsGDIPImage = hImageFlipped
    Exit Function
    
EH: Call MsgBox("Error converting icon:" & vbCrLf & Err.Description & ".", vbOKOnly Or vbExclamation, "getIconAsGDIPImage()")
    Call Err.Clear
    getIconAsGDIPImage = 0

End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateIcon
' Author    : beededea
' Date      : 14/07/2019
' Purpose   : This method creates an icon based on an image handle
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
           .PicType = VBRUN.PictureTypeConstants.vbPicTypeIcon
        End With
        
        Result = OLE_CLSIDFromString(StrPtr(IID_IPicture), VarPtr(IID(0)))
                                                    
        If (Result = OLE_ERROR_CODES.S_OK) Then
        
            ' Creates a new picture object initialized according to a PICTDESC structure.
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
' Procedure : make32BitLong
' Author    : beededea
' Date      : 20/11/2019
' Purpose   : packing variables into a 32bit LONG for an API call
'---------------------------------------------------------------------------------------
'
Private Function make32BitLong(ByVal LoWord As Integer, Optional ByVal HiWord As Integer = 0) As Long
   On Error GoTo make32BitLong_Error
   'If debugflg = 1 Then debugLog "%make32BitLong"

    make32BitLong = CLng(HiWord) * CLng(&H10000) + CLng(LoWord)

   On Error GoTo 0
   Exit Function

make32BitLong_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure make32BitLong of Module Module1"
End Function

