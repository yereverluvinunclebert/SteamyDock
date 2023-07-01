VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c32bppDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' COMPILE THE APPLICATION FOR BEST PERFORMANCE

' Credits/Acknowledgements - Thanx goes to:
'   Paul Caton for his class on calling non VB-Friendly DLLs that use _cdecl calling convention
'       http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=70195&lngWId=1
'   Carles P.V for his pvResize logic
'       Used when manually scaling images with NearestNeighbor or BiLinear interpolation
'   Carles P.V for his GIF LZW encoding routines. Used when saving to GIF
'       http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=45899&lngWId=1
'   Alfred Koppold for his PNG, VB-only, decompression routines. Used when zLib & GDI+ not available
'       http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56537&lngWId=1
'   John Korejwa for his JPG, VB-only, encoding routines. Used when saving as JPG and GDI+ not available
'       http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50065&lngWId=1
'   John Kleinen for example of a method of calling OLE class functions via API (See GetDroppedFileNames)
'       http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49268&lngWId=1
'   Steve McMahon for his example on 24bpp to 8bpp palette reduction using oct trees
'       http://www.vbaccelerator.com/home/VB/Code/vbMedia/Image_Processing/Colour_Depth_Reduction/VB6_Colour_Depth_Sample.asp
'   www.zlib.net for their free zLIB.dll, the standard DLL for compressing/decompressing PNGs
'       Without it, we'd be limited to GDI+ for creating PNGs
'   coders like you that provide constructive criticism to make this class better & more all-inclusive
'       Without your comments, this project probably would have died several versions/updates ago
' For most current updates/enhancements visit the following. Last changed 10 Oct 08
'   Visit http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
' To see a usercontrol applying a version of this class
'   AlphaImage Control. http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=68262&lngWId=1
'   Image List Control. http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=69621&lngWId=1

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                    O V E R V I E W
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' AlphaBlend API :: (msimg32.dll) , GDI+ API (gdiplus.dll)

' About 32bpp pre-multiplied RGB (pARGB) bitmaps, if you are not aware.
'   - These are used specifically for the AlphaBlend API & are GDI+ compatible
'   - Only TGA directly identifies premultiplied pixels. Though they can exist in
'       PNG & BMP also, there is no flag saying: "Hey, I have premultiplied pixels"
'   - The following 4 formats can contain alpha channels: PNG, TGA, BMP, ICO
'   Advantages:
'       - Images can be per-pixel alpha blended
'       - Overall image opacity can be simultaneously adjusted during rendering
'       - AlphaBlend does both BitBlt & StretchBlt for pARGB images.
'       - Speed: AlphaBlend & GDI+ are pretty quick APIs vs manual blending
'   Disadvantages:
'       - The original RGB values are permanently destroyed during pre-multiplying
'           -- Premultiplied formula: preMultipliedRed=(OriginalRed * Alpha) \ 255
'           -- There is no way to convert pARGB back to non-premultiplied RGB values
'              The formula would be: reconstructedRed=(preMultipliedRed * 255) \ Alpha.
'               but because of integer division when pre-multiplying, the result is very
'               close and if this should be premultiplied again & converted again, the
'               calculated colors can eventually be completely different than the original.
'               Fully opaque pixels pixels are not affected. Any transparent pixel is
'               permanently destroyed & can never be returned. Its value becomes vbBlack
'           ** Note: When images are converted to other image formats for saving,
'               removal of premultiplication is performed to meet their specs.
'       - Displaying a pre-multiplied bitmap without AlphaBlend/GDI+ will not result in
'           the image being displayed as expected.
'       - Not ideal for saving due to its size: SizeOf= W x H x 4
'           -- better to save source image instead or compress the DIB bytes using favorite compression utility
'           -- with GDI+ or zLib, image can be converted to PNG for excellent reduced storage
'           -- Saving as a compressed TGA format generally provides very good file sizes
'       - AlphaBlend API is not included/compatible with Win95, NT4 and lower
'       - AlphaBlend on Win9x systems can be buggy, especially when rendering to DIBs vs DDBs

' Note that GDI+ is standard on WinXP+, and can be used on Win98,ME,2K & on NT4 if SP6 is installed
'   :: side note. I have actually used GDI+ on Win95 without problems; have not tried on NT4 SP5 or lower
' Download GDI+ from:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdicpp/GDIPlus/GDIPlus.asp
' Download ZLib from: http://www.zlib.net

' ----------------------------------------------
' About Win95, Win98, WinME, NT3.5 & NT4 support
' ----------------------------------------------
' These routines may not honor AlphaBlend if it exists on those systems. Win98's version,
' for example, has several bugs that can crash the application when AlphaBlending to DIBs.
' NT4, NT3.5 & Win95 do not come with AlphaBlend and I do not have WinME to test with.
' Therefore, to support these older systems, the Render routine will alphablend manually
' regardless if the AlhpaBlend API exists on the older system or not. However, this can
' be overridden by you. See isAlphaBlendFriendly routine. Therefore, AlphaBlend is only
' reliable on Win2K and above. XP & above already have GDI+

' Class Purpose:
' ----------------------------------------------
' This class holds the 32bpp image. It also marshals any new image thru
' the battery of parsers to determine best method for converting the image
' to a 32bpp alpha-compatible image. It handles rendering, rotating, scaling,
' mirroring of DIBs using manual processes, AlphaBlend, and/or GDI+.

' There are several ways an image can be loaded
' 1. From filename (unicode supported) :: LoadPicture_File
' 2. From a byte array :: LoadPicture_Stream
' 3. From stdPicture object or VB's .Picture object :: LoadPicture_StdPicture
' 4. From VB's Resource file :: LoadPicture_Resource
' 5. From a memory image handle or VB's .Picture.Handle property :: LoadPicture_ByHandle
' 6. From the clipboard :: LoadPicture_ClipBoard
' 7. From file drag/drop's Data object (unicode supported) :: LoadPicture_DropedFiles
' 8. From Pasted files (unicode supported) :: LoadPicture_PastedFiles

' The parser order is very important for fastest/best results. See
' spt_LoadPictureEx how the routine marshals the image thru the parsers.
' cPNGparser :: will convert PNG, all bit depths; aborts quickly if not PNG
' cGIFparser :: will convert non-transparent/transparent GIFs; aborts quickly
' cICOpraser :: will convert XP-Alpha, paletted, true color, & Vista PNG icons
'               -- can also convert most non-animated cursors
' cBMPparser :: will convert bitmaps, wmf/emf & jpgs
' cTGAparser :: will convert TGA (Targa, TrueVision) images only
' As a last resort, when GDI+ exists, anything unable to be processed by the
' parsers (i.e., TIFFs) are sent to GDI+. If GDI+ can process the image, then
' the image will be converted, internally, to PNG to enable additional processing.

' The parsers are efficient. Most image formats have a magic number that give
'   a hint to what type of image the file/stream is. However, checks need to
'   be employed because non-image files could feasibly have those same magic
'   numbers. If the image is determined not to be one the parser is designed to
'   handle, the parser rejects it and the next parser takes over. Rejection occurs,
'   generally, in only a few lines of code. The icon parser is slightly different
'   because PNG files can be included into a Vista ico file. When this occurs, the
'   icon parser will pass off the PNG format to the PNG parser automatically.
' And last but not least, the parsers have no advanced knowledge of the image
' format; as far as they are concerned, anything passed is just a byte array.
' Relying on file extensions when available is unreliable

' Class Organization:
' ----------------------------------------------
' Search the class for the words NEW SECTION
' The class routines are organized in the following sections:
'   Class Initialization & Termination Routines
'   Public Properties & Methods (almost 60 and growing)
'       Public Read-Only Properties
'       Public Methods
'       Class to Class Communication Methods
'   Local Support Functions
' ----------------------------------------------

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                       CHANGE HISTORY
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' Accompanying Usage_c32bppSuite.rtf is updated with every change
' Last changed: 16 Feb 09. See change history within the RTF file
'   -- most noticable changes
'   :: fixed icon parsing class which can misparse icon masks
'   :: fixed cPNGreader class which could misparse grayscale png transparency
' 26 Dec 06: First version
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

'._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._
'   IMPORTANT FOR UPDATES/ENHANCEMENTS   IMPORTANT FOR UPDATES/ENHANCEMENTS
' -----------------------------------------------------------------------------
' NOTE TO SELF: Whenever any routine recreates the DIB (new DIB memory handle),
' a call to Set m_GDIplus = Nothing must be made before changes are applied. This is because
' when the optional KeepGDIplusActive property=True, GDI+ is wrapped around the DIB.
' Deleting the DIB and not releasing GDI+ can cause crashes or unexpected results.


' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.
' Note: I did take liberties, changing parameter types, in several APIs throughout

' Used to determine operating system
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As Any) As Long
Private Type OSVERSIONINFOEX
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128 ' up to here is OSVERSIONINFO vs EX
   wServicePackMajor As Integer ' 8 bytes larger than OSVERSIONINFO
   wServicePackMinor As Integer
   wSuiteMask As Integer
   wProductType As Byte
   wReserved As Byte
End Type

' APIs used to manage the 32bpp DIB
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef Pointer As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long



Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef BITS As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
'@Ignore ProcedureNotUsed
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const STRETCH_HALFTONE As Long = &H4&

' APIs used for fonts/text








'@Ignore ProcedureNotUsed
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type Size
    Cx As Long
    Cy As Long
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
    lfFaceName As String * 32
End Type
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Public Enum eTextAlignment
    TA_CENTER = 6
    TA_LEFT = 0
    TA_RIGHT = 2
    TA_RTLREADING = 256
End Enum
Public Enum AlphaTypeEnum
    AlphaNone = 0
    AlphaSimple = 1
    AlphaComplex = 2
End Enum

' APIs used to create files
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1&

' ////////////////////////////////////////////////////////////////
' Unicode-capable Drag and Drop of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, _
    ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, _
    ByVal paCNT As Long, ByRef paTypes As Integer, _
    ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByRef lpString As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

' ////////////////////////////////////////////////////////////////
' Unicode-capable Pasting of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lIndex As Long
    TYMED As Long
End Type
Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type
Private Type STGMEDIUM
    TYMED As Long
    data As Long
    pUnkForRelease As IUnknown
End Type
' ////////////////////////////////////////////////////////////////


' used to create the checkerboard pattern on demand
'@Ignore ProcedureNotUsed
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 3) As Long ' reusable UDT for 1 & 2 dim arrays
    ' array items: 0=1D item count, 1=1D LBound, 2=2D item count, 3=2D LBound
End Type

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette As Long
End Type

Private Enum eOScapability
    osAlphaBlendUsable = 1  ' then AlphaBlend enabled & used when needed
    osGDIplusUsable = 2     ' then GDI+ enabled & used when needed (set in isGDIplusEnabled)
   'oszLIBusable = 4        ' then zLib enabled & can be used to create/read PNGs (no longer used, tested as needed)
    osWin2KorBetter = 8     ' AlphaBlend capable system else it isn't for these classes. See isAlphaBlendFriendly property for more info
    osWin98MEonly = 16      ' Win98 or WinME. When m_OScap includes osWin98MEonly+osAlphaBlendUsable then user overrode isAlphaBlendFriendly
    osGDIplusNotAvail = 32  ' then NT4 w/less than SP6 or Win95. Otherwise system is GDI+ capable else it isn't
    osIsNT = 64             ' NT-based system. Unicode compatible
End Enum

Public Enum eImageFormat    ' source image format
    imgError = -1  ' no DIB has been initialized
    imgNone = 0    ' no image loaded
    imgBitmap = 1  ' standard bitmap or jpg
    imgIcon = 3    ' standard icon
    imgWMF = 2     ' windows meta file
    imgEMF = 4     ' enhanced WMF
    imgCursor = 5  ' standard cursor
    imgBmpARGB = 6  ' 32bpp bitmap where RGB is not pre-multiplied
    imgBmpPARGB = 7 ' 32bpp bitmap where RGB is pre-multiplied
    imgIconARGB = 8 ' XP-type icon; 32bpp ARGB
    imgGIF = 9      ' gif; if class.Alpha=AlphaSimple, then transparent GIF
    imgPNG = 10     ' PNG image
    imgPNGicon = 11 ' PNG in icon file (Vista)
    imgCursorARGB = 12 ' alpha blended cursors? do they exist yet?
    imgCheckerBoard = 64 ' image is displaying own checkerboard pattern; not a true image
    imgTGA = 128    ' if alpha<>AlphaNone, then transparent TGA
End Enum

Public Enum ePngProperties      ' following are recognized "Captions" within a PNG file
    pngProp_Title = 1           ' See cPNGwriter.SetPngProperty for more information
    pngProp_Author = 2
    pngProp_Description = 4
    pngProp_Copyright = 8
    pngProp_CreationTime = 16
    pngProp_Software = 32
    pngProp_Disclaimer = 64
    pngProp_Warning = 128
    pngProp_Source = 256
    pngProp_Comment = 512
    ' special properties
    pngProp_Miscellaneous = 1024 ' this is free-form text can be of any length & contain most any characters
    pngProp_DateTimeModified = 2048  ' date/time of the last image modification (not the time of initial image creation)
    pngProp_DefaultBkgColor = 4096   ' default background color to use if PNG viewer does not do transparency
    pngProp_FilterMethod = 8192        ' one of the eFilterMethods values
    pngProp_ClearProps = -1     ' resets all PNG properties
End Enum

Public Enum eTrimOptions    ' see TrimImage method
    trimAll = 0             ' can be combined using OR
    trimLeft = 1
    trimTop = 2
    trimRight = 4
    trimBottom = 8
End Enum

Public Enum eScaleOptions   ' See ScaleImage method
    ScaleToSize = 0         ' [Default] will always scale
    scaleDownAsNeeded = 1   ' will only scale down if image won't fit
    ScaleStretch = 2        ' wll always stretch/distort
End Enum

Public Enum eGrayScaleFormulas
    gsclNTSCPAL = 0     ' R=R*.299, G=G*.587, B=B*.114 - Default
    gsclCCIR709 = 1     ' R=R*.213, G=G*.715, B=B*.072
    gsclSimpleAvg = 2   ' R,G, and B = (R+G+B)/3
    gsclRedMask = 3     ' uses mostly the Red sample value: RGB = .8Green,.1Red,.1Red
    gsclGreenMask = 4   ' uses mostly the Green sample value: RGB = .1Red,.8Green,.1Red
    gsclBlueMask = 5    ' uses mostly the Blue sample value: RGB = .1Red,.1Red,.8Blue
    gsclRedGreenMask = 6 ' uses Red & Green sample value: RGB = .45Red,.45Green,.1Blue
    gsclBlueGreenMask = 7 ' uses Blue & Green sample value: RGB = .1Red,.45Green,.45Blue
    gsclNone = -1
End Enum

Public Enum eFilterMethods
    filterDefault = 0     ' paletted PNGs will use filterNone while others will use filterPaeth
    filterNone = 1        ' no byte preparation used; else preps bytes using one of the following
    filterAdjLeft = 2     ' see cPNGwriter.EncodeFilter_Sub
    filterAdjTop = 3      ' see cPNGwriter.EncodeFilter_Up
    filterAdjAvg = 4      ' see cPNGwriter.EncodeFilter_Avg
    filterPaeth = 5       ' see cPNGwriter.EncodeFilter_Paeth
    filterAdaptive = 6    ' this is a best guess of the above 4 (can be different for each DIB scanline)
End Enum

Public Enum eRegionStyles     ' See CreateRegion
    regionBounds = 0
    regionEnclosed = 1
    regionShaped = 2
End Enum

Public Enum eConstants      ' See SourceIconSizes
    HIGH_COLOR = &HFFFF00
    TRUE_COLOR = &HFF000000
    TRUE_COLOR_ALPHA = &HFFFFFFFF
End Enum

Private m_Tag As Variant            ' user-defined, user usage only. See TAG property
'Private m_PNGprops As cPNGwriter    ' used for more advanced PNG creation options
Private m_StretchQuality As Boolean ' if true will use BiLinear or better interpolation
Private m_Handle As Long        ' handle to 32bpp DIB
Private m_Pointer As Long       ' pointer to DIB bits
Private m_Height As Long        ' height of DIB
Private m_Width As Long         ' width of DIB
Private m_hDC As Long           ' DC if self-managing one
Private m_prevDCobject As Long  ' object deselected from DC when needed
Private m_osCAP As eOScapability ' See eOScapability enumeration above
Private m_Format As eImageFormat ' type of source image
Private m_ManageDC As Boolean   ' does class manage its own DC?
Private m_AlphaImage As AlphaTypeEnum ' does the DIB contain alpha/transparency?
Private m_GDItoken As Long      ' GDI+ token when GDI+ is initialized
Private m_ImageByteCache() As Byte  ' should you want the DIB class to cache original bytes
' ^^ N/A if image is loaded by handle, stdPicture, or resource

Private m_GDIplus As cGDIPlus   ' maintains GDI+ instance to speed up drawing
Private m_KeepGDIplusActive As Boolean ' see KeepGDIplusActive property

' NEW SECTION *******************************************************************************
'                   CLASS INITIALIZATION & TERMINATION ROUTINES
' *******************************************************************************************

Private Sub Class_Initialize()

    ' Determine operating system for compatibility of 32bpp images
    ' http://vbnet.mvps.org/code/helpers/iswinversion.htm
    ' http://msdn2.microsoft.com/en-gb/library/ms724834.aspx
    
   Dim osType As OSVERSIONINFOEX
   Const VER_PLATFORM_WIN32_WINDOWS As Long = 1

   ' Retrieve version data for OS.
   osType.dwOSVersionInfoSize = Len(osType)
   If GetVersionEx(osType) = 0 Then
      ' The OSVERSIONINFOEX structure is only supported
      ' in NT4/SP6+ and NT5.x, so we're likely running
      ' on an earlier version of Windows. Revert structure
      ' size to OSVERSIONINFO and try again.
      osType.dwOSVersionInfoSize = Len(osType) - 8
      Call GetVersionEx(osType)
   End If
   
    If osType.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then

        If osType.dwMinorVersion = 0 Then ' Win95; can't use AlphaBlend nor GDI+
            m_osCAP = osGDIplusNotAvail   ' fyi. I copied gdi+ onto win95 & it worked! but is it trustworthy

        Else ' flag as Alphablend disabled, but capable & is Win98/ME
            m_osCAP = osWin98MEonly
        End If

    Else

        If osType.dwMajorVersion > 4 Then ' if Win2K or better
            m_osCAP = osAlphaBlendUsable Or osWin2KorBetter ' flag as AlphaBlend enabled (Win2K or better) and capable

        Else ' WinNT4. If SP6 or better than GDI+ capable else not. Regardless, not AlphaBlend capable
            If osType.wServicePackMajor < 6 Then m_osCAP = osGDIplusNotAvail
        End If
        m_osCAP = m_osCAP Or osIsNT

    End If

    ' Note for programmers:  To test the spt_Win9xBlend function on systems with GDI+
    ' 1. Unrem next line
    '   m_osCAP = osGDIplusNotAvail
    ' 2. Rem out the next two lines
    Me.isGDIplusEnabled = True  ' attempt to start GDI+, test system capability, add osGDIplusUsable to m_OScap
    If Me.isGDIplusEnabled Then Me.HighQualityInterpolation = True
    
    m_Format = imgError

End Sub

Private Sub Class_Terminate()
    DestroyDIB ' simply clean up
End Sub

' NEW SECTION *******************************************************************************
'                           PUBLIC PROPERTIES AND METHODS
' *******************************************************************************************


Public Property Let Alpha(ByVal AlphaType As AlphaTypeEnum)
    m_AlphaImage = AlphaType  ' determines the flags used for AlphaBlend API
    ' this flag is set by the various image parsers; setting it yourself
    ' can produce less than desirable effects.
End Property
Public Property Get Alpha() As AlphaTypeEnum
    Alpha = m_AlphaImage
End Property

Public Property Let HighQualityInterpolation(ByVal Value As Boolean)
    ' When possible GDI+ will be used for stretching & rotation.
    ' If GDI+ is used,then high quality equates to BiCubic interpolation
    ' If not used, then BiLinear (manual processing) will be used.
    ' If High Quality is false, then Nearest Neighbor (very fast, less quality) interpolation used
    m_StretchQuality = Value
End Property
Public Property Get HighQualityInterpolation() As Boolean
    HighQualityInterpolation = m_StretchQuality
End Property

Public Property Get ImageType() As eImageFormat
    ImageType = m_Format    ' returns image format of the source image
End Property
Friend Property Let ImageType(ByRef iType As eImageFormat)
    m_Format = iType    ' set by the various image parsers. This is not used in decision making
    ' anywhere in these classes, you can do with it what you want -- for now.
End Property

Public Property Let ManageOwnDC(ByVal bManage As Boolean)
    ' Determines whether or not this class will manage its own DC
    ' If false, then a DC is created each time the image needs to be Rendered
    Dim tDC As Long
    If bManage = False Then     ' removing management of DC
        If Not m_hDC = 0& Then   ' DC does exist, destroy it
            ' first remove the dib, if one exists
            If Not m_Handle = 0& Then SelectObject m_hDC, m_prevDCobject
            m_prevDCobject = 0&
        End If
        DeleteDC m_hDC
        m_hDC = 0&
    Else                        ' allowing creation of dc
        If m_hDC = 0& Then        ' create DC only if we have a dib to put in it
            If Not m_Handle = 0& Then
                tDC = GetDC(0&)
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
        End If
    End If
    m_ManageDC = bManage
End Property
Public Property Get ManageOwnDC() As Boolean
    ManageOwnDC = m_ManageDC
End Property

Public Property Get isAlphaBlendFriendly() As Boolean
    isAlphaBlendFriendly = ((m_osCAP And osAlphaBlendUsable) = osAlphaBlendUsable)
    ' WinNT4 & below and Win95 are not shipped with msimg32.dll (AlphaBlend API)
    ' Win98 has bugs & would believe that WinME is buggy too but don't know for sure
    ' Therefore, the Rendering in this class will not use AlphaBlend on these
    ' operating systems even if the DLL exists, but will use GDI+ if available
    ' Can be overridden by setting this property to True
End Property
Public Property Let isAlphaBlendFriendly(ByVal Enabled As Boolean)
    ' This has been provided to override safety of using AlphaBlend on Win9x systems.
    ' Caution. Only set this when rendering to a known device dependent bitmap (DDB)
    ' Alphablend can crash when rendering DIB to DIB vs DIB to DDB. Be warned.
    If Enabled = True Then
        ' Overriding in play: allow AlphaBlend if system is Win98 or better
        ' By default this is already set for Win2K or better
        If ((m_osCAP And osWin2KorBetter) = osWin2KorBetter) Then
            m_osCAP = m_osCAP Or osAlphaBlendUsable
        ElseIf ((m_osCAP And osWin98MEonly) = osWin98MEonly) Then
            m_osCAP = m_osCAP Or osAlphaBlendUsable
        End If
    Else
        m_osCAP = m_osCAP And Not osAlphaBlendUsable ' disallow AlphaBlend
    End If
End Property
Public Property Get isGDIplusEnabled() As Boolean
    ' identifies if GDI+ is usable on the system.
    ' Before this property is set, GDI+ is tested to ensure it is usable
    isGDIplusEnabled = ((m_osCAP And osGDIplusUsable) = osGDIplusUsable)
End Property
Public Property Let isGDIplusEnabled(ByVal Enabled As Boolean)
    ' Sets the property. If set to False by you, GDI+ will not be used
    ' for any rendering, but still may be used to create PNG/JPG files if needed
    
    ' You can reset it to true at any time. If the system won't support
    ' GDI+, then the True setting will simply be ignored -- no harm, no foul
    ' To test success:  c32class.isGDIplusEnabled=True: If c32class.isGDIplusEnabled=True Then ' success
    
    If Not Enabled = Me.isGDIplusEnabled Then
        m_osCAP = (m_osCAP And Not osGDIplusUsable)
        If Enabled Then
            If (m_osCAP And osGDIplusNotAvail) = 0 Then ' else Win95, NT4 SP5 or lower
                Dim cGDIp As cGDIPlus
                If m_GDIplus Is Nothing Then
                    Set cGDIp = New cGDIPlus
                Else
                    Set cGDIp = m_GDIplus
                End If
                If cGDIp.isGDIplusOk() = True Then m_osCAP = m_osCAP Or osGDIplusUsable
            End If
        End If
    End If
End Property

Public Property Let gdiToken(ByVal Token As Long)
    ' Everytime a GDI+ API function is called, the class calls GDI+ apis to
    ' create a GDI+ token first then destroys the token after the function is called.
    
    ' This occurs quite often. However, you can create your own token by calling
    ' GdiplusStartup and then passing the token to each class for the class to use.
    ' You would call GdiplusShutdown during your main form's Terminate event to
    ' release the Token.
    
    ' When Token is zero, the classes will revert to creating a token on demand.
    ' When the Token is not zero, any other DIB class created internally by this class will
    ' pass the token as needed. The only routine that can create a new instance externally
    ' and returns that new instance is the CreateDropShadow method. You must set the
    ' token for that class at some point for that dropshadow class to use the token.
    
    If m_KeepGDIplusActive Then Set m_GDIplus = Nothing ' remove; may be referencing old token
    m_GDItoken = Token
    
End Property
Public Property Get gdiToken() As Long
    ' returns the GDI+ token if one was created
    gdiToken = m_GDItoken
End Property

Public Property Let Tag(ByRef vValue As Variant)
    On Error Resume Next
    If IsObject(vValue) Then Set m_Tag = vValue Else m_Tag = vValue
End Property
Public Property Set Tag(ByVal vValue As Variant)
    Me.Tag = vValue ' use the object check in the Let property
End Property
Public Property Get Tag() As Variant
    If IsObject(m_Tag) Then Set Tag = m_Tag Else Tag = m_Tag
End Property

' This setting will keep the current cGDI+ class active which
' prevents destroying the GDI+ hImage and Token objects.  This
' can speed up rendering. This option SHOULD NOT be set when the
' class is not compiled. Else crashes will happen when
' user hits END or VB toolbar's STOP button while in IDE
Public Property Get KeepGDIplusActive() As Boolean
    KeepGDIplusActive = m_KeepGDIplusActive
End Property
Public Property Let KeepGDIplusActive(ByVal keepActive As Boolean)
    m_KeepGDIplusActive = keepActive
    If keepActive = False Then
        Set m_GDIplus = Nothing
    ElseIf Me.isGDIplusEnabled = False Then
        ' can't set to True if GDI+ not installed
        m_KeepGDIplusActive = False
    End If
End Property


' NEW SECTION *******************************************************************************
'                               PUBLIC READ-ONLY PROPERTIES
' *******************************************************************************************

Public Property Get Width() As Long
    Width = m_Width     ' width of image in pixels
End Property
Public Property Get Height() As Long
    Height = m_Height   ' height of image in pixels
End Property
Public Property Get BitsPointer() As Long
    BitsPointer = m_Pointer ' pointer to the bits of the image
End Property
Public Property Get scanWidth() As Long
    scanWidth = m_Width * 4&    ' number of bytes per scan line
End Property
Public Property Get Handle() As Long
    Handle = m_Handle   ' the picture handle of the image
End Property
Public Property Get isZlibEnabled() As Boolean
    ' To create PNG files, GDI+ or zLib is required. This property informs
    ' you if zLIB exists in the system's DLL path
    If iparseValidateDLL("zlib1.dll", "crc32") = False Then
        isZlibEnabled = iparseValidateDLL("zlib.dll", "crc32")
    Else
        isZlibEnabled = True
    End If
    
End Property


' NEW SECTION *******************************************************************************
'                                       PUBLIC METHODS
' *******************************************************************************************

Public Function LoadPictureFile(ByVal Filename As String, _
                                Optional ByVal IconCx As Long = 32, _
                                Optional ByVal IconCy As Long = 32, _
                                Optional ByVal SaveFormat As Boolean = False, _
                                Optional ByVal iconBitDepth As Long = 32) As Boolean

    ' PURPOSE: Convert passed image file into a 32bpp image
    
    ' Parameters.
    ' FileName :: full path of file. Validation occurs before we continue
    ' iconCx :: desired width of icon if file is an icon file. Default is 32
    ' iconCy :: desired height of icon if file is an icon file. Default is 32
    ' ^^ the end parameters: 256,256 is just telling the class that
                ' we want that size icon if one exists in the passed resource. If not,
                ' then give us the one closest to it & best quality too.
    
    ' SaveFormat :: if true, then the image will be cached as a byte array only
    '   if the image was successfully loaded. Call GetOrginalFormat to retrieve them.
    ' iconBitDepth :: the desired bit depth of an icon if the resource is an icon file. Default is 32
    
    ' Why would you want to cache the bytes? If this is being used in a usercontrol, saving
    ' the bytes in WriteProperty will almost always be less size than saving the 32bit DIB.
    ' Additionally, these classes have the ability to get different sizes from
    ' the original source (i.e., WMF, icon, cursors) if available, but if the
    ' 32bit DIB is saved, it is a constant size. The potential of different sizes
    ' could allow better resizing of the image vs stretching the DIB.

    On Error Resume Next
    Dim hFile As Long
    
    hFile = iparseGetFileHandle(Filename, True, ((m_osCAP And osIsNT) = osIsNT))
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    
    If GetFileSize(hFile, 0&) > 56 Then
        Dim aDIB() As Byte  ' dummy array
        LoadPictureFile = spt_LoadPictureEx(hFile, Filename, aDIB(), IconCx, IconCy, 0&, 0&, SaveFormat, iconBitDepth)

    End If
    CloseHandle hFile
    
End Function





Public Function LoadPictureStdPicture(ByRef Picture As StdPicture) As Boolean

    ' PURPOSE: Convert passed stdPicture into a 32bpp image
    ' Revised to allow 32bpp stdPicture objects which can be loaded via VB's LoadPicture
    
    Me.DestroyDIB
    If Not Picture Is Nothing Then
        ' simply pass off to other parsers
        If Picture.Type = vbPicTypeIcon Then
            ' pass to icon/cursor parser
            Dim cICO As New cICOparser
            Call cICO.ConvertstdPicTo32bpp(Picture.Handle, Me)
            Set cICO = Nothing
        ElseIf Not Picture.Type = vbPicTypeNone Then
            ' pass to bmp,jpg,wmf parser
            ' Note: transparent GIFs should not be passed as stdPictures
            '   Pass transparent GIFs by Stream or FileName
            Dim cBmp As New cBMPparser
            If Picture.Type = vbPicTypeBitmap Then
                ' pass by handle to ensure 32bpp stdPicture objects are processed correctly
                Call cBmp.ConvertstdPicTo32bpp(Nothing, Picture.Handle, Me, 0&)
            Else ' probably wmf/emf
                Call cBmp.ConvertstdPicTo32bpp(Picture, 0&, Me, 0&)
            End If
            Set cBmp = Nothing
        End If
        LoadPictureStdPicture = Not (m_Handle = 0&)
    End If
    
End Function



'Public Function LoadPictureClipBoard() As Boolean
'
'    ' PURPOSE: Convert clipboard object into a 32bpp image
'
'    On Error Resume Next
'    With Clipboard
'        If (.GetFormat(vbCFBitmap) Or .GetFormat(vbCFDIB) Or .GetFormat(vbCFEMetafile) Or .GetFormat(vbCFMetafile)) Then
'            If Not Err Then LoadPictureClipBoard = LoadPictureStdPicture(.GetData())
'        End If
'    End With
'    If Err Then Err.Clear
'End Function

'Public Function LoadPictureDroppedFiles(ByRef DragDrop_DataObject As DataObject, _
'                                        Optional ByVal FileIndex As Long = 1&, _
'                                        Optional ByVal IconCx As Long = 32, _
'                                        Optional ByVal IconCy As Long = 32, _
'                                        Optional ByVal SaveFormat As Boolean = False, _
'                                        Optional ByVal iconBitDepth As Long = 32) As Boolean
'
'    ' How to load just the 1st file if 1 or more were dropped
'    '   myC32bppDIB.LoadPicture_DroppedFiles Data, 1
'
'    ' How to load multiple files into multiple classes
'    ' Example assumes your classes are in a collection called colImages
'    '   Dim tmpClass As c32bppDIB
'    '   For X = 1 To Data.Files.Count
'    '       Set tmpClass=New c32bppDIB
'    '       If tmpClass.LoadPicture_DroppedFiles(Data, X) Then colImages.Add tmpClass, [key]
'    '   Next
'
'    ' Parameters
'    ' DragDrop_DataObject :: the Data parameter in your form/control's OLEDragDrop event
'    ' FileIndex :: the one-bound file that is to be loaded if more than one file were dropped
'    ' For the oher parameter descriptions, see LoadPicture_File
'
'    ' Note: By calling GetDroppedFileNames yourself, you have more flexibility
'
''    If GetDroppedFileNames(DragDrop_DataObject) Then
''        If FileIndex <= DragDrop_DataObject.Files.count Then
''            LoadPictureDroppedFiles = LoadPictureFile(DragDrop_DataObject.Files(FileIndex), IconCx, IconCy, SaveFormat, iconBitDepth)
''            'LoadPicture_DropedFiles = LoadPicture_File("E:\dean\steampunk theme\icons\icons MKV\SSL-certificate-folder.ico", 256, 256, False, 32)
''        End If
''    End If
'
'End Function





Public Function LoadDIBinDC(ByVal bLoad As Boolean) As Long

    ' Purpose: Select/Unselect the DIB into a DC.
    ' Returns the DC handle when image is loaded
    ' Called by image parser if it needs to paint the image into the DIB
       
    If bLoad = True Then
        Dim tDC As Long
        If Not m_Handle = 0& Then    ' do we have an image?
            If m_hDC = 0& Then        ' do we have a DC?
                tDC = GetDC(0&)     ' if not create one
                m_hDC = CreateCompatibleDC(tDC)
                ReleaseDC 0&, tDC
            End If
            If m_prevDCobject = 0& Then
                m_prevDCobject = SelectObject(m_hDC, m_Handle)
            End If
            LoadDIBinDC = m_hDC
        End If
    Else
        If Not m_prevDCobject = 0& Then
            SelectObject m_hDC, m_prevDCobject
            If m_ManageDC = False Then
                DeleteObject m_hDC
                m_hDC = 0&
            End If
            m_prevDCobject = 0&
        End If
    End If
End Function


'@Ignore FunctionReturnValueNotUsed
Public Function InitializeDIB(ByVal Width As Long, ByVal Height As Long) As Boolean

    ' Purpose: create a blank (all black, all transparent) DIB of requested height & width
    
    Dim tBMPI As BITMAPINFO
    Dim tDC As Long

    
    DestroyDIB ' clear any pre-existing dib
    
    If Width < 0& Then Exit Function
    If Height = 0& Then
        Exit Function
    ElseIf Height < 0& Then
        Height = Abs(Height) ' no top-down dibs
    End If
    
    On Error Resume Next
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40&
        .biSizeImage = .biHeight * .biWidth * 4&
    End With
    If Err Then
        Err.Clear
        ' only possible error would be that Width*Height*4& is absolutely huge
        Exit Function
    End If
    
    tDC = GetDC(0&) ' get screen DC
    m_Handle = CreateDIBSection(tDC, tBMPI, 0&, m_Pointer, 0&, 0&)
    If m_ManageDC = True Then
        ' create a DC if class is managing its own & one isn't created yet
        If m_hDC = 0& Then m_hDC = CreateCompatibleDC(tDC)
    End If
    ' release the screen DC if we captured it
    ReleaseDC 0&, tDC
    
    If Not m_Handle = 0& Then    ' let's hope system resources allowed DIB creation
        m_Width = Width
        m_Height = Height
        
        ' ref point 0001 - only location where LaVolpe's code has been modified
        ' expose the size of the original PNG image to the rest of the program
        origWidth = m_Width
        origHeight = m_Height
        
        'check the size of the preview PNG image and display it
        'rDIconConfigForm.lblWidthHeight.Caption = " width " & m_Width & " height " & m_Width & " (pixels)"

        
        EraseDIB                ' have experienced new dib filled with non-zeroes. Zero it out
        InitializeDIB = True
    End If

End Function

Public Sub DestroyDIB()
    
    ' PURPOSE: Destroy any existing image
    Set m_GDIplus = Nothing
    If Not m_hDC = 0& Then   ' do we have a DC?
        ' do we have an image; if so get it out of the DC
        If Not m_prevDCobject = 0& Then SelectObject m_hDC, m_prevDCobject
        ' destroy our DC, no point in keeping it w/o image
        DeleteObject m_hDC
        m_hDC = 0&
    End If
    ' if we do have an image, destroy it now
    If Not m_Handle = 0& Then
        DeleteObject m_Handle
        Erase m_ImageByteCache
    End If
    ' reset other image attributes
    m_Width = 0&
    m_Height = 0&
    m_Handle = 0&
    m_Pointer = 0&
    m_prevDCobject = 0&
    m_AlphaImage = AlphaNone
    m_Format = imgError
End Sub

Public Sub EraseDIB()
    ' Purpose: clear out an existing DIB pixels, making it 100% transparent/black
    If Not m_Handle = 0& Then
        FillMemory ByVal m_Pointer, m_Width * m_Height * 4&, 0
        m_Format = imgNone
        m_AlphaImage = AlphaSimple
    End If
    
End Sub


'@Ignore FunctionReturnValueNotUsed
Public Function Render(ByVal destinationDC As Long, _
                Optional ByVal destX As Long, Optional ByVal destY As Long, _
                Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, _
                Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal Blend As Boolean = True, _
                Optional ByVal SetHalfTone As Boolean = True, _
                Optional ByRef destHostDIB As c32bppDIB = Nothing, _
                Optional ByVal grayScale As eGrayScaleFormulas = gsclNone, _
                Optional ByVal LightAdjustment As Single = 0!, _
                Optional ByVal Angle As Single = 0!, _
                Optional ByVal CenterOnDestXY As Boolean = False) As Boolean

    ' PURPOSE: Render an existing 32bpp DIB to a target DC
    ' Angle & CenterOnDestXY added to allow this routine to replace the RotateAtCenterPoint & RotateAtTopLeft calls
    
    ' Mirroring: When destWidth & srcWidth are compared, if one is negative and the other positive, mirroring horizontally occurs
    '            When destHeight & srcHeight are compared, if one is negative and the other positive, mirroring vertically occurs
    ' All four of those parameters are optional & any that are not passed will default to the image's width/height as needed
    ' before the mirroring check is applied
    
    ' Parameters. Only destinationDC is required
    ' destinationDC :: target DC to draw to. Ignored if destHostDIB is passed
    ' destX, destY :: the top/left coordinates to draw to, default is 0,0
    ' destWidth, destHeight :: the width and height to draw to, default is the image's width & height
    ' srcX, srcY :: the left & top offset within the DIB
    ' srcWidth, srcHeight :: the amount of DIB to be rendered
    ' Opacity :: how opaque to draw the image, default is 100% opaque
    ' Blend :: no longer used, reserved & left in for backward compatibility
    ' SetHalfTone :: if True, then the destination DC's stretch mode will be modified to
    '       produce better quality results. This option is not available on Win9x systems.
    '       Tip: When AlphaBlending to another DIB set to False
    '            When AlphaBlending to CompatibleBitmap (DDB) or visible DC set to True
    ' destHostDIB :: When rendering from DIB class to DIB class, pass the destination
    '       DIB class to ensure alpha blending occurs correctly on systems that do not
    '       support GDI+ or AlphaBlend APIs. When passed, destinationDC is ignored
    ' grayscale :: one of several formulas to grayscale while rendering (optional)
    ' LightAdjustment :: values between -100 and 100 percent of added pixel darkeness/lightness
    '       -100% will display a black image & 100 percent will display a white image
    ' Angle :: between -360 and 360. Rotation is clockwise
    ' CenterOnDestXY :: If true then rendering is centered on the destX,destY coordinates
    
    Dim lBlendFunc As Long
    Dim tDC As Long
    Dim hOldImage As Long
    
    Dim lStretchMode As Long
    Dim aResizedBytes() As Byte
    Dim aMirrorBytes() As Byte
    
    Dim bStretching As Boolean
    Dim bMirroring As Boolean
    Dim bRotating As Boolean
    Dim bCanUseAlphaBlend As Boolean
    
    Const AC_SRC_OVER = &H0&
    Const AC_SRC_ALPHA = &H1&
    
    ' validate a few things
    If Opacity < 1& Then        ' nothing to render if image is 100% transparent
        Render = Not (m_Handle = 0)
        Exit Function
    ElseIf m_Handle = 0& Then
        Exit Function
    ElseIf destinationDC = 0& Then
        If destHostDIB Is Nothing Then Exit Function
    End If
    
    ' validate optional destination parameters
        If destWidth = 0& Then destWidth = m_Width
        If destHeight = 0& Then destHeight = m_Height
    
    ' validate optional parameters for source image
        If SrcX < 0& Then SrcX = 0&  ' source X,Y cannot be negative
        If SrcY < 0& Then SrcY = 0&  ' but the dest X,Y can be
        If srcWidth = 0& Then
            srcWidth = m_Width
        ElseIf srcWidth < 0& Then        ' locally, only the destWidth can be negative for mirroring, not srcHeight
            destWidth = -destWidth
            srcWidth = -srcWidth
        End If
        If srcHeight = 0& Then
            srcHeight = m_Height
        ElseIf srcHeight < 0& Then      ' locally, only the destHeight can be negative for mirroring, not srcHeight
            destHeight = -destHeight
            srcHeight = -srcHeight
        End If
        If SrcX + srcWidth > m_Width Then srcWidth = m_Width - SrcX
        If SrcY + srcHeight > m_Height Then srcHeight = m_Height - SrcY
        
    
    ' angle is Single. See if passed angle is evenly divisible by 360
    If Int(Angle) = Angle Then
        bRotating = Not ((Angle Mod 360) = 0)
    Else    ' angle has fractional component; therefore can't be Mod 360
        bRotating = True
    End If
    
    ' validate opacity is within range
    Opacity = Abs(Opacity) Mod 100
    If Opacity = 0& Then Opacity = 100&
    
    ' validate light adjustment is within range. Disqualify LigthAdjustmnet if out of range
    If Not LightAdjustment = 0! Then
        If LightAdjustment > 100! Then
            LightAdjustment = 0!
        ElseIf LightAdjustment < -100! Then
            LightAdjustment = 0!
        End If
    End If
    
    If CenterOnDestXY = True Then
'        If bRotating Then
'            destX = destX - Abs(destWidth) \ 2       ' use negative destWidth,destHeight if mirroring - rotation routines expect it
'            destY = destY - Abs(destHeight) \ 2
'            If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
'            'Render = spt_RotateImage(destinationDC, Angle, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, destHostDIB, grayScale, LightAdjustment)
'            If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
'            Exit Function
'        Else
            destX = destX - Abs(destWidth \ 2)  ' not rotating, don't allow negative destWidth,destHeight for calculations
            destY = destY - Abs(destHeight \ 2)
'        End If
'    ElseIf bRotating Then
'        If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
'        'Render = spt_RotateImage(destinationDC, Angle, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, destHostDIB, grayScale, LightAdjustment)
'        If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
'        Exit Function
    End If
    
    ' are we mirroring? rules out AlphaBlend usage if we are
    If destWidth > 0& Then
        bMirroring = (destHeight < 0&)
    Else
        bMirroring = True
    End If
    ' are we resizing? may rule out AlphaBlend usage (stretching DIB to DIB can crash on Win9x)
    If (Abs(destWidth) = srcWidth) Then
        bStretching = Not (Abs(destHeight) = srcHeight)
    Else
        bStretching = True
    End If
    
    ' The following IF tree and above boolean assignments are to determine:
    ' 1. Do we use AlphaBlend
    ' 2. Do we use GDI+
    ' 3. Do we do it manually.
    ' The answer depends on user settings, O/S, graphics manipulation & DLL abilities
    If m_KeepGDIplusActive = False Then
        ' if user provided token, allow GDI+ to override AlphaBlend
        If Me.isGDIplusEnabled = False Or Me.gdiToken = 0& Then
            ' see if alphablend will support the various rendering options
            If Me.isAlphaBlendFriendly Then     ' Win98 or better with AlphaBlend enabled & GDI+ disabled
                If Not bMirroring Then                  ' can't use alphaBlend if mirroring
                    If grayScale = gsclNone Then        ' can't use alphaBlend if gray scaling
                        If LightAdjustment = 0! Then    ' can't use alphaBlend if modifying pixel brightness
                            If bStretching Then
                                bCanUseAlphaBlend = Not m_StretchQuality ' can't use alphaBlend if using Bilinear interpolation
                            Else
                                bCanUseAlphaBlend = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    ' note. If AlphaBlend cannot support the rendering options, we will want to use GDI+
    
    If Me.isGDIplusEnabled = True And bCanUseAlphaBlend = False Then
        ' we will use GDI+ to render when system is not AlphaBlend friendly, user supplied gdi+ token, or AlphaBlend cannot handle rendering options
        Dim cGDIp As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDIp = New cGDIPlus
            If m_KeepGDIplusActive Then Set m_GDIplus = cGDIp
        Else
            Set cGDIp = m_GDIplus
        End If
        If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
        Render = cGDIp.RenderGDIplus(Me, destinationDC, 0&, Opacity, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, m_StretchQuality, grayScale, m_GDItoken, LightAdjustment)
        If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
    
    Else
        
        
        If m_hDC = 0& Then  ' do we have a DC to select our image into?
            tDC = GetDC(0&) ' if not create one, if ManageOwnDC=True, we will have one
            m_hDC = CreateCompatibleDC(tDC)
            ReleaseDC 0&, tDC
            hOldImage = SelectObject(m_hDC, m_Handle)
        Else
            ' we have a DC, but is the image selected into it?
            If m_prevDCobject = 0& Then hOldImage = SelectObject(m_hDC, m_Handle)
        End If
        
        If bCanUseAlphaBlend = False Then
            ' Ruled out use of AlphaBlend (preferred when GDI+ isn't available)
            ' Win95/NT4 - not shipped with AlphaBlend
            ' Mirroring or high quality interpolation stretching - can't use AlphaBlend
            ' Grayscaling/light adjustments on the fly - can't use AlphaBlend
            ' Stretching from DIB to destination - can crash with Win9x
            ' AlphaBlend can't do high quality interpolation
            

            ' doing it completely manually
            ' 1. Mirror and change light intensity as needed
            If bMirroring Then spt_MirrorDIB SrcX, SrcY, srcWidth, srcHeight, destWidth, destHeight, aMirrorBytes(), , LightAdjustment
            ' 2. Resize using Nearest Neighbor or Bi-Linear algorithms & change light intensity as needed, then Render
            If bStretching Then
                If spt_pvResize(destinationDC, aResizedBytes(), aMirrorBytes(), Nothing, SrcX, SrcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight, LightAdjustment) = False Then Exit Function
                ' use custom blending routine, passing aResizedBytes array
                Render = spt_Win9xBlend(destinationDC, aResizedBytes(), SrcX, SrcY, destX, destY, destWidth, destHeight, (255& * Opacity) \ 100&, destHostDIB, grayScale, LightAdjustment)
            Else
                ' use custom blending routine, passing aMirrorBytes array. If LigthAdjustment not applied above, it will be applied in spt_Win9xBlend
                Render = spt_Win9xBlend(destinationDC, aMirrorBytes(), SrcX, SrcY, destX, destY, destWidth, destHeight, (255& * Opacity) \ 100&, destHostDIB, grayScale, LightAdjustment)
            End If
            
        Else ' we can use AlphaBlend
        
            If Not destHostDIB Is Nothing Then destinationDC = destHostDIB.LoadDIBinDC(True)
            
            If SetHalfTone Then ' Stretch_Halftone not compatible with win9x
                If ((m_osCAP And osIsNT) = osIsNT) Then lStretchMode = SetStretchBltMode(destinationDC, STRETCH_HALFTONE)
            End If
            
            ' calculate the opacity required & add it to the BlendFunction variable
            lBlendFunc = AC_SRC_OVER Or (((255& * Opacity) \ 100&) * &H10000)
            ' if the image has transparency, then we add the AC_SRC_ALPHA flag too
            If Me.Alpha > AlphaNone Then lBlendFunc = lBlendFunc Or (AC_SRC_ALPHA * &H1000000)
            Render = Not (AlphaBlend(destinationDC, destX, destY, destWidth, destHeight, m_hDC, SrcX, SrcY, srcWidth, srcHeight, lBlendFunc) = 0&)
            
            If SetHalfTone Then ' Stretch_Halftone not compatible with win9x
                If ((m_osCAP And osIsNT) = osIsNT) Then SetStretchBltMode destinationDC, lStretchMode
            End If
            If Not destHostDIB Is Nothing Then destHostDIB.LoadDIBinDC False
            
        End If
        
        ' remove the image from the DC if necessary
        If Not hOldImage = 0& Then SelectObject m_hDC, hOldImage
        If Not tDC = 0& Then    ' if we created a DC, let's destroy it now
            DeleteDC m_hDC
            m_hDC = 0&
        End If
        
    End If
    
End Function












































Public Function CreateGDIplusToken() As Long

    ' Purpose: create a GDI+ token for you.
    ' Creating a token does not assign it to the class. Call gdiToken property
    Dim cGDI As cGDIPlus
    Dim gToken As Long
    If m_GDIplus Is Nothing Then
        Set cGDI = New cGDIPlus
    Else
        Set cGDI = m_GDIplus
    End If
    If cGDI.isGDIplusOk(gToken, True) = True Then
        CreateGDIplusToken = gToken
    End If
End Function
'@Ignore FunctionReturnValueNotUsed
Public Function DestroyGDIplusToken(ByVal Token As Long) As Boolean

    ' Purpose: release GDI+ for you. If the passed Token is used
    ' by other classes, you must set the other class' gdiToken property to zero
    ' before calling this routine.
    If Not Token = 0& Then
        Dim cGDI As cGDIPlus
        If m_GDIplus Is Nothing Then
            Set cGDI = New cGDIPlus
        Else
            Set cGDI = m_GDIplus
        End If
        DestroyGDIplusToken = cGDI.InitializeGDIplus(Token, True)
    End If
End Function



'Public Function GetDroppedFileNames(ByRef OLEDragDrop_DataObject As DataObject) As Boolean
'
'    ' Purpose: return unicode filename & is only included to make things a little easier
'    '   for those that want to fully support unicode filenames whereas VB only provides
'    '   ANSI filenames. It is designed to be called from your OLEDragDrop event, passing that
'    '   event's Data parameter's object like so: Possible example from your project
'
''    Private Sub picPreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
''        ' simmple OLE drag/drop example, assuming this class declared as cImage
''
''        If cImage.GetDroppedFileNames(Data) = True Then
''            [add code here]. The Data object will contain unicode filenames as necessary
''        Else
''            [add code here]. No file names were dropped, something else was
''        End If
''
''    End Sub
'
'    ' Caution: Editing this routine after it has been called may crash the IDE
'    ' I believe I have fixed that issue but am not 100% positive
'
'    If OLEDragDrop_DataObject Is Nothing Then Exit Function
'    If OLEDragDrop_DataObject.GetFormat(vbCFFiles) = False Then Exit Function
'
'    Dim fmtEtc As FORMATETC
'Dim pMedium As STGMEDIUM
'
'    Dim dFiles As DROPFILES
'    Dim Vars(0 To 1) As Variant
'Dim pVars(0 To 1) As Long
'Dim pVartypes(0 To 1) As Integer
'
'    Dim varRtn As Variant
'    'Dim iFiles As Long
'Dim iCount As Long
'Dim hDrop As Long
'
'    Dim lLen As Long
'Dim sfile As String
'
'
'    Dim IID_IDataObject As Long ' IDataObject Interface ID
'    Const IDataObjVTable_GetData As Long = 12 ' 4th vtable entry
'    Const CC_STDCALL As Long = 4&
'    Const TYMED_HGLOBAL = 1
'    Const DVASPECT_CONTENT = 1
'
'    With fmtEtc
'        .cfFormat = vbCFFiles         ' same as CF_DROP
'        .lIndex = -1                    ' want all data
'        .TYMED = TYMED_HGLOBAL        ' want global ptr to files
'        .dwAspect = DVASPECT_CONTENT  ' no rendering
'    End With
'
'    ' The IDataObject pointer is 16 bytes after VBs DataObject
'    CopyMemory IID_IDataObject, ByVal ObjPtr(OLEDragDrop_DataObject) + 16, 4&
'
'    ' So, if we know the VTable address and we know which function index we want
'    ' to call, we can call it directly using the following OLE API. Otherwise we
'    ' would need to use a TLB to define the IDataObject interface since VB doesn't
'    ' 't expose it. This has some really neat implications if you think about it.
'    ' The IDataObject function we want is GetData which is the 4th function in
'    ' the VTable... http://msdn2.microsoft.com/en-us/library/ms688421.aspx
'
'    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
'    pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))
'
'    ' The variants are required by the OLE API: http://msdn2.microsoft.com/en-us/library/ms221473.aspx
'    If DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, _
'                        vbLong, 2, pVartypes(0), pVars(0), varRtn) = 0 Then
'
'        If pMedium.data = 0 Then
'            Exit Function      ' nothing to do
'        Else
'            ' we have a pointer to the files, kinda sorta
'            CopyMemory hDrop, ByVal pMedium.data, 4&
'            If Not hDrop = 0 Then
'                ' the hDrop is a pointer to a DROPFILES structure
'                ' copy the 20-byte structure for our use
'                CopyMemory dFiles, ByVal hDrop, 20&
'            End If
'        End If
'
'        If dFiles.fWide = 0 Then ' ansi system (Win9x)
'            GlobalFree pMedium.data
'
'        Else
'            ' use the pFiles member to track offsets for file names
'            dFiles.pFiles = dFiles.pFiles + hDrop
'            iCount = OLEDragDrop_DataObject.Files.count
'            OLEDragDrop_DataObject.Files.Clear
'            For iCount = 1 To iCount
'                ' get the length of the current file & multiply by 2 because it is unicode
'                ' lstrLenW is supported in Win9x
'                lLen = lstrlenW(ByVal dFiles.pFiles) * 2
'                sfile = String$(lLen \ 2, 0)    ' build a buffer to hold the file name
'                CopyMemory ByVal StrPtr(sfile), ByVal dFiles.pFiles, lLen ' populate the buffer
'                ' add the file
'                OLEDragDrop_DataObject.Files.Add sfile, iCount
'                ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
'                dFiles.pFiles = dFiles.pFiles + lLen + 2
'            Next
'            GlobalFree pMedium.data
'
'        End If
'
'        GetDroppedFileNames = True
'
'    End If
'
'End Function

'Public Function GetPastedFileNames(ByRef ListOfFiles() As String) As Long
'
'    ' Purpose: support function for unicode filenames. The filenames returned by
'    ' VB's Clipboard object only contains ANSI.
'
'    ' This function will return the number of file names in the clipboard, if any.
'    ' The returned string array will contain the unicode/ansi filenames as needed.
'    ' If the function returns zero, then the string array is not valid
'
'    Dim hDrop As Long
'    Dim lLen As Long
'    Dim iCount As Long
'    Dim dFiles As DROPFILES
'
'   ' Get handle to CF_HDROP if any:
'   If OpenClipboard(0&) = 0 Then Exit Function
'
'    hDrop = GetClipboardData(vbCFFiles)
'    If Not hDrop = 0 Then   ' then copied/cut files exist in memory
'        iCount = DragQueryFile(hDrop, -1&, vbNullString, 0)
'        If iCount Then
'            ' the hDrop is a pointer to a DROPFILES structure
'            ' copy the 20-byte structure for our use
'            CopyMemory dFiles, ByVal hDrop, 20&
'            ' use the pFiles member to track offsets for file names
'            dFiles.pFiles = dFiles.pFiles + hDrop
'
'            ReDim ListOfFiles(1 To iCount)
'
'            For iCount = 1 To iCount
'                If dFiles.fWide = 0 Then   ' ANSI text, use API to get file name
'                   lLen = DragQueryFile(hDrop, iCount - 1, vbNullString, 0&)       ' query length
'                   ListOfFiles(iCount) = String$(lLen, 0)                          ' set up buffer
'                   DragQueryFile hDrop, iCount - 1, ListOfFiles(iCount), lLen + 1  ' populate buffer
'                Else
'                   ' get the length of the current file & multiply by 2 because it is unicode
'                   ' lstrLenW is supported in Win9x
'                   lLen = lstrlenW(ByVal dFiles.pFiles) * 2
'                   ListOfFiles(iCount) = String$(lLen \ 2, 0)    ' build a buffer to hold the file name
'                   CopyMemory ByVal StrPtr(ListOfFiles(iCount)), ByVal dFiles.pFiles, lLen ' populate the buffer
'                   ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
'                   dFiles.pFiles = dFiles.pFiles + lLen + 2
'               End If
'            Next
'            GetPastedFileNames = iCount - 1
'        End If
'    End If
'    CloseClipboard
'
'End Function




' NEW SECTION *******************************************************************************
'                   CLASS TO CLASS COMMUNICATION METHODS
' *******************************************************************************************




' NEW SECTION *******************************************************************************
'                       LOCAL SUPPORT FUNCTIONS FOR THIS CLASS ONLY
' *******************************************************************************************




Private Function spt_LoadPictureEx(ByVal FileHandle As Long, ByRef Filename As String, ByRef aStream() As Byte, _
                            ByRef Cx As Long, ByRef Cy As Long, _
                            ByRef streamOffset As Long, ByRef streamLength As Long, _
                            ByVal SaveFormat As Boolean, ByRef bitDepth As Long) As Boolean
    
    ' PURPOSE: Marshal passed file/array to image classes for conversion to 32bpp image
    ' For parameter information, see LoadPicture_File & LoadPicture_Stream
    
    Me.DestroyDIB
    
    ' various image parsers, in order of precedence
    ' All 4 recognize transparency
    Dim cPNG As cPNGparser  ' very fast to abort if not a PNG file
    Dim cGIF As cGIFparser  ' very fast to abort if not a GIF file
    Dim cICO As cICOparser  ' must parse key parts of a file. handles icons & Vista PNG Icons
    Dim cBmp As cBMPparser  ' catchall. Handles bitmaps, wmf, emf & jpgs
    Dim cGDI As cGDIPlus    ' fallback if cGDI is available
    Dim cTGA As cTGAparser  ' TGA formats; v1 is loose format difficult to verify - so last parser in list
    
    Dim bReturn As Boolean  ' function return value
    Dim rtnRead As Long
    
    ' validate passed desired icon sizes
    If Cx < 0& Then Cx = 0&
    If Cy < 0& Then Cy = 0&
    If bitDepth < 0& Then
        bitDepth = 32
    ElseIf bitDepth > 32 Then
        bitDepth = 32
    End If
    
    Set cPNG = New cPNGparser   ' see if image is a PNG; aborts quickly if not
    If FileHandle = 0 Then
        bReturn = cPNG.LoadStream(aStream(), Me, streamOffset, streamLength, m_GDItoken)
    Else     ' note: processing from file is slightly faster than via array
        bReturn = cPNG.LoadFile(FileHandle, Filename, Me, m_GDItoken)
    End If
    If Err Then MsgBox Err.Description
    Set cPNG = Nothing
    If Not bReturn Then
        If Not FileHandle = 0& Then
            streamOffset = 0&
            streamLength = GetFileSize(FileHandle, 0&)
            ReDim aStream(streamOffset To streamLength - 1&)
            SetFilePointer FileHandle, 0&, 0&, 0&
            ReadFile FileHandle, aStream(streamOffset), streamLength, rtnRead, ByVal 0&
        End If
        Set cGIF = New cGIFparser ' what about a GIF; aborts quickly if not
        bReturn = cGIF.LoadStream(aStream(), Me, streamOffset, streamLength)
        Set cGIF = Nothing
        If Not bReturn Then
            Set cICO = New cICOparser   ' will process Vista PNG icon if needed
            bReturn = cICO.LoadStream(aStream(), Cx, Cy, Me, streamOffset, streamLength, bitDepth, m_GDItoken)
            Set cICO = Nothing
            If Not bReturn Then ' check for bmp, emf, wmf & jpg << last chance
                Set cBmp = New cBMPparser
                bReturn = cBmp.LoadStream(aStream(), Me, streamOffset, streamLength)
                Set cBmp = Nothing
            End If
        End If
    End If
    If m_Handle = 0& Then ' hmmm, not an image file processed here, we can try
                          ' one more thing... Send it GDI+ if possible
        If Me.isGDIplusEnabled Then
            If m_GDIplus Is Nothing Then
                Set cGDI = New cGDIPlus
            Else
                Set cGDI = m_GDIplus
            End If
            If cGDI.GDIplusLoadPNG(Filename, aStream(), Me, m_GDItoken) = True Then
                ' it worked, whatever file type it was. Convert it to PNG
                If SaveFormat = True Then
                    Call cGDI.SaveToPNG(vbNullString, aStream, Me, m_GDItoken)
                    FileHandle = 0& ' prevent next IF from trying to re-load it from file if applicable
                End If
                Me.ImageType = imgPNG
                bReturn = True
            End If
        End If
        If Not bReturn Then
            ' try TGA format
            Set cTGA = New cTGAparser
            bReturn = cTGA.LoadStream(aStream(), Me, streamOffset, streamLength)
            Set cTGA = Nothing
        End If
    End If
    If Not m_Handle = 0 Then
        If SaveFormat = True Then ' we will cache the original bytes
            If iparseIsArrayEmpty(VarPtrArray(aStream)) = 0& And Not FileHandle = 0& Then
                ' we loaded the image from the file and not a stream (PNG), need to get stream
                If streamLength = 0& Then streamLength = GetFileSize(FileHandle, 0&)
                ReDim m_ImageByteCache(0 To streamLength - 1)
                SetFilePointer FileHandle, 0&, 0&, 0&
                ReadFile FileHandle, m_ImageByteCache(0), streamLength, rtnRead, ByVal 0&
            Else
                m_ImageByteCache() = aStream()
            End If
        End If
        spt_LoadPictureEx = True
    End If

End Function

Private Function spt_pvResize(ByVal destDC As Long, _
                        ByRef rSizedBytes() As Byte, ByRef rMirror() As Byte, _
                        Optional ByRef tHost As c32bppDIB, _
                        Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                        Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                        Optional ByVal destX As Long, Optional ByVal destY As Long, _
                        Optional ByRef destWidth As Long, Optional ByRef destHeight As Long, _
                        Optional ByRef LightAdjustment As Single = 0!) As Boolean
                            
    ' Purpose: resize an alpha image, maintaining premultiplied pixels & alpha values
    ' Code originally by Carles P.V. but significantly modified for this project.
    
    ' Parameters:
    ' destDC :: DC being rendered to, may be null
    ' rSizedbytes() : array to hold resized alpha section; not used if tHost is not Nothing
    ' tHost : when resizing to another DIB class, the destination DIB class
    ' srcX,Y : the coordinates of the source image to start resizing from
    ' srcWidth,srcHeight : the width/height of the source image to resize from
    ' destX,Y : the coordinates of the destination image to resize to
    ' destWidth,destHeight : the width/height of the destination image to resize to

    If srcWidth = 0& Then srcWidth = m_Width
    If srcHeight = 0& Then srcHeight = m_Height

    Dim aNewBits() As Byte
    Dim dSA As SafeArray
    ' new size, overlay of DIB pointer
    Dim aOldBits() As Byte
    Dim tSA As SafeArray
    ' old size, overlay of DIB pointer
    
    Dim xLUdbl() As Double                     ' look up table (LUT)
    Dim xRatio As Double
    Dim yRatio As Double
     ' scaled ratios
    Dim srcPixel As Long
    Dim dstPixel As Long
     ' source/destination pixel locations
    Dim lCol As Long
    Dim lRow As Long
             ' loop variables
    
    Dim newWidth As Long
    Dim newHeight As Long
    
    
    ' following used with BiLinear scaling
    Dim fX As Double
    Dim fY As Double
    
    Dim tgtX As Long
    Dim tgtY As Long
    
    Dim edgeOffsetX As Long
    Dim edgeOffsetY As Long
    
    Dim iX As Long
    Dim iY As Long
    
    Dim R As Double
    Dim g As Double
    Dim b As Double
    Dim a As Double
    
    Dim scalerX As Double
    Dim scalerY As Double

    
    ' fill in optional parameters
    If Not tHost Is Nothing Then
        newWidth = tHost.Width
        newHeight = tHost.Height
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight
        If newWidth > tHost.Width Then newWidth = tHost.Width
        If newHeight > tHost.Height Then newHeight = tHost.Height
    Else
        newWidth = Abs(destWidth)
        newHeight = Abs(destHeight)
        
        ' Scaling ratio (ratio of actual image to scaled image)
        xRatio = srcWidth / newWidth
        yRatio = srcHeight / newHeight

        ' safety checks, recalculation of bounding destination size
        ' if not done, we could very easily access unallocated memory.
        If destX < 0 Then   ' negative DC offset
            newWidth = newWidth + destX    ' reduce width to process
            destX = -destX                 ' used to offset LUT; adjust so not processing bytes not used
        Else
            destX = 0&                        ' fits within destination bitmap; no offsetting needed
        End If
        
        ' now to check the vertical
        If destY < 0& Then  ' negative DC offset
            newHeight = newHeight + destY
            destY = 0&
        Else
            destY = 0&
        End If
        
    End If
    If newHeight < 1& Or newWidth < 1& Then Exit Function
    
    ' set DMA overlays on destination & source
    If tHost Is Nothing Then
        ReDim rSizedBytes(0& To newWidth * 4& - 1&, 0& To newHeight - 1&)
        dSA.pvData = VarPtr(rSizedBytes(0&, 0&))
    Else
        dSA.pvData = tHost.BitsPointer ' called by CopyImageTo & Resize routines
    End If
    iparseOverlayHost_Byte aNewBits, VarPtr(dSA), 2, newHeight, newWidth * 4&, dSA.pvData ' overlay DMA array
    
    If iparseIsArrayEmpty(VarPtrArray(rMirror)) = 0& Then
        iparseOverlayHost_Byte aOldBits, VarPtr(tSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
    Else ' source is the mirrored DIB; clipped as needed
        iparseOverlayHost_Byte aOldBits, VarPtr(tSA), 2, UBound(rMirror, 2) + 1&, UBound(rMirror, 1) + 1&, VarPtr(rMirror(0&, 0&))    ' overlay DMA array
    End If
    
    On Error GoTo eh
    
    SrcY = UBound(aOldBits, 2) - srcHeight - SrcY + 1&     ' adjust Y position in source for bottom up DIBs
    
    If (m_StretchQuality = False) Then ' nearest neighbor algorithm (very fast but lowest quality)
        
        ' Scaling LUT, cache actual X position of DIB in relation to scaled X
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        ReDim xLUdbl(0 To newWidth - 1&)
        For lCol = 0 To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUdbl(lCol) = Int(((lCol + destX) * xRatio) + SrcX) * 4&
        Next
        
        For lRow = newHeight - 1& To 0& Step -1&
            '^ current scanline for the scaled image
            ' offset destY is used for negative coordinates
            srcPixel = Int((lRow + destY) * yRatio) + SrcY  ' recalcualted once per scanline
            ' current scanline for the scaled image
            dstPixel = 0&
            For lCol = 0& To newWidth - 1&
                ' copy into resized array the nearest raw/actual pixel
                CopyMemory aNewBits(dstPixel, lRow), aOldBits(Int(xLUdbl(lCol)), srcPixel), 4&
                dstPixel = dstPixel + 4&
            Next lCol
        Next lRow
        
    Else
        
        ' BiLinear interoplation, up to 4 source pixels (16 bytes) are blended for each destination pixel (4 bytes)
 
        ReDim xLUd(0 To newWidth - 1&)   ' work with doubles, we need the decimal portions
        ' Cache one scan line of X coords so we don't have to calculate for every pixel
        For lCol = 0& To newWidth - 1&
            ' offset destX used for negative coordinates, X is location in source to start blending at
            xLUd(lCol) = (((lCol + destX) * xRatio) + SrcX)
        Next
        srcWidth = srcWidth - 1& ' subtract now vs subtracting in loop below
        For lRow = newHeight - 1& To 0& Step -1&
            
            fY = (lRow + destY) * yRatio + SrcY ' get the scaled source row
            tgtY = Int(fY)              ' get whole number of double
            If tgtY = 0& Then         ' for top down images, test for last row vs 0
                ' last row of source image, will use only this row
                edgeOffsetY = 0&
                fY = 0#
            Else
                ' will use this row & next row for blending
                edgeOffsetY = 1&
                fY = Abs(1# - (fY - tgtY)) ' for top down images, use fY = fY-tgtY
                ' ^ Y coordinate fraction; pct of next vertical pixel that is used
            End If
            
            For lCol = 0& To newWidth - 1&
                
                R = 0#: g = 0#: b = 0#: a = 0#
                
                tgtX = Int(xLUd(lCol))  ' coordinate rounded down to whole number
                If tgtX = srcWidth Then
                    ' at far edge of source image, will use only this pixel for blending
                    edgeOffsetX = 0&
                    fX = 0#
                Else
                    ' will use this pixel and next pixel for blending
                    edgeOffsetX = 1&
                    fX = xLUd(lCol) - tgtX
                    ' ^ X coordinate fraction, pct of next horizontal pixel that is used
                End If
                
                For iY = 0& To edgeOffsetY
    
                    scalerY = Abs(iY - fY)   ' percentage of current row's pixel to blend
                    If Not scalerY = 1# Then  ' else result will be zero
                    
                        For iX = 0& To edgeOffsetX
                            scalerX = Abs(fX - iX)   ' percentage of current column's pixel to blend
                            
                            If Not scalerX = 1# Then  ' else result will be zero
                                scalerX = (1# - scalerX) * (1# - scalerY) ' combine percentages
    
                                ' Build the blended RGB values. For top down images Add iY vs subtracting
                                srcPixel = (tgtX + iX) * 4&
                                b = b + aOldBits(srcPixel, tgtY - iY) * scalerX
                                g = g + aOldBits(srcPixel + 1&, tgtY - iY) * scalerX
                                R = R + aOldBits(srcPixel + 2&, tgtY - iY) * scalerX
                                a = a + aOldBits(srcPixel + 3&, tgtY - iY) * scalerX
    
                            End If
                        Next
                    End If
                Next
                iX = lCol * 4&
                ' update destination with adjusted pixel
                aNewBits(iX, lRow) = Int(b)
                aNewBits(iX + 1&, lRow) = Int(g)
                aNewBits(iX + 2&, lRow) = Int(R)
                aNewBits(iX + 3&, lRow) = Int(a)
            Next
        Next
    End If
    
    iparseOverlayHost_Byte aOldBits, 0, 0, 0, 0, 0  ' remove DMA overlay
    iparseOverlayHost_Byte aNewBits, 0, 0, 0, 0, 0  ' remove DMA overlay
    
    ' the passed destWidth,destHeight params are used when rendering; we are just sizing now
    destWidth = newWidth ' the parameter is ByRef, update it now
    destHeight = newHeight ' the parameter is ByRef, update it now
    Erase rMirror()
    
    If Not LightAdjustment = 0! Then
        Call spt_LightenDarken(Nothing, LightAdjustment, rSizedBytes(), SrcX, SrcY, srcWidth, srcHeight)
        LightAdjustment = 0!    ' reset, taken care of now
    End If
    
    spt_pvResize = True
eh:
    If Err Then
        Err.Clear   ' troubleshooting only, should be removed before compiling to final app
        Stop
        Resume
    End If
End Function

Private Sub spt_LightenDarken(ByRef cImage As c32bppDIB, ByVal Lightness As Single, ByRef rtnArray() As Byte, _
            ByRef SrcX As Long, ByRef SrcY As Long, ByVal srcWidth As Long, ByRef srcHeight As Long)

    ' called by Render, Win9xRender, MirrorDB, & spt_pvResize drawing routines as needed
    
    ' Purpose: lighten or darken non-transparent pixels and pass results in rtnArray
    ' Parameters:
    '   cImage :: if Nothing, then rtnArray has source bytes
    '   Lightness :: values between -100 and 100 percent. -100 will produce blackness & 100 will produce whiteness
    '   rtnArray :: the array that will hold the modified bytes
    ' Note: when rtnArray is also the source array, it must contain pre-multiplied bytes
    '       and be a 2-dimensional, zero-bound arrray
    ' Checks on the array dimensions and bounds are not made here. They are guaranteed by calling routines
    
    If Lightness = 0! Then Exit Sub
    If Not cImage Is Nothing Then
        ' ensure an image exists if passing a dib class
        If cImage.Handle = 0 Then Exit Sub
    End If
    
    Dim tSA As SafeArray
    Dim srcBytes() As Byte
    
    Dim dSA As SafeArray
    Dim dstBytes() As Byte
    
    
    Dim x As Long
    Dim y As Long
    Dim b As Long
    
    Dim dstX As Long
    Dim dstY As Long
    
    Dim tVal As Long
    Dim alphaAdj As Long

    Dim srcAlpha As Byte
    
    If cImage Is Nothing Then   ' source is the return array
        iparseOverlayHost_Byte srcBytes, VarPtr(tSA), 2, UBound(rtnArray, 2) + 1, UBound(rtnArray, 1) + 1, VarPtr(rtnArray(0, 0)) ' overlay DMA array
    Else                        ' source is a dib class
        iparseOverlayHost_Byte srcBytes, VarPtr(tSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
    End If
    
    If cImage Is Nothing Then ' source & destination are same
        dSA.pvData = tSA.pvData
    Else
        ReDim rtnArray(0 To srcWidth * 4 - 1, 0 To srcHeight - 1)
        dSA.pvData = VarPtr(rtnArray(0, 0))
    End If
    iparseOverlayHost_Byte dstBytes, VarPtr(dSA), 2, srcHeight, srcWidth * 4&, dSA.pvData ' overlay dest DMA array
    dstY = UBound(dstBytes, 2)
    
    Lightness = 255! * (Lightness / 100!) ' calculate pct of 255 that ligthness equates to
    
    ' when pre-multiplied bytes are in effect, we can't just add Lightness because
    '   the alpha value restricts what percent of the source rgb is actually used in relation to alpha.
    '   So we must modify lightness in relation to the alpha byte too.
    ' Separate loops used to prevent having to compare if the adjusted pixel is > 255 and also < 0
    
    If Lightness > 0! Then ' make lighter
          ' loop thru the rows
        For y = UBound(srcBytes, 2) - SrcY To UBound(srcBytes, 2) - SrcY - srcHeight + 1 Step -1
            dstX = 0&
            For x = SrcX * 4& To (SrcX + srcWidth) * 4& - 1 Step 4&
                srcAlpha = srcBytes(x + 3&, y)
                If srcAlpha = 0 Then
                    dstX = dstX + 4
                Else
                    ' calculate lightness in relation to alpha
                    alphaAdj = (srcAlpha * Lightness) \ 255
                    For b = x To x + 2&
                        tVal = srcBytes(b, y) + alphaAdj
                        If tVal > srcAlpha Then
                            dstBytes(dstX, dstY) = srcAlpha
                        Else
                            dstBytes(dstX, dstY) = (tVal And &HFF)
                        End If
                        dstX = dstX + 1
                    Next
                    dstBytes(dstX, dstY) = srcAlpha   ' ensure alpha is copied too
                    dstX = dstX + 1
                End If
            Next
            dstY = dstY - 1
        Next
    Else    ' make darker
        For y = UBound(srcBytes, 2) - SrcY To UBound(srcBytes, 2) - SrcY - srcHeight + 1 Step -1
            dstX = 0&
            For x = SrcX * 4& To (SrcX + srcWidth) * 4& - 1 Step 4&
                srcAlpha = srcBytes(x + 3&, y)
                If srcAlpha = 0 Then
                    dstX = dstX + 4
                Else
                    ' calculate lightness in relation to alpha
                    alphaAdj = (srcAlpha * Lightness) \ 255
                    For b = x To x + 2&
                        tVal = srcBytes(b, y) + alphaAdj
                        If tVal < 0& Then
                            dstBytes(dstX, dstY) = 0
                        Else
                            dstBytes(dstX, dstY) = (tVal And &HFF)
                        End If
                        dstX = dstX + 1
                    Next
                    dstBytes(dstX, dstY) = srcAlpha ' ensure alpha is copied too
                    dstX = dstX + 1
                End If
            Next
            dstY = dstY - 1
        Next
    End If
    
    ' remove overlays
    iparseOverlayHost_Byte srcBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    iparseOverlayHost_Byte dstBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    SrcX = 0: SrcY = 0
End Sub


Private Function spt_MirrorDIB(ByRef SrcX As Long, ByRef SrcY As Long, _
                    ByVal srcWidth As Long, ByVal srcHeight As Long, _
                    ByRef newWidth As Long, ByRef newHeight As Long, _
                    ByRef mirrorBytes() As Byte, Optional ByRef tHost As c32bppDIB, _
                    Optional ByRef LightAdjustment As Single = 0!) As Boolean

    ' Purpose: mirror an image to an array
    
    ' is image being mirrored?
    If newWidth > 0& And newHeight > 0& Then Exit Function
    
    Dim tSA As SafeArray
    Dim srcBytes() As Byte
    
    Dim dSA As SafeArray
    Dim dstBytes() As Byte
    
    Dim x As Long
    Dim y As Long
    Dim yOffset As Long
    Dim xOffset As Long

    
    iparseOverlayHost_Byte srcBytes, VarPtr(tSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay source DMA array
    If tHost Is Nothing Then    ' destination is mirrorBytes
        ReDim mirrorBytes(0& To srcWidth * 4& - 1&, 0& To srcHeight - 1&)
        dSA.pvData = VarPtr(mirrorBytes(0&, 0&))
    Else                        ' destination is passed DIB class
        dSA.pvData = tHost.BitsPointer ' called by CopyImageTo & Resize routines
    End If
    iparseOverlayHost_Byte dstBytes, VarPtr(dSA), 2, srcHeight, srcWidth * 4&, dSA.pvData ' overlay dest DMA array
    
    If newHeight < 0& Then
        If newWidth > 0& Then   ' flipping vertically only, faster/easier
            xOffset = srcWidth * 4& ' number of bytes to flip at once
            y = SrcY    ' when flipping, our Y is adjusted from the bottom
            x = SrcX * 4&   ' starting X position in source
            For yOffset = srcHeight - 1& To 0& Step -1&
                CopyMemory dstBytes(0&, yOffset), srcBytes(x, y), xOffset
                y = y + 1&   ' move source Y to next row
            Next
        Else   ' flipping both vertically/horizontally
            y = SrcY ' when flipping vertically, adjust from the bottom
            For yOffset = srcHeight - 1& To 0& Step -1&
                x = (m_Width - SrcX) * 4& - 4&  ' X adjusted from right when flipping
                For xOffset = 0 To srcWidth * 4 - 4& Step 4&
                    CopyMemory dstBytes(xOffset, yOffset), srcBytes(x, y), 4&
                    x = x - 4&  ' move source X to next pixel
                Next
                y = y + 1&   ' move source Y to next row
            Next
        End If
    
    Else  ' flipping horizontally only, same comments as above
        y = m_Height - SrcY - 1& ' not fipping vertically, so read from bottom up
        For yOffset = srcHeight - 1& To 0& Step -1&
            x = (srcWidth + SrcX) * 4& - 4&
            For xOffset = 0 To srcWidth * 4 - 4& Step 4&
                CopyMemory dstBytes(xOffset, yOffset), srcBytes(x, y), 4&
                x = x - 4&
            Next
            y = y - 1&
        Next
    End If
    iparseOverlayHost_Byte srcBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    iparseOverlayHost_Byte dstBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    
    ' rendering routines will now use the mirrored byte array
    ' so we ensure the bounds parameters match the mirrored array bounds
    SrcX = 0&
    SrcY = 0&
    newWidth = Abs(newWidth)
    newHeight = Abs(newHeight)
    
    ' if light adjustments, preprocess bytes now
    If Not LightAdjustment = 0! Then
        Call spt_LightenDarken(Nothing, LightAdjustment, mirrorBytes(), SrcX, SrcY, newWidth, newHeight)
        LightAdjustment = 0!    ' reset, taken care of now
    End If
    spt_MirrorDIB = True
    
End Function

Private Function spt_Win9xBlend(ByVal destinationDC As Long, ByRef aResizedBytes() As Byte, _
                            ByVal SrcX As Long, ByVal SrcY As Long, _
                            ByVal destX As Long, ByVal destY As Long, _
                            ByVal destWidth As Long, ByVal destHeight As Long, _
                            ByVal GlobalAlpha As Long, ByRef tHost As c32bppDIB, _
                            ByVal grayScale As eGrayScaleFormulas, _
                            ByVal lightAdj As Single) As Boolean
    
    ' Purpose: manually blend an alpha bitmap to a target DC
    
    ' Never called when GDI+ is available unless user forced isGDIplusEnabled=False.
    ' Used when AlphaBlend is not available or when AlphaBlend is available but cannot
    '   perform the graphic manipulation required
    
    ' Parameters identify the destination more than anything else. The source was already pre-processed if needed
    
    ' destinationDC :: DC to blend to
    ' aResizedBytes() :: array of bytes sized to target destination blend area.
    '   if array is null, then the destination size is same size as our DIB's image
    ' srcX,Y :: the position on source to begin blending
    ' destX,Y :: the position on destination where blending starts
    ' destWidth,Height :: the amount of columns/rows to blend
    ' globalAlpha :: the AlphaBlend global alpha value: between 0 and 255
    ' tHost :: blending will occur DIB to DIB vs DIB to DC
    ' grayScale :: if simultaneously grayscaling then the grayscale formula
    ' ligthAdj :: light adjustment value adds/subtracts -100 to 100 percent intensity per pixel
    
    ' Special note. Having problems rendering 32bpp DIBs to WinME; artifacts are being rendered
    ' from the alpha channel. Therefore, to completely eliminate this problem (hopefully), the
    ' destination array will be 24bpp vs 32bpp which will then be updated and rendered onto the
    ' destination DC.  32bpp would be easier, but oh well.

    ' through trial and error, the rule, rendering appears to be
    ' flip first, then stretch and/or rotate, then clip as needed
    
    Dim srcBytes() As Byte
    Dim srcSA As SafeArray
    
    Dim dstBytes() As Byte
    Dim dstSA As SafeArray
    
    Dim srcCol As Long
    Dim srcRow As Long
    
    Dim srcAlpha As Long
    Dim dstAlpha As Long
    
    'Dim y As Long
    'Dim x As Long
    
    Dim sX As Long
    'Dim sY As Long
    
    Dim dX As Long
    Dim dY As Long
    
    Dim dDC As Long
    Dim tDC As Long
    Dim hOldBmp As Long
    Dim hDIB As Long
    
    Dim Rg As Single
    Dim Gg As Single
    Dim Bg As Single

    Dim gScaleByte As Long
    
    Dim BMPI As BITMAPINFO
    
    ' The following is just a wee bit confusing.
    ' Our source can be 2 different objects:
    '   1) Our host DIB
    '   2) The passed aResizedBytes() array if pre-processing was required
    ' Likewise, the destination can be 2 different objects
    '   1) A passed DC handle (DIB to 24bpp bitmap (render 32bpp to 24bpp)
    '   2) Another DIB class if tHost is passed (render 32bpp to 32bpp)
    ' So to use a common base for all possibilities, we use SafeArrays
    
    If iparseIsArrayEmpty(VarPtrArray(aResizedBytes)) = 0& Then
        If lightAdj = 0! Then
            srcSA.pvData = m_Pointer
            ' need to tweak for negative offsets
            If destX < 0& Then
                SrcX = SrcX - destX ' less area that needs to be rendered
                destWidth = destWidth + destX ' adjust amount of destination we copy
                destX = 0&                    ' set destination offset to zero
            End If
            If destY < 0& Then
                SrcY = SrcY - destY ' less area that needs to be rendered
                destHeight = destHeight + destY ' adjust amount of destination we copy
                destY = 0&                      ' set destination offset to zero
            End If
            SrcY = m_Height - SrcY  ' set DIB offset for 1st row to be blended
        Else
            ' if light adjustments, preprocess bytes
            spt_LightenDarken Me, lightAdj, aResizedBytes(), SrcX, SrcY, destWidth, destHeight
            srcSA.pvData = VarPtr(aResizedBytes(0, 0))
            SrcX = 0&: SrcY = destHeight
        End If
     Else ' source is the resized array
         srcSA.pvData = VarPtr(aResizedBytes(0, 0))
         If destX < 0& Then 'destX = 0& ' Abs(destX) ' when negative target coords are used, we set to
            destWidth = destWidth + destX
            SrcX = -destX: destX = 0&
         Else
            SrcX = 0&
         End If
         If destY < 0& Then 'destY = 0& 'Abs(destY) ' zero to match the resized array (zero-based)
            SrcY = destHeight + destY
            destHeight = destHeight + destY
            destY = 0&
        Else
            SrcY = destHeight
         End If
    End If
     
     If Not tHost Is Nothing Then
        ' need to ensure we won't be copying memory to unallocated memory
        If destWidth > tHost.Width Then destWidth = tHost.Width
        If destHeight > tHost.Height Then destHeight = tHost.Height
     End If
     If destWidth < 1& Or destHeight < 1& Then Exit Function  ' nothing to Blt, passed bad params
    
    If tHost Is Nothing Then
        ' we need the contents of the target DC, describe its bitmap
        With BMPI.bmiHeader
            .biSize = 40
            .biHeight = destHeight
            .biWidth = destWidth
            .biPlanes = 1
            .biBitCount = 24    ' see Special note above
        End With
        dDC = GetDC(0&)
        tDC = CreateCompatibleDC(dDC)
        If Not tDC = 0& Then
            hDIB = CreateDIBSection(dDC, BMPI, 0&, dstSA.pvData, 0&, 0&)
        End If
        ReleaseDC 0&, dDC
        If tDC = 0& Or hDIB = 0& Then Exit Function
        hOldBmp = SelectObject(tDC, hDIB)
        BitBlt tDC, 0&, 0&, destWidth, destHeight, destinationDC, destX, destY, vbSrcCopy
        SelectObject tDC, hOldBmp
        DeleteDC tDC
        iparseOverlayHost_Byte dstBytes, VarPtr(dstSA), 2, destHeight, iparseByteAlignOnWord(24, destWidth), dstSA.pvData ' overlay DMA array
    
    Else
        ' DIB class to DIB class blending
        iparseOverlayHost_Byte dstBytes, VarPtr(dstSA), 2, tHost.Height, tHost.scanWidth, tHost.BitsPointer ' overlay DMA array
        dstSA.rgSABound(1) = -(tHost.Height - destY - destHeight)
        dstSA.rgSABound(3) = -destX * 4
    End If
    
    If srcSA.pvData = m_Pointer Then ' using DIB as source
        iparseOverlayHost_Byte srcBytes, VarPtr(srcSA), 2, m_Height, m_Width * 4&, m_Pointer ' overlay DMA array
    Else                        ' using resized array as source
        iparseOverlayHost_Byte srcBytes, VarPtr(srcSA), 2, UBound(aResizedBytes, 2) + 1, UBound(aResizedBytes, 1) + 1, srcSA.pvData ' overlay DMA array
    End If
    
    On Error Resume Next ' expected errors? using a corrupted pre-multiplied image possibly, but will just draw wrong
    
    ' this loop is broken into several different loops to enhance speed when less options are used
    ' 1. separate loop when rendering DIB to DC or DIB to DIB
    ' 2. separate loop when using global alpha of 255 or using less than 255
    ' 3. separate loop when grayscaling and not grayscaling
    
    SrcX = SrcX * 4&                    ' first pixel to be procesed
    For srcRow = 0& To destHeight - 1&
    
        ' offset our DIB row as we go
        SrcY = SrcY - 1                 ' current source row being processed
        sX = SrcX                       ' 1st column of source row
        dY = destHeight - srcRow - 1    ' next row for destination image
        dX = 0&                         ' 1st column of destination row
        
        If tHost Is Nothing Then        ' DIB to DC rendering (4 pixel source to 3 pixel target)
            
            If grayScale = gsclNone Then
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             CopyMemory dstBytes(dX, dY), srcBytes(sX, SrcY), 3&
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                         End If
                        dX = dX + 3&
                        sX = sX + 4&
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                            ' following formula is for already pre-multiplied bytes
                            dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                            dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                            dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                            dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                         End If
                         dX = dX + 3&
                         sX = sX + 4&
                     Next
                End If
            Else    ' gray scaling
                Call iparseGrayScaleRatios(grayScale, Rg, Gg, Bg)
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             dstBytes(dX, dY) = (srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)
                             dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            ' calculate green,red channel grayscale value
                            gScaleByte = ((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)) And &HFF
                            ' blend grayscale to target
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + gScaleByte
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + gScaleByte
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + gScaleByte
                         End If
                        dX = dX + 3&
                        sX = sX + 4&
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                            ' following formula is for already pre-multiplied bytes
                            dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                            ' calculate green,red channel grayscale value
                            gScaleByte = (((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)) And &HFF) * GlobalAlpha
                            ' calculate green,red channels
                            dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + gScaleByte) \ &HFF
                            dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + gScaleByte) \ &HFF
                            dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + gScaleByte) \ &HFF
                         End If
                         dX = dX + 3&
                         sX = sX + 4&
                     Next
                End If
            End If
    
        Else    ' DIB class to DIB class rendering  (4 pixel source to 4 pixel target)
        
            If grayScale = gsclNone Then
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             CopyMemory dstBytes(dX, dY), srcBytes(sX, SrcY), 4&
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + srcBytes(sX, SrcY)
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + srcBytes(sX + 1&, SrcY)
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + srcBytes(sX + 2&, SrcY)
                            dstBytes(dX + 3&, dY) = (dstAlpha * dstBytes(dX + 3&, dY)) \ &HFF + srcBytes(sX + 3&, SrcY)
                         End If
                        dX = dX + 4&
                        sX = sX + 4&
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + (srcBytes(sX, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + (srcBytes(sX + 1&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + (srcBytes(sX + 2&, SrcY) * GlobalAlpha)) \ &HFF
                             dstBytes(dX + 3&, dY) = (dstAlpha * (dstBytes(dX + 3&, dY)) + (srcBytes(sX + 3, SrcY) * GlobalAlpha)) \ &HFF
                         End If
                         dX = dX + 4&
                         sX = sX + 4&
                     Next
                End If
            Else
                Call iparseGrayScaleRatios(grayScale, Rg, Gg, Bg)
                If GlobalAlpha = &HFF& Then
                    ' with full opaqueness, use separate loop, less calculations
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If (srcAlpha = &HFF&) Then
                             ' copy pixel to destination, adjusting for destination row/column as needed
                             dstBytes(dX, dY) = ((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg))
                             dstBytes(dX + 1&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 2&, dY) = dstBytes(dX, dY)
                             dstBytes(dX + 3&, dY) = &HFF
                         ElseIf Not srcAlpha = 0& Then
                             '-- Blend
                            dstAlpha = &HFF& - srcAlpha ' calculate dest alpha value
                            gScaleByte = ((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg) And &HFF)
                            dstBytes(dX, dY) = (dstAlpha * dstBytes(dX, dY)) \ &HFF + gScaleByte
                            dstBytes(dX + 1&, dY) = (dstAlpha * dstBytes(dX + 1&, dY)) \ &HFF + gScaleByte
                            dstBytes(dX + 2&, dY) = (dstAlpha * dstBytes(dX + 2&, dY)) \ &HFF + gScaleByte
                            dstBytes(dX + 3&, dY) = (dstAlpha * dstBytes(dX + 3&, dY)) \ &HFF + (srcBytes(sX + 3, SrcY))
                         End If
                        dX = dX + 4&
                        sX = sX + 4&
                     Next
                Else
                    ' global alpha and per-pixel blending
                    For srcCol = 0& To destWidth - 1&
                         srcAlpha = srcBytes(sX + 3&, SrcY)     ' get its alpha value
                         If Not srcAlpha = 0& Then
                             ' following formula is for already pre-multiplied bytes
                             dstAlpha = &HFF& - ((srcAlpha * GlobalAlpha) \ &HFF&)
                             gScaleByte = (((srcBytes(sX, SrcY) * Bg) + (srcBytes(sX + 1&, SrcY) * Gg) + (srcBytes(sX + 2&, SrcY) * Rg)) And &HFF) * GlobalAlpha
                             dstBytes(dX, dY) = (dstAlpha * (dstBytes(dX, dY)) + gScaleByte) \ &HFF
                             dstBytes(dX + 1&, dY) = (dstAlpha * (dstBytes(dX + 1&, dY)) + gScaleByte) \ &HFF
                             dstBytes(dX + 2&, dY) = (dstAlpha * (dstBytes(dX + 2&, dY)) + gScaleByte) \ &HFF
                             dstBytes(dX + 3&, dY) = (dstAlpha * (dstBytes(dX + 3&, dY)) + (srcBytes(sX + 3, SrcY) * GlobalAlpha)) \ &HFF
                         End If
                         dX = dX + 4&
                         sX = sX + 4&
                     Next
                End If
            End If
        End If
    Next
    ' remove overlay
    iparseOverlayHost_Byte srcBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    Erase aResizedBytes()
    If Err Then Err.Clear
    
    ' transfer results
    If tHost Is Nothing Then
        SetDIBitsToDevice destinationDC, destX, destY, destWidth, destHeight, 0&, 0&, 0&, destHeight, dstBytes(0, 0), BMPI, 0&
    End If
    iparseOverlayHost_Byte dstBytes, 0, 0, 0, 0, 0  ' remove DMA overlay
    If Not hDIB = 0& Then DeleteObject hDIB
    
    spt_Win9xBlend = True

End Function







































































