Attribute VB_Name = "mdlSdMain"
' .01 DAEB mdlmain.bas DAEB 24/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
' .02 DAEB mdlmain.bas DAEB 27/01/2021 Modified the menu text to incorporate the user-defined key and the hiding time
' .03 DAEB mdlMain.bas 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
' .04 DAEB 26/10/2020 mdlMain.bas dock DAEB removed declarations required by IsRunning since the move of this function to common.bas STARTS.
' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list STARTS
' .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
' .07 DAEB 19/04/2021 mdlMain.bas  added a new type link for determining shortcuts
' .08 DAEB 19/04/2021 mdlMain.bas Added new function to identify an icon to assign to the entry
' .09 DAEB 30/04/2021 mdlMain.bas deleteThisIcon created by extracting from the menu form so it can be used elsewhere
' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
' .11 DAEB 01/05/2021 mdlMain.bas load a transparent 128 x 128 image into the collection, used to highlight the position of a drag/drop
' .12 DAEB 11/05/2021 mdlMain.bas new bounceZone public variable to be loaded by dockSettings
' .13 DAEB 11/05/2021 mdlMain.bas renamed the old bounceCounter to bounceHeight
' .14 DAEB 11/05/2021 mdlMain.bas new bounceCounter now only records the count
' .15 DAEB 20/05/2021 mdlMain.bas Added new check box to allow a quick launch of the chosen app
' .16 DAEB 12/07/2021 mdlMain.bas Add the BounceZone as a configurable variable.

Option Explicit

'------------------------------------------------------------
' mdlMain
'
' Global variables, function and APIs that appear in just this program alone, as an included module mdlMain.bas.
'
'
'------------------------------------------------------------

' APIs and variables for querying processes START
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private uProcess As PROCESSENTRY32
Private hSnapshot As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long ' Alias "CreateToolhelp32Snapshot"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' APIs for querying processes END

'Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal filename As String, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
Public Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single) As Long
'Public Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal filename As Long, GpImage As Long) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClassName As String, ByVal lpszWindowName As String) As Long

' Private APIs for useful functions START
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long

' Public APIs for useful functions START

Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
' API to obtain correct screen width (to correct VB6 bug)
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Public Declare Function PrivateExtractIcons Lib "user32" _
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

' APIs for useful functions END

'Private APIs for GDI+
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As Any, grayMatrix As Any, ByVal flags As ColorMatrixFlags) As GpStatus
'Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Image As Long, hBmp As Long, ByVal BGColor As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Context As Long, ByVal image As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Context As Long, ByVal PixOffsetMode As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal img As Long, Context As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal dx As Long, ByVal dy As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal pScanData As Long, image As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus
' APIs image cropping
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal image As Long, ByRef PixelFormat As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus

'Public APIs for GDI+

Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As Long, image As Long) As Long
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As GDIPLUS_FONTSTYLE, ByVal Unit As GDIPLUS_UNIT, createdfont As Long) As Long
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As String, ByVal fontCollection As Long, fontFamily As Long) As Long
Public Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hdc As Long, GpGraphics As Long) As Long
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As Long
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal Graphics As Long) As Long
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Public Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal image As Long) As Long
Public Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Public Declare Function GdipDrawImageRectRect Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal callbackData As Long) As Long
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal Str As String, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As Long
Public Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal image As Long, Height As Long) As Long
Public Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal image As Long, Width As Long) As Long
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Public Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, gdipInput As GDIPLUS_STARTINPUT, GdiplusStartupOutput As Long) As Long
Public Declare Function GdipReleaseDC Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal hdc As Long) As Long
Public Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Public Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal SmoothingMode As Long) As Long
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As GDIPLUS_ALIGNMENT) As Long

'.nn

 
'Private Declare Function GdipSetCompositingQuality Lib "gdiplus" _
'    (ByVal graphics As Long, ByVal CompositingQuality As _
'    CompositingQualityMode) As GpStatus
'
'    Public Const QualityModeInvalid As Long = -1&
' Public Const QualityModeDefault As Long = 0&
' Public Const QualityModeLow As Long = 1&
' Public Const QualityModeHigh As Long = 2&
'
'    Private Enum CompositingQualityMode
'    CompositingQualityInvalid = QualityModeInvalid
'    CompositingQualityDefault = QualityModeDefault
'    CompositingQualityHighSpeed = QualityModeLow
'    CompositingQualityHighQuality = QualityModeHigh
'    CompositingQualityGammaCorrected = QualityModeHigh + 1
'    CompositingQualityAssumeLinear = QualityModeHigh + 2
'End Enum

'.nn
    
' APIs for GDI+ END

' Private APIs and vars for enumerating running windows START
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function IsTopWIndow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' Private APIs and vars for enumerating running windows END

' Public APIs and vars for enumerating running windows START
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

' .38 DAEB 18/03/2021 frmMain.frm utilised SetActiveWindow to give window focus without bringing it to fore
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
    
' .25 DAEB frmMain.bas 10/02/2021 added API and vars to test to see if a window is zoomed
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Integer) As Boolean
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
' .39 DAEB 18/03/2021 frmMain.frm utilised BringWindowToTop instead of SetWindowPos & HWND_TOP as that was used by a C program that worked perfectly.
Public Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long

'APIs and vars for enumerating running windows ENDS

Public Const DI_NORMAL = 3
Public Const LR_LOADFROMFILE As Long = &H10
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE As Long = 6 ' .25 DAEB frmMain.bas 10/02/2021 added API and vars to test to see if a window is zoomed

Public Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Private Const MAX_PATH = 260

' APIs and vars for enumerating running windows STARTS
Private Const GW_OWNER = 4
Private Const WS_EX_TOOLWINDOW = &H80&
'Private Const WS_EX_TOOLWINDOW = &H80&
Private Const WS_EX_APPWINDOW = &H40000
Private Const GW_HWNDNEXT = 2
Private Const GA_ROOT = 2&
' APIs and vars for enumerating running windows END

' global Windows+ constants START
Public Const ULW_ALPHA = &H2
Public Const DIB_RGB_COLORS As Long = 0
Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE As Long = -20
Public Const HWND_TOP As Long = 0
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_BOTTOM As Long = 1
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE  As Long = &H2
Public Const SWP_NOZORDER  As Long = &H4
Public Const SWP_HIDEWINDOW  As Long = &H80
Public Const SWP_ACTIVATE  As Long = &H10
Public Const SWP_NOACTIVATE  As Long = &H20
Public Const SWP_SHOWWINDOW  As Long = &H40
Public Const SWP_NOOWNERZORDER  As Long = &H200 ' .40 DAEB 18/03/2021 frmMain.frm Added SWP_NOOWNERZORDER as an additional flag as that was used by a C program that worked perfectly, fixing the z-order position problems

Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Public Const OUT_DEFAULT_PRECIS = 0
' Windows+ constants END

Public Const PixelFormat32bppPARGB = &HE200B
Public Const PixelFormat32bppARGB = &H26200A

' global GDI+ Types START
Public Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type GDIPLUS_STARTINPUT
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmpHeader As BITMAPINFOHEADER
    bmpColors As RGBQUAD
End Type

Public Type CLSID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Public Type ImageCodecInfo
   ClassID As CLSID
   FormatID As CLSID
   CodecName As Long      ' String Pointer; const WCHAR*
   DllName As Long        ' String Pointer; const WCHAR*
   FormatDescription As Long ' String Pointer; const WCHAR*
   FilenameExtension As Long ' String Pointer; const WCHAR*
   MimeType As Long       ' String Pointer; const WCHAR*
   flags As ImageCodecFlags   ' Should be a Long equivalent
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long      ' Byte Array Pointer; BYTE*
   SigMask As Long         ' Byte Array Pointer; BYTE*
End Type
' global GDI+ Types END

' .07 DAEB 19/04/2021 mdlMain.bas  added a new type link for determining shortcuts
Public Type Link
    Attributes As Long
    Filename As String
    Description As String
    RelPath As String
    WorkingDir As String
    Arguments As String
    CustomIcon As String
End Type


Public Type PictDesc
    cbSizeofStruct  As Long
    PicType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type


Private Type GpMatrix
    matrix(5) As Single
End Type
 
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type
' vars for GDI+ colour matrix ENDS


' global GDI+ Enums START
Public Enum GDIPLUS_ALIGNMENT
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Public Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Enum GDIPLUS_UNIT
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

'Public Enum TASKBAR_POSITION
'    vbbottom
'    vbLeft
'    vbRight
'    vbtop
'End Enum

    
Public Enum TASKBAR_POSITION
    vbtop
    vbBottom
    vbLeft
    vbRight
End Enum
    

' NOTE: Enums evaluate to a Long
Public Enum GpStatus   ' aka Status
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

' Information flags about image codecs
Public Enum ImageCodecFlags
   ImageCodecFlagsEncoder = &H1
   ImageCodecFlagsDecoder = &H2
   ImageCodecFlagsSupportBitmap = &H4
   ImageCodecFlagsSupportVector = &H8
   ImageCodecFlagsSeekableEncode = &H10
   ImageCodecFlagsBlockingDecode = &H20

   ImageCodecFlagsBuiltin = &H10000
   ImageCodecFlagsSystem = &H20000
   ImageCodecFlagsUser = &H40000
End Enum

Public Enum eInterpolationMode
  ipmDefault = &H0
  ipmLow = &H1
  ipmHigh = &H2
  ipmBilinear = &H3
  ipmBicubic = &H4
  ipmNearestNeighbor = &H5
  ipmHighQualityBilinear = &H6
  ipmHighQualityBicubic = &H7
End Enum

'Public Enum WrapMode
'   WrapModeTile = &H8
'   WrapModeTileFlipX = &H9
'   WrapModeTileFlipY = &H10
'   WrapModeTileFlipXY = &H11
'   WrapModeClamp = &H12
'End Enum

Public Enum SmoothingModeEnum
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = 4&
    SmoothingModeAntiAlias8x8 = 5&
End Enum

' vars for GDI+ colour matrix STARTS
Private Enum ColorAdjustType
   ColorAdjustTypeDefault
   ColorAdjustTypeBitmap
   ColorAdjustTypeBrush
   ColorAdjustTypePen
   ColorAdjustTypeText
   ColorAdjustTypeCount
   ColorAdjustTypeAny
End Enum
 
Private Enum ColorMatrixFlags
   ColorMatrixFlagsDefault = 0
   ColorMatrixFlagsSkipGrays = 1
   ColorMatrixFlagsAltGray = 2
End Enum

' global GDI+ Enums END


Public gdipInit As GDIPLUS_STARTINPUT
Public rctText As RECTF


' GDI+ globals variables START
Public iconBitmap As Long
Public gdipFullScreenBitmap As Long
Public lngGDI As Long
Public lngReturn As Long
Public dockPosition As TASKBAR_POSITION
' GDI+ globals variables END


'vars for the mouse position
Public apiWindow As POINTAPI
Public apiPoint As POINTAPI
Public apiMouse As POINTAPI
Public newPoint As POINTAPI

Public funcBlend32bpp As BLENDFUNCTION
Public bmpInfo As BITMAPINFO


' collection objects
Private collTemporaryIcons As Object
Public collLargeIcons As Object
Public collSmallIcons As Object

Public dcMemory As Long
Public bmpMemory As Long

'vars for the animation
Public iconSizeLargePxls As Byte
Public iconSizeSmallPxls As Byte


' Steamydock global configuration variables START
Public rdThemeSkinFile As String
Public rdThemeSeparatorFile As String
Public validTheme As Boolean
Public animatedIconsRaised As Boolean
Public selectedIconIndex As Integer
Public prevIconIndex As Integer
Public bounceHeight As Integer ' .13 DAEB 11/05/2021 mdlMain.bas renamed the old bounceCounter to bounceHeight
Public bounceCounter As Integer ' .14 DAEB 11/05/2021 mdlMain.bas new bounceCounter now only records the count
Public inc As Boolean
Public bounceTimerRun As Integer
Public fcount As Integer

Public processCheckArray() As Boolean
Public explorerCheckArray() As Boolean

Public fileNameArray() As String
Public initiatedProcessArray() As String
Public initiatedExplorerArray() As String

Public dictionaryLocationArray() As String ' array to store the location (index) for the images in the dictionary collection
Public namesListArray() As String ' the name assigned to each icon
Public sCommandArray() As String ' the command assigned to each icon, used for quick access
Public targetExistsArray() As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
Public disabledArray() As Integer

Public WindowsVer As String
Public rdIconMaximum As Integer
Public theCount As Integer
Public dockOpacity As Integer
Public userLevel As String

Public autoFadeOutTimerCount As Integer
Public autoFadeInTimerCount As Integer ' .01 mdlmain.bas DAEB 24/01/2021 Added new parameter autoFadeInTimerCount for the new fade in timer
Public autoSlideInTimerCount As Integer ' .nn DAEB 03/03/2021 new separate timer for the slide in feature
Public autoSlideOutTimerCount As Integer ' .nn DAEB 03/03/2021 new separate timer for the slide out feature
Public autoHideRevealTimerCount As Integer
Public animationFlg As Boolean
Public dockLoweredTime As Date
Public dockHidden As Boolean
Public debugflg As Integer
Public readEmbeddedIcons As Boolean
Public dragToDockOperating As Boolean
Public dragFromDockOperating As Boolean  '.nn DAEB 30/04/2021 frmMain.frm Added a response to a drag and drop operating from the dock
Public dragInsideDockOperating As Boolean '.nn '.nn new check for dragInsideDockOperating
Public dragImageToDisplay As String
Public hideDockForNMinutes As Boolean
Public forceRunNewAppFlag As Boolean

Public insideDockFlg As Boolean '.nn Added to allow a MouseUp to capture a drag from one part of the dock to another

' vars to obtain correct screen width (to correct VB6 bug)

Public screenWidthPixels As Long
Public screenHeightPixels As Long

Public oldScreenWidthPixels As Long
Public oldScreenHeightPixels As Long

' vars to store the position of each icon

Public iconStoreLeftPixels() As Double ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
Public iconStoreRightPixels() As Double ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'Public iconStoreTopPixels() As Double ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
Public iconStoreBottomPixels() As Double ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon

' Left  Right
' +-----+ Top
' |-----|
' |-----|
' +-----+ Bottom

' using iconStoreXxxxPixels we can derive all the rectangle's co-ordinates

Public iconArrayUpperBound As Single

' array used to store the location of the dictionary image for each icon
Public dictionaryLocationArrayUpperBound As Single

Public iconWidthPxls As Single

'APIs and constants to read embedded icons. Most of these APIs are not really used yet, they are there just in case anyone wants to
'complete the reading of embedded icons from binaries/DLLs


Public bounceZone As Integer ' .16 DAEB 12/07/2021 mdlMain.bas Add the BounceZone as a configurable variable.
Public smallDockBeenDrawn As Boolean

Public hdcScreen As Long

Public lHotKey As Long




'---------------------------------------------------------------------------------------
' Steamydock global configuration variables END
'---------------------------------------------------------------------------------------



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

Public Function GetEncoderClsid(strMimeType As String, ClassID As CLSID)
   Dim num As Long
   Dim Size As Long
   Dim i As Long

   Dim ICI() As ImageCodecInfo
   Dim Buffer() As Byte
   
   On Error GoTo GetEncoderClsid_Error

   GetEncoderClsid = -1 'Failure flag

   ' Get the encoder array size
   Call GdipGetImageEncodersSize(num, Size)
   If Size = 0 Then Exit Function ' Failed!

   ' Allocate room for the arrays dynamically
   ReDim ICI(1 To num) As ImageCodecInfo
   ReDim Buffer(1 To Size) As Byte

   ' Get the array and string data
   Call GdipGetImageEncoders(num, Size, Buffer(1))
   ' Copy the class headers
   Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * num))

   ' Loop through all the codecs
   For i = 1 To num
      ' Must convert the pointer into a usable string
      If StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0 Then
         ClassID = ICI(i).ClassID   ' Save the class id
         GetEncoderClsid = i        ' return the index number for success
         Exit For
      End If
   Next
   ' Free the memory
   Erase ICI
   Erase Buffer

   On Error GoTo 0
   Exit Function

GetEncoderClsid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetEncoderClsid of Module mdlMain"
End Function



'---------------------------------------------------------------------------------------
' Procedure : PtrToStrW
' Author    : From www.mvps.org/vbnet...i think
' Date      : 21/08/2020
' Purpose   : '   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
'---------------------------------------------------------------------------------------
'
Public Function PtrToStrW(ByVal lpsz As Long) As String
    Dim sOut As String
    Dim lLen As Long

   On Error GoTo PtrToStrW_Error

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If

   On Error GoTo 0
   Exit Function

PtrToStrW_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PtrToStrW of Module mdlMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Convert_Dec2RGB
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Convert_Dec2RGB(ByVal myDECIMAL As Long) As String
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




'---------------------------------------------------------------------------------------
' Procedure : Color_RGBtoARGB
' Author    : lavolpe
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Color_RGBtoARGB(ByVal RGBColor As Long, ByVal opacity As Long) As Long

    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

   On Error GoTo Color_RGBtoARGB_Error

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    Color_RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    If opacity < 128 Then
        If opacity < 0& Then opacity = 0&
        Color_RGBtoARGB = Color_RGBtoARGB Or opacity * &H1000000
    Else
        If opacity > 255& Then opacity = 255&
        Color_RGBtoARGB = Color_RGBtoARGB Or ((opacity And Not &H80) * &H1000000) Or &H80000000
    End If
    

   On Error GoTo 0
   Exit Function

Color_RGBtoARGB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Color_RGBtoARGB of Module mdlMain"
    
End Function




'---------------------------------------------------------------------------------------
' Procedure : ColorToGDIplus
' Author    : lavolpe
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ColorToGDIplus(ByVal lColor As Long, Optional ByVal AlphaByte As Byte = 255) As Long

    ' helper function to convert RGB to BGR for GDI+ usage
   On Error GoTo ColorToGDIplus_Error

    If lColor < 0& Then lColor = GetSysColor(lColor And &HFF)
    ColorToGDIplus = (lColor And &HFF00&) _
                    Or (lColor And &HFF) * &H10000 _
                    Or (lColor And &HFF0000) \ &H10000 _
                    Or (AlphaByte And &H7F) * &H1000000
    If (AlphaByte And &H80) Then ColorToGDIplus = ColorToGDIplus Or &H80000000

   On Error GoTo 0
   Exit Function

ColorToGDIplus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ColorToGDIplus of Module mdlMain"

End Function

'---------------------------------------------------------------------------------------
' Procedure : checkDockProcessesRunning
' Author    : beededea
' Date      : 08/07/2020
' Purpose   : it used to test all icon binaries against all processes in the task list
'             now it just populates an array that lets other programs test to see if a program is running
'             this routine is used to identify an item in the dock as currently running even if not triggered by the dock
'---------------------------------------------------------------------------------------
'
Public Sub checkDockProcessesRunning()
    
    On Error GoTo checkDockProcessesRunning_Error
    
    Dim useloop As Integer: useloop = 0
        
    For useloop = 0 To rdIconMaximum
        ' instead of looping through all elements in the docksettings.ini file, we now store all the current commands in an sCommandArray
        ' we loop through the array much quicker than looping through the temporary settings file and extracting the commands from each
        ' we must remember to populate the array whenever an icon is added or deleted
        If sCommandArray(useloop) <> "" Then
            processCheckArray(useloop) = IsRunning(sCommandArray(useloop), vbNull)
        End If

    Next useloop
   On Error GoTo 0
   Exit Sub

checkDockProcessesRunning_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkDockProcessesRunning of Module mdlMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : checkExplorerRunning
' Author    : beededea
' Date      : 08/07/2020
' Purpose   : it used to test all currently running explorer windows against all icon's target as stored
'             in the sCommandArray. This routine is used to identify an Explorer Window in the dock as currently
'             being open even if not triggered by the dock.
'---------------------------------------------------------------------------------------
'
Public Sub checkExplorerRunning()
    
    On Error GoTo checkExplorerRunning_Error

    Dim windowCount As Integer: windowCount = 0
    Dim openExplorerPathArray() As String
    Dim windowLoop As Integer: windowLoop = 0
    Dim sCommandLoop As Integer: sCommandLoop = 0
    Dim NameProcess As String: NameProcess = vbNullString
    
    Call enumerateExplorerWindows(openExplorerPathArray(), windowCount)
    
    If windowCount = 0 Then Call enumerateExplorerWindows(openExplorerPathArray(), windowCount)
    
    ' instead of looping through all elements in the docksettings.ini file, we now store all the current commands in an sCommandArray
    ' we loop through the array much quicker than looping through the temporary settings file and extracting the commands from each.
    
    For windowLoop = 0 To windowCount - 1
        For sCommandLoop = 0 To rdIconMaximum
            If sCommandArray(sCommandLoop) <> vbNullString Then
                If LCase$(sCommandArray(sCommandLoop)) = LCase$(openExplorerPathArray(windowLoop)) Then
                    explorerCheckArray(sCommandLoop) = True
                End If
            End If
        Next sCommandLoop
    Next windowLoop
        
   On Error GoTo 0
   Exit Sub

checkExplorerRunning_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkExplorerRunning of Module mdlMain"

End Sub


''---------------------------------------------------------------------------------------
'' Procedure : getWindowHWndForPid
'' Author    : dee-u Candon City, Ilocos https://www.vbforums.com/showthread.php?561413-getting-hwnd-from-process
'' Date      : 03/09/2020
'' Purpose   : Return the window handle for a PID.
''---------------------------------------------------------------------------------------
''
'Public Function getWindowHWndForPid(ByVal PID As Long) As Long
'
'    Dim lHwnd       As Long
'    Dim test_pid    As Long
'    Dim Thread_ID   As Long
'    Dim lExStyle    As Long
'    Dim bNoOwner    As Boolean
'
'    ' Get the first window handle.
'    On Error GoTo getWindowHWndForPid_Error
'
'    lHwnd = FindWindow(vbNullString, vbNullString)
'    ' Loop until we find the target or we run out
'    ' of windows. Much easier than using the enumerateWindows function
'    Do While lHwnd <> 0
'        ' check if window is visible or not - not a good test as some windows are top but still hidden such as GPU-z that minimises to the systray.
'        If IsWindowVisible(lHwnd) Then
'
'            ' This is a top-level window. See if
'            ' it has the target instance handle.
'            Thread_ID = GetWindowThreadProcessId(lHwnd, test_pid)
'
'            If test_pid = PID Then
'                ' See if this window has a parent. If not,
'                ' it is a top-level window.
'                If GetParent(lHwnd) = 0 Then
'
'                    bNoOwner = (GetWindow(lHwnd, GW_OWNER) = 0)
'                    'get current window style of a window
'                    lExStyle = GetWindowLong(lHwnd, GWL_EXSTYLE) '33554568
'
'                    If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
'                        ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
'                            'MsgBox "WS_EX_APPWINDOW " & WS_EX_APPWINDOW & " lExStyle " & lExStyle & " bNoOwner " & bNoOwner
'                            lHwnd = GetAncestor(lHwnd, GA_ROOT)
'                            getWindowHWndForPid = lHwnd
'                            Exit Function
'                    End If
'
'
'                End If
'            End If
'       End If
'
'        ' Examine the next window.
'        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
'    Loop
'
'   On Error GoTo 0
'   Exit Function
'
'getWindowHWndForPid_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getWindowHWndForPid of Module mdlMain"
'End Function

' .03 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
'---------------------------------------------------------------------------------------
' Procedure : fEnumWindowsCallBack
' Author    : beededea
' Date      : 08/02/2021
' Purpose   : call back routine that returns to fEnumWindows
'
' This callback function is called by Windows itself (using the EnumWindows API call) for EVERY window that
' exists even a limited list of Windows that exist in the systray
'
' Windows to display are those that:
'   -   do not belong to this app
'   -   are visible or invisible and in the systray list
'   -   do not have a parent
'   -   have no owner and are not Tool windows OR
'       have an owner and are App windows
'
'---------------------------------------------------------------------------------------
'
Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String
Dim test_pid    As Long
Dim Thread_ID   As Long
Dim pid   As Long

' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list STARTS
' .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
'Dim appSystray() As String
'Dim i As Integer
'appSystray = Split(appSystrayTypes, "|")
' .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list ENDS

On Error GoTo fEnumWindowsCallBack_Error

pid = lParam

If hwnd <> dock.hwnd Then
        ' check if window is visible or not
        If IsWindowVisible(hwnd) Then
            ' This is a top-level window. See if it has the target instance handle.
            ' test_pid is the process ID returned for the window handle
            
            ' GetWindowThreadProcessId finds the process ID given for the thread which owns the window
            Thread_ID = GetWindowThreadProcessId(hwnd, test_pid)
                      
            If test_pid = pid Then
                If GetParent(hwnd) = 0 Then
                    bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
                    lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)

                        If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                            ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
        
                                hwnd = GetAncestor(hwnd, GA_ROOT)
        
                                storeWindowHwnd = hwnd ' a bit of a kludge, a global var that carries the window handle to the calling function
                                Exit Function
                        End If
                End If
            End If
            ' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list STARTS
'           .06 DAEB 03/03/2021 mdlMain.bas  removed the appSystrayTypes feature, no longer needed to access the systray apps
'        Else ' not IsWindowVisible(hwnd)
'
'            ' Some windows are top level but not visible, such as GPU-z that minimise to the systray.
'            ' this section is for these types of apps.
'            ' The trouble is that there are a lot of invisible windows for each process that we don't want to bring to the fore
'            ' we cannot currently identify a process in the systray, so we have a kludge that is a temporary list of
'            ' apps that can minimise to the systray. We use the program's captions to compare.
'
'            'GetWindowThreadProcessId finds the process ID given for the thread which owns the window
'            Thread_ID = GetWindowThreadProcessId(hwnd, test_pid)
'
'            If test_pid = pid Then
'                If GetParent(hwnd) = 0 Then
'                    bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
'                    lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
'
'
'                    sWindowText = Space$(256) ' pad the string to 256 chars
'                    lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText)) ' obtain the caption
'
'                    For i = 0 To UBound(appSystray) ' search through all the potential systray apps in the manually populated array
'                        If InStr(sWindowText, appSystray(i)) Then
'
'                            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
'                                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
'
'                                    hwnd = GetAncestor(hwnd, GA_ROOT)
'
'                                    storeWindowHwnd = hwnd ' a bit of a kludge, a global var that carries the window handle to the calling function
'                                    Exit Function
'                            End If
'                        End If
'                    Next
'                End If
'            End If
        End If
        ' .05 DAEB mdlMain.bas 10/02/2021 changes to handle invisible windows that exist in the known apps systray list ENDS
End If


fEnumWindowsCallBack = True

   On Error GoTo 0
   Exit Function

fEnumWindowsCallBack_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fEnumWindowsCallBack of Module mdlMain"
End Function


' .03 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
'---------------------------------------------------------------------------------------
' Procedure : fEnumWindows
' Author    : http://www.thescarms.com
' Date      : 08/02/2021
' Purpose   : enumerates all top-level windows
'---------------------------------------------------------------------------------------
'
Public Function fEnumWindows(ByVal processID As Long) As Long
    'Dim retVal As Long
'
' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
    ' the process id is passed as the 2nd param lpdata but the return value is passed back as a global variable
   On Error GoTo fEnumWindows_Error

    fEnumWindows = EnumWindows(AddressOf fEnumWindowsCallBack, processID)

   On Error GoTo 0
   Exit Function

fEnumWindows_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fEnumWindows of Module mdlMain"

End Function




'---------------------------------------------------------------------------------------
' Procedure : readDockConfiguration
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : Read the dock configuration file from one of three potential sources
'---------------------------------------------------------------------------------------
'
Public Sub readDockConfiguration()
    Dim useloop As Integer
    
    On Error GoTo readDockConfiguration_Error
    If debugflg = 1 Then debugLog "%" & " sub readDockConfiguration"
            
    'the RD settings.ini configuration option
    'origSettingsFile = rdAppPath & "\settings.ini" ' Rocketdock 's settings file
    'rdSettingsFile = App.Path & "\rdSettings.ini" ' a copy of the settings file that we work on
    
    rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", dockSettingsFile)
    rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", dockSettingsFile)
    
    'final check to be sure that we aren't using an incorrectly set dockSettings.ini file when RD has never been installed
'    If rocketDockInstalled = False And RDregistryPresent = False Then
         rDGeneralReadConfig = "True"
'    End If
    
    'the 3rd new configuration option
'    If rDGeneralReadConfig = "True" Then

        ' read the rocketdock settings.ini and find the very last icon
        theCount = Val(GetINISetting("Software\SteamyDock\IconSettings\Icons", "count", dockSettingsFile))
        'theCount = 72 ' debug
        rdIconMaximum = theCount - 1

        ' read the Rocketdock dock settings from the new configuration file
        Call readDockSettingsFile("Software\SteamyDock\DockSettings", dockSettingsFile)
        Call validateInputs
        Call adjustControls
        
        'assign some variables values according to those validated inputs
        dock.animateTimer.Interval = Val(rDAnimationInterval)
        
        ' copy the original ICON configs out of the dockSettingsFile and into the temporary settings file that we will operate upon
        For useloop = 0 To rdIconMaximum
             ' get the relevant icon entries from the 3rd config
             readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
             ' note we are copying from the dock settings as "Software\SteamyDock\DockSettings" and into the temporary settings file as "software\rocketdock"
             'Call writeIconSettingsIni("Software\RocketDock\Icons", useloop, rdSettingsFile) ' the alternative settings.ini exists
        Next useloop
        


         
'    Else
'
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'
'            ' copy the original settings file to a duplicate that we will operate upon
'            FileCopy origSettingsFile, rdSettingsFile
'
'            ' read the rocketdock settings.ini and find the very last icon
'            theCount = Val(GetINISetting("Software\RocketDock\Icons", "count", rdSettingsFile))
'            rdIconMaximum = theCount - 1
'
'            ' we only need to read the dock settings from the temporary settings file
'            Call readDockSettingsFile("Software\RocketDock", rdSettingsFile)
'            Call validateInputs
'            Call adjustControls
            
            'the icon settings do not need to be read now as the dock takes its icon config straight from the file we copied above

'        Else
'            'the RD registry configuration option
'
'            ' read the rocketdock ICON registry entry and find the last icon
'            theCount = Val(getstring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count"))
'            rdIconMaximum = theCount - 1
'
'            ' read the DOCK configuration from the registry
'            Call readRegistry ' this does the reading and the validation
'            Call adjustControls
'
'            ' copy the original ICON configs out of the registry and into a settings file that we will operate upon
'            readIconRegistryWriteSettings rdSettingsFile
'
'        End If
'    End If
            
        
        
   On Error GoTo 0
   Exit Sub

readDockConfiguration_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readDockConfiguration of module mdlMain.bas"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : adjustControls
' Author    : beededea
' Date      : 17/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub adjustControls()

    On Error GoTo adjustControls_Error

    Dim MyPath  As String: MyPath = vbNullString
    Dim themePresent As Boolean: themePresent = False
    Dim myName As String: myName = vbNullString
    Dim toggleText As String: toggleText = vbNullString

    MyPath = sdAppPath & "\Skins\" '"E:\Program Files (x86)\RocketDock\Skins\"

    If Not fDirExists(MyPath) Then
        MsgBox "WARNING - The skins folder is not present in the correct location " & sdAppPath
    End If

    myName = Dir$(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(MyPath & myName) And vbDirectory) = vbDirectory Then
             'debugLog MyName   ' Display entry only if it
          End If   ' it represents a directory.
       End If
       myName = Dir$   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> vbNullString Then
        If myName = rDtheme Then themePresent = True
       End If
    Loop

    ' if the theme is not in the list then make it none to ensure no corruption *1
    If themePresent = False Then rDtheme = "Blank"
    
    ' .02 mdlmain.bas STARTS DAEB 27/01/2021 Modified the menu text to incorporate the user-defined key and the hiding time
    If Val(sDContinuousHide) = 1 Then
        toggleText = "Hide for the next minute "
    Else
        toggleText = "Hide for the next " & sDContinuousHide & " minutes "
    End If
    
    If rDHotKeyToggle <> "Disabled" Then toggleText = toggleText & "(" & rDHotKeyToggle & " to restore)"
    menuForm.mnuHideTwenty.Caption = toggleText
    ' .02 mdlmain.bas ENDS DAEB 27/01/2021 Modified the menu text to incorporate the user-defined key and the hiding time
    
    If rDLockIcons = 1 Then
        menuForm.mnuLockIcons.Checked = True
        menuForm.mnuDeleteIcon.Enabled = False
    Else
        menuForm.mnuLockIcons.Checked = False
        menuForm.mnuDeleteIcon.Enabled = True
    End If
    menuForm.mnuTop.Checked = False
    menuForm.mnuBottom.Checked = False
    menuForm.mnuLeft.Checked = False
    menuForm.mnuRight.Checked = False

    If rDSide = vbtop Then
        menuForm.mnuTop.Checked = True
        dockPosition = vbtop
    End If
    If rDSide = vbBottom Then
        menuForm.mnuBottom.Checked = True
        dockPosition = vbBottom
    End If
    If rDSide = vbLeft Then
        menuForm.mnuLeft.Checked = True
        dockPosition = vbLeft
    End If
    If rDSide = vbRight Then
        menuForm.mnuRight.Checked = True
        dockPosition = vbRight
    End If


    menuForm.mnuAutoHide.Checked = False
    If rDAutoHide = "1" Then
        menuForm.mnuAutoHide.Checked = True
        dock.autoHideChecker.Enabled = True
    End If
    
    If sDDefaultEditor <> vbNullString Then menuForm.mnuEditWidget.Caption = "Edit Program using " & sDDefaultEditor
   
    If sDDebugFlg = "1" Then
        menuForm.mnuDebug.Caption = "Turn Debugging OFF"
        menuForm.mnuAppFolder.Visible = True
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuDebug.Caption = "Turn Debugging ON"
        menuForm.mnuAppFolder.Visible = False
        menuForm.mnuEditWidget.Visible = False
    End If
    
   On Error GoTo 0
   Exit Sub

adjustControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustControls of Module Module1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : testDockRunning
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : if the process already exists then kill it
'---------------------------------------------------------------------------------------
'
Public Sub testDockRunning()
    On Error GoTo testDockRunning_Error
    
    Dim NameProcess As String: NameProcess = vbNullString
    Dim AppExists As Boolean: AppExists = False
   
    If debugflg = 1 Then debugLog "% sub testDockRunning"

    AppExists = App.PrevInstance
    If AppExists = True Then
        NameProcess = "steamydock.exe"
        checkAndKillPutWindowBehind NameProcess, False, True
    End If

   On Error GoTo 0
   Exit Sub

testDockRunning_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testDockRunning of Form dock"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : insertNewIconDataIntoCurrentPosition
' Author    : beededea
' Date      : 18/08/2019
' Purpose   : Add something to the dock called by all the menuAdd functions that follow

    ' instead of reordering the images within the dictionary (which is difficult as you can't just add and
    ' replace objects into an existing collection, also, it does not release memory) instead we simply
    ' add a new icon reference to the settings file, add a new image to the end of the collection and then we
    ' manipulate an array of index numbers indicating which image in the collection to use. This persists until the
    ' dock is restarted. Then the images are loaded and numbered sequentially.
    
'---------------------------------------------------------------------------------------
'.nn Modified the function inputs to add the missing icon characteristics that are needed when dragging and dropping an icon within the dock.
' .15 DAEB 20/05/2021 mdlMain.bas Added new check box to allow a quick launch of the chosen app

Public Sub insertNewIconDataIntoCurrentPosition(ByVal thisFilename As String, ByVal thisTitle As String, _
    ByVal thisCommand As String, _
    ByVal thisArguments As String, ByVal thisWorkingDirectory As String, _
    ByVal thisShowCmd As String, ByVal thisOpenRunning As String, _
    ByVal thisSeparator As String, ByVal thisDockletFile As String, _
    ByVal thisUseContext As String, ByVal thisUseDialog As String, _
    ByVal thisUseDialogAfter, ByVal thisQuickLaunch, ByVal thisDisabled As String)
    
    Dim useloop As Integer
    Dim thisIcon As Integer

    On Error GoTo insertNewIconDataIntoCurrentPosition_Error
    'If debugflg = 1 Then debugLog "%" & "insertNewIconDataIntoCurrentPosition"

    ' starting at the end of the steamydock map, scroll backward and increment the number
    ' until we reach the current position.
    
    For useloop = rdIconMaximum To selectedIconIndex Step -1
         Call zeroAllIconCharacteristics
         
         readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
        
         Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop + 1, dockSettingsFile)
    
    Next useloop
    
    'increment the new icon count
    theCount = theCount + 1
    rdIconMaximum = rdIconMaximum + 1 '
    iconArrayUpperBound = rdIconMaximum
    
    'amend the count in the alternative rdSettings.ini
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, dockSettingsFile
    
    'resize all arrays used for storing icon information
    ReDim Preserve fileNameArray(rdIconMaximum) As String ' the file location of the original icons
    ReDim Preserve dictionaryLocationArray(rdIconMaximum) As String ' the dictionary location of the original icons
    ReDim Preserve namesListArray(rdIconMaximum) As String ' the name assigned to each icon
    ReDim Preserve sCommandArray(rdIconMaximum) As String ' the command assigned to each icon
    ReDim Preserve targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
    ReDim Preserve processCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    ReDim Preserve explorerCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    
    ReDim Preserve initiatedProcessArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
    ReDim Preserve initiatedExplorerArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track explorer processes initiated from the dock
    
    ReDim Preserve disabledArray(rdIconMaximum) As Integer '
   
   ' dynamically extend the number of picture boxes by one
    
    thisIcon = useloop + 1
    
    Call zeroAllIconCharacteristics
    
    'when we arrive at the original position then add a blank item
    ' with the following blank characteristics
    sFilename = thisFilename ' the default Rocketdock filename for a blank item
    
    sTitle = thisTitle
    sCommand = thisCommand
    sArguments = thisArguments
    sWorkingDirectory = thisWorkingDirectory
    sDockletFile = thisDockletFile
    sIsSeparator = thisSeparator
    
    sShowCmd = thisShowCmd
    sOpenRunning = thisOpenRunning
    sUseContext = thisUseContext
    
    sUseDialog = thisUseDialog
    sUseDialogAfter = thisUseDialogAfter
    sQuickLaunch = thisQuickLaunch ' .15 DAEB 20/05/2021 mdlMain.bas Added new check box to allow a quick launch of the chosen app
    
    sDisabled = thisDisabled
            
    Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", thisIcon, dockSettingsFile)

    ' then re-read the config for every icon
    For useloop = rdIconMaximum To selectedIconIndex Step -1
        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
        ' read the two main icon variables into arrays, one for each
        fileNameArray(useloop) = sFilename
        namesListArray(useloop) = sTitle
        sCommandArray(useloop) = sCommand
        targetExistsArray(useloop) = 0
        
        ' check to see if each process is running and store the result away
        explorerCheckArray(useloop) = isExplorerRunning(sCommand)
        processCheckArray(useloop) = IsRunning(sCommand, vbNull)

        If sDisabled = "1" Then
            disabledArray(useloop) = 1
        Else
            disabledArray(useloop) = 0
        End If
    
    Next useloop

    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, dockSettingsFile
    
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", selectedIconIndex, dockSettingsFile
    
    
    '.nn new check for dragInsideDockOperating
    If dragInsideDockOperating = False Then '.nn for performance reason, disabled when dragging and dropping as it is carried out during the delete operation as well
        'Call saveIconConfigurationToSource ' final write to the docksettings file
    End If
    
   On Error GoTo 0
   Exit Sub

insertNewIconDataIntoCurrentPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertNewIconDataIntoCurrentPosition of module mdlMain.bas"
    
End Sub


 

'---------------------------------------------------------------------------------------
' Procedure : saveIconConfigurationToSource
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : writes to the registry, SETTINGS.INI or the 3rd config.
'---------------------------------------------------------------------------------------
'
Public Sub saveIconConfigurationToSource()

    Dim useloop As Integer
    Dim location As String
    Dim dockSettingsCount As Integer
    
    useloop = 0
    dockSettingsCount = 0
    location = vbNullString
     
    ' save the current fields to the settings file or registry
    On Error GoTo btnSaveRestart_Click_Error
    
'    If debugflg = 1 Then debugLog "%" & "saveIconConfigurationToSource"
    
'    If rDGeneralWriteConfig = "True" Then ' the 3rd option, steamydock compatibility, writes to the new config file
         
        'first step is to cleardown the third settings file icon data
        location = "Software\SteamyDock\IconSettings\Icons"
         
        'read the old count from the dockSettingsFile
        dockSettingsCount = Val(GetINISetting(location, "count", dockSettingsFile))
         
        'Delete all icon keys - Note that when you write a null string to a record in an ini file it removes the key, deleting it.
        For useloop = 0 To dockSettingsCount
            ' write the steamydock dockSsettings.ini
            PutINISetting location, useloop & "-FileName", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-FileName2", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-Title", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-Command", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-Arguments", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-WorkingDirectory", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-ShowCmd", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-OpenRunning", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-IsSeparator", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-UseContext", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-DockletFile", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-UseDialog", vbNullString, dockSettingsFile
            PutINISetting location, useloop & "-UseDialogAfter", vbNullString, dockSettingsFile '.nn Add the two missing icon characteristics
        Next useloop
        
        ' write the 3rd settings file with real data
            For useloop = 0 To rdIconMaximum
                ' get the relevant entries from the intermediate settings file
                readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
                'readIconSettingsIni "Software\RocketDock\Icons", useloop, rdSettingsFile
                
                
                ' write the steamydock dockSsettings.ini
                Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile) ' the settings.ini only exists when RD is set to use it
             Next useloop
         ' when RD compatibility is finally removed we could do without the intermediate file and just work from the dockSettings.ini
         ' but not yet...
         
         'now write the count to the settings file
         PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, dockSettingsFile
         
'    Else ' rocketdock compatibility
'        origSettingsFile = rdAppPath & "\settings.ini"
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'
'            ' we don't need to write anything else to the intermediate rdsettings file as it has already been done in insertNewIconDataIntoCurrentPosition
'
'            'using the intermediate option is much faster just requiring a file copy
'            ' all we need to do is copy the duplicate settings file to the original
'            FileCopy rdSettingsFile, origSettingsFile
'        Else
'            ' just as for the new 3rd option, we have to transpose data from the temporary settings file to the registry, so we have to do them all in one go.
'            For useloop = 0 To rdIconMaximum
'                 ' read the rocketdock alternative settings.ini
'                 'readIconSettingsIni (useloop) ' the alternative settings.ini exists when RD is set to use it
'                 'readIconSettingsIni "Software\RocketDock\Icons", useloop, rdSettingsFile
'                 readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
'
'                 ' write the rocketdock registry
'                 writeRegistryOnce (useloop)
'             Next useloop
'             '0-IsSeparator
'             'now write the count to the registry
'             'Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count", Str$(theCount))
'
'        End If
'    End If

   On Error GoTo 0
   Exit Sub

btnSaveRestart_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSaveRestart_Click of module mdlMain.bas"
            
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readIconData
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : read the icon settings file
'---------------------------------------------------------------------------------------
'
Public Sub readIconData(ByVal iconCount As Integer)

    'if it is a good icon then read the data
    On Error GoTo readIconData_Error
'    If debugflg = 1 Then debugLog "%" & "readIconData"

    'If fFExists(rdSettingsFile) Then ' does the alternative settings.ini exist? '.nn removed for performance reasons
        'get the rocketdock alternative settings.ini for this icon alone
        'readIconSettingsIni "Software\RocketDock\Icons", iconCount, rdSettingsFile
    'End If

   On Error GoTo 0
   Exit Sub

readIconData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readIconData of module mdlMain.bas"
    
End Sub






'---------------------------------------------------------------------------------------
' Procedure : addProgramDLLorEXE
' Author    : beededea
' Date      : 13/04/2020
' Purpose   :
' the file dialog would not display when the code for the dialog was under the dock_form
' this may be because the dock_form is not visible at any time. Moving the file dialog form to the
' main dock form caused the dialog to display.
'---------------------------------------------------------------------------------------
'
Public Sub addProgramDLLorEXE()

     Dim iconImage As String
     Dim iconFileName As String
     Dim retFileName As String
     Dim retfileTitle As String
     Dim dialogInitDir As String
     Dim qPos As Integer
     Dim filestring As String
     Dim suffix As String
     Dim thisTitle As String: thisTitle = vbNullString
     
     Const x_MaxBuffer = 256
    
    On Error GoTo addProgramDLLorEXE_Error
    
     dialogInitDir = App.Path 'start dir, might be "C:\" or so also
    
     With x_OpenFilename
       .hInstance = App.hInstance
       .lpstrTitle = "Select a File Target for this icon to call"
       .lpstrInitialDir = dialogInitDir
       
       .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
       .nFilterIndex = 2
       
       .lpstrFile = String(x_MaxBuffer, 0)
       .nMaxFile = x_MaxBuffer - 1
       .lpstrFileTitle = .lpstrFile
       .nMaxFileTitle = x_MaxBuffer - 1
       .lStructSize = Len(x_OpenFilename)
     End With
    
    
     If GetOpenFileName(x_OpenFilename) <> 0 Then
        If x_OpenFilename.lpstrFile = "*.*" Then
            'txtTarget.Text = savLblTarget
        Else
            retfileTitle = x_OpenFilename.lpstrFileTitle
            retFileName = x_OpenFilename.lpstrFile
        End If
    Else
     'The CANCEL button was pressed
     'MsgBox "Cancel"
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
      
    'If fFExists(retFileName) Then
    
  ' if the user drags an icon to the dock then RD takes a icon link of the following form:
    'FileName = "C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\vbexpress.exe?62453184"
    
    If InStr(sFilename, "?") And readEmbeddedIcons = True Then  ' Note: the question mark is an illegal character and test for a valid file will fail in VB.NET despite working in VB6 so we test it as a string instead
        ' does the string contain a ? if so it probably has an embedded .ICO
        qPos = InStr(1, sFilename, "?")
        If qPos <> 0 Then
            ' extract the string before the ? (qPos)
            filestring = Mid$(sFilename, 1, qPos - 1)
        End If
        
        ' test the resulting filestring exists
        If fFExists(filestring) Then
            ' extract the suffix
            suffix = ExtractSuffixWithDot(filestring)

            suffix = Right(filestring, Len(filestring) - InStr(1, filestring, "."))
            ' test as to whether it is an .EXE or a .DLL
            If InStr(1, ".exe,.dll", LCase(suffix)) <> 0 Then
                'FileName = txtCurrentIcon.Text ' revert to the relative path which is what is expected
                'Call displayEmbeddedIcons(filestring, picBox, icoPreset)

            Else
                ' the file may have a ? in the string but does not match otherwise in any useful way
                'FileName = rdAppPath & "\icons\" & "help.png"
            End If
        Else ' the file doesn't exist in any form with ? or otherwise as a valid path
            iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-EXE.png"
            If fFExists(iconFileName) Then
                iconImage = iconFileName
            Else
                iconImage = App.Path & "\iconSettings\Icons\help.png"
            End If
        End If
    Else
    
        ' .08 DAEB 20/04/2021 mdlMain.bas Added new function to identify an icon to assign to the entry
        
        ' we should, if it is a EXE dig into it to determine the icon using privateExtractIcon
                             
        ' However, we do not extract the icon from the shortcut as it will be useless for steamydock
        ' VB6 not being able to extract and handle a transparent PNG form
        ' even if it was we have no current method of making a transparent PNG from a bitmap or ICO that
        ' I can easily transfer to the GDI collection - but I am working on it...
        ' the vast majority of default icons are far too small for steamydock in any case.
        ' the result of the above is that there is currently no icon extracted, though that may change.
        
        ' instead we have a list of apps that we can match the shortcut name against, it exists in an external comma
        ' delimited file. The list has two identification factors that are used to find a match and then we find an
        ' associated icon to use with a relative path.
           
        iconFileName = identifyAppIcons(retFileName) ' .54 DAEB 19/04/2021 frmMain.frm Added new function to identify an icon to assign to the entry
                    
        If fFExists(iconFileName) Then
          iconImage = iconFileName
        Else
            iconFileName = App.Path & "\iconSettings\my collection\steampunk icons MKVI" & "\document-EXE.png"
            If fFExists(iconFileName) Then
                iconImage = iconFileName
            Else
                iconImage = App.Path & "\iconSettings\Icons\help.png"
            End If
        End If

    End If
    
    thisTitle = getFileNameFromPath(retFileName)
    
    dock.Refresh
    
    Call insertNewIconDataIntoCurrentPosition(iconImage, thisTitle, retFileName, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
    Call menuForm.addImageToDictionaryAndCheckForRunningProcess(iconImage, retFileName)
        
    ' .13 DAEB 01/04/2021 menu.frm calls mnuIconSettings_Click_Event to start up the icon settings tools and display the properties of the new icon.
    If sDShowIconSettings = "1" And dragInsideDockOperating <> True Then ' do not show when dragging an icon inside the dock to a new location
        Call menuForm.mnuIconSettings_Click_Event
    End If
    On Error GoTo 0
    Exit Sub

addProgramDLLorEXE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addProgramDLLorEXE of module mdlMain.bas"

End Sub

'' .09 DAEB 30/04/2021 mdlMain.bas cloneThisIcon created by extracting from the menu form so it can be used elsewhere
''---------------------------------------------------------------------------------------
'' Procedure : cloneThisIcon
'' Author    : beededea
'' Date      : 01/05/2021
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Sub cloneThisIcon()
'
'    Dim useloop As Integer
'    Dim thisIcon As Integer
'    Dim notQuiteTheTop As Integer
'    Dim answer As VbMsgBoxResult
'    Dim itemName As String
'    Dim dMessage As String
'
'    On Error GoTo cloneThisIcon_Error
'
''    If debugflg = 1 Then debugLog "%" & "cloneThisIcon"
'
'    itemName = namesListArray(selectedIconIndex)
'
'
'
'
'    If selectedIconIndex < rdIconMaximum Then 'if not the top icon loop through them all and reassign the values
'
'        For useloop = selectedIconIndex To rdIconNumber Step -1
'            ' read the rocketdock alternative settings.ini
'             'readSettingsIni (useloop) ' the settings.ini only exists when RD is set to use it
'             readIconSettingsIni "Software\RocketDock\Icons", useloop, rdSettingsFile
'
'            ' and increment the identifier by one
'             'writeSettingsIni (useloop + 1)
'             Call writeIconSettingsIni("Software\RocketDock" & "\Icons", useloop + 1, rdSettingsFile) ' the alternative settings.ini exists when RD is set to use it
'
'        Next useloop
'
'    End If
'
'    'increment the new icon count
'    theCount = theCount + 1
'
'    'amend the count in both the alternative rdSettings.ini
'    PutINISetting "Software\RocketDock\Icons", "count", theCount, rdSettingsFile
'
'    'must go here
'    rdIconMaximum = rdIconMaximum + 1
'
'    If selectedIconIndex > rdIconMaximum Then selectedIconIndex = rdIconMaximum
'
'    Call saveIconConfigurationToSource
'
'    dock.Refresh
'
'    Call insertNewIconDataIntoCurrentPosition(iconImage, retFileName, retFileName, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
'    Call menuForm.addImageToDictionaryAndCheckForRunningProcess(iconImage, retFileName)
'
'    Call checkDockProcessesRunning ' trigger a test of running processes in half a second
'
'    ' if that fails, spit out an error.
'    ' no point in changing this to a non-modal message box as the dock will not restart until the modal menu has completed its work.
'    'MsgBox (itemName & " Dock item deleted at position " & selectedIconIndex)
'    If insideDockFlg = False Then MessageBox dock.hWnd, itemName & " Dock item deleted at position " & selectedIconIndex, "SteamyDock Confirmation Message", vbOKOnly
'
'    On Error GoTo 0
'    Exit Sub
'
'cloneThisIcon_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cloneThisIcon of module mdlMain.bas"
'End Sub


' .09 DAEB 30/04/2021 mdlMain.bas deleteThisIcon created by extracting from the menu form so it can be used elsewhere
'---------------------------------------------------------------------------------------
' Procedure : deleteThisIcon
' Author    : beededea
' Date      : 01/05/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub deleteThisIcon()

    Dim useloop As Integer
    Dim thisIcon As Integer
    Dim notQuiteTheTop As Integer
    Dim answer As VbMsgBoxResult
    Dim itemName As String
    Dim dMessage As String
    
    On Error GoTo deleteThisIcon_Error
    
'    If debugflg = 1 Then debugLog "%" & "deleteThisIcon"

    itemName = namesListArray(selectedIconIndex)
    
    disabledArray(selectedIconIndex) = 0
    
    'If chkConfirmSaves.Value = 1 Then
    
    '.nn Added a check to see if the operation is happening during a drag and drop inside the dock
    If insideDockFlg = False Then
        If dragFromDockOperating = True Then
            ' .12 DAEB 11/05/2021 mdlMain.bas Added function to align and centre a string so it can appear in a msgbox neatly.
            dMessage = "You have dragged the currently selected entry from the dock, " & vbCr & align(itemName, 90, " ", "both") & vbCr & " This will delete it permanently -  are you sure?"
            dragFromDockOperating = False
        Else
            dMessage = "This will delete the currently selected entry from the dock, " & vbCr & align(itemName, 90, " ", "both") & vbCr & " It will remove it permanently -  are you sure?"
        End If
        
        answer = msgBoxA(dMessage, vbQuestion + vbYesNo, "Deleting an Icon", True, "dragAndDeleteThisIcon")
        If answer = vbNo Then
            Exit Sub
        End If
    Else
        
        If dragInsideDockOperating = False Then
            dMessage = "This will delete the currently selected entry from the dock, " & vbCr & align(itemName, 90, " ", "both") & vbCr & " It will remove it permanently -  are you sure?"
            answer = msgBoxA(dMessage, vbQuestion + vbYesNo, "Deleting an Icon", True, "deleteThisIcon")
            If answer = vbNo Then
                Exit Sub
            End If
        End If
    
    End If
    
    dragInsideDockOperating = False '.nn new check for dragInsideDockOperating '.nn reset
    
    If selectedIconIndex < rdIconMaximum Then 'if not the top icon loop through them all and reassign the values
        ' read the steamyDock settings one item up in the list
        ' then write the new item at the current location effectively overwriting it
        For useloop = selectedIconIndex + 1 To rdIconMaximum
            Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile)
            Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop - 1, dockSettingsFile)
'            If useloop = 73 Then
'                Dim a As Integer
'                a = 1
'            End If
        Next useloop
        
 
        ' for the final, now unused icon, we need to delete the data and write it.
'        Call zeroAllIconCharacteristics
'        Call writeIconSettingsIni("Software\SteamyDock\IconSettings\Icons", rdIconMaximum, dockSettingsFile)
'        dictionaryLocationArray(selectedIconIndex) = dictionaryLocationArrayUpperBound

        
        ' then re-read the config for every icon
        For useloop = selectedIconIndex To rdIconMaximum
            Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile)
            
            ' read the two main icon variables into arrays, one for each
            fileNameArray(useloop) = sFilename
            namesListArray(useloop) = sTitle
            sCommandArray(useloop) = sCommand
            targetExistsArray(useloop) = 0
            
            ' check to see if each process is running and store the result away
            explorerCheckArray(useloop) = isExplorerRunning(sCommand)
            processCheckArray(useloop) = IsRunning(sCommand, vbNull)
            
            ' then copy the image array contents one location down
            If useloop + 1 <= rdIconMaximum Then
                
                dictionaryLocationArray(useloop) = dictionaryLocationArray(useloop + 1)
                disabledArray(useloop) = disabledArray(useloop + 1)
                
                ' unused vars now
                iconStoreLeftPixels(useloop) = iconStoreLeftPixels(useloop + 1)
                iconStoreRightPixels(useloop) = iconStoreRightPixels(useloop + 1)
                'iconStoreTopPixels(useloop) = iconStoreTopPixels(useloop + 1)
                iconStoreBottomPixels(useloop) = iconStoreBottomPixels(useloop + 1)
            End If
            
'            If useloop = 73 Then
'                Dim b As Integer
'                b = 1
'            End If
            
        Next useloop
    End If
        
    ' to tidy up we need to overwrite the final data from the rdsettings.ini, we will write sweet nothings to it
    removeSettingsIni (rdIconMaximum)
        
    'decrement the icon count and the maximum icon
    theCount = theCount - 1

    'amend the count in both the alternative rdSettings.ini
    PutINISetting "Software\SteamyDock\IconSettings\Icons", "count", theCount, dockSettingsFile
    
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "steamyDock", dockSettingsFile
    PutINISetting "Software\SteamyDock\DockSettings", "lastIconChanged", selectedIconIndex, dockSettingsFile
    
    'must go here
    rdIconMaximum = rdIconMaximum - 1

    If selectedIconIndex > rdIconMaximum Then selectedIconIndex = rdIconMaximum

    ' Call removeImageFromDictionary(selectedIconIndex) ' no longer needed
    
    ' instead of reordering the images within the dictionary, which is difficult as you can't just add and
    ' replace objects into an existing collection, also it does not release memory. So, instead we simply
    ' remove the icon reference from the settings file, leave the collection alone and manipulate the index number
    ' indicating which image in the collection to use.
    
    'resize all arrays used for storing icon information
    ReDim Preserve fileNameArray(rdIconMaximum) As String ' the file location of the original icons
    ReDim Preserve dictionaryLocationArray(rdIconMaximum) As String ' the dictionary location of the original icons
    ReDim Preserve namesListArray(rdIconMaximum) As String ' the name assigned to each icon
    ReDim Preserve sCommandArray(rdIconMaximum) As String ' the command assigned to each icon
    ReDim Preserve targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
    ReDim Preserve processCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    ReDim Preserve explorerCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    ReDim Preserve initiatedProcessArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
    ReDim Preserve initiatedExplorerArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track explorer processes initiated from the dock
    
    ReDim Preserve disabledArray(rdIconMaximum) As Integer '
    
    iconArrayUpperBound = rdIconMaximum
    
    Call checkDockProcessesRunning ' trigger a test of running processes in half a second
    
    ' if that fails, spit out an error.
    ' no point in changing this to a non-modal message box as the dock will not restart until the modal menu has completed its work.
    'MsgBox (itemName & " Dock item deleted at position " & selectedIconIndex)
    'If insideDockFlg = False Then MessageBox dock.hwnd, itemName & " Dock item deleted at position " & selectedIconIndex, "SteamyDock Confirmation Message", vbOKOnly

    On Error GoTo 0
    Exit Sub

deleteThisIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteThisIcon of module mdlMain.bas"
End Sub




' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : loadAdditionalImagestoDictionary
' Author    : beededea
' Date      : 29/08/2020
' Purpose   : the dictionary is rebuilt after an icon add or delete and the additional images need to be re-added back to the dictionary
'---------------------------------------------------------------------------------------
'
Public Sub loadAdditionalImagestoDictionary()

    Dim themeName As String
    Dim imageOpacity As Integer: imageOpacity = 0
    
    On Error GoTo loadAdditionalImagestoDictionary_Error
    
'    If debugflg = 1 Then debugLog "%" & "loadAdditionalImagestoDictionary"
    
    themeName = vbNullString
            
'    If key = "sDSkinLeft" Or key = "sDSkinRight" Or key = "sDSkinMid" Then
'        imageOpacity = Val(rDThemeOpacity)
'    Else
'    End If

    If rDtheme <> vbNullString And rDtheme <> "Blank" Then
        imageOpacity = Val(rDThemeOpacity)
        ' load the theme background image into the collection sDSkinLeft is the unique key
        themeName = App.Path & "\skins\" & rDtheme & "\" & rDtheme & "SDleft.png"
        If fFExists(themeName) Then
            resizeAndLoadImgToDict collLargeIcons, "sDSkinLeft", themeName, sDisabled, (0), (0), sDSkinSize, sDSkinSize, , imageOpacity
        End If
    '
    '    ' load the theme background image into the collection sDSkinMid is the unique key
        themeName = App.Path & "\skins\" & rDtheme & "\" & rDtheme & "SDmiddle.png"
        If fFExists(themeName) Then
            resizeAndLoadImgToDict collLargeIcons, "sDSkinMid", themeName, sDisabled, (0), (0), sDSkinSize, sDSkinSize, , imageOpacity
        End If

    '    ' load the theme background image into the collection sDSkinRight is the unique key
        themeName = App.Path & "\skins\" & rDtheme & "\" & rDtheme & "SDright.png"
        If fFExists(themeName) Then
            resizeAndLoadImgToDict collLargeIcons, "sDSkinRight", themeName, sDisabled, (0), (0), sDSkinSize, sDSkinSize, , imageOpacity
        End If
        
        ' load the theme separator image into the collection sDSeparator is the unique key
    '    If fFExists(App.path & "\skins\" & rDtheme & "\" & rDThemeImage) Then
    '        resizeAndLoadImgToDict collLargeIcons, "sDSeparator", App.path & "\skins\" & rDtheme & "\" & rDThemeImage, vbNullString, CLng(0), CLng(0), CLng(128), CLng(128)
    '    End If
    
    End If
    
    imageOpacity = Val(rDIconOpacity)
    
    ' load a transparent 128 x 128 image into the collection, used to stop click-throughs
    If fFExists(App.Path & "\blankSquare.png") Then
        resizeAndLoadImgToDict collLargeIcons, "blank", App.Path & "\blankSquare.png", sDisabled, (0), (0), (128), (128), , 1
    End If
    
    ' .11 DAEB 01/05/2021 mdlMain.bas load a transparent 128 x 128 image into the collection, used to highlight the position of a drag/drop
    If fFExists(App.Path & "\red.png") Then
        resizeAndLoadImgToDict collLargeIcons, "red", App.Path & "\red.png", sDisabled, (0), (0), (256), (256), , imageOpacity
    End If
    
    ' load a small circle image into the collection, used to signify running process
    '                           thisDictionary, key ,strFilename, strName,thisDisabled ,Left, ByVal Top As Long,Width, Height,fullStringKey)
    If fFExists(App.Path & "\tinyCircle.png") Then
        resizeAndLoadImgToDict collLargeIcons, "tinycircle", App.Path & "\tinyCircle.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    
    If fFExists(App.Path & "\busycog.png") Then
        resizeAndLoadImgToDict collLargeIcons, "busycog", App.Path & "\busycog.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    
    If fFExists(App.Path & "\smallGoldCoin.png") Then
        resizeAndLoadImgToDict collLargeIcons, "smallgoldCoin", App.Path & "\smallGoldCoin.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    
    ' load a small circle image into the collection, used to signify running process
    If fFExists(App.Path & "\red-X.png") Then
        resizeAndLoadImgToDict collLargeIcons, "redx", App.Path & "\red-X.png", sDisabled, (0), (0), (64), (64), , imageOpacity
    End If
    
    ' .63 DAEB 29/04/2021 frmMain.frm load a small rotating hourglass image into the collection, used to signify running actions
    If fFExists(App.Path & "\busy-F1-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass1", App.Path & "\busy-F1-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    If fFExists(App.Path & "\busy-F2-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass2", App.Path & "\busy-F2-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    If fFExists(App.Path & "\busy-F3-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass3", App.Path & "\busy-F3-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    If fFExists(App.Path & "\busy-F4-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass4", App.Path & "\busy-F4-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    If fFExists(App.Path & "\busy-F5-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass5", App.Path & "\busy-F5-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If
    If fFExists(App.Path & "\busy-F6-32x32x24.png") Then
        resizeAndLoadImgToDict collLargeIcons, "hourglass6", App.Path & "\busy-F6-32x32x24.png", sDisabled, (0), (0), (128), (128), , imageOpacity
    End If

    
   On Error GoTo 0
   Exit Sub

loadAdditionalImagestoDictionary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAdditionalImagestoDictionary of module mdlMain.bas"
    
End Sub

'' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
''---------------------------------------------------------------------------------------
'' Procedure : addNewImageFromDictionary
'' Author    : beededea
'' Date      : 18/06/2020
'' Purpose   : only used when a single icon is to be added to the dock
''             this routine is a workaround to the memory leakage problem in resizeAndLoadImgToDict
''             where if run twice the RAM usage doubled as the vars are not clearing their contents when the routine ends
''
'' When an icon is added it should no longer call the routine to recreate the arrays and collections
'' instead it calls this routine, previously there was one dictionary.
''
'' there is now a separate dictionary for the smaller icons
'' there is another dictionary for the larger icons
'' there is a third temporary dictionary that is used as temporary storage whilst resizing the above
'' when a new icon is added to the dock
''
'' we use the existing resizeAndLoadImgToDict to read the larger icon format
'' the icons to the left are written to the 3rd temporary dictionary with existing keys, the new icon is then written using the current location as part of the key
'' the icons to the right are then read from the old dictionary and then written to the new temporary dictionary with updated keys
'' the larger image dictionary is cleared down readied for population
'' the temporary dictionary is used to repopulate the larger image dictionary, a clone
'' the temporary dictionary is cleared down, ready for re-use
'
'' then we do the same for the smaller icon format images
''---------------------------------------------------------------------------------------
''
'Public Sub oldAddNewImageToDictionary(ByVal newFileName As String, ByVal newName As String)
'
'    Dim useloop As Integer
'    Dim thiskey As String
'    Dim newKey As String
'
'    On Error GoTo addNewImageToDictionary_Error
'
''    If debugflg = 1 Then debugLog "%" & "addNewImageToDictionary "
'
'    'resize all arrays used for storing icon information
'    ReDim fileNameArray(rdIconMaximum) As String ' the file location of the original icons
'    ReDim dictionaryLocationArray(rdIconMaximum) As String ' the file location of the original icons
'    ReDim namesListArray(rdIconMaximum) As String ' the name assigned to each icon
'    ReDim sCommandArray(rdIconMaximum) As String ' the command assigned to each icon
'    ReDim targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
'    ReDim processCheckArray(rdIconMaximum) As String ' the process name assigned to each icon
'    ReDim initiatedProcessArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
'                                                         ' but I feel that it does not really matter so I am going to not bother at the moment, this is something that could be done later!
'
'    ' assuming that the details have already been written to the configuration file
'    ' extract filenames from steamydock registry, settings.ini or user data area
'    ' we reload the arrays that store pertinent icon information
'    For useloop = 0 To rdIconMaximum
'        'readIconData (useloop)
'        readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile
'
'        ' read the two main icon variables into arrays, one for each
'        fileNameArray(useloop) = sFilename
'        namesListArray(useloop) = sTitle
'        sCommandArray(useloop) = sCommand
'        targetExistsArray(useloop) = 0
'
'        ' check to see if each process is running and store the result away
'        'processCheckArray(useloop) = isProcessInTaskList(sCommand)
'        processCheckArray(useloop) = IsRunning(sCommand, vbNull)
'
'    Next useloop
'
'    'redimension the array that is used to store all of the icon current positions in twips
'    ReDim Preserve iconStoreLeftPixels(theCount) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
'    ReDim Preserve iconStoreRightPixels(theCount) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    ReDim Preserve iconStoreTopPixels(theCount) ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
'    ReDim Preserve iconStoreBottomPixels(theCount) ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon
'
'    iconArrayUpperBound = rdIconMaximum '<*
'
'    ' populate the array element containing the final icon position - 31/05/2021 removed unnecessary code
''    iconStoreLeftPixels(rdIconMaximum) = iconStoreLeftPixels(rdIconMaximum - 1) + (iconWidthPxls) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
''    iconStoreRightPixels(rdIconMaximum) = iconStoreLeftPixels(rdIconMaximum) + (iconWidthPxls)   '.nn
''    iconStoreTopPixels(rdIconMaximum) =
''    iconStoreBottomPixels(rdIconMaximum) =
'
'    ' re-order the large icons in the collLargeIcons dictionary collection
'    Call incrementCollection(collLargeIcons, iconSizeLargePxls, newFileName, newName)
'
'    ' re-order the small icons in the collSmallIcons dictionary collection
'    Call incrementCollection(collSmallIcons, iconSizeSmallPxls, newFileName, newName)
'
'    Call loadAdditionalImagestoDictionary ' the additional images need to be re-added back to the dictionary
'
'   On Error GoTo 0
'   Exit Sub
'
'addNewImageToDictionary_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addNewImageToDictionary of module mdlMain.bas"
'
'End Sub
'---------------------------------------------------------------------------------------
' Procedure : addNewImageFromDictionary
' Author    : beededea
' Date      : 18/06/2020
' Purpose   : only used when a single icon is to be added to the dock
'             this routine is a workaround to the memory leakage problem in resizeAndLoadImgToDict
'             where if run twice the RAM usage doubled as the vars are not clearing their contents when the routine ends
'
' When an icon is added it should no longer call the routine to recreate the arrays and collections
' instead it calls this routine, previously there was one dictionary.
'
' there is now a separate dictionary for the smaller icons
' there is another dictionary for the larger icons
' there is a third temporary dictionary that is used as temporary storage whilst resizing the above
' when a new icon is added to the dock
'
' we use the existing resizeAndLoadImgToDict to read the larger icon format
' the icons to the left are written to the 3rd temporary dictionary with existing keys, the new icon is then written using the current location as part of the key
' the icons to the right are then read from the old dictionary and then written to the new temporary dictionary with updated keys
' the larger image dictionary is cleared down readied for population
' the temporary dictionary is used to repopulate the larger image dictionary, a clone
' the temporary dictionary is cleared down, ready for re-use

' then we do the same for the smaller icon format images
'---------------------------------------------------------------------------------------
'
Public Sub addNewImageToDictionary(ByVal newFileName As String, ByVal newName As String)

    Dim useloop As Integer
    Dim thiskey As String
    Dim newKey As String
    Dim partialStringKey As String: partialStringKey = ""
    Dim imageOpacity As Integer: imageOpacity = 0

    On Error GoTo addNewImageToDictionary_Error
    
    dictionaryLocationArrayUpperBound = rdIconMaximum

    'resize all arrays used for storing icon information
    ReDim Preserve fileNameArray(rdIconMaximum) As String ' the file location of the original icons
    ReDim Preserve dictionaryLocationArray(rdIconMaximum) As String ' the dictionary location of the original icons
    ReDim Preserve namesListArray(rdIconMaximum) As String ' the name assigned to each icon
    ReDim Preserve sCommandArray(rdIconMaximum) As String ' the command assigned to each icon
    ReDim Preserve targetExistsArray(rdIconMaximum) As Integer ' .88 DAEB 08/12/2022 frmMain.frm Array for storing the state of the target command
    ReDim Preserve processCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    ReDim Preserve explorerCheckArray(rdIconMaximum) As Boolean ' the process name assigned to each icon
    
    ReDim Preserve initiatedProcessArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
    ReDim Preserve initiatedExplorerArray(rdIconMaximum) As String ' if we redim the array without preserving the contents nor re-sorting and repopulating again we lose the ability to track processes initiated from the dock
    
    ReDim Preserve disabledArray(rdIconMaximum) As Integer '
    
    ReDim Preserve iconStoreLeftPixels(rdIconMaximum) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    ' 01/06/2021 DAEB frmMain.frm Added to capture the right X co-ords of each icon
    ReDim Preserve iconStoreRightPixels(rdIconMaximum) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    'ReDim Preserve iconStoreTopPixels(rdIconMaximum) ' 01/06/2021 DAEB frmMain.frm Added to capture the top Y co-ords of each icon
    ReDim Preserve iconStoreBottomPixels(rdIconMaximum) ' 01/06/2021 DAEB frmMain.frm Added to capture the bottom Y co-ords of each icon

'    If debugflg = 1 Then debugLog "%" & "addNewImageToDictionary "

    imageOpacity = Val(rDIconOpacity)

    'now we add the new icon image giving it a key relating to the topmost position in the dictionary locating array
    partialStringKey = LTrim$(Str$(dictionaryLocationArrayUpperBound))
    If fFExists(newFileName) Then
        ' we use the existing resizeAndLoadImgToDict to read the icon format and load into the two dictionaries
         resizeAndLoadImgToDict collLargeIcons, partialStringKey, newFileName, sDisabled, (0), (0), (iconSizeLargePxls), (iconSizeLargePxls), , imageOpacity
         resizeAndLoadImgToDict collSmallIcons, partialStringKey, newFileName, sDisabled, (0), (0), (iconSizeSmallPxls), (iconSizeSmallPxls), , imageOpacity
    End If
  
    'If selectedIconIndex <= rdIconMaximum Then 'if not the top icon loop through them all and reassign the values
        ' read the steamyDock settings one item up in the list
        ' then write the the new item at the current location effectively overwriting it
        For useloop = rdIconMaximum To (selectedIconIndex + 1) Step -1
            ' copy the array contents one location down
            If useloop < rdIconMaximum Then
                dictionaryLocationArray(useloop) = dictionaryLocationArray(useloop - 1)
            End If
        Next useloop
                                
        ' then re-read the config
        For useloop = selectedIconIndex To rdIconMaximum
            Call readIconSettingsIni("Software\SteamyDock\IconSettings\Icons", useloop, dockSettingsFile)
            
            ' read the two main icon variables into arrays, one for each
            fileNameArray(useloop) = sFilename
            namesListArray(useloop) = sTitle
            sCommandArray(useloop) = sCommand
            targetExistsArray(useloop) = 0
            
            ' check to see if each process is running and store the result away
            explorerCheckArray(useloop) = isExplorerRunning(sCommand)
            processCheckArray(useloop) = IsRunning(sCommand, vbNull)
        Next useloop
        
        dictionaryLocationArray(selectedIconIndex) = dictionaryLocationArrayUpperBound

    'End If
    
    ' reads from the last icon to the current one and for each it writes it one step up
'    For useloop = rdIconMaximum To selectedIconIndex Step -1
'        thiskey = useloop & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        newKey = useloop + 1 & "ResizedImg" & LTrim$(Str$(thisByteSize))
'        If thisCollection.Exists(thiskey) Then
'            thisCollection(newKey) = thisCollection(thiskey)
'        End If
'    Next useloop
    

    
    ' populate the array element containing the final icon position - 31/05/2021 removed unnecessary code
'    iconStoreLeftPixels(rdIconMaximum) = iconStoreLeftPixels(rdIconMaximum - 1) + (iconWidthPxls) ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
'    iconStoreRightPixels(rdIconMaximum) = iconStoreLeftPixels(rdIconMaximum) + (iconWidthPxls)   '.nn
'    iconStoreTopPixels(rdIconMaximum) =
'    iconStoreBottomPixels(rdIconMaximum) =

    ' re-order the large icons in the collLargeIcons dictionary collection
'    Call incrementCollection(collLargeIcons, iconSizeLargePxls, newFileName, newName)
'
'    ' re-order the small icons in the collSmallIcons dictionary collection
'    Call incrementCollection(collSmallIcons, iconSizeSmallPxls, newFileName, newName)
'
'    Call loadAdditionalImagestoDictionary ' the additional images need to be re-added back to the dictionary

   On Error GoTo 0
   Exit Sub

addNewImageToDictionary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addNewImageToDictionary of module mdlMain.bas"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : resizeAndLoadImgToDict
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'Uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes that consumes memory 200K -800k approx each run.
'Creates a stream object stored in global memory using the location address of the variable where the data resides
'Creates a GDI+ Image object based on the stream, using GdipLoadImageFromStream that consumes memory 300k approx.
'Finally, uses a function GdipCreateBitmapFromScan0 to both create and resize the image that consumes memory 100k approx.
'
'This occurs for each image read into the collection but the memory is not being released.
'
'Tried releasing the memory by setting the variablles to erase or set to nothing or by assigning them to an empty object
'to no avail. Instead there is a workaround that combines two collections to form a new one, see removeImageFromDictionary directly below this routine


' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
Public Function resizeAndLoadImgToDict(ByRef thisDictionary As Object, ByVal key As String, ByVal strFilename As String, ByVal thisDisabled As String, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal fullStringKey As String = "", Optional ByVal imageOpacity As Integer) As Long

    Dim thiskey As String
    Dim saveStatus As Boolean
    Dim encoderCLSID As CLSID
    Dim bytesFromFile() As Byte
    Dim Strm As stdole.IUnknown
    Dim img As Long
    Dim imgCrop As Long
    Dim imgCrop2 As Long
    Dim dx As Long
    Dim dy As Long
    Dim dockSkinWidth As Long
    Dim cropWidth As Long
    
    Dim action As String
    Dim lngPixelFormat As Long
    Dim stat As GpStatus
    'Dim imageOpacity As Integer
    
    'Dim clearBytes() As Byte
    
    On Error GoTo resizeAndLoadImgToDict_Error

    ' Get the CLSID of the PNG encoder
    Call GetEncoderClsid("image/png", encoderCLSID)
    
    ' uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes
    bytesFromFile = ReadBytesFromFile(strFilename)  ' <consumes memory 200K -800k approx.

    ' creates a stream object stored in global memory using the location address of the variable where the data resides, Olaf Schmidt
    CreateStreamOnHGlobal VarPtr(bytesFromFile(0)), 0, Strm
    
    ' Creates a GDI+ Image object based on the stream, loads it into img - Olaf Schmidt
    Call GdipLoadImageFromStream(ObjPtr(Strm), img)        ' <consumes memory 300k approx.
    If img = 0 Then Err.Raise vbObjectError, , "Could not load image with GDIPlus"

    'GDI+ API to determine image dimensions, Olaf Schmidt
    Call GdipGetImageWidth(img, dx)
    If Width <= 0 Then Width = dx
    
    Call GdipGetImageHeight(img, dy)
    If Height <= 0 Then Height = dy
        
    ' a bit of a bodge but we need to handle the background image by cropping it
    ' Rocketdock has a background theme image in a single image, it is cropped left and right to extract the ends whilst the middle is both cropped and stretched.

    
    ' uses a function createScaledImg extracted from Olaf Schmidt's code in gdiPlusCacheCls to create and resize the image
    If key = "sDSkinMid" Then
        dockSkinWidth = (rdIconMaximum * iconSizeSmallPxls) + iconSizeLargePxls * 2

        ' Get the current image pixel format. The C++ SDK clone example used PixelFormatDontCare, but this can limit what you can do with the image.
        Call GdipGetImagePixelFormat(img, lngPixelFormat)

        ' when cloning a bitmap area beyond the original size of the image, it fails and creates nothing.
        ' this ensures the cropped width is never greater than the original image.
        
        cropWidth = dockSkinWidth
        If dockSkinWidth >= 2000 Then cropWidth = 2000
        
        ' Create a new Bitmap object by cropping a portion of the long 2000px bitmap to the calculated dock width - x, y, width, height
        Call GdipCloneBitmapAreaI(0, 0, cropWidth, dy, lngPixelFormat, img, imgCrop) '
        iconBitmap = createScaledImg(imgCrop, cropWidth, dy, cropWidth, Height, imageOpacity)
    Else
        iconBitmap = createScaledImg(img, dx, dy, Width, Height, imageOpacity) ' <consumes memory 100k approx.
    End If
    
    ' Save as a PNG file no longer required but retained here for reference purposes
'    If key = "sDSkinMid" Then
'        saveStatus = GdipSaveImageToFile(iconBitmap, StrConv(App.path & "\cache\" & LTrim$(str$(Width)) & strName, vbUnicode), encoderCLSID, ByVal 0)
'    End If
    
    'override the key
    If fullStringKey = "" Then
        ' create a unique key string
        thiskey = key & "ResizedImg" & LTrim$(Str$(Width))
    Else
        thiskey = fullStringKey
    End If
    
    ' add the bitmap to the dictionary collection
    If thisDictionary.Exists(thiskey) Then
        thisDictionary.Remove thiskey
    End If
    thisDictionary.Add thiskey, iconBitmap

   On Error GoTo 0
   Exit Function

resizeAndLoadImgToDict_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure resizeAndLoadImgToDict of module mdlMain.bas"
        
End Function



' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : ReadBytesFromFile
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : Credit to Olaf Schmidt
'---------------------------------------------------------------------------------------
'
Public Function ReadBytesFromFile(ByVal Filename As String) As Byte()
   On Error GoTo ReadBytesFromFile_Error

    Dim ab As Object
    
'  With CreateObject("ADODB.Stream")
'    .Open
'      .Type = 1 'adTypeBinary
'      .LoadFromFile filename
'      ReadBytesFromFile = .Read
'    .Close
'  End With

  Set ab = CreateObject("ADODB.Stream")
  
  With ab
    .Open
      .Type = 1 'adTypeBinary
      .LoadFromFile Filename
      ReadBytesFromFile = .Read
    .Close
  End With
  
   On Error GoTo 0
   Exit Function

ReadBytesFromFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadBytesFromFile of module mdlMain.bas"
End Function


' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : createScaledImg
' Author    : Credit to Olaf Schmidt for the original
'             also to Joaquim https://www.vbforums.com/showthread.php?840601-RESOLVED-how-use-ColorMatrix
' Date      : 07/04/2020
' Purpose   : Creates the scaled image with quality and opacity attributes
'---------------------------------------------------------------------------------------
'
Public Function createScaledImg(SrcImg As Long, dxSrc, dySrc, dxDst, dyDst, opacity As Integer) As Long
    Dim img As Long
    Dim Ctx As Long
    Dim imgQuality As Long
    Dim SmoothingMode As Long
    Dim stat As GpStatus
    Dim imgAttr As Long
    Dim clrMatrix As ColorMatrix
    Dim graMatrix As ColorMatrix

    On Error GoTo createScaledImg_Error
        
' prepare the quality vars according to config.
'    SmoothingModeDefault = 0&
'    SmoothingModeHighSpeed = 1&
'    SmoothingModeHighQuality = 2&
'    SmoothingModeNone = 3&
'    SmoothingModeAntiAlias8x4 = 4&
'    SmoothingModeAntiAlias = 4&
'    SmoothingModeAntiAlias8x8 = 5&

    imgAttr = &H11
    
    'Setup the transform matrix for alpha adjustment
    clrMatrix.m(0, 0) = 1
    clrMatrix.m(1, 1) = 1
    clrMatrix.m(2, 2) = 1
    clrMatrix.m(3, 3) = 1 * opacity / 100 ' 0.5 'Alpha transform (50%)
    clrMatrix.m(4, 4) = 1

    If rDIconQuality = "0" Then
        imgQuality = &H1 '    ipmNearestNeighbor = &H5
        SmoothingMode = SmoothingModeNone
    End If
    If rDIconQuality = "1" Then
        imgQuality = &H6 '    ipmHighQualityBiLinear = &H6
        SmoothingMode = SmoothingModeHighSpeed
    End If
    If rDIconQuality = "2" Then
        imgQuality = &H7 '    ipmHighQualityBicubic = &H7
        SmoothingMode = SmoothingModeHighQuality
    End If
    
    'Creates a Bitmap object based on an array of bytes along with the destination size and format information.
    Call GdipCreateBitmapFromScan0(dxDst, dyDst, dxDst * 4, PixelFormat32bppPARGB, 0, img)
    
    If img Then
        createScaledImg = img ' set the return value to the bitmap object
        'Creates a Graphics object that is associated with an Image bitmap object ie. the hw context of the image
        Call GdipGetImageGraphicsContext(img, Ctx)
    Else
        Err.Raise vbObjectError, , "unable to create scaled Img-Resource"
    End If
    
    If Ctx Then
        ' set the quality
        Call GdipSetPixelOffsetMode(Ctx, 3)            '     4=Half, 3=None
        Call GdipSetInterpolationMode(Ctx, imgQuality) ' three levels of quality
        Call GdipSetSmoothingMode(Ctx, SmoothingMode)  '          ditto
        'Call GdipSetCompositingQuality(Ctx, CompositingQualityHighQuality)  ' CompositingQualityHighSpeed
        ' Sets the compositing quality of this Graphics object when alpha blended. Speed vs quality. USed in conjunction with GdipSetCompositingMode
                                
                
        'Create storage for the image attributes struct used below
        Call GdipCreateImageAttributes(imgAttr)

        'Setup the image attributes using the color matrix  'ColorAdjustTypeDefault
        Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeBitmap, 1, clrMatrix, graMatrix, ColorMatrixFlagsDefault)

        ' draw the loaded source image onto a generated image to the desired scale
        If SrcImg <> 0 Then
            GdipDrawImageRectRectI Ctx, SrcImg, 0, 0, dxDst, dyDst, 0, 0, dxSrc, dySrc, 2, imgAttr, 0, 0
        End If
        
        
        ' delete the now unwanted graphics context
        Call GdipDeleteGraphics(Ctx)
    End If

   On Error GoTo 0
   Exit Function

createScaledImg_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createScaledImg of module mdlMain.bas"
End Function

        
'---------------------------------------------------------------------------------------
' Procedure : updateDisplayFromDictionary
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : This utility displays using GDI+, one of several icon images stored in a dictionary collection by key.
'---------------------------------------------------------------------------------------
'
Public Function updateDisplayFromDictionary(thisCollection As Object, strFilename As String, ByVal key As String, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
   'Dim h As Long
   On Error GoTo updateDisplayFromDictionary_Error

    If thisCollection(key) <> 0 Then
        iconBitmap = thisCollection(key) ' get the stored image from the collection
    Else
        'MsgBox "help - no bitmap for " & key
        'End
        Exit Function
    End If
    
    ' the old method, retained for documentation was to load a disc file into a bitmap
    'GdipLoadImageFromFile StrPtr(strFilename), iconBitmap

    If Width = -1 Or Height = -1 Then
        Call GdipGetImageHeight(iconBitmap, Height)
        Call GdipGetImageWidth(iconBitmap, Width)
    End If

'    Dim opacity As String
'    opacity = "100"
'    If opacity <> "100" Then
'        Dim imgAttr As Long
'        Dim clrMatrix As ColorMatrix
'        Dim graMatrix As ColorMatrix
'
'        imgAttr = &H11
'
'        'Setup the transform matrix for alpha adjustment
'        clrMatrix.m(0, 0) = 1
'        clrMatrix.m(1, 1) = 1
'        clrMatrix.m(2, 2) = 1
'        clrMatrix.m(3, 3) = 1 * Val(opacity) / 100 ' 0.5 'Alpha transform (50%)
'        clrMatrix.m(4, 4) = 1
'
''       Dim iconBitmap2 As Long
'
''       'Create storage for the image attributes struct used below
'        Call GdipCreateImageAttributes(imgAttr)
''
''       'Setup the image attributes using the color matrix  'ColorAdjustTypeDefault
'        Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeBitmap, 1, clrMatrix, graMatrix, ColorMatrixFlagsDefault)
''
'        Call GdipDrawImageRectRect(gdipFullScreenBitmap, iconBitmap, Left, Top, Width, Height, 0, 0, Width, Height, 2, imgAttr, 0, 0)
'    Else
        'draws a GDIP image on the
        Call GdipDrawImageRectI(gdipFullScreenBitmap, iconBitmap, Left, Top, Width, Height)  ' shrinks the bitmap into the image object
'    End If
    
   Exit Function

updateDisplayFromDictionary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateDisplayFromDictionary of Form dock"
End Function


' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : createNewGDIPBitmap
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : Create a gdi bitmap with width and height of what we are going to draw into it. This is the entire drawing area for everything,
'             creating a bitmap in memory that our VB6/GDIP application writes to directly. Called each animation interval.
'---------------------------------------------------------------------------------------
'
Public Function createNewGDIPBitmap()
        
    On Error GoTo createNewGDIPBitmap_Error
    ''If debugflg = 1 Then debugLog "%" & "createNewGDIPBitmap" ' commented out to avoid too many debug errors
    
    ' create a device independent bitmap and return a handle bmpMemory, giving it a handle to dcMemory and providing any attributes to the new bitmap
    ' (dcMemory previously created with CreateCompatibleDC)
    bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
    
    ' Make the device context use the bitmap.
    Call SelectObject(dcMemory, bmpMemory)
    
    ' Creates a GDIP graphic object and provides a pointer 'gdipFullScreenBitmap' using a handle to the bitmap graphic section assigned to the device context
    Call GdipCreateFromHDC(dcMemory, gdipFullScreenBitmap)

   On Error GoTo 0
   Exit Function

createNewGDIPBitmap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createNewGDIPBitmap of module mdlMain.bas"
    
End Function

' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
'---------------------------------------------------------------------------------------
' Procedure : setWindowCharacteristics
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : update some characteristics for the window we will be updating using UpdateLayeredWindow API
'---------------------------------------------------------------------------------------
'
Public Function setWindowCharacteristics()

   On Error GoTo setWindowCharacteristics_Error
    If debugflg = 1 Then debugLog "% sub setWindowCharacteristics"
    
    'set the transparency of the underlying form with click through
    lngReturn = GetWindowLong(dock.hwnd, GWL_EXSTYLE)
    SetWindowLong dock.hwnd, GWL_EXSTYLE, lngReturn Or WS_EX_LAYERED
    
    ' determine the z position of the dock with respect to other application and o/s windows.
    ' this also changes the window positioning and size:
    ' x The x coordinate of where to put the upper-left corner of the window.
    ' Y The y coordinate of where to put the upper-left corner of the window.
    ' cx The x coordinate of where to put the lower-right corner of the window.
    ' cy The y coordinate of where to put the lower-right corner of the window.
    
    ' we may have to set GDI to the width of the whole virtual screen
    
    If rDzOrderMode = "0" Then
        SetWindowPos dock.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    ElseIf rDzOrderMode = "1" Then
        SetWindowPos dock.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    ElseIf rDzOrderMode = "2" Then
        SetWindowPos dock.hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE
    End If
    
    ' point structure that specifies the location of the layer updated in UpdateLayeredWindow
    apiPoint.X = 0
    apiPoint.Y = 0
    
    ' point structure that specifies the size of the window in pixels
    apiWindow.X = screenWidthPixels ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    apiWindow.Y = screenHeightPixels  ' .59 DAEB 26/04/2021 frmMain.frm changed to use pixels alone, removed all unnecesary twip conversion
    
    ' the third parameter to UpdateLayeredWindow is a pointer to a structure that specifies the new screen position of the layered window.
    ' If the current position is not changing, pptDst can be NULL. It is null.
    
    ' point structure that specifies the position of the new layer
    'newPoint.X = 0
    'newPoint.Y = 0
    
    ' blending characteristics for opacity
    funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
    funcBlend32bpp.BlendFlags = 0
    funcBlend32bpp.BlendOp = AC_SRC_OVER
  
    ' set the opacity of the whole dock, used to display solidly and for instant autohide
    funcBlend32bpp.SourceConstantAlpha = 255 * Val(dockOpacity) / 100 ' this calc can be done elsewhere and we just use a passed var
    ' the above line is also replicated where the dock opacity requires dynamic modification, ie. during an autohide and reveal

    GdipDeleteGraphics gdipFullScreenBitmap 'The graphics may now be deleted
            
    'Update the specified window handle (hwnd) with a handle to our bitmap (dc) passing all the required characteristics
    UpdateLayeredWindow dock.hwnd, hdcScreen, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA
    
    ' The UpdateLayeredWindow API call above does not need really to be run here as it is run repeatedly by the animate timer and the function to draw the icons small
    
   On Error GoTo 0
   Exit Function

setWindowCharacteristics_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setWindowCharacteristics of module mdlMain.bas"
End Function



'---------------------------------------------------------------------------------------
' Procedure : checkWindowIconisationZorder
' Author    : beededea
' Date      : 19/01/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function checkWindowIconisationZorder(ByVal thisCommand As String, ByVal selectedIconIndex As Integer, ByVal commandOverride As String, ByVal runAction As String) As Boolean
    Dim processID As Long:  processID = 0
    Dim lngRetVal As Long: lngRetVal = 0

    If processCheckArray(selectedIconIndex) = True Or commandOverride <> vbNullString Then
        'the array check above is the quick way to check process is already running
        'but we need to run IsRunning again to get the process PID
        If IsRunning(thisCommand, processID) Then ' it checks again that the process is still running, as the check process timer that populates the processCheckArray is too infrequent to be relied upon
            
            lngRetVal = handleWindowConditionAndZorder(processID, runAction)
            checkWindowIconisationZorder = True ' return
            Exit Function
            'If lngRetVal = 0 Then

        End If ' IsRunning(thisCommand, processID)
    End If ' processCheckArray(selectedIconIndex)
End Function



Public Function handleWindowConditionAndZorder(ByVal processID As Long, ByVal runAction As String) As Long

    Dim windowHwnd As Long:  windowHwnd = 0
    Dim CurrentForegroundThreadID As Long: CurrentForegroundThreadID = 0
    Dim NewForegroundThreadID As Long: NewForegroundThreadID = 0
    Dim lngRetVal As Long: lngRetVal = 0
    Dim hTray As Long: hTray = 0 ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    Dim hOverflow As Long: hOverflow = 0 ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    ' .22 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function STARTS
    
    hTray = FindWindow_NotifyTray() ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
    hOverflow = FindWindow_NotifyOverflow() ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas

    
    'windowHwnd = getWindowHWndForPid(processID) ' old method of enumerating all windows and find the associated pid of each, returning the hWnd of the window associated with the PID
    
    'The EnumWindows function is more reliable than calling the GetWindow function in a loop as we used to do.
    'ie. An application that calls GetWindow to perform this task risks being caught in an infinite
    'loop or referencing a handle to a window that has been destroyed.
    
    ' enumerate all windows and find the associated pid of each, returning the hWnd of the window associated with the given PID
    Call fEnumWindows(processID)
    windowHwnd = storeWindowHwnd
    
    ' .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas STARTS
    ' if the hwnd is zero then a matching process has not been found, in this case search the systray
    If windowHwnd = 0 Then
    
        'Me.Print "Tray Handle: 0x" & Hex(hTray)
        isSysTray hTray, processID, windowHwnd
    
        'Me.Print "Overflow Handle: 0x" & Hex(hOverflow)
        isSysTray hOverflow, processID, windowHwnd
    End If
    ' .33 DAEB 03/03/2021 frmMain.frm New systray code from DragokasENDS
    
     'GetWindowRect windowHwnd, pRect unused
    
    ' Get the thread for the current window that is to fore now (the dock)
    CurrentForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    
    ' Get the thread ID for the window we are trying to bring to the fore
    NewForegroundThreadID = GetWindowThreadProcessId(windowHwnd, ByVal 0&)
    
    'AttachThreadInput is used to ensure SetForegroundWindow will work
    'even if our application isn't currently the foreground window
    '(e.g. a minimised application running in the background)
    
    If CurrentForegroundThreadID <> NewForegroundThreadID Then
        ' Attach shared keyboard input to the thread we are raising
        Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, True)
        
        ' Make the raised window the foreground window.
        If runAction = "back" Then
            ' .38 DAEB 18/03/2021 frmMain.frm utilised SetActiveWindow to give window focus without bringing it to fore
            
            '    The SetActiveWindow function activates a window, but not if the application is in the background.
            '    The window will be brought into the foreground (top of Z order) if the application is in the foreground when it sets the activation.
            lngRetVal = SetActiveWindow(windowHwnd)
        Else
            '     Brings the thread that created the specified window into the foreground AND activates the window. Keyboard input is
            '     directed to the window, and various visual cues are changed for the user. The system assigns a slightly higher
            '     priority to the thread that created the foreground window than it does to other threads.
            lngRetVal = SetForegroundWindow(windowHwnd)
        End If
        
        ' break the thread's attachment to the newly raised window, breaking the association
        ' effectively passing control to the raised window.
        Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, False)
    Else
       lngRetVal = SetForegroundWindow(windowHwnd) ' bring window to the fore
    End If
    
    ' .22 DAEB frmMain.frm 08/02/2021 changes to replace old method of enumerating all windows with enumerate improved Windows function ENDS
    If lngRetVal <> 0 Then
                          
        If IsIconic(windowHwnd) Then
            Call ShowWindow(windowHwnd, SW_RESTORE) ' if a minimised window, bring to fore as a standard window
            Call setWindowZorder(windowHwnd, runAction)
        ElseIf IsZoomed(windowHwnd) Then
            Call ShowWindow(windowHwnd, SW_MINIMIZE) ' if a full size window, minimise
            Call setWindowZorder(windowHwnd, runAction)
        ElseIf (Not IsIconic(windowHwnd) And Not IsZoomed(windowHwnd)) Then ' a normal window
            Call setWindowZorder(windowHwnd, runAction)
            If prevIconIndex <> selectedIconIndex Then ' .27 DAEB frmMain.frm 11/02/2021 now operates like the standard Windows dock on a click, minimising then restoring
                
                ' .34 DAEB frmMain.frm 08/02/2021  - commented out the extra unwanted ShowWindow(windowHwnd, SW_RESTORE)
                ' bringing this window to the fore, not needed, SetForegroundWindow does the job already - the showWindow causes a z-order problem if it is included.
                ''''Call ShowWindow(windowHwnd, SW_RESTORE) ' < do not comment back in, leave all the commands in this if...else section commented out
                
                ' lngRetVal = SetForegroundWindow(windowHwnd) ' trial bring window to the fore
                'SetWindowPos windowHwnd, HWND_TOPMOST, pRect.Left, pRect.Top, 0, 0, SWP_NOSIZE' trial bring window to the fore
                
            Else ' if the icon clicked is the same as the one before then
                If runAction <> "focus" And runAction <> "back" Then
                    'Call ShowWindow(windowHwnd, SW_MINIMIZE)   ' minimise the window
                    Call ShowWindowAsync(windowHwnd, SW_MINIMIZE) ' .41 DAEB 18/03/2021 frmMain.frm utilised ShowWindowAsync instead of ShowWindow as the C program utilised it and it seemed to make sense to do so too
                End If
            End If
            
            ' I was not able to obtain a handle of a window with focus as it never matched
            ' the selected window. It seems that you cannot check whether the chosen window already has focus as the
            ' second you click an icon on the dock, the dock itself
            ' seems to acquire focus.
    
        ' .26 DAEB frmMain.frm 10/02/2021 added test to check window state and alter it accordingly
    
        End If
        prevIconIndex = selectedIconIndex ' .27 DAEB frmMain.frm 11/02/2021 now operates like the standard Windows dock on a click, minimising then restoring
        
    End If
    handleWindowConditionAndZorder = lngRetVal

End Function

Private Sub setWindowZorder(ByVal windowHwnd As Long, ByVal runAction As String)

    If runAction = "focus" Then
        BringWindowToTop windowHwnd ' .39 DAEB 18/03/2021 frmMain.frm utilised BringWindowToTop instead of SetWindowPos & HWND_TOP as that was used by a C program that worked perfectly.
        'SetWindowPos windowHwnd, HWND_TOP, 0, 0, 0, 0, SWP_ACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    ' .42 DAEB 03/03/2021 frmMain.frm To support new receive focus menu option
    If runAction = "normal" Then
        ' .40 DAEB 18/03/2021 frmMain.frm Added SWP_NOOWNERZORDER as an additional flag as that was used by a C program that worked perfectly, fixing the z-order position problems
        SetWindowPos windowHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER
    End If
    
     ' .42 DAEB 03/03/2021 frmMain.frm To support new receive focus menu option
    If runAction = "back" Then
        ' .40 DAEB 18/03/2021 frmMain.frm Added SWP_NOOWNERZORDER as an additional flag as that was used by a C program that worked perfectly, fixing the z-order position problems
        SetWindowPos windowHwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : isSysTray
' Author    : beededea
' Date      : 20/02/2021
' Purpose   : .33 DAEB 03/03/2021 frmMain.frm New systray code from Dragokas
'---------------------------------------------------------------------------------------
'
Public Function isSysTray(hTray As Long, ByRef processID As Long, ByRef hwnd As Long)

    Dim Count As Long: Count = 0
    Dim hIcon() As Long: 'hIcon() = 0
    Dim i As Long: i = 0
    Dim pid As Long: pid = 0

    On Error GoTo isSysTray_Error

    Count = GetIconCount(hTray)

    If Count <> 0 Then
        Call GetIconHandles(hTray, Count, hIcon)
    End If

    For i = 0 To Count - 1
        pid = GetPidByWindow(hIcon(i))
        'if the extracted pid matches the supplied processID then we have the window handle
        If pid = processID Then
            hwnd = hIcon(i)
            Exit Function
        End If
    Next

   On Error GoTo 0
   Exit Function

isSysTray_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure isSysTray of Form dock"
End Function

'---------------------------------------------------------------------------------------
' Procedure : confirmEachKillPutWindowBehind
' Author    : beededea
' Date      : 20/12/2022
' Purpose   : This routine is an analog of confirmEachKill. It is more or less identical and you should keep them in synch.
'             This version has calls to routines that require additional API calls bringing Windows to front/back.
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for SteamyDock to run that it would not share with the other utilities.
'
'---------------------------------------------------------------------------------------
'
Public Function confirmEachKillPutWindowBehind(ByVal binaryName As String, ByVal procId As Long, ByVal processToKill As String, ByVal confirmEachProcessKill As Boolean, ByRef ExitCode As Long) As Boolean
    Dim goAheadAndKill As Boolean: goAheadAndKill = False
    Dim rmessage As String: rmessage = ""
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim a As Long

    On Error GoTo confirmEachKillPutWindowBehind_Error

    If confirmEachProcessKill = True Then
        ' send application to back so the msgbox appears on top
        a = handleWindowConditionAndZorder(procId, "back")
        rmessage = "A matching process has been found. Kill this application? - " & binaryName & " with process ID " & procId
        'answer = MsgBox(rmessage, vbYesNo)
        answer = msgBoxA(rmessage, vbYesNo, "Killing this application", True, "confirmEachKillPutWindowBehind")
        If answer = vbNo Then
            goAheadAndKill = False
        Else
            goAheadAndKill = True
        End If
    Else
        goAheadAndKill = True
    End If
    
    If goAheadAndKill = True Then
        confirmEachKillPutWindowBehind = TerminateProcess(processToKill, ExitCode)
        Call CloseHandle(processToKill)
    Else
        a = handleWindowConditionAndZorder(procId, "focus")
    End If

    On Error GoTo 0
    Exit Function

confirmEachKillPutWindowBehind_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure confirmEachKillPutWindowBehind of Module common"
            Resume Next
          End If
    End With
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkAndKillPutWindowBehind
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : Find and kill any given process name
'           : This routine is an analog of checkAndKill. It is more or less identical and you should keep them in synch.
'             This version has calls to routines that require additional API calls bringing Windows to front/back.
'             I could have used compile time references (#) to bypass these but it seemed more appropriate to create
'             separate copy for SteamyDock to run that it would not share with the other utilities.
'---------------------------------------------------------------------------------------
'
Public Function checkAndKillPutWindowBehind(ByRef NameProcess As String, ByVal checkForFolder As Boolean, ByVal confirmEachProcessKill As Boolean) As Boolean

    ' variables declared
    Dim AppCount As Integer: AppCount = 0
    Dim RProcessFound As Long: RProcessFound = 0
    Dim SzExename As String: SzExename = vbNullString
    Dim MyProcess As Long: MyProcess = 0
    Dim i As Integer: i = 0
    Dim binaryName As String: binaryName = vbNullString
    Dim folderName As String: folderName = vbNullString
    Dim procId As Long: procId = 0
    Dim runningProcessFolder As String: runningProcessFolder = vbNullString
    Dim processToKill As Long: processToKill = 0
    Dim ExitCode As Long: ExitCode = 0
    
    On Error GoTo checkAndKillPutWindowBehind_Error
    'If debugflg = 1 Then debugLog "%checkAndKillPutWindowBehind"

    checkAndKillPutWindowBehind = False
    MyProcess = GetCurrentProcessId()
    
    If NameProcess <> vbNullString Then
          AppCount = 0
          
          binaryName = getFileNameFromPath(NameProcess)
          'Aborted
          If binaryName = vbNullString Then Exit Function ' catchall to prevent closure of unknown processes if the name is malformed
          
          folderName = getFolderNameFromPath(NameProcess)
          
          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

          'hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            'WinDirEnv = Environ("Windir") + "\"
            'WinDirEnv = LCase$(WinDirEnv)

            If Right$(SzExename, Len(binaryName)) = LCase$(binaryName) Then

                    AppCount = AppCount + 1
                    processToKill = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                    If uProcess.th32ProcessID = MyProcess Then
                       'MsgBox "hmmm" & MyProcess ' we never want to kill our own process...
                    Else
                        If checkForFolder = True Then ' only check the process actual run folder when killing an app from the dock
                            procId = uProcess.th32ProcessID ' actual PID
                            runningProcessFolder = getFolderNameFromPath(getExePathFromPID(procId))
                            If LCase$(runningProcessFolder) = LCase$(folderName) Then
                                ' checkAndKillPutWindowBehind = TerminateProcess(processToKill, ExitCode)
                                ' Call CloseHandle(processToKill)
'                                If App.Title = "SteamyDock" Then
'                                    checkAndKillPutWindowBehind = confirmEachKillPutWindowBehind(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
'                                Else
                                checkAndKillPutWindowBehind = confirmEachKillPutWindowBehind(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                            End If
                        Else ' just go ahead and kill whatever process I say must go
                            ' checkAndKillPutWindowBehind = TerminateProcess(processToKill, ExitCode)
                            ' Call CloseHandle(processToKill)
                            checkAndKillPutWindowBehind = confirmEachKillPutWindowBehind(binaryName, procId, processToKill, confirmEachProcessKill, ExitCode)
                        End If
                    End If
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
            
          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
    End If


   On Error GoTo 0
   Exit Function

checkAndKillPutWindowBehind_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkAndKillPutWindowBehind of Module Common"

End Function

'---------------------------------------------------------------------------------------
' Procedure : restartSteamydock
' Author    : beededea
' Date      : 27/02/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub restartSteamydock()
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo restartSteamydock_Error

    thisCommand = App.Path & "\restartSteamyDock.exe"
    
    If fFExists(thisCommand) Then
        If userLevel <> "runas" Then userLevel = "open"
        Call dock.runCommand("focus", thisCommand)
    Else
         MessageBox dock.hwnd, thisCommand & " is missing", "SteamyDock Confirmation Message", vbOKOnly + vbExclamation
    End If


    On Error GoTo 0
    Exit Sub

restartSteamydock_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure restartSteamydock of Form menuForm"
            Resume Next
          End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : toggleDebugging
' Author    : beededea
' Date      : 07/07/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub toggleDebugging()

    On Error GoTo toggleDebugging_Error

    debugLog "% sub toggleDebugging"
    
    If debugflg = 0 Then
        debugflg = 1
        menuForm.mnuDebug.Caption = "Turn Debugging OFF"
        menuForm.mnuAppFolder.Visible = True
        menuForm.mnuEditWidget.Visible = True
    Else
        debugflg = 0
        menuForm.mnuDebug.Caption = "Turn Debugging ON"
        menuForm.mnuAppFolder.Visible = False
        menuForm.mnuEditWidget.Visible = False
    End If
    
    PutINISetting "Software\SteamyDock\DockSettings", "debugFlg", debugflg, dockSettingsFile
    

   On Error GoTo 0
   Exit Sub

toggleDebugging_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure toggleDebugging of Module mdlMain"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : thisFormUnload
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub thisFormUnload()
   On Error GoTo thisFormUnload_Error

    Call dock.shutdwnGDI

    ' unload the other native VB6forms
    
    Unload about
    Unload licence
    Unload frmMessage
    Unload hiddenForm
    Unload menuForm
    Unload showAndTell
    Unload splashForm
    Unload dock
    
    ' remove all variable references to each form in turn
    
    Set about = Nothing
    Set frmMessage = Nothing
    Set hiddenForm = Nothing
    Set menuForm = Nothing
    Set showAndTell = Nothing
    Set splashForm = Nothing
    Set licence = Nothing
    Set dock = Nothing
    
    RemoveHotKey lHotKey

    ' .13 DAEB frmMain.frm 27/01/2021 Added system wide keypress support
    
    End

   On Error GoTo 0
   Exit Sub

thisFormUnload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisFormUnload of Form dock"
End Sub




