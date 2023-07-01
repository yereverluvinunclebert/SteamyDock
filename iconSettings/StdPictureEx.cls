VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdPictureEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this class was used in addition to La Volpe's later graphics code as the later version had a bug when displaying
' GIFs

'========================================================================================================
' Class purpose and usage
'========================================================================================================
' The class is defined as predeclared. This means you do not create an instance of it. An instance is
'   created automatically when your project is run, similar to a .bas module. You do not unload it or set
'   it to Nothing. You can access its properties/methods directly by its name, i.e., StdPictureEx.LoadPicture

' Requirements: This class requires GDI+ and XP operating system or higher. If the class fails to load
'   GDI+ or runs on a system lower than XP, then all functions may fail over to VB's standard picture
'   functions as if this class does not even exist. Unicode support may still be provided, but no
'   additional enhancements will be available. If you wish to code this for lower operating systems,
'   you will need to ensure CreateIconFromResourceEx API usage does not fail. 32bpp icons were introduced
'   with XP and will not be supported on lower operating systems. The GDI+ requirement remains.

' Enhancements: These formats are supported: BMP,ICO,CUR,WMF,EMF,JPG,PNG,TIF,GIF.
'   Additionally, unicode support is provided for path and file names. Proper 32bpp icon rendering is
'   supported. PNG-encoded icons are supported. Caching of supported image formats can be performed and
'   saved to file. Navigation of TIF pages and GIF frames are supported. This means you can animate GIFs.
'   Depending on system settings, bitmap formats that include color management are supported. In any case,
'   bitmaps with both true and premultiplied alpha channels are supported. The LoadPicture and SavePicture
'   methods also will accept byte arrays and existing stdPicture objects, along with file names.

' When to refresh picture's container object? When calling the SubImage() method to change the frame of
'   an animated GIF or page of a multipage TIF, the container may need to be refreshed. If the return
'   result from SubImage() is assigned to a VB image control, generally refresh otherwise not needed.
'   When in doubt, you can always refresh the container. If the SubImage() picture's Handle property
'   value is same as the picture passed to that method, refreshing is usually required. When this
'   class changes the frame/page of a picture, it attempts to re-use the previous frame/page and simply
'   update the pixel data. This may not trigger a change in picture to VB, thus refreshing may be needed.

' The term 'managed', within this class, indicates that the class is rendering the image. This class will
'   render the 32bpp bitmaps it creates and also 32bpp icons since VB does not do these. When the picture
'   is eventually set to Nothing or replaced by another image, this class will be informed via the thunks
'   it created and will release any cached data at that time and dispose of any related GDI+ image.

' In order to support alpha channel image formats, any returned stdPicture object may contain an
'   alpha-blended bitmap having its bits premultiplied for use with AlphaBlend API. That API is used to
'   render semi-transparency in bitmaps. When the class' IsManaged property returns true, then if the
'   Picture.Type value is vbPicTypeBitmap, the bitmap bits are premultiplied. Because of this bitmap
'   format, you should not pass the stdPicture outside of your project. The receiving end may not know
'   how to render the image. If needed, you can call the StdPictureEx.SavePicture method and convert
'   the bitmap to a standard bitmap. Any transparency in the bitmap will be rendered over a background
'   color you provided to SavePicture. That method, in that case, always produces a 24bpp bitmap.

' When possible, the cached data is done so that VB also has access to it via the stdPicture itself.
'   This applies to non-CMYK JPGs, single-frame GIFs, many icons and most bitmaps. For all other cases,
'   the cached data is stored by this class and is not accessible by VB. The StdPictureEx.SavePicture
'   method can be used to save this cached data to file. Bitmap and metafile formats are not cached.
'   General rule of thumb. If caching original data is wanted and VB can load the image without color
'   loss, then the image will be loaded via GDI+ and data will be cached by VB/COM. If this cannot be
'   done, then this class will cache the data and manage the picture so the class will know when to
'   dispose of the cached data.

' The class creates two assembly thunks (executable code in memory). These thunks are only used to
'   prevent crashes, during IDE, when subclassing VB picture objects. Their purposes are briefly
'   described next. By existing in memory unknown to VB, the thunks are not destroyed when your project
'   closes due to executing an End statement via code, IDE toolbar button or debug message box. Since
'   the thunk isn't released in that scenario, its code remains alive, crashes are prevented unlike
'   standard subclassing.
'--------------------------------------------------------------------------------------------------------
'   Thunk #1: Management Window. This window is created the first time this class loads and remains alive
'       until the project or IDE completely unloads. The thunk is the window's subclassed window procedure.
'       When the window is destroyed (when project/VB closes), then it properly shuts down GDI+ after
'       first releasing any GDI+ image objects that may have been created. It releases Thunk#2 and the
'       copy of the COM IPicture/IPictureDisp virtual table. The closure of the project's thread releases
'       this thunk.

'   Thunk #2: IPicture/IPictureDisp subclasser. This is also only created once. Whenever pictures are
'       subclassed, just 4 of the many picture functions are subclassed, the others are untouched. The
'       thunk will do the drawing of any subclassed pictures. It also informs this class when a subclassed
'       picture is eventually destroyed. The four picture functions that are subclassed/tweaked are:
'           1. IPictureDisp's IUnknown:Release. Tracks when picture is set to nothing or replaced
'           2. IPicture's IUnknown:Release. Tracks when picture is set to nothing or replaced
'           3. IPicture:Get_Attributes. Forces VB to see our subclassed image as transparent. Why?
'               VB will only refresh/redraw the area behind the image if it believes the image has
'               transparency. By default, only GIFs are recognized 'bitmap' formats that support
'               transparency. Subclassing this function/property ensures we can fool VB into believing
'               whatever we want has transparency.
'           4. IPicture:Render. Uses AlphaBlend API to draw bitmaps and DrawIconEx API to draw
'               icons,cursors.
'       FYI: Reason both IUnknown:Release functions are subclassed is that when a picture is set to
'       nothing or is replaced or goes out of scope, COM does not expose the zero-reference count in
'       both interfaces. It does this in one or the other. So both are tracked to ensure we know when
'       it reaches zero.

'   VTable Copy: COM appears to handle all pictures via a single interface instance within the current
'       thread. All created pictures are processed by this single VTable. To subclass the VTable, this
'       class makes a copy of that VTable, then tweaks that copy. All subclassed pictures are rerouted
'       to this tweaked VTable. All other picture objects are left untouched. Reversing this can be done
'       via the UnManage method.
'========================================================================================================

Option Explicit
'///// GDI functions
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetObjectA Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32.dll" (ByVal hEMF As Long) As Long
'///// User32 functions
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As Any) As Long
'Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECTI, ByVal hBrush As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DestroyCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long
Private Const LR_CREATEDIBSECTION As Long = &H2000
'///// Kernel32 functions
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualQuery Lib "kernel32.dll" (ByRef lpAddress As Any, ByRef lpBuffer As Any, ByVal dwLength As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
'Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32.dll" (ByVal lpfn As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'///// GDI+ functions
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (ByRef Token As Long, ByRef inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
'Private Declare Function GdipGetImageType Lib "GdiPlus.dll" (ByVal Image As Long, ByRef pType As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus.dll" (ByVal Stream As Long, ByRef Image As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RECTI, ByVal mFlags As Long, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByRef scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal nImage As Long, ByRef srcRect As RECTF, ByRef srcUnit As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GdiPlus.dll" (ByVal hImage As Long, ByRef PixelFormat As Long) As Long
'Private Declare Function GdipImageSelectActiveFrame Lib "GdiPlus.dll" (ByVal Image As Long, ByRef dimensionID As Any, ByVal FrameIndex As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "GdiPlus.dll" (ByVal Image As Long, ByRef dimensionID As Any, ByRef count As Long) As Long
Private Declare Function GdipImageGetFrameDimensionsCount Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef count As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "GdiPlus.dll" (ByVal hImage As Long, ByVal GUID As Long) As Long
Private Declare Function GdipImageGetFrameDimensionsList Lib "GdiPlus.dll" (ByVal pImage As Long, ByRef dimensionIDs As Any, ByVal count As Long) As Long
Private Declare Function GdipEmfToWmfBits Lib "GdiPlus.dll" (ByVal hEMF As Long, ByVal cbData16 As Long, ByVal pData16 As Long, ByVal iMapMode As Long, ByVal eFlags As Long) As Long
Private Declare Function GdipGetHemfFromMetafile Lib "GdiPlus.dll" (ByVal metafile As Long, ByRef hEMF As Long) As Long


Private Const ImageLockModeUserInputBuf     As Long = &H4&
Private Const ImageLockModeRead             As Long = &H1&
Private Const PixelFormat32bppPremultiplied As Long = &HE200B
Private Const PixelFormat32bppAlpha         As Long = &H26200A
Private Const PixelFormat32bpp              As Long = &H262000
Private Const PixelFormat24bpp              As Long = &H21808
'///// Misc functions
Private Declare Function OleLoadPicture Lib "OLEPRO32.DLL" (ByRef pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As Any, ByRef ppvObj As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (ByRef lpPictDesc As Any, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef iPic As IPicture) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long

Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal ppstm As Long, ByRef hGlobal As Long) As Long
Private Declare Function CoTaskMemRealloc Lib "ole32.dll" (ByVal pv As Long, ByVal cb As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, ByRef lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByRef pBlock As Any, ByVal lpSubBlock As String, ByRef lplpBuffer As Any, ByRef puLen As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIML As Long, ByRef Cx As Long, ByRef Cy As Long) As Long
Private Declare Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef riid As Long, ByVal ppv As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long

Private Type RECTF                      ' GDI+ rectangle w/Single vartypes
    nLeft As Single
    nTop As Single
    nWidth As Single
    nHeight As Single
End Type
Private Type RECTI                      ' GDI/GDI+ rectangle w/Long vartypes
    nLeft As Long
    nTop As Long
    nWidth As Long
    nHeight As Long
End Type
Private Type BitmapData                 ' GDI+ lock/unlock bits structure
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    Scan0Ptr As Long
    ReservedPtr As Long
End Type
Private Type ICONDIR
    idReserved As Integer               ' per msdn: must be zero
    idType As Integer                   ' per msdn: must be 1 or 2
    idCount As Integer                  ' unsigned
End Type
Private Type ICONDIRENTRY
    bWidth As Byte                      ' will be 0 when width > 255
    bHeight As Byte                     ' will be 0 when height > 255
    bColorCount As Byte                 ' not used in this class; not applicable if > 8bpp
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer                ' not used in this class; extracted from image data instead
    dwBytesInRes As Long                ' how many bytes are used by the image data
    dwImageOffset As Long               ' where in icon resource, image data begins (bmp/png header)
End Type
Private Type BITMAPV5HEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    ' up to this point: BITMAPCOREHEADER 12 bytes (biWidth/biHeight INTEGER vs LONG, bitcount=1,4,8,24 only)
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
    ' v2 starts next(52 bytes)
    bV2RedMask As Long                  ' offset 40
    bV2GreenMask As Long
    bV2BlueMask As Long
    ' v3 starts next (56 bytes)
    bV3AlphaMask As Long                ' offset 52
    ' v4 starts next (108 bytes)
    bV4CSType As Long
    bV4Endpoints(0 To 8) As Long        ' offset 60 << actually bv4Endpoints declared as CIEXYZTRIPLE
    bV4GammaRed As Long
    bV4GammaGreen As Long
    bV4GammaBlue As Long
    ' v5 starts next (124 bytes)
    bV5Intent As Long                   ' offset 108
    bV5ProfileData As Long              ' offset 112
    bV5ProfileSize As Long              ' offset 116
    bV5Reserved As Long
End Type

Private Type PICREF     ' Used only to track GDI+ image creation/disposal
    flags As Long
    ' 0x0000000F        ' PictureTypeConstantsEx value
    ' 0x0000FFF0        ' current frame (GIF/TIF)
    ' 0x0FFF0000        ' current frame count (GIF/TIF), max 4095
    ' 0xF0000000        ' reserved
    pIPicture As Long   ' ObjPtr(IPicture)
    pIPicDisp As Long   ' ObjPtr(IPictureDisp)
    pHandle As Long     ' GDI+ image handle reference (multi-frame/page images only)
    oStream As IUnknown ' Cached data (IStream). Optional
End Type

Public Enum PictureTypeConstantsEx      ' Used as return value for PictureType property
    ptcNone = vbPicTypeNone
    ptcBitmap = vbPicTypeBitmap         ' BMP & if original format not cached: converted TIF,JPG,PNG,GIF
    ptcIcon = vbPicTypeIcon             ' VB does not distinguish btwn icon and cursor
    ptcMetafile = vbPicTypeMetafile
    ptcEMetafile = vbPicTypeEMetafile
    ptcJPEG = vbPicTypeEMetafile + 1&
    ptcPNG = vbPicTypeEMetafile + 2&
    ptcGIF = vbPicTypeEMetafile + 3&    ' frame count/navigation available. See LoadPicture method
    ptcTIF = vbPicTypeEMetafile + 4&    ' page count/navigation available. See LoadPicture method
End Enum
Public Enum LoadPictureSizeConstantsEx
    lpsSmall = vbLPSmall                ' system small icon size, 16x16
    lpsLarge = vbLPLarge                ' system large icon size, usually 32x32
    lpsSmallShell = vbLPSmallShell      ' shell small icon size, usually 16x16 (DPI scalale)
    lpsLargeShell = vbLPLargeShell      ' shell large icon size, usually 32x32 (DPI scalable)
    lpsCustom = vbLPCustom              ' user-defined size. LoadPicture DesiredIconCx/Cy parameters apply
    lpsDefault = -1&                    ' if just 1 icon exists, actual size else lpsLarge
    lpsXtraLargeShell = -2&             ' shell extra large size, usually 48x48 (DPI scalable) XP+
    lpsJumboShell = -3&                 ' shell jump size, 256x256 (Vista+)
End Enum                                ' note: if shell sizes fail to be retrieved, above sizes are used instead

Private m_PageGUID(0 To 3, 0 To 1) As Long ' used to change frames/pages of GIF/TIF
Private m_hDC As Long                   ' general usage. Is zero if class failed to initialize properly
Private m_RefCount As Long              ' number of m_PicRef() items
Private m_PicRefs() As PICREF           ' collection of managed picture data (-1 LBound)
Private m_Primary As StdPictureEx       ' subclassing class instance if not this class instance
Private m_Hwnd As Long                  ' management window handle
Private m_ThunkPtr As Long              ' address of IPicture/IPictureDisp subclassing thunk
Private m_VistaPlus As Boolean          ' are we running in Vista or better?

Public Function LoadPicture(Optional ByRef ImageSource As Variant, _
                            Optional ByVal IconSize As LoadPictureSizeConstantsEx = lpsDefault, _
                            Optional ByVal IconColorDepth As Long, _
                            Optional ByVal DesiredIconCx As Long, _
                            Optional ByVal DesiredIconCy As Long, _
                            Optional ByVal KeepOriginalFormat As Boolean = False, _
                            Optional ByVal RequiredFormat As PictureTypeConstants = vbPicTypeNone) As StdPicture
                           
'========================================================================================================
' Key notes. Method replicates VB's LoadPicture function and adds more options
'========================================================================================================
'    Returned Picture object is always a new object, never a reference to source, with one exception.
'       If ImageSource is a stdPicture object and this method fails, then it will return a reference to
'       the passed stdPicture vs. Nothing.
'    When assigning returned picture to image control, you should refresh image control.
'    Unicode file names are supported.
'     GDI+ is used to support TIF, PNG, CMYK JPGs, non-placeable WMF, and GIF.
'     note: Different versions of GDI+ support different TIF compression schemes. Therefore, a TIF loaded
'     on Win7 may fail to load on Vista or XP.
'========================================================================================================

'========================================================================================================
' Parameters
'========================================================================================================
'   ImageSource. One of the following can be provided. Anything else returns Nothing
'       Path/file name of image, as a String. Unicode supported
'       Byte array of complete image format/data, one dimensional only
'       Existing stdPicture/IPictureDisp object (basically any VB icon/picture property).
'          A copy of the passed Picture's data/format will be used. You should not pass that picture
'          unless there is a need to duplicate it, waste of system resources. Here are some reasons
'          to pass an existing Picture object:
'          1. It is unmanaged and contains a 32bpp alpha bitmap and transparency is wanted
'          2. You want to return the current frame/page of a mutli-frame/page GIF/TIF
'          3. You want a copy of the passed Picture and its cached data
'          So, after what was just said, this is what you can expect if passing a stdPicture object:
'          - Managed, multi-Page/Frame GIF/TIF.
'               If KeepOriginalFormat=True then a copy else single frame/page
'          - Any other scenario returns a copy of image & cached data if KeepOriginalFormat=True
'
'     IconSize. Applies if loading icons,cursors or if RequiredFormat is vbPicTypeIcon
'       lpsDefault. If only 1 icon exists in ImageSource, it is uses as-is else a lpsLarge is the size used.
'       lpsLarge. Uses the system's large icon size, usually 32x32 but can be larger
'       lpsSmall. Uses the system's small icon size, currently fixes at 16x16
'       lpsSmallShell. Windows shell small icon size, usually 16x16 (scales with DPI)
'       lpsLargeShell. Windows shell large icon size, usually 32x32 (scales with DPI)
'       lpsExtraLargeShell. Windows shell extra large icon size, usually 48x48 (scales with DPI) (XP or better)
'       lpsJumboShell. Windows shell jumbo icon size, currently fixed at 256x256 (Vista or better)
'       lpsCustom. The DesiredCx,DesiredCy parameters are used else always ignored
'
'     IconColorDepth. Applies if loading icons,cursors or if RequiredFormat is vbPicTypeIcon
'       0 = use best color depth that matches the system's color depth
'       1,4,8,16,24,32 = specific color depth. Next highest is used if no match is found
'
'     DesiredIconCx,DesireIconCy. Applies if IconSize is lpsCustom and if ImageSource is icon,cursor
'       0 = use system's large icon size
'       anything else will result in the icon matching the size being used else resized as needed
'
'     KeepOriginalFormat. Applies to all image formats except BMP, WMF, EMF
'       Setting to True requires more memory to cache the format. Cached data is always the bits.
'       Caching original format data allows StdPictureEx.SavePicture to save the original data
'          Tip: do not set this to true if loading from a resource file unless animating a GIF
'              or navigating multipage TIF. You have it cached in resource file, no need to re-cache.
'          If needed, you can remove the cached data by calling this method again & set this parameter to false
'       WMF,EMF,BMP: Ignored. VB/COM can save to these formats without caching original format
'       ICO,CUR: VB can save to icon format, but reduces to 4bpp, degrading quality, therefore parameter applies.
'           If more than 1 icon/cursor exists in passed data, only the selected image data is cached, not all.
'       GIF,TIF: If multiple frames/pages exists, setting parameter to true allows selection of the frames/pages
'       If this parameter is False, the only image formats returned are: BMP,ICO,WMF,EMF
'       If this parameter is True, then PictureProperty of this class can return: BMP,ICO,WMF,EMF,JPG,PNG,TIF,GIF
'
'     RequiredFormat. This option forces the returned picture to one of two formats
'       Only vbPicTypeBitmap and vbPicTypeIcon are accepted. KeepOriginalFormat does not apply if bitmap
'       is chosen. All parameters apply if icon is chosen This can be useful to force any loaded image to
'       icon and simultaneously resize the result by supplying the icon-related size parameters. Most VB
'       image properties accept icons but most icon properties do not accept bitmaps
'========================================================================================================

    Dim pPicRef As PICREF
    Dim newPic As IPictureDisp
    Dim bKeepData As Boolean

   
    Call pvVerifyInitialization
    ' note:if above call fails, then this class cannot be used. Actions default to VB methods only
    '   if failure, then these variables will contain following values & may be tested elsewhere:
    '   m_Primary is Nothing, m_hDC = 0, m_ThunkPtr = 0 or -1
    If m_Primary Is Nothing Then
        Select Case RequiredFormat
            Case vbPicTypeIcon      ' prevent caching original data until icon is created
                bKeepData = KeepOriginalFormat: KeepOriginalFormat = False
            Case vbPicTypeBitmap    ' bitmaps do not cache original data
                KeepOriginalFormat = False
            Case vbPicTypeNone
            Case Else: RequiredFormat = vbPicTypeNone
        End Select
        If DesiredIconCx < 0& Then DesiredIconCx = 0& Else DesiredIconCx = DesiredIconCx And &HFFF&
        If DesiredIconCy < 0& Then DesiredIconCy = 0& Else DesiredIconCy = DesiredIconCy And &HFFF&
        If IconColorDepth < 0& Then IconColorDepth = 0&
       
        If IsMissing(ImageSource) = False Then
            If IsEmpty(ImageSource) = False Then
                If IsObject(ImageSource) Then
                    If Not ImageSource Is Nothing Then
                        If TypeOf ImageSource Is StdPicture Then
                            If Not ImageSource.handle = 0& Then    ' send off to be processed
                                Set newPic = ImageSource
                                Set newPic = pvLoadStdPicture(newPic, IconSize, DesiredIconCx, DesiredIconCy, _
                                                              KeepOriginalFormat, pPicRef)
                                If newPic Is Nothing Then Set newPic = ImageSource
                            End If
                        End If
                    End If
                ElseIf VarType(ImageSource) = vbString Then
                    If Not ImageSource = vbNullString Then          ' send off to be processed
                        Set newPic = pvLoadFile(CStr(ImageSource), IconSize, IconColorDepth, DesiredIconCx, DesiredIconCy, _
                                                KeepOriginalFormat, pPicRef)
                    End If
                ElseIf VarType(ImageSource) = (vbArray Or vbByte) Then
                    Dim iData() As Byte
                    On Error Resume Next
                    pPicRef.flags = LBound(ImageSource, 2)
                    If Err Then     ' not multi-dimensional array
                        Err.Clear   ' check for uninitialized/invalid arrays
                        If UBound(ImageSource) > LBound(ImageSource) Then
                            If Err.Number = 0& Then                 ' send off to be processed
                                On Error GoTo 0
                                iData() = ImageSource
                                Set newPic = pvLoadPicture(iData(), IconSize, IconColorDepth, DesiredIconCx, DesiredIconCy, _
                                                         KeepOriginalFormat, pPicRef)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        Set LoadPicture = m_Primary.LoadPicture(ImageSource)
    End If
   
    If Not newPic Is Nothing Then
        If RequiredFormat = vbPicTypeBitmap Then
            Set newPic = pvCvAnytoBMP(newPic, pPicRef)
        ElseIf RequiredFormat = vbPicTypeIcon Then
            Set newPic = pvCvAnyToICO(newPic, pPicRef, bKeepData, IconSize, DesiredIconCx, DesiredIconCy)
        End If
        ' when above pvLoad[xxx] calls return, the pPicRef.Flags is non-null if class will manage it
        If pPicRef.flags Then
            Dim nPic As IPicture
            Set nPic = newPic: pPicRef.pIPicture = ObjPtr(nPic): Set nPic = Nothing
            pPicRef.pIPicDisp = ObjPtr(newPic)
            Call pvAddGDIpItem(pPicRef)
            CopyMemory ByVal pPicRef.pIPicDisp, GetProp(m_Hwnd, "IPicDispAddr"), 4&
            CopyMemory ByVal pPicRef.pIPicture, GetProp(m_Hwnd, "IPicAddr"), 4&
        End If
        Set LoadPicture = newPic
    End If

End Function



Public Property Get SubImageCount(ByRef Picture As StdPicture) As Long
    ' return number of pages/frames of multi-frame/page GIF/TIF
    If pvVerifyInitialization Then
        If m_Primary Is Nothing Then
            SubImageCount = (m_PicRefs(pvPictureToIndex(Picture)).flags And &HFFF0000) \ &H10000
        Else
            SubImageCount = m_Primary.SubImageCount(Picture)
        End If
    End If
End Property

Public Property Get SubImageIndex(ByRef Picture As StdPicture) As Long
    ' return current frame/page index of multi-frame/page GIF/TIF
    If pvVerifyInitialization Then
        If m_Primary Is Nothing Then
            SubImageIndex = (m_PicRefs(pvPictureToIndex(Picture)).flags And &HFFF0&) \ &H10
        Else
            SubImageIndex = m_Primary.SubImageIndex(Picture)
        End If
    End If
End Property





Public Property Get IsManaged(ByRef Picture As StdPicture) As Boolean
    ' returns whether VB or this class is rendering/managing the passed Picture
    If pvVerifyInitialization Then
        If m_Primary Is Nothing Then
            If Not Picture Is Nothing Then
                If pvPictureToIndex(Picture) > -1& Then
                    Dim pTable As Long
Dim vTable As Long

                    CopyMemory pTable, ByVal pvSafePointerAdd(m_ThunkPtr, 8&), 4&  ' get pointer to VTable thunk
                    CopyMemory vTable, ByVal pvSafePointerAdd(pTable, 4&), 4&      ' get original VTable address
                    CopyMemory pTable, ByVal ObjPtr(Picture), 4&                   ' VTable address for Picture object
                    IsManaged = Not (pTable = vTable)
                End If
            End If
        Else
            IsManaged = m_Primary.IsManaged(Picture)
        End If
    End If
End Property

Public Property Get PictureType(ByRef Picture As StdPicture) As PictureTypeConstantsEx
    ' replacement for VB's Picture.Type property
    ' managed images can return GIF, JPG, TIF, PNG as picture types
    If pvVerifyInitialization Then
        If m_Primary Is Nothing Then
            If Not Picture Is Nothing Then
                Dim lSize As Long
Dim oPic As IPicture
Dim oStream As IUnknown

                Dim hMem As Long
Dim hLock As Long

               
                With m_PicRefs(pvPictureToIndex(Picture))
                    If .flags = 0 Then                              ' passed unmanaged Picture
                        PictureType = Picture.Type                  ' set default return value
                        Set oPic = Picture
                        If oPic.KeepOriginalFormat Then             ' can peek to see if GIF or JPG
                            hMem = GlobalAlloc(&H2&, 0&)
                            If hMem Then
                                Call CreateStreamOnHGlobal(hMem, 1&, oStream)
                            Else
                                GlobalFree hMem
                            End If
                            If Not oStream Is Nothing Then
                                oPic.SaveAsFile ByVal ObjPtr(oStream), 0&, lSize
                                If lSize < 4& Then
                                    Set oStream = Nothing
                                Else
                                    hLock = GlobalLock(hMem)
                                    If hLock Then
                                        CopyMemory lSize, ByVal hLock, 4&
                                        GlobalUnlock hMem
                                    Else
                                        lSize = 0&
                                    End If
                                    If (lSize And &HFFFFFF) = &H464947 Then
                                        PictureType = ptcGIF
                                    ElseIf Not (lSize And &HFFFF&) = &H4D42& Then
                                        PictureType = ptcJPEG
                                    End If
                                End If
                                Set oStream = Nothing
                            End If
                        End If
                    Else
                        PictureType = .flags And &HF
                    End If
                End With
            End If
        Else
            PictureType = m_Primary.PictureType(Picture)
        End If
    End If
End Property



Public Property Get HasOriginalFormat(ByRef Picture As StdPicture) As Boolean
   
    ' returns whether original image format/data is maintained. If maintained and not needed,
    '   you can call this class LoadPicture and pass the KeepOriginalFormat parameter as false to
    '   reload image without caching the data, reducing memory usage, i.e.,
    '   Set Image1.Picture = StdPictureEx.LoadPicture(Image1.Picture)
    ' Note: When your project is uncompiled, VB ensures loaded pictures cache original data/format.
    '   But when compiled, those same Pictures no longer have their data cached. In other words do
    '   not assume that this property will return the same result for design-time loaded images
    '   when your project is compiled vs. uncompiled. If original format is needed, highly recommend
    '    you store your images in a resource file and load them from there.
   
    Dim oIPic As IPicture
    Call pvVerifyInitialization
    If m_Primary Is Nothing Then
        With m_PicRefs(pvPictureToIndex(Picture))
            If .oStream Is Nothing Then
                Set oIPic = Picture
                HasOriginalFormat = CBool(oIPic.KeepOriginalFormat)
            Else
                HasOriginalFormat = True
            End If
        End With
    Else
        HasOriginalFormat = m_Primary.HasOriginalFormat(Picture)
    End If
   
End Property





Private Function pvPictureToIndex(ByVal Picture As IPictureDisp) As Long
    ' returns index for passed managed picture
    Dim pObj As Long
    If Picture Is Nothing Then
        pvPictureToIndex = -1&
    Else
        pObj = ObjPtr(Picture)
        For pvPictureToIndex = m_RefCount - 1& To 0& Step -1&
            If m_PicRefs(pvPictureToIndex).pIPicDisp = pObj Then Exit For
        Next
    End If

End Function
   
Private Function pvVerifyInitialization() As Boolean
    ' helper function that ensures a) thunks created and/or b) primary subclasser identified
    If Not m_ThunkPtr = -1& Then
        If m_ThunkPtr = 0& Then
            If pvCreateThunks() = True Then
                If m_Primary Is Nothing Then
                    m_hDC = CreateCompatibleDC(0&)
                    ReDim m_PicRefs(-1 To -1)
                End If
                pvVerifyInitialization = True
            End If
        Else
            pvVerifyInitialization = True
        End If
    End If
End Function

Private Function pvLoadPicture(ByRef SourceData() As Byte, ByRef Size As Long, ByRef IconColorDepth As Long, _
                            ByRef DesiredIconCx As Long, ByRef DesiredIconCy As Long, _
                            ByRef KeepFormat As Boolean, _
                            ByRef pPicRef As PICREF) As StdPicture
                           
    ' workhorse for the class
                           
    Dim hBmp As Long
    Dim lRead As Long
    Dim lFlags As Long
    
    Dim hBits As Long
    Dim hImage As Long
    'Dim hToken As Long
    Dim LB As Long
    
    Dim BIH As BITMAPV5HEADER
    Dim uBMP As BitmapData
    Dim rBounds As RECTI

    Dim IStream As IUnknown
    Const ImageLockModeWrite = &H2
   
    LB = LBound(SourceData)
    If m_hDC Then                               ' else class not initialized
        CopyMemory lFlags, SourceData(LB), 4&
        If (lFlags And &HFFFF&) = &H4D42& Then  ' bitmaps handled separately to support 32bpp
            KeepFormat = False                  ' and also to support v4 & v5 of the bitmap info header
            CopyMemory BIH.biSize, SourceData(LB + 14&), 4&       ' size of header
            If Not (BIH.biSize < 40& Or BIH.biSize > 124&) Then   ' quick validation
                CopyMemory BIH.biWidth, SourceData(LB + 18&), 36& ' get next 36 bytes of header
                If BIH.biBitCount > 0 And BIH.biCompression < 4& Then  ' else likely BI_PNG,BI_JPG & not supported
                    hBmp = CreateDIBSection(m_hDC, SourceData(LB + 14&), 0&, hBits, 0&, 0&) ' use GDI to parse bitmap header & create DIB
                    If hBmp Then
                        lRead = BIH.biSize              ' determine where pixel data starts
                        If BIH.biCompression = 3& And BIH.biSize = 40& Then ' BI_BITFIELDS
                            If (BIH.biBitCount = 16 Or BIH.biBitCount = 32) Then lRead = lRead + 12&
                        End If
                        If BIH.biClrUsed Then
                            lRead = lRead + 4& * BIH.biClrUsed
                        ElseIf BIH.biBitCount <= 8& Then
                            lRead = lRead + 4& * 2 ^ BIH.biBitCount
                        End If
                        If SetDIBits(m_hDC, hBmp, 0&, Abs(BIH.biHeight), SourceData(LB + 14& + lRead), SourceData(LB + 14&), 0&) = 0& Then
                            DeleteObject hBmp: hBmp = 0&
                        ElseIf BIH.biBitCount = 32& Then ' othewise, we'll let VB deal with it
                            lFlags = pvValidateAlphaChannel(hBits, BIH.biWidth, Abs(BIH.biHeight)) ' test alpha-usage
                            If lFlags = PixelFormat32bppAlpha Then  ' create GDI+ image for conversion
                                With uBMP                       ' will be destroying DIB, need to upload its bits
                                    .Height = Abs(BIH.biHeight)
                                    .Width = BIH.biWidth
                                    .PixelFormat = lFlags
                                    .stride = -.Width * 4&
                                    .Scan0Ptr = pvSafePointerAdd(hBits, -.stride * (.Height - 1&))
                                    rBounds.nHeight = .Height: rBounds.nWidth = .Width
                                End With                        ' create GDI+ bitmap & upload DIB bits
                                GdipCreateBitmapFromScan0 BIH.biWidth, uBMP.Height, uBMP.stride, uBMP.PixelFormat, ByVal 0&, hImage
                                If hImage Then
                                    If GdipBitmapLockBits(hImage, rBounds, ImageLockModeUserInputBuf Or ImageLockModeWrite, uBMP.PixelFormat, uBMP) Then
                                        If GdipBitmapLockBits(hImage, rBounds, ImageLockModeUserInputBuf Or ImageLockModeWrite, uBMP.PixelFormat, uBMP) Then uBMP.Scan0Ptr = 0&
                                    End If
                                    If uBMP.Scan0Ptr Then
                                        GdipBitmapUnlockBits hImage, uBMP
                                        DeleteObject hBmp: hBmp = 0&
                                        pPicRef.flags = ptcBitmap: lFlags = 0&
                                        Set pvLoadPicture = pvConvertFromGDIplus(hImage, IStream, KeepFormat, pPicRef)
                                    Else
                                        GdipDisposeImage hImage: hImage = 0&
                                    End If
                                End If
                            ElseIf lFlags = PixelFormat32bppPremultiplied Then ' if premultiplied, then managed
                                pPicRef.flags = ptcBitmap
                            End If
                        End If
                    End If
                End If
            End If
            If hBmp Then
                Set pvLoadPicture = pvHandleToStdPicture(hBmp, vbPicTypeBitmap)
                If pvLoadPicture Is Nothing Then DeleteObject hBmp: lFlags = 0&
            Else
                lFlags = 0&                                     ' flag for default load
            End If
        ElseIf (lFlags And &HFFFFFF) = &H464947 Then            ' test for GIF
            If KeepFormat = False Then lFlags = 0&
        ElseIf Not lFlags = &H474E5089 Then                     ' test for PNG
            If Not (lFlags And &HFFFF&) = &H4949& Then          ' test for Tiff two ways
                If Not (lFlags And &HFFFF&) = &H4D4D& Then
                    ' test for icons; they are handled separately
                    If (lFlags And &HFFFF&) = 0& Then ' this may want to be relaxed? Per MSDN must be zero
                        If ((lFlags And &HFFFF0000) = &H10000 Or (lFlags And &HFFFF0000) = &H20000) Then
                            Set pvLoadPicture = pvLoadhIcon(SourceData(), DesiredIconCx, DesiredIconCy, Size, IconColorDepth, KeepFormat, pPicRef)
                        End If
                    End If
                    lFlags = 0&
                End If
            End If
        End If
    End If
   
    ' when we get here, lFlags=0 if attempting to load using just VB
    If pvLoadPicture Is Nothing Then
        lRead = Abs(UBound(SourceData) - LB) + 1&
        hBits = 0&: Call pvIStreamFromPointer(0&, lRead, hBits)
        If hBits Then
            CopyMemory ByVal hBits, SourceData(LB), lRead
            Set IStream = pvIStreamFromPointer(0&, 0&, hBits)
        End If
        On Error Resume Next
        If lFlags = 0& Then
            Set pvLoadPicture = pvIStreamToPicture(IStream, KeepFormat)
            If pvLoadPicture Is Nothing Then
                If m_hDC = 0& Then Exit Function    ' else re-load stream & try GDI+
                ' Note: if failed to load some formats, OLE does something to the stream that
                '   prevents GDI+ from loading the image (guessing locking). So to get past
                '   that potential problem, we simply reload the stream
                Set IStream = Nothing
                hBits = 0&: Call pvIStreamFromPointer(0&, lRead, hBits)
                CopyMemory ByVal hBits, SourceData(LB), lRead
                Set IStream = pvIStreamFromPointer(0&, 0&, hBits)
            End If
        End If
        On Error GoTo 0
        If pvLoadPicture Is Nothing Then                  ' failed, try GDI+ (CMYK JPGs supported)
            If m_hDC Then Set pvLoadPicture = pvConvertFromGDIplus(0&, IStream, KeepFormat, pPicRef)
        End If
    End If

End Function

Private Function pvLoadFile(ByVal Filename As String, ByRef Size As Long, ByRef IconColorDepth As Long, _
                            ByRef DesiredIconCx As Long, ByRef DesiredIconCy As Long, _
                            ByRef KeepFormat As Boolean, _
                            ByRef pPicRef As PICREF) As StdPicture

    ' nothing is loaded by file, everything loaded by array
    Dim lRead As Long
Dim lSize As Long
Dim hFile As Long

    Dim aPix() As Byte
   
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&
    
    lRead = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
            Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    hFile = CreateFileW(StrPtr(Filename), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, lRead, 0&)
    If Not hFile = -1& Then
        lSize = GetFileSize(hFile, ByVal 0&)            ' load via array
        If lSize > 0& Then                              ' else > 2GB & not dealing with it
            SetFilePointer hFile, 0&, ByVal 0&, 0&
            ReDim aPix(0 To lSize - 1&)
            ReadFile hFile, aPix(0), lSize, lRead, ByVal 0&
            CloseHandle hFile
            If lSize = lRead Then                       ' pass off to array handler
                Set pvLoadFile = pvLoadPicture(aPix(), Size, IconColorDepth, DesiredIconCx, DesiredIconCy, KeepFormat, pPicRef)
            End If
        Else
            CloseHandle hFile
        End If
    End If
   
End Function

Private Function pvLoadStdPicture(ByRef Picture As StdPicture, ByRef IconSize As Long, _
                            ByRef DesiredIconCx As Long, ByRef DesiredIconCy As Long, _
                            ByRef KeepFormat As Boolean, ByRef pPicRef As PICREF, _
                            Optional ByVal ReturnDataOnly As Boolean) As StdPicture
                           
    ' Note: ReturnDataOnly is True when this is called from the SavePicture method only
    '   and in that case, bitmap is the only format this method will not return pPicRef.oStream

    ' method copies the passed picture object to a new stdPicture object (duplicates handle/data)
    ' kind of lengthy routine, but source data can exist in more than one place or may not exist at all
    '   1. If COM cached source data for picture, use that source data; may or may not be managed
    '   2. If managed and source data cached locally, use that data & manage
    '   3. If no source data at all, depending on the image type, create source data one of a few ways
    '       GIF will be processed uniquely to preserve any transparency, will be managed
    '       ICO,BMP,JPG: copy the handle and use that as the new source, may or may not be managed
    '       For any other formats, have VB create source similar to SavePicture, not managed
    Dim hGlobal As Long
Dim lSize As Long
Dim hLock As Long

    Dim pUDT() As Long
Dim icData() As Byte
Dim Index As Long

    Dim BIH As BITMAPV5HEADER
Dim pRECT As RECTI
Dim oIPic As IPicture

    Dim transColor As Long
Dim lColor As Long
Dim hImage As Long

    ' pPicRef.flags:
    '  > 0  managed returned picture
    '    0  unmanaged returned Picture
    '  < 0  flags for this routine, changed to 0+ on exit
    '   -1  source not yet determined
    '   -2  Stream from COM cached data
    '   -3  Stream from COM created data
    '   -4  Stream from locally cached data
    '   -5  Handle source (gif)
    '   -6  Handle source (icon)
    '   -7  Handle source (bitmap)
    pPicRef.flags = -1&
    Set oIPic = Picture
   
    '/// Step 1: Determine where source info is coming from
    If oIPic.KeepOriginalFormat Then          ' has original data cached by COM, retrieve it
        pPicRef.flags = -2&
    ElseIf m_hDC = 0& Then                          ' class was never initialized properly
        pPicRef.flags = -3&
    Else
        Index = pvPictureToIndex(Picture)
        If Index < 0& Then                          ' unmanaged
            Select Case Picture.Type
            Case vbPicTypeMetafile, vbPicTypeEMetafile
                pPicRef.flags = -3&
            Case vbPicTypeBitmap
                If (oIPic.Attributes And 2&) Then   ' GIF without source data
                    pPicRef.flags = -5&
                Else                                ' bmp or jpg
                    pPicRef.flags = -7&
                End If
            Case Else                               ' icon
                If ReturnDataOnly Then pPicRef.flags = -3& Else pPicRef.flags = -6&
            End Select
            KeepFormat = False
       
        Else ' managed picture objects... PNG,TIF,GIF,CMYK-JPEG,32bpp ICO/BMP & possibly 8+bpp icons/cursors
            If Picture.Type = vbPicTypeIcon Then
                If ReturnDataOnly Then pPicRef.flags = -3& Else pPicRef.flags = -6&
            Else
                pPicRef.flags = -7&
            End If
            If Not m_PicRefs(Index).oStream Is Nothing Then
                If KeepFormat = True Then pPicRef.flags = -4&
            End If
        End If
    End If
   
    '/// Step 2: Process the source info
    Select Case pPicRef.flags
    Case -2&, -3&, -4&                                  ' stream sources
        If pPicRef.flags = -4& Then                     ' use locally cached data
            pPicRef.flags = 0&
            If ReturnDataOnly Then
                Set pPicRef.oStream = m_PicRefs(Index).oStream: Exit Function
            End If
            Call GetHGlobalFromStream(ObjPtr(m_PicRefs(Index).oStream), hGlobal)
            If hGlobal Then lSize = GlobalSize(hGlobal)
        Else                                            ' have COM save it to stream
            hGlobal = GlobalAlloc(&H2&, 0&): Call CreateStreamOnHGlobal(hGlobal, 1&, pPicRef.oStream)
            If pPicRef.oStream Is Nothing Then GlobalFree hGlobal: GoTo ExitRoutine
            If pPicRef.flags = -2& Then                 ' from COM-cached data
                oIPic.SaveAsFile ByVal ObjPtr(pPicRef.oStream), 0&, lSize
            Else                                        ' from COM-created data
                oIPic.SaveAsFile ByVal ObjPtr(pPicRef.oStream), 1&, lSize
            End If
            pPicRef.flags = 0&
            If ReturnDataOnly Then Exit Function
        End If
        If lSize = 0& Then GoTo ExitRoutine
        
        hLock = GlobalLock(hGlobal)                     ' transfer stream to byte array
        If hLock = 0& Then GoTo ExitRoutine             ' then process byte array
        ReDim icData(0 To lSize - 1&)
        CopyMemory icData(0), ByVal hLock, lSize
        GlobalUnlock hGlobal
        pPicRef.flags = 0&: Set pPicRef.oStream = Nothing
        Set pvLoadStdPicture = pvLoadPicture(icData(), IconSize, 0&, DesiredIconCx, DesiredIconCy, KeepFormat, pPicRef)
   
    Case -6&                                                ' icon,cursor
        If ReturnDataOnly Then Exit Function
        pPicRef.flags = 0&: ReDim pUDT(0 To 4)              ' determine if 32bpp blended icon or not
        GetIconInfo Picture.handle, pUDT(0)                 ' blended icons are managed others aren't unless caching data
        If pUDT(3) Then DeleteObject pUDT(3)                ' destroy any mask bitmap
        If pUDT(4) Then                                     ' color bitmap handle
            BIH.biSize = 40&: hImage = pUDT(4)
            GetDIBits m_hDC, hImage, 0&, 0&, ByVal 0&, BIH, 0&
            If BIH.biBitCount = 32 Then                     ' manage if not simple transparency
                If BIH.biSizeImage = 0& Then BIH.biSizeImage = pvDWordAlign(BIH.biBitCount, BIH.biWidth, BIH.biHeight)
                ReDim pUDT(0 To BIH.biSizeImage \ 4 - 1)    ' simple transparency: all alpha values are zero
                GetDIBits m_hDC, hImage, 0&, Abs(BIH.biHeight), pUDT(0), BIH, 0&
                If pvValidateAlphaChannel(VarPtr(pUDT(0)), BIH.biWidth, Abs(BIH.biHeight)) Then
                    pPicRef.flags = vbPicTypeIcon
                End If
            End If
            DeleteObject hImage
        End If
        Erase pUDT()
        Call pvSetIconSize(IconSize, DesiredIconCx, DesiredIconCy, 1&)
        hImage = CopyImage(Picture.handle, 1&, DesiredIconCx, DesiredIconCy, 0&)
        Set pvLoadStdPicture = pvHandleToStdPicture(hImage, vbPicTypeIcon)
        If pvLoadStdPicture Is Nothing Then DestroyIcon hImage
        KeepFormat = False
   
    Case -7&                                                ' bitmap
        pPicRef.flags = 0&: If ReturnDataOnly Then Exit Function
        hImage = CopyImage(Picture.handle, 0&, 0&, 0&, LR_CREATEDIBSECTION)
        Set pvLoadStdPicture = pvLoadhBitmap(hImage, pPicRef)
        If pvLoadStdPicture Is Nothing Then DeleteObject hImage
        KeepFormat = False
       
    Case -5&                                                ' GIF frame without source
        pPicRef.flags = 0&: If ReturnDataOnly Then Exit Function
        BIH.biSize = 40&: GetDIBits m_hDC, Picture.handle, 0&, 0&, ByVal 0&, BIH, 0&
        If BIH.biBitCount <= 8 Then                         ' preserve transparency
            ReDim pUDT(0 To 265)                            ' size large enough for header + 256 palette
            CopyMemory pUDT(0), BIH, 40&                    ' then call again to fill in the palette
            GetDIBits m_hDC, Picture.handle, 0&, 0&, ByVal 0&, pUDT(0), 0&
            For transColor = 1& To 258&
                For lSize = 10& To 265&                     ' header is 40 bytes & 40 bytes = 10 Longs
                    If pUDT(lSize) = transColor Then Exit For ' see if transColor in palette
                Next
                If lSize = 266& Then Exit For               ' if transColor not in palette, we're done
            Next
            Erase pUDT()
            With BIH
                .biBitCount = 32: .biClrUsed = 0: .biClrImportant = 0: .biCompression = 0
                .biHeight = Abs(.biHeight): .biSizeImage = pvDWordAlign(.biBitCount, .biWidth, .biHeight)
            End With
            hImage = CreateDIBSection(m_hDC, BIH, 0&, hLock, 0&, 0&)
            If hImage Then                                  ' render GIF over our transColor
                hImage = SelectObject(m_hDC, hImage)        ' create brush from BGR vs RGB format
                lColor = CreateSolidBrush((transColor And &HFF00&) Or (transColor And &HFF) * &H10000)
                pRECT.nWidth = BIH.biWidth: pRECT.nHeight = BIH.biHeight
                FillRect m_hDC, pRECT, lColor: DeleteObject lColor
                With Picture
                    .Render (m_hDC), 0&, 0&, (BIH.biWidth), (BIH.biHeight), 0&, (.Height - 1&), .Width, -.Height, ByVal 0&
                End With
                hImage = SelectObject(m_hDC, hImage)
                For lSize = 0& To BIH.biSizeImage - 1& Step 4&
                    hGlobal = pvSafePointerAdd(hLock, lSize)    ' now any color that is transColor is 100%
                    CopyMemory lColor, ByVal hGlobal, 4&        '   transparent, else it is 100% opaque
                    If lColor = transColor Then lColor = 0& Else lColor = lColor Or &HFF000000
                    CopyMemory ByVal hGlobal, lColor, 4&
                Next
                Set pvLoadStdPicture = pvHandleToStdPicture(hImage, vbPicTypeBitmap)
                If pvLoadStdPicture Is Nothing Then
                    DeleteObject hImage
                Else
                    pPicRef.flags = vbPicTypeBitmap
                End If
            End If
        End If
        If pvLoadStdPicture Is Nothing Then
            hImage = CopyImage(Picture.handle, 0&, 0&, 0&, 0&)
            Set pvLoadStdPicture = pvHandleToStdPicture(hImage, vbPicTypeBitmap)
            If pvLoadStdPicture Is Nothing Then DeleteObject hImage
            pPicRef.flags = 0&: KeepFormat = False
        End If
    End Select
   
ExitRoutine:
    '/// Step 3: Clean up
    Set oIPic = Nothing
    If pvLoadStdPicture Is Nothing Then
        Set pvLoadStdPicture = Picture
        Set pPicRef.oStream = Nothing
        pPicRef.flags = 0&
    ElseIf KeepFormat = False Then
        Set pPicRef.oStream = Nothing
    End If
   
End Function

Private Function pvLoadhBitmap(ByRef hBitmap As Long, ByRef pPicRef As PICREF) As StdPicture

    ' Determines if passed image is 32bpp or not. If it is, a new version (premultiplied RGB) is returned
    ' Source will be a stdPicture object passed from LoadPicture or a bitmap handle passed from pvLoadFromFile
   
    Dim BIH As BITMAPV5HEADER
Dim tPic As StdPicture

    Dim hImage As Long
Dim lFlags As Long
Dim aPix() As Byte

   
    If m_hDC Then                                   ' else class not initialized correctly
        BIH.biSize = 40&
        GetDIBits m_hDC, hBitmap, 0&, 0&, ByVal 0&, BIH, 0&  ' get bit count
        If BIH.biBitCount = 32 Then
            ReDim aPix(0 To BIH.biSizeImage - 1)        ' get pixel data
            BIH.biHeight = -BIH.biHeight                ' want bits in top-down vs bottom-up order
            If GetDIBits(m_hDC, hBitmap, 0&, Abs(BIH.biHeight), aPix(0), BIH, 0&) Then
                lFlags = pvValidateAlphaChannel(VarPtr(aPix(0)), BIH.biWidth, Abs(BIH.biHeight))
                If lFlags = PixelFormat32bppAlpha Then
                    ' taking easy way out, load bitmap in GDI+ then convert to premultiplied
                    ' else we'd have to test whether passed handle is DIB and if not (is DDB) then create
                    '   DIB & copy bits; then we can loop thru the pixels and premultiply them
                    GdipCreateBitmapFromScan0 BIH.biWidth, Abs(BIH.biHeight), BIH.biWidth * 4&, PixelFormat32bppAlpha, aPix(0), hImage
                    If hImage Then Set tPic = pvConvertFromGDIplus(hImage, Nothing, False, pPicRef)
                ElseIf lFlags = PixelFormat32bppPremultiplied Or lFlags = PixelFormat32bpp Then ' premultiplied already
                    Set tPic = pvHandleToStdPicture(hBitmap, vbPicTypeBitmap)
                    If Not tPic Is Nothing Then pPicRef.flags = ptcBitmap
                Else
                    Set tPic = pvHandleToStdPicture(hBitmap, vbPicTypeBitmap)
                    pPicRef.flags = 0&
                End If
                Set pvLoadhBitmap = tPic
           End If
            Erase aPix()
        End If
    End If
    If pvLoadhBitmap Is Nothing Then
        Set pvLoadhBitmap = pvHandleToStdPicture(hBitmap, vbPicTypeBitmap)
        pPicRef.flags = 0&
    End If
 
End Function

Private Function pvLoadhIcon(ByRef icData() As Byte, ByRef dCx As Long, ByRef dCy As Long, _
                             ByRef Size As LoadPictureSizeConstantsEx, ByRef dDepth As Long, _
                             ByVal KeepFormat As Boolean, ByRef pPicRef As PICREF) As StdPicture

    ' method returns stdPicture containing a GDI icon. If the icon is 32bpp then the class will render it
    ' the icon data can contain 32bpp and PNG-encoded icons

    Dim LB As Long
    Dim lMax As Long
    Dim lVal As Long
    
    Dim Index As Long
    Dim lFlags As Long
    Dim hHandle As Long
    
    Dim icSorted() As Long
    Dim lSize As Long
    
    Dim ICD As ICONDIR
    Dim ICDE As ICONDIRENTRY
    
    Dim xDepth As Long
    Dim xCx As Long
    Dim xCy As Long

    Const png_Signature1 As Long = 1196314761
   
    LB = LBound(icData)
    lMax = Abs(UBound(icData) - LB) + 1&
    If lMax < 62& Then Exit Function     ' can't be icon data
   
    '/// VALIDATE INTEGRITY OF THE FORMAT
    On Error GoTo ExitRoutine
    CopyMemory ICD, icData(LB), 6&      ' sanity checks
    If (ICD.idCount < 1& Or ICD.idType < 1& Or ICD.idType > 2&) Then Exit Function
    captureIconCount = ICD.idCount ' capture the icon count for the rest of the program
    lSize = ICD.idCount * 16& + 6&
    ICD.idCount = (ICD.idCount And &H7FF&)
    If lSize > lMax Then Exit Function
   
    ReDim icSorted(0 To ICD.idCount - 1&)
    ' validate array has enough bytes, sort by size & depth
    For ICD.idCount = 0& To ICD.idCount - 1&
        CopyMemory ICDE, icData(LB + ICD.idCount * 16& + 6&), 16&
        If ICDE.dwBytesInRes < 1& Or ICDE.dwBytesInRes < 1& Then Exit Function ' corrupt or not icon data
        If ICDE.dwBytesInRes + ICDE.dwImageOffset > lMax Then Exit Function
        lSize = lSize + ICDE.dwBytesInRes
        If lSize > lMax Then Exit Function           ' corrupt or not icon data
        CopyMemory lFlags, icData(LB + ICDE.dwImageOffset), 4&
        If lFlags = png_Signature1 Then
            ' Offset + 16& is PNG width, big endian
            ' Offset + 20& is PNG height, big endian
            ' Offset + 24& is PNG bit depth
            ' Offset + 25& is PNG color type
            CopyMemory lFlags, icData(LB + ICDE.dwImageOffset + 16&), 4&
            lVal = (pvReverseLong(lFlags) And &H7FF)
            CopyMemory lFlags, icData(LB + ICDE.dwImageOffset + 20&), 4&
            lVal = (((pvReverseLong(lFlags) And &H7FF) + lVal) \ 2) * &H100&
            lFlags = icData(LB + ICDE.dwImageOffset + 24&)
            Select Case icData(LB + ICDE.dwImageOffset + 25&)
                Case 2&: lFlags = 24&   ' RGB format
                Case 4&: lFlags = 9&    ' Alpha grayscale (fudge 8bit so this 8bit better than normal 8bit)
                Case 6&: lFlags = 32&   ' Alpha RGB
                Case 0&: If lFlags = 16& Then lFlags = 8&
                Case Else ' use reported depth
            End Select
            lVal = lVal Or lFlags Or &H80000000 Or ICD.idCount * &H100000
        Else    ' non-PNG
            CopyMemory lVal, icData(LB + ICDE.dwImageOffset + 4&), 4&
            CopyMemory lFlags, icData(LB + ICDE.dwImageOffset + 8&), 4&
            lVal = (((lVal And &H7FF) + (Abs(lFlags \ 2) And &H7FF)) \ 2) * &H100&
            CopyMemory lFlags, icData(LB + ICDE.dwImageOffset + 14), 2&
            lVal = lVal Or (lFlags And &HFF) Or ICD.idCount * &H100000
        End If
        For Index = 0& To ICD.idCount - 1&
            ' we sort icon metrics for a best pick strategy, icons are sorted by size & depth
            ' not a perfect algo, but since most icons in same resource are same shape, should be ok
            ' 0x000000FF    bit depth
            ' 0x0007FF00    (width+height)\2
            ' 0x7FF00000    index into the icon directory
            ' 0x80000000    PNG resource
            If (lVal And &H7FF00) < (icSorted(Index) And &H7FF00) Then
                CopyMemory icSorted(Index + 1&), icSorted(Index), (ICD.idCount - Index) * 4&
                Exit For
            End If
        Next
        icSorted(Index) = lVal
    Next
   
    '/// DETERMINE SIZE OF ICON TO RETURN
    Call pvSetIconSize(Size, dCx, dCy, (ICD.idCount))
   
    '/// GET BEST MATCH FOR RETURNED ICON SIZE & DEPTH
    If ICD.idCount > 1& Then
        If dDepth < 1& Then dDepth = GetDeviceCaps(m_hDC, 12&) ' 12 = constant BITSPERPIXEL
        lMax = (((dCx And &H7FF) + (dCy And &H7FF)) \ 2) * &H100&   ' desired value
        For Index = 0& To ICD.idCount - 1&
            If (icSorted(Index) And &H7FF00) >= lMax Then           ' size-match found
                lMax = icSorted(Index) And &H7FF00                  ' find best-match for depth (sorted descending)
                For lVal = Index + 1& To ICD.idCount - 1&
                    If (icSorted(lVal) And &H7FF00) > lMax Then     ' landed on a different size, use previous entry
                        lVal = lVal - 1&
                        Exit For
                    ElseIf dDepth < (icSorted(lVal) And &HFF) Then  ' use this size & depth
                        Exit For
                    End If
                Next
                Index = lVal: Exit For                              ' algo is done
            End If
        Next
        If Index = ICD.idCount Then Index = Index - 1&              ' adjust if Index out of bounds, uses largest size/depth
        CopyMemory ICDE, icData(LB + 6& + 16& * (icSorted(Index) And &H7FF00000) \ &H100000), 16& ' get chosen ICDE
    End If
   
    '/// CREATE HICON WHETHER PNG-ENCODED OR NOT
    If (m_VistaPlus Or icSorted(Index) > -1&) Then ' If PNG, then icSorted(n) high bit is set
        ' we can't use this on lower than Vista if PNG
        If ICD.idType = 1 Then  ' icon
            hHandle = CreateIconFromResourceEx(icData(LB + ICDE.dwImageOffset), ICDE.dwBytesInRes, 1&, &H30000, dCx, dCy, 0&)
        Else ' Cursors are a bit different...
            CopyMemory lVal, icData(LB + ICDE.dwImageOffset - 4&), 4& ' back up those 4 bytes in source data
            CopyMemory icData(LB + ICDE.dwBytesInRes - 4&), ICDE.wPlanes, 4& ' copy the X/Y hotspot to that location
            hHandle = CreateIconFromResourceEx(icData(LB + ICDE.dwImageOffset - 4&), ICDE.dwBytesInRes + 4&, 0&, &H30000, dCx, dCy, 0&)
            CopyMemory icData(LB + ICDE.dwImageOffset - 4&), lVal, 4& ' restore those 4 bytes to source data
        End If
   
    '/// CREATE ICON FROM PNG ON XP AND LOWER
    Else
        Dim rBounds As RECTI
        Dim sBounds As RECTF
          ' variables for creating PNG, icon/cursor and mask
        Dim uBMP As BitmapData
        Dim BIH As BITMAPV5HEADER

        Dim IStream As IUnknown
        Dim pngData() As Byte
        
        Dim mskShift As Long
        Dim mskScanWidth As Long
        
        Dim clrPtr As Long
        Dim mskPtr As Long
        
        Dim clrPtrBase As Long
        Dim mskPtrBase As Long

       
        Set IStream = pvIStreamFromPointer(VarPtr(icData(LB + ICDE.dwImageOffset)), ICDE.dwBytesInRes, 0&)
        If Not IStream Is Nothing Then  ' get PNG loaded via GDI+
            GdipLoadImageFromStream ObjPtr(IStream), hHandle
            If hHandle Then
                GdipGetImageBounds hHandle, sBounds, 2& ' 2=UnitPixel
                With uBMP   ' set variables for extracting PNG data, creating icon mask
                    .Height = sBounds.nHeight: .Width = sBounds.nWidth
                    rBounds.nHeight = .Height: rBounds.nWidth = .Width
                    GdipGetImagePixelFormat hHandle, .PixelFormat
                    Select Case .PixelFormat
                    Case PixelFormat32bppAlpha, PixelFormat32bppPremultiplied, &H61007, &H34400D, &H1C400E ' all known GDI+ alpha formats
                        .PixelFormat = PixelFormat32bppAlpha ' return 32bpp Alpha
                        icSorted(Index) = 32&
                    Case Else
                        .PixelFormat = PixelFormat24bpp  ' return 24bpp format
                        icSorted(Index) = 24&
                    End Select
                    If ICD.idType = 1 Then LB = 0& Else LB = 4&                             ' Cursor has 4 more bytes
                    BIH.biBitCount = (.PixelFormat And &HFF00&) \ &H100
                    .stride = -pvDWordAlign(BIH.biBitCount, .Width)                         ' color pixel scanwidth
                    mskScanWidth = pvDWordAlign(1, .Width)                                  ' mask scanwidth
                    clrPtrBase = 40& + LB: mskPtrBase = -.stride * .Height + clrPtrBase     ' offsets for pixel/mask data
                    lSize = mskPtrBase + mskScanWidth * .Height                             ' want GDI+ to give us pixels in bottom-up order
                    ReDim pngData(0 To lSize - 1&)                                          ' resize array & set position where
                    .Scan0Ptr = VarPtr(pngData(mskPtrBase + .stride))                       '   GDI+ will cache the pixel data
                End With
                If GdipBitmapLockBits(hHandle, rBounds, ImageLockModeRead Or ImageLockModeUserInputBuf, uBMP.PixelFormat, uBMP) Then
                    If GdipBitmapLockBits(hHandle, rBounds, ImageLockModeRead Or ImageLockModeUserInputBuf, uBMP.PixelFormat, uBMP) Then lSize = 0&
                End If
                If lSize Then GdipBitmapUnlockBits hHandle, uBMP ' done with GDI+ PNG, destroy it
                GdipDisposeImage hHandle: Set IStream = Nothing: hHandle = 0&
                If lSize Then   ' create the bitmap info header
                    With BIH
                        .biHeight = uBMP.Height * 2&
                        .biWidth = uBMP.Width
                        .biSize = 40&: .biPlanes = 1
                        .biSizeImage = lSize - .biSize - LB     ' image size minus header & cursor XY hotspot
                    End With: CopyMemory pngData(LB), BIH, BIH.biSize ' copy header to array, offset for cusory XY hotspot as needed
                    If BIH.biBitCount = 32 Then                 ' otherwise, mask is all zeroes, fully opaque
                        ' build a true mask. Though not used by Windows, it could be used by whoever
                        ' called this routine. Just a matter of doing it right vs. doing it quick
                        mskPtr = mskPtrBase: clrPtr = clrPtrBase + 3&: xDepth = 0&
                        mskShift = 128&: xCy = Abs(uBMP.Height): xCx = -uBMP.stride
                        For xCy = 0& To uBMP.Height - 1&
                            For clrPtr = clrPtr To clrPtr + xCx - 4& Step 4&
                                ' mask bits: 0=opaque, 1=transparent
                                If pngData(clrPtr) < 128 Then xDepth = xDepth Or mskShift
                                If mskShift = 1& Then   ' filled mask byte; write it, reset it
                                    pngData(mskPtr) = xDepth: mskPtr = mskPtr + 1&
                                    xDepth = 0&: mskShift = 128&
                                Else    ' set next shift position
                                    mskShift = mskShift \ 2&
                                End If
                            Next
                            If mskShift < 128& Then ' 99% of time, won't test true. Most icons are DWord aligned
                                ' write remainder, last mask bit. Set up for next scanline
                                pngData(mskPtr) = xDepth: xDepth = 0&: mskShift = 128&
                                mskPtr = xCy * mskScanWidth + mskScanWidth + mskPtrBase
                                clrPtr = xCy * xCx + xCx + clrPtrBase + 3&
                            End If
                        Next
                    End If  ' create icon from bitmap data & then create stdPicture from it
                     ' Cursors are a bit different... transfer XY hotspot to array as needed
                    If Not ICD.idType = 1 Then CopyMemory pngData(0), ICDE.wPlanes, 4&
                    hHandle = CreateIconFromResourceEx(pngData(0), lSize, ICD.idType And 1&, &H30000, dCx, dCy, 0&)
                    Erase pngData()
                End If
            End If
        End If
    End If
   
    '/// CLEAN UP & CACHE ORIGINAL DATA IF REQUESTED
    If hHandle Then
        Set pvLoadhIcon = pvHandleToStdPicture(hHandle, vbPicTypeIcon)
        If pvLoadhIcon Is Nothing Then
            If ICD.idType = 1& Then DestroyIcon hHandle Else DestroyCursor hHandle
        ElseIf m_hDC Then                               ' else class not initialized correctly
            If KeepFormat = True Then
                hHandle = 0&: Call pvIStreamFromPointer(0&, 22& + ICDE.dwBytesInRes, hHandle)
                If hHandle Then
                    CopyMemory ByVal pvSafePointerAdd(hHandle, 22&), icData(LBound(icData) + ICDE.dwImageOffset), ICDE.dwBytesInRes
                    ICDE.dwImageOffset = 22&: CopyMemory ByVal pvSafePointerAdd(hHandle, 6&), ICDE, 16&
                    ICD.idCount = 1: CopyMemory ByVal hHandle, ICD, 6&
                    Set pPicRef.oStream = pvIStreamFromPointer(0&, 0&, hHandle)
                    pPicRef.flags = ptcIcon ' manage if caching original format
                End If
            End If
            If (icSorted(Index) And &HFF) = 32& Then pPicRef.flags = ptcIcon ' always manage 32bpp
        End If
    End If
   
ExitRoutine:
End Function

Private Function pvCvAnyToICO(ByRef Picture As StdPicture, ByRef pPicRef As PICREF, ByRef KeepFormat As Boolean, ByRef IconSize As Long, ByRef DesiredIconCx As Long, ByRef DesiredIconCy As Long) As StdPicture

    ' called to convert bitmap, wmf, emf to icon
    ' if the loaded image was some other type (gif,tif,png), the LoadPicture method ensured that
    '   a bitmap would be returned before getting here. See LoadPicture

    Dim BHI As BITMAPV5HEADER
Dim icData() As Byte
Dim icHdr() As Long
Dim X As Long

    Dim mskShift As Long
Dim mskPtr As Long
Dim mskBits As Long
Dim lAlpha As Long


    If Picture.handle = 0& Then
        Set pvCvAnyToICO = Picture
       
    ElseIf Picture.Type = vbPicTypeIcon Then
        Set pvCvAnyToICO = Picture
       
    ElseIf Picture.Type = vbPicTypeBitmap Then
        ReDim icHdr(0 To 265)                               ' includes 256 color table
        icHdr(0) = 40&: GetDIBits m_hDC, Picture.handle, 0&, 0&, ByVal 0&, icHdr(0), 0&
        CopyMemory BHI, icHdr(0), 40&: Erase icHdr()        ' copy the header & release the array
        With BHI
            If .biBitCount = 0& Then
                Set pvCvAnyToICO = Picture
                Exit Function
            End If                                          ' begin calculating icon format size
            mskShift = pvDWordAlign(1, .biWidth)
            .biCompression = 0&: .biClrImportant = 0&: .biClrUsed = 0&
            If .biBitCount < 32 Then
                If Not .biBitCount = 24 Then .biBitCount = 24: .biSizeImage = 0&
            End If
            If .biSizeImage = 0& Then .biSizeImage = pvDWordAlign(.biBitCount, .biWidth, .biHeight)
            ReDim icData(0 To .biSizeImage + Abs(.biHeight) * mskShift + 61&)
        End With
        GetDIBits m_hDC, Picture.handle, 0&, Abs(BHI.biHeight), icData(62), BHI, 0&
        BHI.biHeight = Abs(BHI.biHeight)
        If BHI.biBitCount = 32 Then
            If pvValidateAlphaChannel(VarPtr(icData(62)), BHI.biWidth, BHI.biHeight) = PixelFormat32bppPremultiplied Then
                mskShift = 128: mskPtr = 62& + BHI.biSizeImage
                For X = 65& To mskPtr - 1& Step 4&
                    lAlpha = icData(X)
                    If lAlpha = 0& Then
                        mskBits = mskBits Or mskShift
                    ElseIf lAlpha < 255& Then
                        icData(X - 1&) = (icData(X - 1&) * lAlpha) \ 255
                        icData(X - 2&) = (icData(X - 2&) * lAlpha) \ 255
                        icData(X - 3&) = (icData(X - 2&) * lAlpha) \ 255
                    End If
                    If mskShift = 1& Then
                        icData(mskPtr) = mskBits
                        mskBits = 0&: mskShift = 128: mskPtr = mskPtr + 1&
                    Else
                        mskShift = mskShift \ 2&
                    End If
                Next
            End If
        End If
        With BHI
            icData(2) = 1: icData(4) = 1                    ' set icon directory type & count
            If .biWidth < 256& Then icData(6) = .biWidth
            If .biHeight < 256& Then icData(7) = .biHeight
            CopyMemory icData(14), CLng(UBound(icData) - 21&), 4&
            icData(10) = 1: icData(12) = BHI.biBitCount: icData(18) = 22 ' planes, bitcount & offset to where icon begins
            .biHeight = .biHeight + .biHeight
        End With
        CopyMemory icData(22), BHI, 40&                     ' copy header
        Set pvCvAnyToICO = pvLoadhIcon(icData(), DesiredIconCx, DesiredIconCy, lpsDefault, 0&, KeepFormat, pPicRef)
    Else
        Set pvCvAnyToICO = pvCvMetaFileToAny(Picture, pPicRef, vbPicTypeIcon, KeepFormat, IconSize, DesiredIconCx, DesiredIconCy)
    End If
   
    If pvCvAnyToICO Is Nothing Then Set pvCvAnyToICO = Picture

End Function

Private Function pvCvAnytoBMP(ByRef Picture As StdPicture, ByRef pPicRef As PICREF) As StdPicture

    ' convert bmp, wmf, emf, png, tiff, etc to bitmap

    Dim ICI() As Long
Dim aMask() As Byte
Dim hDIB As Long
Dim hBits As Long

    Dim X As Long
Dim Y As Long
Dim dPtr As Long
Dim hMask As Long

    Dim BIH As BITMAPV5HEADER
Dim mScanWidth As Long

   
    If Picture.handle = 0& Then
        Set pvCvAnytoBMP = Picture
       
    ElseIf Picture.Type = vbPicTypeBitmap Then
        If pPicRef.flags Then pPicRef.flags = vbPicTypeBitmap
        Set pvCvAnytoBMP = Picture
       
    ElseIf Picture.Type = vbPicTypeIcon Then
        ' icons are a bit of a pain to convert to bmp while maintaining transparency
       
        ReDim ICI(0 To 4)   ' equivalent to ICONINFO structure; (3)=mask,(4)=color
        GetIconInfo Picture.handle, ICI(0)
        If ICI(4) Then                              ' has color bitmap else B&W only
            BIH.biSize = 40&: GetDIBits m_hDC, ICI(4), 0&, 0&, ByVal 0&, BIH, 0&
            With BIH                                ' prepare a 32bpp format for DIB creation
                If .biBitCount Then
                    .biBitCount = 32: .biCompression = 0
                    hDIB = CreateDIBSection(m_hDC, BIH, 0&, hBits, 0&, 0&)
                End If
            End With
            If hDIB = 0& Then GoTo ExitRoutine      ' create DIB, abort on failure
                                                    ' transfer bitmap to our DIB
            GetDIBits m_hDC, ICI(4), 0&, BIH.biHeight, ByVal hBits, BIH, 0&
            DeleteObject ICI(4): ICI(4) = 0&        ' test for complex transparency
            Select Case pvValidateAlphaChannel(hBits, BIH.biWidth, BIH.biHeight)
            Case PixelFormat32bppAlpha              ' need to convert to premultiplied alpha
                Set pvCvAnytoBMP = pvLoadhBitmap(hDIB, pPicRef)
                GoTo ExitRoutine
            Case PixelFormat32bpp, PixelFormat32bppPremultiplied
                Set pvCvAnytoBMP = pvHandleToStdPicture(hDIB, vbPicTypeBitmap)
                pPicRef.flags = vbPicTypeBitmap
                GoTo ExitRoutine
            End Select                              ' otherwise, process the mask bits
        End If
       
        If ICI(3) Then                              ' has mask bits else oops
            hMask = ICI(3)                          ' cash mask bitmap
            ReDim ICI(0 To 265)                     ' BitmapInfoHeader + 256 color palette
            With BIH
                If .biBitCount = 0& Then            ' no color bitmap; therefore, this is B&W icon
                    .biSize = 40: GetDIBits m_hDC, hMask, 0&, 0&, ByVal 0&, BIH, 0&
                    If .biBitCount = 0 Then GoTo ExitRoutine
                    Y = .biHeight \ 2&              ' prepare a 32bpp format for DIB creation
                    .biBitCount = 32: .biClrImportant = 0: .biClrUsed = 0&
                    .biSizeImage = 0&: .biCompression = 0: .biHeight = Y
                    hDIB = CreateDIBSection(m_hDC, BIH, 0&, hBits, 0&, 0&)
                    If hDIB = 0& Then GoTo ExitRoutine ' create the DIB
                    .biHeight = Y + Y: X = Y        ' transfer top 1/2 of mask to our DIB
                    GetDIBits m_hDC, hMask, 0&, Y, ByVal hBits, BIH, 0&
                Else
                    Y = .biHeight: X = 0&
                End If                              ' get the mask bits
                .biBitCount = 8: .biClrImportant = 0: .biClrUsed = 256: .biSizeImage = 0
                mScanWidth = pvDWordAlign(.biBitCount, .biWidth)
                ReDim aMask(0 To mScanWidth * Y - 1&)
                CopyMemory ICI(0), BIH, 40&: ICI(11) = vbWhite
                GetDIBits m_hDC, hMask, X, Y, aMask(0), ICI(0), 0&
                Erase ICI()
            End With
            DeleteObject hMask: ICI(3) = 0&
            dPtr = 3&: For Y = 0& To Y - 1&         ' apply the mask to our bitmap
                X = Y * mScanWidth
                For X = X To X + BIH.biWidth - 1&
                    If aMask(X) = 0& Then
                        CopyMemory ByVal pvSafePointerAdd(hBits, dPtr), 255, 1&
                    Else
                        hMask = 0&                  ' contains simple transparency
                    End If
                    dPtr = dPtr + 4&
                Next
            Next
            Set pvCvAnytoBMP = pvHandleToStdPicture(hDIB, vbPicTypeBitmap)
            If hMask Then pPicRef.flags = vbPicTypeBitmap Else pPicRef.flags = hMask
        End If
   
    Else    ' metafile conversion to bitmap
        Set pvCvAnytoBMP = pvCvMetaFileToAny(Picture, pPicRef, vbPicTypeBitmap, False)
    End If

ExitRoutine:
    If Picture.Type = vbPicTypeIcon Then
        If ICI(4) Then DeleteObject ICI(4)
        If ICI(3) Then DeleteObject ICI(3)
    End If
    If pvCvAnytoBMP Is Nothing Then Set pvCvAnytoBMP = Picture
   
End Function

Private Function pvCvMetaFileToAny(ByRef Picture As StdPicture, ByRef pPicRef As PICREF, _
                                    ByVal PicType As Long, ByRef KeepFormat As Boolean, _
                                    Optional ByRef IconSize As Long, _
                                    Optional ByRef IconCx As Long, Optional ByRef IconCy As Long) As StdPicture

    ' convert wmf, emf to icon or bitmap

    Dim BIH As BITMAPV5HEADER
    Dim nPic As IPicture
    
    Dim hDIB As Long
    Dim hBits As Long
    Dim X As Long
    
    Dim bAlpha As Byte
    Dim dPtr As Long
    'Dim hIcon As Long
    Dim icData() As Byte
    
    Dim mskPtr As Long
    Dim mskShift As Long
    Dim mskBits As Long

   
    With BIH                                ' create DIB to work on
        Set nPic = Picture
        If (nPic.Attributes And 2&) Then .biBitCount = 32 Else .biBitCount = 24
        Set nPic = Nothing                  ' determine target size, metafiles scale exceptionally well
        .biWidth = (1440! / Screen.TwipsPerPixelX) * Picture.Width / 2540!
        .biHeight = (1440! / Screen.TwipsPerPixelY) * Picture.Height / 2540!
        If PicType = vbPicTypeIcon Then
            Call pvSetIconSize(IconSize, IconCx, IconCy, 1&)
            If IconCx Then .biWidth = IconCx
            If IconCy Then .biHeight = IconCy
        Else
        End If
        .biPlanes = 1: .biSize = 40         ' finish the header
        .biSizeImage = pvDWordAlign(.biBitCount, .biWidth, .biHeight)
    End With                                ' create DIB & abort on error
    hDIB = CreateDIBSection(m_hDC, BIH, 0&, hBits, 0&, 0&)
    If hDIB = 0& Then Exit Function
   
    ' WMF/EMF can be difficult to determine transparent areas. When Picture.Render is called,
    '   each pixel the metafile touches, clears the alpha byte. So, if we preset the entire
    '   image's bits to 255, then any that remain after rendering will be determined transparent
    '   because the metafile didn't touch those bits.
    If BIH.biBitCount = 32& Then FillMemory ByVal hBits, BIH.biSizeImage, 255
    hDIB = SelectObject(m_hDC, hDIB)
    With Picture
        .Render (m_hDC), 0&, 0&, (BIH.biWidth), (BIH.biHeight), 0&, (.Height - 1&), .Width, -.Height, ByVal 0&
    End With
    hDIB = SelectObject(m_hDC, hDIB)
   
    If PicType = vbPicTypeBitmap Then                   ' loop thru alpha bytes & toggle them
        If BIH.biBitCount = 32& Then
            For X = 3& To BIH.biSizeImage - 1& Step 4
                dPtr = pvSafePointerAdd(hBits, X)
                CopyMemory bAlpha, ByVal dPtr, 1
                bAlpha = bAlpha Xor 255
                If bAlpha = 255 Then
                    CopyMemory ByVal dPtr, bAlpha, 1
                Else
                    CopyMemory ByVal pvSafePointerAdd(hBits, X - 3&), 0&, 4&
                End If
            Next
            pPicRef.flags = PicType
        Else
            pPicRef.flags = 0&
        End If
        Set pvCvMetaFileToAny = pvHandleToStdPicture(hDIB, vbPicTypeBitmap)
        If pvCvMetaFileToAny Is Nothing Then
            DeleteObject hDIB: pPicRef.flags = 0&
        End If
    Else
        With BIH                                        ' prepare array to hold icon format
            mskShift = pvDWordAlign(1, .biWidth)
            ReDim icData(0 To .biSizeImage + mskShift * .biHeight + 21&)
            icData(2) = 1: icData(4) = 1                    ' set icon directory type & count
            If .biWidth < 256& Then icData(6) = .biWidth
            If .biHeight < 256& Then icData(7) = .biHeight
            CopyMemory icData(14), CLng(UBound(icData) - 21&), 4&
            icData(10) = 1: icData(12) = 32: icData(18) = 22 ' planes, bitcount & offset to where icon begins
            .biHeight = .biHeight + .biHeight
        End With
        CopyMemory icData(22), BIH, 40&                     ' copy header
        CopyMemory icData(62), ByVal hBits, BIH.biSizeImage ' copy bits
        DeleteObject hDIB
        If BIH.biBitCount = 32& Then                        ' loop and toggle alpha + build icon mask
            mskPtr = BIH.biSizeImage + 62&: mskShift = 128&
            For X = 65& To mskPtr - 1& Step 4&
                icData(X) = icData(X) Xor 255
                If icData(X) = 0 Then mskBits = mskBits Or mskShift
                If mskShift = 1& Then
                    icData(mskPtr) = mskBits
                    mskBits = 0&: mskShift = 128&
                    mskPtr = mskPtr + 1&
                Else
                    mskShift = mskShift \ 2&
                End If
            Next
        End If
        Set pvCvMetaFileToAny = pvLoadhIcon(icData(), 0&, 0&, lpsDefault, 0&, KeepFormat, pPicRef)
    End If

End Function

Private Sub pvSetIconSize(ByRef Size As Long, ByRef dCx As Long, ByRef dCy As Long, ByVal iconCount As Long)

    ' determine icons size from passed options

    Dim hIML As IUnknown
    Dim GUID(0 To 3) As Long

    Const IID_IImageList As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
   
    If Size = lpsCustom Then                    ' DesiredCx,Cy should not be zero
        If dCx < 1& And dCy < 1& Then
            Size = lpsDefault                   ' if they are, use Default
        Else
            If dCx < 1& Then dCx = dCy
            If dCy < 1& Then dCy = dCx
            Exit Sub                            ' done
        End If
    End If
    If Size = lpsDefault Then                   ' if only 1 icon, use actual size
        If iconCount > 1& Then
            Size = lpsLarge
        Else
            dCx = 0&: dCy = dCx
            Exit Sub                            ' done
        End If
    ElseIf Abs(Size) = lpsSmallShell Or Abs(Size) = lpsLargeShell Then
        If Size = lpsJumboShell Then            ' Jumbo available on Vista+ only
            If m_VistaPlus = False Then
                dCx = 256&: dCy = dCx
                Exit Sub                        ' done
            End If
        End If
        IIDFromString StrPtr(IID_IImageList), GUID(0)
    End If
    On Error Resume Next
    Select Case Size
    Case lpsLarge                               ' generally 32x32 can be 48x48
        dCx = GetSystemMetrics(11): dCy = GetSystemMetrics(12)
    Case lpsSmall                               ' always 16x16 as of Win10
        dCx = GetSystemMetrics(49): dCy = GetSystemMetrics(50)
    Case lpsSmallShell                          ' 16x16 but can scale due to DPI
        SHGetImageListXP 1&, GUID(0), VarPtr(hIML)
    Case lpsLargeShell                          ' 32x32 but can scale due to DPI
        SHGetImageListXP 0&, GUID(0), VarPtr(hIML)
    Case lpsXtraLargeShell                      ' 48x48 but can scale due to DPI
        SHGetImageListXP 2&, GUID(0), VarPtr(hIML)
    Case lpsJumboShell                          ' always 256x256 as of Win10
        SHGetImageListXP 4&, GUID(0), VarPtr(hIML)
    End Select
    If Not hIML Is Nothing Then ImageList_GetIconSize ObjPtr(hIML), dCx, dCy
    If Err Then
        Select Case Size
        Case lpsLarge: dCx = 32&: dCy = dCx
        Case lpsSmall: dCx = 16&: dCy = dCx
        Case lpsSmallShell
            dCx = GetSystemMetrics(49): dCy = GetSystemMetrics(50)
        Case lpsLargeShell
            dCx = GetSystemMetrics(11): dCy = GetSystemMetrics(12)
        Case lpsXtraLargeShell: dCx = 48&: dCy = dCx
        Case lpsJumboShell: dCx = 256&: dCy = dCx
        End Select
    End If

End Sub

Private Function pvConvertFromGDIplus(ByRef hImage As Long, ByRef IStream As IUnknown, ByVal KeepFormat As Boolean, ByRef pPicRef As PICREF) As StdPicture

    ' takes a GDI+ image (bitmap, PNG, TIFF, etc) and converts it to a GDI image then to a stdPicture
    ' the GDI+ image is only ever kept under 2 conditions
    '   1) It is a multi-frame GIF or multi-page TIFF and
    '   2) Navigation is desired

    Dim uBMP As BitmapData
Dim srcRect As RECTF
Dim rBounds As RECTI

    Dim BIH As BITMAPV5HEADER
Dim hLock As Long
Dim gStream As IUnknown

    Dim vParamPtr() As Long
Dim vParamType() As Integer

    Dim vParams() As Variant
Dim paramCount As Long
Dim hHandle As Long

   
    If hImage = 0& Then
        GdipLoadImageFromStream ObjPtr(IStream), hImage
        If hImage = 0& Then Exit Function
    End If
    pPicRef.flags = pvQueryFormat(hImage)
    Select Case pPicRef.flags
    Case ptcBitmap: KeepFormat = False
    Case ptcPNG, ptcJPEG
    Case ptcGIF
        pPicRef.flags = pPicRef.flags Or pvGetFrameCount(hImage, pPicRef.flags) * &H10000 Or 1&
        If (pPicRef.flags And &HFFF0000) > &H10000 And KeepFormat = True Then
            pPicRef.pHandle = hImage ' allow GDI+ to manage this for frame navigation
        Else    ' if single frame or not caching format, let VB manage it
            GdipDisposeImage hImage: hImage = 0&
            ' specifically set the stream seek pointer to 0, otherwise pvIStreamToPicture fails here.
            ' IStream:Seek is VTable offset of 20, params are  bytesToMove@,origin&,*newPosition@
            paramCount = 3&
            ReDim vParams(0 To paramCount)
            ReDim vParamPtr(0 To paramCount)
            ReDim vParamType(0 To paramCount)
            vParams(0) = 0@: vParams(1) = 0&: vParams(2) = 0&
            For paramCount = 0& To paramCount - 1&
                vParamPtr(paramCount) = VarPtr(vParams(paramCount))
                vParamType(paramCount) = VarType(vParams(paramCount))
            Next
            Call DispCallFunc(ObjPtr(IStream), 20&, 4&, vbLong, paramCount, _
                    VarPtr(vParamType(0)), VarPtr(vParamPtr(0)), vParams(paramCount))
            Set pvConvertFromGDIplus = pvIStreamToPicture(IStream, KeepFormat)
            pPicRef.flags = 0&
            Exit Function
        End If
    Case ptcTIF ' if keeping format, allow GDI+ to manage this for page navigation
        If KeepFormat Then
            pPicRef.flags = pPicRef.flags Or pvGetFrameCount(hImage, pPicRef.flags) * &H10000 Or 1&
            If (pPicRef.flags And &HFFF0000) > &H10000 And KeepFormat = True Then
                pPicRef.pHandle = hImage
            End If
        End If
    Case ptcEMetafile   ' likely only reason we got here is non-placeable WMF was loaded
        GdipGetHemfFromMetafile hImage, hHandle                ' convert to GDI WMF
        GdipDisposeImage hImage: hImage = 0&: pPicRef.flags = 0&
        paramCount = GdipEmfToWmfBits(hHandle, 0&, 0&, 8&, 0&) ' 8& = MM_ANISOTROPIC; get size of bits
        If paramCount Then
            Call pvIStreamFromPointer(0&, paramCount + 22&, hLock)
            If hLock Then                                       ' transfer bits, offsetting for placeable header
                If GdipEmfToWmfBits(hHandle, paramCount, pvSafePointerAdd(hLock, 22&), 8&, 0&) Then
                    DeleteEnhMetaFile hHandle: hHandle = 0&
                    If screenWidthTwips > Screen.Height Then
                        rBounds.nWidth = 256& * Screen.TwipsPerPixelX
                        rBounds.nHeight = (Screen.Height / screenWidthTwips) * rBounds.nWidth
                    Else
                        rBounds.nHeight = 256& * Screen.TwipsPerPixelY
                        rBounds.nWidth = (screenWidthTwips / Screen.Height) * rBounds.nHeight
                    End If                                                              ' add placeable header
                    CopyMemory ByVal hLock, &H9AC6CDD7, 4&                              ' WMF magic number
                    CopyMemory ByVal pvSafePointerAdd(hLock, 4&), 0@, 6&                ' Handle, Left, Top
                    CopyMemory ByVal pvSafePointerAdd(hLock, 10&), rBounds.nWidth, 2&   ' Right
                    CopyMemory ByVal pvSafePointerAdd(hLock, 12&), rBounds.nHeight, 2&  ' Bottom
                    CopyMemory ByVal pvSafePointerAdd(hLock, 14&), 1440, 2&             ' units per inch
                    CopyMemory ByVal pvSafePointerAdd(hLock, 16&), 0&, 4&               ' reserved
                    paramCount = 22289& Xor rBounds.nWidth Xor rBounds.nHeight Xor 1440&
                    ' ^^ 22289 = 0 Xor &HCDD7& Xor &H9AC6&
                    CopyMemory ByVal pvSafePointerAdd(hLock, 20&), paramCount, 2&       ' checksum
                    Set gStream = pvIStreamFromPointer(0&, 0&, hLock): hLock = 0&
                    Set pvConvertFromGDIplus = pvIStreamToPicture(gStream, False)
                End If
                If hLock Then Set gStream = pvIStreamFromPointer(0&, 0&, hLock)
                Set gStream = Nothing
            End If
        End If
        If hHandle Then DeleteEnhMetaFile hHandle
        Exit Function
    Case Else
        GdipDisposeImage hImage: hImage = 0&
        Exit Function
    End Select
    
    GdipGetImageBounds hImage, srcRect, 2& ' 2 = UnitPixel
    With uBMP
        .Height = srcRect.nHeight: .Width = srcRect.nWidth
        rBounds.nHeight = .Height: rBounds.nWidth = .Width
        If KeepFormat Then
            .PixelFormat = PixelFormat32bppPremultiplied ' return premultiplied RGB
        Else
            GdipGetImagePixelFormat hImage, .PixelFormat
            Select Case .PixelFormat
            Case PixelFormat32bppAlpha, PixelFormat32bppPremultiplied, &H61007, &H34400D, &H1C400E ' all known GDI+ alpha formats
                .PixelFormat = PixelFormat32bppPremultiplied ' return premultiplied RGB
                pPicRef.flags = ptcBitmap
            Case Else
                .PixelFormat = PixelFormat24bpp  ' return 24bpp format
            End Select
        End If
    End With
   
    If GdipBitmapLockBits(hImage, rBounds, ImageLockModeRead, uBMP.PixelFormat, uBMP) Then
        Call GdipBitmapLockBits(hImage, rBounds, ImageLockModeRead, uBMP.PixelFormat, uBMP)
    End If
    If uBMP.Scan0Ptr Then
        With BIH                        ' build a bitmap info header structure
            .biHeight = -uBMP.Height
            .biWidth = uBMP.Width
            .biPlanes = 1: .biSize = 40
            If uBMP.PixelFormat = PixelFormat32bppPremultiplied Then .biBitCount = 32 Else .biBitCount = 24
            .biSizeImage = Abs(uBMP.stride * .biHeight)
        End With
        hLock = 0&: Call pvIStreamFromPointer(0&, 54& + BIH.biSizeImage, hLock)
        If hLock Then                   ' wrap pixel data in bitmap file & info headers
            CopyMemory ByVal hLock, &H4D42, 2&
            CopyMemory ByVal pvSafePointerAdd(hLock, 2), CLng(54& + BIH.biSizeImage), 4&
            CopyMemory ByVal pvSafePointerAdd(hLock, 6), 0&, 4&
            CopyMemory ByVal pvSafePointerAdd(hLock, 10), 54&, 4&
            CopyMemory ByVal pvSafePointerAdd(hLock, 14), BIH, 40&
            CopyMemory ByVal pvSafePointerAdd(hLock, 54), ByVal uBMP.Scan0Ptr, BIH.biSizeImage
            GdipBitmapUnlockBits hImage, uBMP
            Set gStream = pvIStreamFromPointer(0&, 0&, hLock)
            Set pvConvertFromGDIplus = pvIStreamToPicture(gStream, False)
        Else
            GdipBitmapUnlockBits hImage, uBMP
        End If
    End If
   
    If pvConvertFromGDIplus Is Nothing Then
        If hImage Then GdipDisposeImage hImage
        Set pPicRef.oStream = Nothing
        pPicRef.flags = 0&
    Else
        If pPicRef.pHandle = 0& Then GdipDisposeImage hImage
        If KeepFormat Then Set pPicRef.oStream = IStream
    End If
   
End Function

Private Function pvIStreamToPicture(ByRef IStream As IUnknown, ByVal KeepFormat As Boolean) As IPicture
   
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty before this was called
   
    Dim aGUID(0 To 3) As Long
    On Error Resume Next
    If Not IStream Is Nothing Then
        aGUID(0) = &H7BF80980    ' GUID for IPICTURE
        aGUID(1) = &H101ABF32
        aGUID(2) = &HAA00BB8B
        aGUID(3) = &HAB0C3000
        Call OleLoadPicture(ByVal ObjPtr(IStream), 0&, Abs(Not KeepFormat), aGUID(0), pvIStreamToPicture)
    End If
End Function

Private Function pvHandleToStdPicture(ByVal hImage As Long, ByRef PicType As PictureTypeConstants) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)
   
    Dim lpPictDesc(0 To 3) As Long
    Dim aGUID(0 To 3) As Long

   
    lpPictDesc(0) = 16&
    lpPictDesc(1) = PicType
    lpPictDesc(2) = hImage
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, pvHandleToStdPicture)
   
End Function

Private Function pvIStreamFromPointer(ByVal MemoryPtr As Long, ByRef Length As Long, ByRef hLock As Long) As stdole.IUnknown
   
    ' Purpose: Create an IStream-compatible IUnknown interface containing the
    ' passed byte aray. This IUnknown interface can be passed to GDI+ functions
    ' that expect an IStream interface -- neat hack
   
    ' Can be called either once or twice
    ' If calling once: MemoryPtr cannot be null, hLock must be null
    ' If calling twice:
    '   1st call: pass MemoryPtr as zero, valid Length of data, hLock as null
    '   2nd call: pass MemoryPtr & Length as zero, hLock as value returned from 1st call
   
    On Error GoTo HandleError
    Dim o_lpMem  As Long
    'Dim tUnk As IUnknown

    Static o_hMem As Long
    If hLock Then
        Call GlobalUnlock(o_hMem)
        Call CreateStreamOnHGlobal(o_hMem, 1&, pvIStreamFromPointer)
    Else
        hLock = 0&
        If Length > 0& Then
            o_hMem = GlobalAlloc(&H2&, Length)
            If o_hMem Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem Then
                    If MemoryPtr Then
                        CopyMemory ByVal o_lpMem, ByVal MemoryPtr, Length
                        Call GlobalUnlock(o_hMem)
                        Call CreateStreamOnHGlobal(o_hMem, 1&, pvIStreamFromPointer)
                    Else
                        hLock = o_lpMem
                    End If
                Else
                    GlobalFree o_hMem: o_hMem = 0&
                End If
            End If
        End If
    End If
   
HandleError:
End Function

Private Function pvQueryFormat(ByRef hImage As Long) As Long
    ' function returns: 0, 1=gif, 2=tif

    Dim GUID(0 To 3) As Long
    If GdipGetImageRawFormat(hImage, VarPtr(GUID(0))) = 0& Then
        Select Case GUID(0)
            Case &HB96B3CB0: pvQueryFormat = ptcGIF
            Case &HB96B3CB1: pvQueryFormat = ptcTIF
            Case &HB96B3CAB: pvQueryFormat = ptcBitmap ' bitmap file
            Case &HB96B3CAA: pvQueryFormat = ptcBitmap ' memory bitmap (Scan0)
            Case &HB96B3CAE: pvQueryFormat = ptcJPEG
            Case &HB96B3CAF: pvQueryFormat = ptcPNG
            Case &HB96B3CAC: pvQueryFormat = ptcEMetafile
        End Select
    End If
End Function

Private Function pvGetFrameCount(ByRef hImage As Long, ByVal PicType As Long) As Long
   ' returns the number of frames/pages of the passed GIF/TIF
    If PicType = ptcGIF Then
        Dim lSize As Long
        If m_PageGUID(0, 0) = 0& Then
            GdipImageGetFrameDimensionsCount hImage, lSize
            If lSize > 0& Then GdipImageGetFrameDimensionsList hImage, m_PageGUID(0, 0), 1&
        End If
        GdipImageGetFrameCount hImage, m_PageGUID(0, 0), pvGetFrameCount
    ElseIf PicType = ptcTIF Then
        Const FrameDimensionPage As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
        If m_PageGUID(0, 1) = 0& Then IIDFromString StrPtr(FrameDimensionPage), m_PageGUID(0, 1)
        GdipImageGetFrameCount hImage, m_PageGUID(0, 1), pvGetFrameCount
    End If
   
End Function

Private Sub pvDisposeGDIpImages()
    ' can be called during both Class_Initialize & Class_Terminate
    ' clears any cached GDI+ images
    Dim lCount As Long
Dim pImages As Long
Dim hImage As Long

   
    If m_ThunkPtr Then
        CopyMemory pImages, ByVal pvSafePointerAdd(m_ThunkPtr, 12&), 4&
        If pImages Then                     ' If we have an array, walk it
            CopyMemory lCount, ByVal pImages, 4&
            For lCount = 0& To lCount - 1&
                CopyMemory hImage, pvSafePointerAdd(pImages, lCount * 4& + 8&), 4&
                If hImage Then GdipDisposeImage hImage ' destroy any handle in the array
            Next                            ' release the array & update pointer on thunk
            CopyMemory ByVal pvSafePointerAdd(m_ThunkPtr, 12&), 0&, 4&
            CoTaskMemRealloc pImages, 0&
        End If
    End If

End Sub

Private Function pvReverseLong(ByVal inLong As Long) As Long

    ' fast function to reverse a long value from big endian to little endian
    ' PNG files contain reversed longs, as do ID3 v3,4 tags
    pvReverseLong = _
      (((inLong And &HFF000000) \ &H1000000) And &HFF&) Or _
      ((inLong And &HFF0000) \ &H100&) Or _
      ((inLong And &HFF00&) * &H100&) Or _
      ((inLong And &H7F&) * &H1000000)
    If (inLong And &H80&) Then pvReverseLong = pvReverseLong Or &H80000000
End Function

Private Function pvValidateAlphaChannel(ByVal PixelPointer As Long, ByVal Width As Long, _
                                        ByVal Height As Long) As Long
       
    ' Method determines if the alhpa channel is used and how it is used
    ' Only supports 32 bpp pixel data. Passing any other format will result in a crash
    ' Assumption is that you will only call this routine to test 32 bpp data
    ' Speedy? For complex transparency, should exit routine very fast. For anything else, all
    '   pixels will be scanned unfortuntely. Just can't exit until format is known for sure.
   
    Dim bAlpha As Long
Dim lFormat As Long

    Dim lColor As Long
Dim lPrevColor As Long

    Dim bScanLine() As Long
Dim Y As Long
Dim X As Long

    Const ZEROES As Long = 256&
   
    ' lFormat will contain one or more of these
    '    acuAllBlack = 0             ' image can be interpreted as 100% black or 100% transparent
    '    acuOpaqueAssumed = 1        ' all alpha values are zero, assuming image is not meant to be 100% transparent
    '    acuOpaque = 2               ' alpha channel is used, but all alpha values are 255
    '    acuSimpleTransparency = 4   ' alpha channel is used and contains simple transparency only
    '    acuComplexTransparency = 8  ' alpha channel is used and contains complex transparency
    '    acuPremultipliedRGB = 16    ' R,G,B components are multiplied against the alpha channel
    On Error GoTo ExitRoutine
   
    ' ///// test the alpha channel. Loop aborts early if it can
    ReDim bScanLine(0 To Width - 1&)
    CopyMemory lPrevColor, ByVal PixelPointer, 4&
    lPrevColor = lPrevColor Xor 1&
   
    For Y = 0& To Height - 1&
        X = Width * 4&
        CopyMemory bScanLine(0), ByVal pvSafePointerAdd(PixelPointer, X * Y), X
        For X = 0& To Width - 1&
            lColor = bScanLine(X)               ' get 32bit color
            If Not lColor = lPrevColor Then     ' and extact the alpha byte
                If lColor = 0& Then
                    lFormat = lFormat Or ZEROES ' entire value is zero
                    ' all zeroes indicates 100% transparent or 100% black image
                    ' mix of zero & non-zero alpha values indicates transparency
                Else
                    bAlpha = (lColor And &HFF000000) \ &H1000000 And &HFF
                    If bAlpha = 0& Then
                        If (lColor And &HFFFFFF) Then  ' RGB value is non-zero
                            If (lFormat And Not ZEROES) > 2& Then
                                ' at least one other alpha value was > 0 and < 255
                                ' since this alpha is zero & RGB non-zero. Done:
                                lFormat = 8&: Y = Height: Exit For
                            End If
                            lFormat = lFormat Or 1& ' keep going, maybe all alphas are zero
                        End If
                    ElseIf bAlpha = 255& Then
                        If (lFormat And 1&) Then
                            ' already seen alpha zero & non-zero RGB. Here we have 255 alpha. Done:
                            lFormat = 8&: Y = Height: Exit For
                        End If
                        lFormat = lFormat Or 2&
                       
                    ' else if any RGB values > alpha then not-premultiplied
                    ElseIf bAlpha < (lColor And &HFF&) Then
                        lFormat = 8&: Y = Height: Exit For ' definitly ARGB
                    ElseIf bAlpha < (lColor And &HFF00&) \ &H100& Then
                        lFormat = 8&: Y = Height: Exit For ' definitly ARGB
                    ElseIf bAlpha < (lColor And &HFF0000) \ &H10000 Then
                        lFormat = 8&: Y = Height: Exit For ' definitly ARGB
                    Else
                        lFormat = lFormat Or 16& ' likely pARGB, but not sure yet
                    End If
                End If
                lPrevColor = lColor
            End If
        Next
    Next
    ' ///// Step 4: Analyze result
    If (lFormat And Not ZEROES) = 8& Then   ' alpha, not premultiplied
        pvValidateAlphaChannel = PixelFormat32bppAlpha
    ElseIf lFormat = (ZEROES Or 2&) Then    ' simple transparency
        pvValidateAlphaChannel = PixelFormat32bpp
    ElseIf (lFormat And 16&) Then           ' premultiplied RGB
        pvValidateAlphaChannel = PixelFormat32bppPremultiplied
    End If                                  ' else assume alpha channel not used
   
ExitRoutine:
End Function

Private Sub pvAddGDIpItem(ByRef pPicRef As PICREF)

    ' called when the class will be managing a picture object
   
    Dim pNew As Long
Dim pImages As Long

    Dim lCount As Long
Dim lSize As Long


    If pPicRef.pHandle Then
        CopyMemory pImages, ByVal pvSafePointerAdd(m_ThunkPtr, 12&), 4&
        If pImages Then                                     ' get size of array & count of handles
            CopyMemory lSize, ByVal pvSafePointerAdd(pImages, 4&), 4&
            CopyMemory lCount, ByVal pImages, 4&
            If lCount = lSize Then                          ' array is full
                lSize = lSize + 3&                          ' increase array with buffer of 2 spare slots
            Else
                lSize = 0&                                  ' array does not need to be resized
            End If
            lCount = lCount + 1&                            ' increment count
        Else
            lSize = 3&: lCount = 1&                         ' new array, init size for 3 handle slots
        End If
        If lSize Then                                       ' resizing.... attempt it
            pNew = CoTaskMemRealloc(pImages, lSize * 4& + 8&)
            If pNew = 0& Then Exit Sub                      ' if failure, we can't add it
            pImages = pNew                                  ' otherwise...
            CopyMemory ByVal pvSafePointerAdd(pImages, 4&), lSize, 4&       ' update the array size
            CopyMemory ByVal pvSafePointerAdd(m_ThunkPtr, 12&), pImages, 4& ' update the array pointer
        End If
        CopyMemory ByVal pImages, lCount, 4&                ' update the handle count, add new handle
        CopyMemory ByVal pvSafePointerAdd(pImages, lCount * 4& + 4&), pPicRef.pHandle, 4&
    End If
    If m_RefCount = 0& Then                                 ' resize local array as needed
        ReDim m_PicRefs(-1 To 2)                            ' FYI: local array is lookup & IStream cache
    ElseIf m_RefCount > UBound(m_PicRefs) Then
        ReDim Preserve m_PicRefs(-1 To m_RefCount + 2&)
    End If
    m_PicRefs(m_RefCount) = pPicRef                         ' add to our local array
    m_RefCount = m_RefCount + 1&                            ' increment local array count

End Sub

Private Function pvCreateThunks() As Boolean

    ' This class method is only ever called once
    ' It will set the m_ThunkPtr to one of these values:
    '   -1: indicating class failed to create thunk or some other error prevented it
    ' non-zero: The thunk address created by this class or existing in another instance of this class

    '@Ignore ConstantNotUsed
    Const THUNK_LONGS   As Long = 110&
    Const PAGE_RWX      As Long = &H40&
    Const MEM_COMMIT    As Long = &H1000&
    Const MEM_RELEASE   As Long = &H8000&
    Const GWL_WNDPROC   As Long = -4&
    Const WS_CHILD As Long = &H40000000

    Dim z_Code(0 To THUNK_LONGS) As Long
    Dim hMod As Long
    Dim vThunks(0 To 2) As Long

    Dim MEM_LEN As Long
    Dim MBI() As Long

    Dim hGDIplus As Long
    Dim hMsImg As Long
    Dim hToken As Long
    
    Dim tPic As StdPicture
    Dim nIPic As IPicture
    
    Dim vIPic As Long
    Dim vIPicDisp As Long
    Dim Index As Long
    
    Dim hWnd As Long
    Dim tObj As Object

   
    If m_Hwnd = 0& Then                     ' complete failure. Could not locate the VB owner window
        m_ThunkPtr = -1&
        Exit Function
    End If
   
    hWnd = m_Hwnd: m_Hwnd = 0&              ' look for a child window we created on the owner window
    m_Hwnd = FindWindowEx(hWnd, 0&, "Static", "IPIC+Thunker")
    If m_Hwnd = 0& Then                     ' if not found, create it now & abort if failure
        m_Hwnd = CreateWindowEx(0&, "Static", "IPIC+Thunker", WS_CHILD, 0&, 0&, 0&, 0&, hWnd, 0&, 0&, ByVal 0&)
        If m_Hwnd = 0& Then m_ThunkPtr = -1&: Exit Function
        ' thunks will now be created below
    Else                                    ' otherwise, get the thunk pointer and other info
        m_ThunkPtr = GetProp(m_Hwnd, "ThunkAddr")
        If m_ThunkPtr = -1& Then Exit Function
        CopyMemory hToken, ByVal m_ThunkPtr, 4&
        If hToken Then                      ' is there an active thunk class out there?
            CopyMemory tObj, hToken, 4&     ' if so, create a reference to it
            Set m_Primary = tObj            ' and if it's this class, clear the extra reference
            CopyMemory tObj, 0&, 4&
            If m_Primary Is Me Then Set m_Primary = Nothing
        Else                                ' else this class will be the active one. Update thunk variables
            CopyMemory ByVal m_ThunkPtr, ObjPtr(Me), 4&
            CopyMemory ByVal pvSafePointerAdd(m_ThunkPtr, 4&), zAddressOf(Me, 1), 4&
        End If
        pvCreateThunks = True
        Exit Function
    End If
   
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Thunk #1: Management Window's new window procedure (Subclass.asm file included with project)
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Thunk starts @ 15*4 bytes after vThunks(0)
    MEM_LEN = 4& * 65& ' 65 = last filled z_Code() member plus 1
    vThunks(0) = VirtualAlloc(0&, MEM_LEN, MEM_COMMIT, PAGE_RWX)
    If vThunks(0) = 0& Then GoTo ExitRoutine
   
    z_Code(15) = &HD231C031: z_Code(16) = &HBBE58960: z_Code(18) = &HFF90F631: z_Code(19) = &H75FF3075: z_Code(20) = &H2875FF2C: z_Code(21) = &H682475FF
    z_Code(23) = &H891C53FF: z_Code(24) = &H1FB81C45: z_Code(25) = &H39000000: z_Code(26) = &H31752845: z_Code(27) = &H1F307D83: z_Code(28) = &H8A850F
    z_Code(29) = &H33390000: z_Code(30) = &H82840F: z_Code(31) = &H73390000: z_Code(32) = &HFF107410: z_Code(33) = &HF8831053: z_Code(34) = &H85087401
    z_Code(35) = &H897175C0: z_Code(36) = &HFF6DEB33: z_Code(37) = &H33FF2C75: z_Code(38) = &HEB0453FF: z_Code(39) = &H287D8363: z_Code(40) = &H8B5D7502
    z_Code(41) = &H78B087B: z_Code(42) = &H2453FF50: z_Code(43) = &H680C478B: z_Code(44) = &H8000&: z_Code(45) = &H53FF5056: z_Code(46) = &H80006814
    z_Code(47) = &HFF560000: z_Code(48) = &H53FF0873: z_Code(49) = &HC733914: z_Code(50) = &H7B8B1F74: z_Code(51) = &H390F8B0C: z_Code(52) = &H8B0E74F1
    z_Code(53) = &H39048F44: z_Code(54) = &H510674F0: z_Code(55) = &H2853FF50: z_Code(56) = &HFFF0E259: z_Code(57) = &H53FF0C73: z_Code(58) = &H30733920
    z_Code(59) = &H73FF0674: z_Code(60) = &H2C53FF30: z_Code(61) = &HFF3473FF: z_Code(62) = &H73FF1853: z_Code(63) = &H1853FF38: z_Code(64) = &H10C261

    z_Code(0) = ObjPtr(Me)
    z_Code(1) = zAddressOf(Me, 1)
    'z_Code(2)  Thunk#2 address, to be filled in later
    'z_Code(3)  filled in dynamically, as needed: gdi+ image handle table
    On Error Resume Next
    DebugPrint 1 / 0   ' when in IDE, prevent calling back to this class
    If Err Then         ' when the IDE is stopped, paused, or ended
        Err.Clear       ' any VB5 users still out there? If so, change vba6 below to vba5
        z_Code(4) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
    End If
    On Error GoTo 0
    hMod = GetModuleHandle("Kernel32.dll")  ' function pointers used by the thunk
        z_Code(5) = GetProcAddress(hMod, "VirtualFree")
        z_Code(6) = GetProcAddress(hMod, "FreeLibrary")
    z_Code(7) = GetProcAddress(GetModuleHandle("User32.dll"), "CallWindowProcA")
    z_Code(8) = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemFree")
    z_Code(9) = GetProcAddress(GetModuleHandle("gdi32.dll"), "DeleteDC")
    ReDim MBI(0 To 3): MBI(0) = 1&      ' MBI() equivalent to GDI+ Startup structure
    Call GdiplusStartup(hToken, MBI(0))
    If hToken Then
        hGDIplus = LoadLibrary("GdiPlus.dll") ' Load GDI+ into process & cache its instance
        z_Code(10) = GetProcAddress(hGDIplus, "GdipDisposeImage")
        z_Code(11) = GetProcAddress(hGDIplus, "GdiplusShutdown")
        z_Code(12) = hToken
        z_Code(13) = hGDIplus
    End If
    hMsImg = LoadLibrary("msimg32.dll")       ' Load msimg32 into process and cache its instance
    z_Code(14) = hMsImg
    z_Code(17) = vThunks(0)
    z_Code(22) = GetWindowLong(m_Hwnd, GWL_WNDPROC)
   
    ' sanity checks
    For Index = 0& To 14&
        If z_Code(Index) = 0& Then
            If Index < 2& Or Index > 4& Then ' z_Code(2)(3) done later, z_Code(4) applies to IDE only
                VirtualFree vThunks(0), 0&, MEM_RELEASE
                vThunks(0) = 0&
                GoTo ExitRoutine
            End If
        End If
    Next
    CopyMemory ByVal vThunks(0), z_Code(0), MEM_LEN
   
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Thunk #2: IPicture & IPictureDisp thunks (Vtable.asm file included with project)
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    MEM_LEN = 4& * 105& ' 105 = last filled z_Code() member plus 1
    vThunks(1) = VirtualAlloc(0&, MEM_LEN, MEM_COMMIT, PAGE_RWX)
    If vThunks(1) = 0& Then GoTo ExitRoutine
   
    Erase z_Code()
    ' IPictureDisp IUnknown:Release thunk. Starts @ 14*4 bytes after vThunks(1)
    z_Code(14) = &HD231C031: z_Code(15) = &HBBE58960: z_Code(17) = &H7B8BF631: z_Code(18) = &H2475FF04: z_Code(19) = &H890857FF
        z_Code(20) = &H75391C45: z_Code(21) = &H6A0D751C: z_Code(22) = &H2475FF1F: z_Code(23) = &H73FF1F6A: z_Code(24) = &H1C53FF10: z_Code(25) = &H4C261
    ' IPicture IUnknown:Release thunk. Starts @ 26*4 bytes after vThunks(1)
    z_Code(26) = &HD231C031: z_Code(27) = &HBBE58960: z_Code(29) = &H7B8BF631: z_Code(30) = &H90CEEB08
    ' IPicture:GetAttributes thunk. Starts @ 31*4 bytes after vThunks(1)
    z_Code(31) = &HD231C031: z_Code(32) = &HBBE58960: z_Code(34) = &H7539F631: z_Code(35) = &H8B0B7428: z_Code(36) = &H7C7287D
        z_Code(37) = &H2&: z_Code(38) = &H45C707EB: z_Code(39) = &H700571C: z_Code(40) = &H8C26180: z_Code(41) = &H90909000
    ' IPicture:Render thunk. Starts @ 42*4 bytes after vThunks(1)
    z_Code(42) = &HD231C031: z_Code(43) = &HBBE58960: z_Code(45) = &H738BF631: z_Code(46) = &H2C7B8D08: z_Code(47) = &H2475FF57: z_Code(48) = &H830C56FF
        z_Code(49) = &HF002C7B: z_Code(50) = &HD484&: z_Code(51) = &H347B8D00: z_Code(52) = &H2475FF57: z_Code(53) = &H8D1C56FF: z_Code(54) = &HFF57307B
        z_Code(55) = &H56FF2475: z_Code(56) = &H30638314: z_Code(57) = &H307B8307: z_Code(58) = &H97840F03: z_Code(59) = &H83000000: z_Code(60) = &HF01307B
        z_Code(61) = &HA885&: z_Code(62) = &H75FF5700: z_Code(63) = &H2856FF24: z_Code(64) = &H7339F631: z_Code(65) = &H8B117530: z_Code(66) = &H30438903
        z_Code(67) = &HFF2C73FF: z_Code(68) = &H53FF3073: z_Code(69) = &H2C438924: z_Code(70) = &HFF000068: z_Code(71) = &H2BC03101: z_Code(72) = &H52E84845
        z_Code(73) = &H50000000: z_Code(74) = &HE844458B: z_Code(75) = &H49&: z_Code(76) = &H34438B50: z_Code(77) = &H29404D8B: z_Code(78) = &H3BE8C8
        z_Code(79) = &H39500000: z_Code(80) = &H5753C75: z_Code(81) = &HEB3C75FF: z_Code(82) = &H3C458B09: z_Code(83) = &H28E8&: z_Code(84) = &H73FF5000
        z_Code(85) = &H3875FF30: z_Code(86) = &HFF3475FF: z_Code(87) = &H75FF3075: z_Code(88) = &H2875FF2C: z_Code(89) = &H8B1853FF: z_Code(90) = &H30433903
        z_Code(91) = &H73FF3375: z_Code(92) = &H3073FF2C: z_Code(93) = &HEB2453FF: z_Code(94) = &H9EC6828: z_Code(95) = &H73FF0000: z_Code(96) = &H53FF5014
        z_Code(97) = &HF631C328: z_Code(98) = &H5656036A: z_Code(99) = &HFF3875FF: z_Code(100) = &H73FF3475: z_Code(101) = &H3075FF2C: z_Code(102) = &HFF2C75FF
        z_Code(103) = &H53FF2875: z_Code(104) = &H2CC26120
        
    ' fill in thunk constants/variables
    z_Code(0) = CreateCompatibleDC(0&)              ' Render's DC
    Set tPic = New StdPicture: Set nIPic = tPic
    CopyMemory vIPicDisp, ByVal ObjPtr(tPic), 4&    ' Original IPictureDisp vTable
    CopyMemory vIPic, ByVal ObjPtr(nIPic), 4&       ' Original IPicture vTable
    Set nIPic = Nothing: Set tPic = Nothing
    z_Code(1) = vIPicDisp                           ' COM's IPictureDisp used for all stdPictures
    z_Code(2) = vIPic                               ' COM's IPicture used for all stdPictures
    'z_Code(3) to be filled in a little later. Address for copy of COM Vtable
    z_Code(4) = m_Hwnd                              ' management window's handle
    z_Code(5) = 1440 / Screen.TwipsPerPixelX        ' VB's DPI, virutalized or not. Used in MulDiv call
    z_Code(6) = GetProcAddress(hMsImg, "AlphaBlend") ' function pointers used in the thunks
    hMod = GetModuleHandle("User32.dll")
        z_Code(7) = GetProcAddress(hMod, "PostMessageA")
        z_Code(8) = GetProcAddress(hMod, "DrawIconEx")
    z_Code(9) = GetProcAddress(GetModuleHandle("gdi32.dll"), "SelectObject")
    z_Code(10) = GetProcAddress(GetModuleHandle("kernel32.dll"), "MulDiv")
    ' address of constants/variables for each of the 4 subclassed functions above
    z_Code(16) = vThunks(1): z_Code(28) = vThunks(1): z_Code(33) = vThunks(1): z_Code(44) = vThunks(1)
   
    ' sanity checks
    For Index = 0& To 10&   ' z_Code(11) & (12) are always zero; thunk use only
        If z_Code(Index) = 0& Then
            If Not Index = 3& Then
                VirtualFree vThunks(1), 0&, MEM_RELEASE
                vThunks(1) = 0&
                GoTo ExitRoutine
            End If
        End If
    Next
    CopyMemory ByVal vThunks(1), z_Code(0), MEM_LEN
    CopyMemory ByVal pvSafePointerAdd(vThunks(0), 8&), vThunks(1), 4&  ' place thunk address in management window's thunk
   
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Copy of COM's VTable for IPicture & IPictureDisp
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    ReDim MBI(0 To 6)                               ' get memory used by COM for IPicture/IPictureDisp
    VirtualQuery ByVal vIPic, MBI(0), 28&           ' MBI() equivalent to MEMORY_BASIC_INFORMATION udt
    vThunks(2) = VirtualAlloc(0&, MBI(3), MEM_COMMIT, PAGE_RWX) ' reserve memory for same size
    If vThunks(2) = 0& Then GoTo ExitRoutine
   
    CopyMemory ByVal vThunks(2), ByVal MBI(0), MBI(3) ' copy VTable memory and we will hack this copy only
    CopyMemory ByVal pvSafePointerAdd(vThunks(1), 12&), vThunks(2), 4& ' update thunk #2 with address of this copy
   
    vIPicDisp = pvSafePointerAdd(vThunks(2), vIPicDisp - MBI(0))
    ' subclass IPictureDisp IUnknown:Release
    CopyMemory ByVal pvSafePointerAdd(vIPicDisp, 8&), pvSafePointerAdd(vThunks(1), 14& * 4&), 4&
   
    vIPic = pvSafePointerAdd(vThunks(2), vIPic - MBI(0)): Erase MBI()
    ' subclass IPicture IUnknown:Release
    CopyMemory ByVal pvSafePointerAdd(vIPic, 8&), pvSafePointerAdd(vThunks(1), 26& * 4&), 4&
    ' subclass IPicture:Render
    CopyMemory ByVal pvSafePointerAdd(vIPic, 32&), pvSafePointerAdd(vThunks(1), 42& * 4&), 4&
    ' subclass IPicture:GetAttributes
    CopyMemory ByVal pvSafePointerAdd(vIPic, 64&), pvSafePointerAdd(vThunks(1), 31& * 4&), 4&
   
   
ExitRoutine:
    ' clean up if failure occurred above...
    For Index = 2& To 0& Step -1&
        If vThunks(Index) = 0& Then Exit For    ' if any thunk failed creation, rewind everything
    Next
    If Index < 0& Then
        m_ThunkPtr = vThunks(0)
        SetProp m_Hwnd, "ThunkAddr", m_ThunkPtr
        SetProp m_Hwnd, "IPicAddr", vIPic
        SetProp m_Hwnd, "IPicDispAddr", vIPicDisp
        ' start subclassing the management window
        SetWindowLong m_Hwnd, GWL_WNDPROC, pvSafePointerAdd(m_ThunkPtr, 15& * 4&)
        pvCreateThunks = True
       
    Else                        ' not all thunks were created, failure & abort
        m_ThunkPtr = -1&
        For Index = 0& To 2&    ' ensure all thunks are released
            If vThunks(Index) Then VirtualFree vThunks(Index), 0&, MEM_RELEASE
        Next                    ' ensure any manually loaded DLLs are released
        If hMsImg Then FreeLibrary hMsImg
        If hGDIplus Then        ' ensure GDI+ is stopped if it was started
            If hToken Then GdiplusShutdown hToken
            FreeLibrary hGDIplus
        End If                  ' set flag on management window so this routine is effectively disabled
        SetProp m_Hwnd, "ThunkAddr", m_ThunkPtr
    End If
   
End Function

Private Function pvSafePointerAdd(ByVal thePointer As Long, ByVal AmountToAdjust As Long) As Long
    ' ensure no pointer addition/subtraction wraps from postiive to negative or vice versa
    Const SIGN_BIT = &H80000000
    pvSafePointerAdd = ((thePointer Xor SIGN_BIT) + AmountToAdjust) Xor SIGN_BIT
End Function

Private Function pvDWordAlign(ByVal Depth As Long, ByVal Width As Long, Optional ByVal Height As Long) As Long
    ' determine DWord aligned scan width for bitmaps. total bmp bits includes multiplying by rows (height)
    ' https://support.microsoft.com/en-us/kb/80080
    pvDWordAlign = ((Width * Depth + &H1F&) And Not &H1F&) \ &H8&
    If Height Then pvDWordAlign = pvDWordAlign * Abs(Height)
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
 
  CopyMemory nAddr, ByVal ObjPtr(oCallback), 4&                             'Get the address of the callback object's instance
  If zProbe(nAddr + &H1C, i, bSub) = False Then Exit Function               'Probe for a Class method
 
  i = pvSafePointerAdd(i, 4&)                                               'Bump to the next entry
  Do While j < 256                                                          'Set a reasonable limit, scan 256 vTable entries
    CopyMemory nAddr, ByVal i, 4&                                           'Get the address stored in this vTable entry
   
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      CopyMemory zAddressOf, ByVal pvSafePointerAdd(i, -(nOrdinal * 4&)), 4& 'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    CopyMemory bVal, ByVal nAddr, 1&                                        'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      CopyMemory zAddressOf, ByVal pvSafePointerAdd(i, -(nOrdinal * 4&)), 4& 'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    j = j + 1&
    i = pvSafePointerAdd(i, 4&)                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
 
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32&                                                      'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4&                                      'Get the vTable entry
   
    If nEntry <> 0 Then                                                     'If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1&                                     'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
   
    nAddr = nAddr + 4&                                                      'Next vTable entry
  Loop
End Function

Private Sub Class_Initialize()
    ' locates top-level VB owner window. This applies both during IDE and compiled exe
    ' determine if Vista or better being run
    ' cleans up any GDI+ image handles that may have been created from a previous instance
    '   where that instance was not shut down properly, i.e., END executed in IDE
    Dim lWnd As Long
Dim pObj As Long

    Const WS_SYSMENU As Long = &H80000
    Const WS_VISIBLE As Long = &H10000000
    Const GW_OWNER As Long = 4
    Const GW_HWNDNEXT As Long = 2
    Const GWL_STYLE As Long = -16
   
    Dim sBuffer()    As Byte
    Dim nBufferLen   As Long
    Dim lplpBuffer   As Long
    Dim vsffi(0 To 25) As Integer ' equivalent to VS_FIXEDFILEINFO
    Dim puLen        As Long
     
    nBufferLen = GetFileVersionInfoSize("kernel32.dll", ByVal VarPtr(vsffi(0)))
    If nBufferLen > 0 Then
         ReDim sBuffer(nBufferLen) As Byte
         Call GetFileVersionInfo("kernel32.dll", 0&, nBufferLen, sBuffer(0))
         Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
         Call CopyMemory(vsffi(0), ByVal lplpBuffer, 52&)
         m_VistaPlus = CBool(vsffi(5) > 5&) ' major version; want to know if Vista or higher O/S
     End If
   
    lWnd = FindWindow(vbNullString, vbNullString)
    Do Until lWnd = 0
        If GetWindowThreadProcessId(lWnd, ByVal 0&) = App.ThreadID Then
            If (GetWindowLong(lWnd, GWL_STYLE) And (WS_SYSMENU Or WS_VISIBLE)) = (WS_SYSMENU Or WS_VISIBLE) Then
                If GetWindow(lWnd, GW_OWNER) = 0 Then
                    m_Hwnd = lWnd
                    Exit Do
                End If
            End If
        End If
        lWnd = GetWindow(lWnd, GW_HWNDNEXT)
    Loop
    If m_Hwnd Then
        ' if previous instance of this class was terminated via END, then any GDI+ images
        ' that were created were not destroyed; still hanging out in memory & is a crash potential.
        ' Clean that up as needed
        lWnd = FindWindowEx(m_Hwnd, 0&, "Static", "IPIC+Thunker")
        If lWnd Then
            m_ThunkPtr = GetProp(lWnd, "ThunkAddr")
            If Not (m_ThunkPtr = 0& Or m_ThunkPtr = -1&) Then   ' thunk exists
                CopyMemory pObj, ByVal m_ThunkPtr, 4&           ' subclasser instance handle?
                If pObj = 0& Then Call pvDisposeGDIpImages
            End If
            m_ThunkPtr = 0&                         ' reset; key flag to determine if pvInitialize is ever called
        End If
    End If
End Sub

Private Sub Class_Terminate()

    ' free up any GDI+ related images, update a couple thunk variables
    If m_Primary Is Nothing Then                    ' we are the primary
        If Not (m_ThunkPtr = 0 Or m_ThunkPtr = -1&) Then ' else class has not been fully initialized
            CopyMemory ByVal m_ThunkPtr, 0&, 4&          ' prevetnt callbacks to this class
            ' Safe to walk routine after above line is executed
            Call pvDisposeGDIpImages    ' destroy any GDI+ handles; if those handles were IStream-based,
            Erase m_PicRefs()           '   then erasing array, removes IStream references
        End If
        If m_hDC Then DeleteDC m_hDC
    Else
        Set m_Primary = Nothing
    End If
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////////
' WARNING: This is always the final method in the class, nothing can follow it.
'          Ignoring this warning will result in a crash every time
'/////////////////////////////////////////////////////////////////////////////////////////////////////
'@Ignore ProcedureNotUsed
Private Sub pvReleased(ByVal pUnk As Long)
    ' This method is called by the Management thunk. It is called when either a
    '   subclassed IPicture or IPictureDisp interface is being set to Nothing.
    ' Note to self: Reason for subclassing both IPicture & IPictureDisp interfaces'
    '   IUnknown:Release functions is because one or the other would show a zero
    '   reference count. Both would not. In other words, it appears when the picture
    '   is being dereferenced, one of the interfaces is called and the other is
    '   reset internally by COM. Just a guess.
   
    ' Only purpose of this method is to release any GDI+ image handle that is associated with
    ' the passed interface. Unless the picture being destroyed was loaded as a multi-page
    ' TIFF or multi-frame GIF, there will be no action taken in this method
   
    Dim Index As Long
    Dim hImage As Long
    Dim lCount As Long
    
    Dim pImages As Long
    Dim gImage As Long
    'Dim gIndex As Long

   
'    DebugPrint "releasing "; pUnk
   
    For Index = 0& To m_RefCount - 1&                           ' locate the pUnk object in our array
        If pUnk = m_PicRefs(Index).pIPicture Then
            hImage = m_PicRefs(Index).pHandle                   ' set reference & exit loop
            Exit For
        ElseIf pUnk = m_PicRefs(Index).pIPicDisp Then
            hImage = m_PicRefs(Index).pHandle                   ' set reference & exit loop
            Exit For
        End If
    Next
    If Index < m_RefCount Then
        If hImage Then
            GdipDisposeImage hImage                             ' dispose of the image
            CopyMemory pImages, ByVal pvSafePointerAdd(m_ThunkPtr, 12&), 4& ' get array of GDI+ handles
            If pImages Then                                     ' could be zero
                CopyMemory lCount, ByVal pImages, 4&            ' get handle count in array
                For Index = 0& To lCount - 1&                   ' start the search
                    CopyMemory gImage, ByVal pvSafePointerAdd(pImages, Index * 4& + 8&), 4&
                    If gImage = hImage Then
                        lCount = lCount - 1&
                        If lCount = 0& Then                     ' release the array
                            CopyMemory ByVal pvSafePointerAdd(m_ThunkPtr, 12&), 0&, 4&
                            CoTaskMemRealloc pImages, 0&
                        Else                                    ' update array & shift as needed
                            If Index < lCount Then
                                CopyMemory ByVal pvSafePointerAdd(pImages, Index * 4& + 8&), _
                                           ByVal pvSafePointerAdd(pImages, lCount * 4& + 8&), 4&
                            End If
                            CopyMemory ByVal pImages, lCount, 4& ' update the count & exit loop
                        End If
                        Exit For
                    End If
                Next
            End If
        End If
        Set m_PicRefs(Index).oStream = Nothing                  ' release the IStream as needed
        m_RefCount = m_RefCount - 1&                            ' update count & shift array as needed
        If Index < m_RefCount Then m_PicRefs(Index) = m_PicRefs(m_RefCount)
        Set m_PicRefs(m_RefCount).oStream = Nothing
    End If

End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////
' WARNING: No executable code can follow this warning banner
'/////////////////////////////////////////////////////////////////////////////////////////////////////















