VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPlusCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a GDI+ based Loader+Cache-Class for Png, Gif and Ico-Alpha-Resources, later accessible per String-Key...
'(think: ImageList-Replacement) - high-quality-scaling of Alpha-Content to any hDC) ... [Olaf Schmidt 2015]
Option Explicit

Public Enum eInterpolationMode
  ipmDefault
  ipmLow
  ipmHigh
  ipmBilinear
  ipmBicubic
  ipmNearestNeighbor
  ipmHighQualityBilinear
  ipmHighQualityBicubic
End Enum

Public Enum eShellImageFactoryFlags
  SIIGBF_RESIZETOFIT = 0   'Shrink the bitmap as necessary to fit, preserving its aspect ratio.
  SIIGBF_BIGGERSIZEOK = 1  'Passed by callers if they want to stretch the returned image themselves. For example, if the caller passes an icon size of 80x80, a 96x96 thumbnail could be returned. This action can be used as a performance optimization if the caller expects that they will need to stretch the image. Note that the Shell implementation of IShellItemImageFactory performs a GDI stretch blit. If the caller wants a higher quality image stretch than provided through that mechanism, they should pass this flag and perform the stretch themselves.
  SIIGBF_MEMORYONLY = 2    'Return the item only if it is already in memory. Do not access the disk even if the item is cached. Note that this only returns an already-cached icon and can fall back to a per-class icon if an item has a per-instance icon that has not been cached. Retrieving a thumbnail, even if it is cached, always requires the disk to be accessed, so GetImage should not be called from the UI thread without passing SIIGBF_MEMORYONLY.
  SIIGBF_ICONONLY = 4      'Return only the icon, never the thumbnail.
  SIIGBF_THUMBNAILONLY = 8 'Return only the thumbnail, never the icon. Note that not all items have thumbnails, so SIIGBF_THUMBNAILONLY will cause the method to fail in these cases.
  SIIGBF_INCACHEONLY = 16  'Allows access to the disk, but only to retrieve a cached item
End Enum

Public Enum eImageWrapMode
  WrapModeNone = -1
  WrapModeTile = 0
  WrapModeTileFlipX
  WrapModeTileFlipY
  WrapModeTileFlipXY
  WrapModeClamp
End Enum

Public Enum eDrawText
  DT_TOPLEFT = &H0
  DT_CENTER = &H1
  DT_RIGHT = &H2
  DT_VCENTER = &H4
  DT_BOTTOM = &H8
  DT_WORDBREAK = &H10
  DT_SINGLELINE = &H20
  DT_EXPANDTABS = &H40
  DT_TABSTOP = &H80
  DT_CALCRECT = &H400
  DT_NOPREFIX = &H800
  DT_EDITCONTROL = &H2000
  DT_PATH_ELLIPSIS = &H4000
  DT_END_ELLIPSIS = &H8000
  DT_MODIFYSTRING = &H10000
  DT_RTLREADING = &H20000
  DT_WORD_ELLIPSIS = &H40000
  DT_HIDEPREFIX = &H100000
  DT_PREFIXONLY = &H200000
End Enum

Private Const PixelFormat32bppPARGB = &HE200B, PixelFormat32bppARGB = &H26200A
 
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, Inbuf As Long, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As Long, Image As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal pStream As Long, ClsIdEnc As Any, EncParms As Any) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal img As Long, Context As Long) As Long
Private Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal Context As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal CombineMode As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageAttributes As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageAttributes As Long, ByVal ColorAdjust As Long, ByVal Enable As Long, M5x5Color As Any, M5x5Gray As Any, ByVal Flags As Long) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageAttributes As Long, ByVal WrapMode As Long, ByVal Color As Long, Optional ByVal Clamp As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageAttributes As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Context As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal Context As Long, ByVal Image As Long, ByVal dstx As Single, ByVal dsty As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcx As Single, ByVal srcy As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Context As Long, ByVal Color As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Context As Long, ByVal Mode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Context As Long, ByVal PixOffsetMode As Long) As Long
Private Declare Function GdipBitmapConvertFormat Lib "gdiplus" (ByVal Image As Long, ByVal Format As Long, ByVal ADitherType As Long, ByVal APaletteType As Long, ByVal pPalette As Long, ByVal AlphaThresholdPercent As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal dx As Long, ByVal dy As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal pScanData As Long, Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBmp As Long, ByVal hPal As Long, Image As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal Image As Long, hIcon As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Image As Long, hBmp As Long, ByVal BGColor As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, FrDimID As Any, Count As Long) As Long
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, FrDimID As Any, ByVal FrIdx As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Context As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
'and the recently added block of GDIP-Functions, used here for simple antialiased Alpha-Drawings
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Context As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Context As Long, Optional ByVal SmoothingMode As Long = 4) As Long
Private Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Context As Long, State As Long) As Long
Private Declare Function GdipRestoreGraphics Lib "gdiplus" (ByVal Context As Long, ByVal State As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Context As Long, ByVal dx As Single, ByVal dy As Single, Optional ByVal MatrixOrderAfter As Long = 1) As Long
Private Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal Context As Long, ByVal sx As Single, ByVal sy As Single, Optional ByVal MatrixOrderAfter As Long = 1) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Context As Long, ByVal Angle As Single, Optional ByVal MatrixOrderAfter As Long = 1) As Long
Private Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal Context As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal ARGBColor As Long, ByVal Width As Single, ByVal Unit As Long, gdipPen As Long) As Long
Private Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal gdipPen As Long, ByVal Width As Single) As Long
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal gdipPen As Long, ByVal ARGBColor As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal Context As Long, ByVal gdipPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal Context As Long, ByVal gdipPen As Long, pSingleArr As Any, ByVal XYPairsCount As Long, ByVal Tension As Single) As Long
Private Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal Context As Long, ByVal gdipPen As Long, pSingleArr As Any, ByVal XYPairsCount As Long, ByVal Tension As Single) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal Context As Long, ByVal gdipPen As Long, ByVal X As Single, ByVal Y As Single, ByVal dx As Single, ByVal dy As Single) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal Context As Long, ByVal gdipPen As Long, ByVal X As Single, ByVal Y As Single, ByVal dx As Single, ByVal dy As Single) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal gdipPen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGBColor As Long, gdipSolidBrush As Long) As Long
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal gdipBrush As Long, ByVal ARGBColor As Long) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal Context As Long, ByVal gdipBrush As Long, ByVal X As Single, ByVal Y As Single, ByVal dx As Single, ByVal dy As Single) As Long
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal Context As Long, ByVal gdipBrush As Long, ByVal X As Single, ByVal Y As Single, ByVal dx As Single, ByVal dy As Single) As Long
Private Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal Context As Long, ByVal gdipBrush As Long, pSingleArr As Any, ByVal XYPairsCount As Long, ByVal Tension As Single, ByVal FillMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal gdipBrush As Long) As Long
Private Declare Function GdipCreateFontFromLogfontW Lib "gdiplus" (ByVal hDC As Long, pLOGFONTW As Any, gdipFont As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal Context As Long, ByVal pWString As Long, ByVal sLen As Long, ByVal gdipFont As Long, RectFLayout As Any, ByVal gdipStringFormat As Long, ByVal gdipBrush As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal Context As Long, ByVal pWString As Long, ByVal sLen As Long, ByVal gdipFont As Long, RectFLayout As Any, ByVal gdipStringFormat As Long, RectFMeasured As Any, ByVal pCodepointsFitted As Long, LinesFilled As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal gdipFont As Long) As Long
Private Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (gdipFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal gdipFormat As Long, ByVal StringAlignment As Long) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal gdipFormat As Long, ByVal StringAlignment As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal gdipFormat As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
 
Private Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon&, IconInfo As Any) As Long
Private Declare Function CreateIconIndirect Lib "user32" (IconInfo As Any) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal hIcon&, ByVal dx&, ByVal dy&, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GdiAlphaBlend& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal hdcSrc&, ByVal srcx&, ByVal srcy&, ByVal SrcdX&, ByVal SrcdY&, ByVal lBlendFunction&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function GetDIBits& Lib "gdi32" (ByVal aHDC&, ByVal hBM&, ByVal nStartSL&, ByVal nNumSL&, lpBits As Any, lpBI As Any, ByVal wUsage&)
Private Declare Function CreateSolidBrush& Lib "gdi32" (ByVal Color&)
Private Declare Function CreateDIBSection& Lib "gdi32" (ByVal hDC&, pBitmapInfo As Any, ByVal un&, ppBits&, ByVal Hdl&, ByVal dw&)
Private Declare Function DrawTextW& Lib "user32" (ByVal hDC&, ByVal pStr&, ByVal nCount&, pRect As Any, ByVal wFormat&)
Private Declare Function GetObjectW& Lib "gdi32" (ByVal hObj&, ByVal nCount&, pData As Any)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC&, ByVal hObject&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObj&)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC&)

Private Declare Function CLSIDFromString Lib "ole32" (ByVal psGUID As Long, id As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
Private Declare Function DispCallFunc& Lib "oleaut32" (ByVal ppv&, ByVal oVft&, ByVal CallingConvention As Long, ByVal rtTYP%, ByVal paCount&, paTypes%, paValues&, fuReturn)

Private Declare Function IStream_Size Lib "shlwapi" (ByVal pStream As Long, SizeLongLong As Any) As Long
Private Declare Function IStream_Reset Lib "shlwapi" (ByVal pStream As Long) As Long
Private Declare Function IStream_Read Lib "shlwapi" (ByVal pStream As Long, pBytes As Any, ByVal CB As Long) As Long

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal CB&)
Private Declare Function SHGetItemFromObject& Lib "shell32" (ByVal Unk As stdole.IUnknown, riid As Any, ppv As Any)
Private Declare Function SHCreateItemFromParsingName& Lib "shell32" (ByVal pPath&, ByVal pBC&, riid As Any, ppv As Any)

Private mImages As Object
Private mToken As Long, mCtx&, mCTXDC&, mPen&, mBrush&, mStackIdx&, mStack&(0 To 255)
Private mFrDim(0 To 15) As Byte, mPngID(0 To 15) As Byte, mFacID(0 To 15) As Byte
Private mhDC&, mOldBmp&, mBBCtx&, mBBImg&, mBBdx&, mBBdy&, mDibPtr& '<- Variables for the BackBuffer-Handling
Public HalfPixelOffsets As Boolean

Private Sub Class_Initialize()
Set mImages = CreateObject("Scripting.Dictionary")
    mImages.Comparemode = 1 'case-insenitive Key-Comparisons
Dim StartupInput&(0 To 3): StartupInput(0) = 1
    GdiplusStartup mToken, StartupInput(0)
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), mFrDim(0) 'FrameDimensionTime (needed for animated Gifs)
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), mPngID(0) 'PngWriter-ClsID
    CLSIDFromString StrPtr("{BCC18B79-BA16-442F-80C4-8A59C30C463B}"), mFacID(0) 'ShellItemImageFactory-IID
End Sub
Private Sub Class_Terminate()
  Dim img: For Each img In mImages.Items: GdipDisposeImage img: Next
  DestroyBackBuf
  ContextClear
  If mToken Then GdiplusShutdown mToken
End Sub

Private Sub PrepareBackBuf(ByVal dx As Long, ByVal dy As Long)
  If mBBdx < dx Or mBBdy < dy Then
    If mBBdx < dx Then mBBdx = dx
    If mBBdy < dy Then mBBdy = dy
    
    Dim BI(0 To 9) As Long, hDIB As Long
        BI(0) = 40
        BI(1) = mBBdx
        BI(2) = -mBBdy
        BI(3) = 32 * 65536 + 1 '32bpp
    If mhDC = 0 Then mhDC = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(0, BI(0), 0, mDibPtr, 0, 0)
    If mOldBmp Then DeleteObject SelectObject(mhDC, hDIB) Else mOldBmp = SelectObject(mhDC, hDIB)
    If mBBImg Then GdipDisposeImage mBBImg
    GdipCreateBitmapFromScan0 mBBdx, mBBdy, mBBdx * 4, PixelFormat32bppPARGB, mDibPtr, mBBImg
    If mBBImg Then GdipGetImageGraphicsContext mBBImg, mBBCtx
    If mBBCtx Then GdipSetPixelOffsetMode mBBCtx, 4 '4=Half
  End If
End Sub
Private Sub PrepareAndClearBackBuf(ByVal dx As Long, ByVal dy As Long, Optional ByVal ClearColor As Long)
  PrepareBackBuf dx, dy
  GdipSetClipRectI mBBCtx, 0, 0, dx, dy, 0
  GdipGraphicsClear mBBCtx, ClearColor
End Sub
Private Sub DestroyBackBuf()
  If mBBCtx Then GdipDeleteGraphics mBBCtx
  If mBBImg Then GdipDisposeImage mBBImg
  If mOldBmp Then DeleteObject SelectObject(mhDC, mOldBmp)
  If mhDC Then DeleteDC mhDC
End Sub
 
Private Function CreateScaledImg(SrcImg As Long, dxSrc, dySrc, dxDst, dyDst) As Long
  Dim img As Long, Ctx As Long
  GdipCreateBitmapFromScan0 dxDst, dyDst, dxDst * 4, PixelFormat32bppPARGB, 0, img
  If img Then CreateScaledImg = img Else Err.Raise vbObjectError, , "unable to create scaled Img-Resource"
  If img Then GdipGetImageGraphicsContext img, Ctx
  If Ctx Then GdipSetPixelOffsetMode Ctx, 4   '4=Half, 3=None
  If Ctx Then GdipSetInterpolationMode Ctx, eInterpolationMode.ipmHighQualityBicubic
  If Ctx <> 0 And SrcImg <> 0 Then GdipDrawImageRectRectI Ctx, SrcImg, 0, 0, dxDst, dyDst, 0, 0, dxSrc, dySrc, 2, 0, 0, 0
  If Ctx Then GdipDeleteGraphics Ctx
End Function

Public Function ReadBytesFromFile(ByVal FileName As String) As Byte()
  With CreateObject("ADODB.Stream")
    .Open
      .Type = 1 'adTypeBinary
      .LoadFromFile FileName
      ReadBytesFromFile = .Read
    .Close
  End With
End Function
Public Sub WriteBytesToFile(ByVal FileName As String, B() As Byte)
  With CreateObject("ADODB.Stream")
    .Open
      .Type = 1 'adTypeBinary
      .Write B
      .SaveToFile FileName, 2 'Create+Overwrite
    .Close
  End With
End Sub

Public Function Exists(key) As Boolean
  Exists = mImages.Exists(key)
End Function
Public Sub Remove(key)
  GdipDisposeImage mImages(key)
  mImages.Remove key
End Sub

Public Function Keys()
  Keys = mImages.Keys
End Function

Public Sub AddImage(key, FileNameOrBytesOrStdPicture, Optional ByVal DesiredWidth& = 0, Optional ByVal DesiredHeight& = 0, Optional ByVal CacheMultiFrames As Boolean)
Dim B() As Byte, Strm As stdole.IUnknown, img&, Frm&, dx&, dy&, FrameCount&, i&
  If IsObject(FileNameOrBytesOrStdPicture) Then
    GdipCreateBitmapFromHBITMAP FileNameOrBytesOrStdPicture.handle, 0, img
  Else
    If VarType(FileNameOrBytesOrStdPicture) = vbString Then B = ReadBytesFromFile(FileNameOrBytesOrStdPicture) Else B = FileNameOrBytesOrStdPicture
    CreateStreamOnHGlobal VarPtr(B(0)), 0, Strm
    GdipLoadImageFromStream ObjPtr(Strm), img
  End If
  If img = 0 Then Err.Raise vbObjectError, , "Could not load image with GDIPlus"
  
  GdipGetImageWidth img, dx:  If DesiredWidth <= 0 Then DesiredWidth = dx
  GdipGetImageHeight img, dy: If DesiredHeight <= 0 Then DesiredHeight = dy
  
  If Exists(key) Then Remove key
  mImages.Add key, CreateScaledImg(img, dx, dy, DesiredWidth, DesiredHeight)
  If img <> 0 And CacheMultiFrames Then
    GdipImageGetFrameCount img, mFrDim(0), FrameCount
    For i = 0 To FrameCount - 1
      GdipImageSelectActiveFrame img, mFrDim(0), i
      If Exists(key & "|" & i) Then Remove key & "|" & i
      mImages.Add key & "|" & i, CreateScaledImg(img, dx, dy, DesiredWidth, DesiredHeight)
    Next
  End If
  
  GdipDisposeImage img
End Sub

Public Sub AddIcon(key, FileNameOrBytes, ByVal DesiredWidth As Long, ByVal DesiredHeight As Long, Optional HotSpotX!, Optional HotSpotY!)
Dim B() As Byte, sz As Long, Offs As Long, hIcon As Long, bpp, Y As Long, X As Long, L() As Long
  If VarType(FileNameOrBytes) = vbString Then B = ReadBytesFromFile(FileNameOrBytes) Else B = FileNameOrBytes
  
  For Each bpp In Array(32, 24, 16, 8, 1)
    Offs = GetIcoHdrOffs(B, bpp, DesiredWidth, sz, HotSpotX, HotSpotY)
    If Offs Then Exit For
  Next
  If Offs = 0 Then Err.Raise vbObjectError, , "the Resource-Content doesn't contain any valid Icons or Cursors"
  
  Offs = B(Offs) + 256& * B(Offs + 1) + 65536 * B(Offs + 2)
  hIcon = CreateIconFromResourceEx(B(Offs), UBound(B) + 1 - Offs, 1, &H30000, sz, sz, 0)
  If hIcon = 0 Then Err.Raise vbObjectError, , "Could not load icon"
  
  PrepareAndClearBackBuf sz, sz, IIf(bpp = 32, 0, vbMagenta Or &HFF000000)
  DrawIconEx mhDC, 0, 0, hIcon, sz, sz, 0, 0, 3 '<- DI_Normal = DI_Image Or DI_Mask
  DestroyIcon hIcon 'not needed anymore, so we destroy it here
  
  If bpp <> 32 And mDibPtr <> 0 Then 'correct the Alpha-Bytes for all Icons or Cursors with a BitDepth below 32
    ReDim L(0 To mBBdx - 1, 0 To mBBdy - 1)
    MemCopy L(0, 0), ByVal mDibPtr, mBBdy * mBBdx * 4
    For Y = 0 To sz - 1: For X = 0 To sz - 1
      If (L(X, Y) And vbMagenta) = vbMagenta Then L(X, Y) = 0 Else L(X, Y) = L(X, Y) Or &HFF000000
    Next X, Y
    MemCopy ByVal mDibPtr, L(0, 0), mBBdy * mBBdx * 4
  End If
  
  If Exists(key) Then Remove key
  mImages.Add key, CreateScaledImg(mBBImg, sz, sz, DesiredWidth, DesiredHeight)
End Sub

Public Sub AddFromShell(key, PathOrShellItemObj, Optional ByVal DesiredSize& = 32, Optional ByVal ShellFlags As eShellImageFactoryFlags)
  Dim HRes As Long, Unk As stdole.IUnknown, hBmp&, P() As Long, dx&, dy&, img&
  
  If IsObject(PathOrShellItemObj) Then
    HRes = SHGetItemFromObject(PathOrShellItemObj, mFacID(0), Unk)
  ElseIf VarType(PathOrShellItemObj) = vbString Then
    HRes = SHCreateItemFromParsingName(StrPtr(PathOrShellItemObj), 0, mFacID(0), Unk)
  Else
    Err.Raise vbObjectError, , "We need either a ShellItem-Obj or a FilePath here..."
  End If
  If HRes Then Err.Raise HRes
  
  HRes = vtblCall(ObjPtr(Unk), 3, DesiredSize, DesiredSize, ShellFlags, VarPtr(hBmp))
  If HRes Then Err.Raise HRes
 
  P = Hdl2PxlArr32(hBmp)
  dx = UBound(P, 1) + 1
  dy = UBound(P, 2) + 1
  If dx <= 0 Then Err.Raise vbObjectError, , "Could not retrieve PixelData for the Shell-Image"
  
  GdipCreateBitmapFromScan0 dx, dy, dx * 4, PixelFormat32bppARGB, VarPtr(P(0, 0)), img
  If img = 0 Then Err.Raise vbObjectError, , "Could not load ARGB-image with GDIPlus"
  
  If Exists(key) Then Remove key
  mImages.Add key, CreateScaledImg(img, dx, dy, dx, dy) 'just for conversion from ARGB to PreMultiplied Alpha
  GdipDisposeImage img
End Sub

'a helper, which deals with raw-icon or -cursor-formats (looping the Dir-entries for best matching sizes for a given BPP-BitDepth)
Private Function GetIcoHdrOffs(B() As Byte, ByVal bpp&, ByVal dw&, sz&, hsx!, hsy!) As Long
  Dim i As Long, MaxW(1 To 256) As Integer, Offs As Long
 
  For i = 0 To B(4) - 1
    If B(2) = 1 And B(12 + i * 16) = bpp Then MaxW((511 + B(6 + i * 16)) Mod 256 + 1) = i + 1
    If B(2) = 2 Then 'it's a Cursor-resource apparently
      Offs = B(18 + i * 16) + 256& * B(19 + i * 16) + 65526 * B(20 + i * 16)
      If B(Offs + 14) = bpp Then MaxW((511 + B(6 + i * 16)) Mod 256 + 1) = i + 1
    End If
  Next
  
  For i = 1 To 256
    If MaxW(i) Then sz = i: GetIcoHdrOffs = (MaxW(i) - 1) * 16 + 18
    If GetIcoHdrOffs > 0 And i >= dw Then Exit For
  Next
  If sz Then hsx = B((MaxW(sz) - 1) * 16 + 10) / sz: hsy = B((MaxW(sz) - 1) * 16 + 12) / sz
End Function

Public Property Get FrameCount(key) As Long
  Do While Exists(key & "|" & FrameCount): FrameCount = FrameCount + 1: Loop
End Property
 
Public Property Get Width(key) As Long
  GdipGetImageWidth GetImage(key), Width
End Property
Public Property Get Height(key) As Long
  GdipGetImageHeight GetImage(key), Height
End Property

Public Function GetHIconFromImage(key) As Long
  GdipCreateHICONFromBitmap GetImage(key), GetHIconFromImage
End Function
Public Sub DestroyHIcon(ByVal hIcon As Long)
  If hIcon Then DestroyIcon hIcon
End Sub

Public Function GetHCursorFromImage(key, Optional ByVal HSpotX As Long, Optional ByVal HSpotY As Long) As Long
Dim hIcon As Long, IconInfo(0 To 4) As Long
    GdipCreateHICONFromBitmap GetImage(key), hIcon
    If hIcon Then GetIconInfo hIcon, IconInfo(0)
    IconInfo(0) = 0: IconInfo(1) = HSpotX: IconInfo(2) = HSpotY
    GetHCursorFromImage = CreateIconIndirect(IconInfo(0))
    If IconInfo(3) Then DeleteObject IconInfo(3)
    If IconInfo(4) Then DeleteObject IconInfo(4)
    DestroyIcon hIcon
End Function
Public Sub DestroyHCursor(ByVal hCursor As Long)
  If hCursor Then DestroyCursor hCursor
End Sub

Public Function GetHBmpFromImage(key, Optional ByVal BGColor As Long = vbWhite) As Long
  GdipCreateHBITMAPFromBitmap GetImage(key), GetHBmpFromImage, BGColor
End Function
Public Sub DestroyHBmp(ByVal hBmp As Long)
  If hBmp Then DeleteObject hBmp
End Sub

'note, that the GDIPlusImageHandle-Instance returned here, does not need to be freed on the outside, as long as the Image remains in the mImages-Collection
Public Function GetImage(ByVal key, Optional ByVal UseGifDirection As Boolean) As Long
  If Right$(key, 1) = "<" Or Right$(key, 1) = ">" Then
    If Not UseGifDirection Then key = Left$(key, Len(key) - 1)
    If UseGifDirection Then key = Left$(key, Len(key) - 1) & "|" & GetNextGifIndex(key)
  End If
  GetImage = mImages(key)
End Function

'normally just a Helper-Function inside here, but defined Public, so it can be used from the outside as well
Public Function Hdl2PxlArr32(ByVal hBmp As Long, Optional ByVal DeleteHdl As Boolean = True) As Long()
  If mhDC = 0 Then PrepareBackBuf 32, 32
  Dim BI(0 To 9) As Long: BI(0) = 40
  Dim P() As Long: ReDim P(-1 To -1, -1 To -1) 'return -1 Bounds in case of an error-caused early exit

  GetDIBits mhDC, hBmp, 0, 0, ByVal 0&, BI(0), 0 'first call will retrieve "infos only"
  If BI(1) + BI(2) Then ReDim P(0 To BI(1) - 1, 0 To BI(2) - 1) Else GoTo 1 'goto cleanup in case of no success

  BI(2) = -BI(2) 'negate the Height-Member (addressing bottom-up behaviour)
  BI(3) = 1 + 65536 * 32 '1 Plane and 32BitsPerPixel
  BI(4) = 0 'enforce a zero (no compression) in this member (get rid of potential residues, as e.g. BI_BITFIELDS from the call above)
  BI(5) = 4 * BI(1) * Abs(BI(2)) 'tell GetDIBits the size we expect (in case the source-size differed)
  
  GetDIBits mhDC, hBmp, 0, -BI(2), P(0, 0), BI(0), 0
  
1 If DeleteHdl Then DeleteObject hBmp
  Hdl2PxlArr32 = P '<- being the last instruction, this will avoid a copy - returning the P()-Pointer directly
End Function

Private Function GetNextGifIndex(key, Optional ByVal BackWards As Boolean) As Long
Static i As Long, V(0 To 255, 0 To 2)
  For i = 0 To 255
    If V(i, 0) = Empty Then
       V(i, 0) = key
       V(i, 1) = FrameCount(Left(key, Len(key) - 1)) - 1
       Exit For
    ElseIf V(i, 0) = key Then
       V(i, 2) = V(i, 2) + IIf(Right$(key, 1) = ">", 1, -1)
       If V(i, 2) > V(i, 1) Then V(i, 2) = 0 Else If V(i, 2) < 0 Then V(i, 2) = V(i, 1)
       GetNextGifIndex = V(i, 2)
       Exit For
    End If
  Next
End Function

Public Function SaveImageToPngByteArray(key) As Byte()
  SaveImageToPngByteArray = "" 'return an initialized, byt empty (0 to -1) ByteArray in case of an error
  Dim Strm As stdole.IUnknown, SizeLL(0 To 1) As Long, B() As Byte
  CreateStreamOnHGlobal 0, 1, Strm
  If Strm Is Nothing Then Exit Function
  
  GdipSaveImageToStream mImages(key), ObjPtr(Strm), mPngID(0), ByVal 0&
  IStream_Size ObjPtr(Strm), SizeLL(0)
  If SizeLL(0) > 0 Then ReDim B(0 To SizeLL(0) - 1) Else Exit Function
  
  IStream_Reset ObjPtr(Strm)
  IStream_Read ObjPtr(Strm), B(0), SizeLL(0)
  SaveImageToPngByteArray = B
End Function

'the WorkHorse-function for Alpha-renderings to outside hDCs
Public Sub AlphaRenderTo(ByVal hDC As Long, key, Optional ByVal X As Long, Optional ByVal Y As Long, _
                                                 Optional ByVal dx As Long, Optional ByVal dy As Long, _
                                                 Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                                                 Optional ByVal dxSrc As Long, Optional ByVal dySrc As Long, _
                                                 Optional ByVal Alpha As Double = 1, Optional ByVal StretchMode As eInterpolationMode = ipmHigh)
  If dxSrc = 0 Then dxSrc = Width(key)
  If dySrc = 0 Then dySrc = Height(key)
  If dx = 0 Then dx = Width(key)
  If dy = 0 Then dy = Height(key)

  PrepareAndClearBackBuf dx, dy
  GdipSetInterpolationMode mBBCtx, StretchMode
  GdipDrawImageRectRectI mBBCtx, GetImage(key, True), 0, 0, dx, dy, xSrc, ySrc, dxSrc, dySrc, 2, 0, 0, 0
  GdiAlphaBlend hDC, X, Y, dx, dy, mhDC, 0, 0, dx, dy, &H1000000 + &H10000 * Int(255 * Alpha)
End Sub

'Just to cover basic Unicode-Text-Rendering for the outside as well (not only Icon- and Image-Resource-Renderings)
Public Function DrawText(ByVal hDC As Long, ByVal S As String, X, Y, dx, dy, Optional ByVal DTFlags As eDrawText = DT_SINGLELINE Or DT_VCENTER) As Long
  Dim R(0 To 3) As Long
      R(0) = X: R(1) = Y: R(2) = X + dx: R(3) = Y + dy
  DrawText = DrawTextW(hDC, StrPtr(S), Len(S), R(0), DTFlags)
  If DTFlags And DT_CALCRECT Then X = R(0): Y = R(1): dx = R(2) - X: dy = R(3) - Y
End Function

'a block of simple routines, to allow for antialiased Alpha-Drawings as well
Public Sub ContextBindTo(ByVal hDC As Long, Optional ByVal Interpolation As eInterpolationMode = ipmBicubic)
  If mCTXDC = hDC Then GdipSetPixelOffsetMode mCtx, IIf(HalfPixelOffsets, 4, 0): Exit Sub 'early exit (still the same hDC)
  ContextClear
  mCTXDC = hDC
  GdipCreateFromHDC hDC, mCtx
  If mCtx = 0 Then Err.Raise vbObjectError, , "Couldn't bind GDIP-Context to hDC"
  GdipSetSmoothingMode mCtx, 4 'antialiased Mode by default
  If HalfPixelOffsets Then GdipSetPixelOffsetMode mCtx, 4 '4=Half
  GdipSetInterpolationMode mCtx, Interpolation
  
  GdipCreatePen1 0, 1, 2, mPen
  GdipCreateSolidFill 0, mBrush
  If mPen = 0 Or mBrush = 0 Then ContextClear: Err.Raise vbObjectError, , "Couldn't create GDIP-DrawingObjects"
End Sub
Public Sub ContextClear()
  If mBrush Then GdipDeleteBrush mBrush: mBrush = 0
  If mPen Then GdipDeletePen mPen:       mPen = 0
  If mCtx Then GdipDisposeImage mCtx:    mCtx = 0
  mCTXDC = 0: mStackIdx = 0
End Sub

Public Function Save(ByVal hDC As Long) As Long
  ContextBindTo hDC
  If mStackIdx < 255 Then mStackIdx = mStackIdx + 1 Else Exit Function
  GdipSaveGraphics mCtx, mStack(mStackIdx)
  Save = mStack(mStackIdx)
End Function
Public Function Restore(ByVal hDC As Long, Optional ByVal State As Long) As Long
  ContextBindTo hDC
  If State = 0 And mStackIdx > 0 Then
     If mStack(mStackIdx) Then GdipRestoreGraphics mCtx, mStack(mStackIdx)
  Else
     Dim i As Long
     For i = mStackIdx To 1 Step -1
       If mStack(i) = State Then GdipRestoreGraphics mCtx, State: mStackIdx = i: Exit For
     Next
  End If
  If mStackIdx > 0 Then mStackIdx = mStackIdx - 1
End Function

Public Sub TranslateDrawings(ByVal hDC As Long, ByVal x0 As Single, ByVal y0 As Single)
  ContextBindTo hDC: GdipTranslateWorldTransform mCtx, x0, y0, 0
End Sub
Public Sub ScaleDrawings(ByVal hDC As Long, ByVal sx As Single, ByVal sy As Single)
  ContextBindTo hDC: GdipScaleWorldTransform mCtx, sx, sy, 0
End Sub
Public Sub RotateDrawings(ByVal hDC As Long, ByVal AngleRad As Single)
  ContextBindTo hDC: GdipRotateWorldTransform mCtx, AngleRad * 57.2957795130823, 0
End Sub
Public Sub RotateDrawingsDeg(ByVal hDC As Long, ByVal AngleDeg As Single)
  ContextBindTo hDC: GdipRotateWorldTransform mCtx, AngleDeg, 0
End Sub
Public Sub ResetTransforms(ByVal hDC As Long)
  ContextBindTo hDC: GdipResetWorldTransform mCtx
End Sub

Public Sub DrawLine(ByVal hDC&, ByVal x1!, ByVal y1!, ByVal x2!, ByVal y2!, Optional ByVal LineWidth! = 1, Optional ByVal LineColor&, Optional ByVal Alpha! = 1)
  Prepare hDC, LineColor, Alpha, LineWidth
  GdipDrawLine mCtx, mPen, x1, y1, x2, y2
End Sub
Public Sub DrawRect(ByVal hDC&, ByVal X!, ByVal Y!, ByVal dx!, ByVal dy!, Optional ByVal BorderWidth! = 1, Optional ByVal BorderColor&, Optional ByVal Alpha! = 1)
  Prepare hDC, BorderColor, Alpha, BorderWidth
  GdipDrawRectangle mCtx, mPen, X, Y, dx, dy
End Sub
Public Sub DrawElps(ByVal hDC&, ByVal X!, ByVal Y!, ByVal dx!, ByVal dy!, Optional ByVal BorderWidth! = 1, Optional ByVal BorderColor&, Optional ByVal Alpha! = 1)
  Prepare hDC, BorderColor, Alpha, BorderWidth
  GdipDrawEllipse mCtx, mPen, X, Y, dx, dy
End Sub
Public Sub DrawPolygon(ByVal hDC&, ByVal LineWidth!, ByVal LineColor&, ByVal Alpha!, ByVal SplineTension!, ByVal Closed As Boolean, ParamArray SingleArrayOrXYPairs())
  If UBound(SingleArrayOrXYPairs) < 0 Then Err.Raise vbObjectError, , "No PolygonData was passed in SingleArrayOrXYPairs()"
  Prepare hDC, LineColor, Alpha, LineWidth
  
  Dim i As Long, Arr() As Single
  If VarType(SingleArrayOrXYPairs(0)) = (vbArray Or vbSingle) Then
    Arr = SingleArrayOrXYPairs(0)
  ElseIf IsArray(SingleArrayOrXYPairs(0)) Then
    Err.Raise vbObjectError, , "Only Arrays of Type Single should be passed"
  Else
    ReDim Arr(0 To UBound(SingleArrayOrXYPairs))
    For i = 0 To UBound(Arr): Arr(i) = SingleArrayOrXYPairs(i): Next 'copy over
  End If
  If UBound(Arr) < 3 Then Err.Raise vbObjectError, , "We need at least 2 Points in SingleArrayOrXYPairs()"
  If Not Closed Or UBound(Arr) = 3 Then GdipDrawCurve2 mCtx, mPen, Arr(0), (UBound(Arr) + 1) \ 2, SplineTension
  If Closed Then GdipDrawClosedCurve2 mCtx, mPen, Arr(0), (UBound(Arr) + 1) \ 2, SplineTension
End Sub

Public Sub FillRect(ByVal hDC&, ByVal X!, ByVal Y!, ByVal dx!, ByVal dy!, ByVal FillColor&, Optional ByVal Alpha! = 1)
  Prepare hDC, FillColor, Alpha, 0
  GdipFillRectangle mCtx, mBrush, X, Y, dx, dy
End Sub
Public Sub FillElps(ByVal hDC&, ByVal X!, ByVal Y!, ByVal dx!, ByVal dy!, ByVal FillColor&, Optional ByVal Alpha! = 1)
  Prepare hDC, FillColor, Alpha, 0
  GdipFillEllipse mCtx, mBrush, X, Y, dx, dy
End Sub
Public Sub FillPolygon(ByVal hDC&, ByVal FillColor&, ByVal Alpha!, ByVal SplineTension!, ByVal FillModeWinding As Boolean, ParamArray SingleArrayOrXYPairs())
  If UBound(SingleArrayOrXYPairs) < 0 Then Err.Raise vbObjectError, , "No PolygonData was passed in SingleArrayOrXYPairs()"
  Prepare hDC, FillColor, Alpha, 0
  
  Dim i As Long, Arr() As Single
  If VarType(SingleArrayOrXYPairs(0)) = (vbArray Or vbSingle) Then
    Arr = SingleArrayOrXYPairs(0)
  ElseIf IsArray(SingleArrayOrXYPairs(0)) Then
    Err.Raise vbObjectError, , "Only Arrays of Type Single should be passed"
  Else
    ReDim Arr(0 To UBound(SingleArrayOrXYPairs))
    For i = 0 To UBound(Arr): Arr(i) = SingleArrayOrXYPairs(i): Next 'copy over
  End If
  If UBound(Arr) < 5 Then Err.Raise vbObjectError, , "We need at least 3 Points in SingleArrayOrXYPairs()"
  GdipFillClosedCurve2 mCtx, mBrush, Arr(0), (UBound(Arr) + 1) \ 2, SplineTension, IIf(FillModeWinding, 1, 0)
End Sub

Public Sub DrawImage(ByVal hDC As Long, key, Optional ByVal X As Single, Optional ByVal Y As Single, _
                                             Optional ByVal dx As Single, Optional ByVal dy As Single, _
                                             Optional ByVal xSrc As Single, Optional ByVal ySrc As Single, _
                                             Optional ByVal dxSrc As Single, Optional ByVal dySrc As Single, _
                                             Optional ByVal Alpha As Single = 1, Optional ByVal StretchMode As eInterpolationMode = ipmBicubic, _
                                             Optional ByVal WrapMode As eImageWrapMode = WrapModeNone)
  ContextBindTo hDC
  If dxSrc = 0 Then dxSrc = Width(key)
  If dySrc = 0 Then dySrc = Height(key)
  If dx = 0 Then dx = Width(key)
  If dy = 0 Then dy = Height(key)
  
  GdipSetPixelOffsetMode mCtx, HalfPixelOffsets
  GdipSetInterpolationMode mCtx, StretchMode
    
  Dim ImgAttr As Long, GM!(0 To 24), CM!(0 To 24): CM(0) = 1: CM(6) = 1: CM(12) = 1: CM(24) = 1
  GdipCreateImageAttributes ImgAttr
  If Alpha < 1 Then CM(18) = Alpha: GdipSetImageAttributesColorMatrix ImgAttr, 0, 1, CM(0), GM(0), 0
  If WrapMode <> WrapModeNone Then GdipSetImageAttributesWrapMode ImgAttr, WrapMode, 0
  GdipDrawImageRectRect mCtx, GetImage(key, True), X, Y, dx, dy, xSrc, ySrc, dxSrc, dySrc, 2, ImgAttr, 0, 0
  If ImgAttr Then GdipDisposeImageAttributes ImgAttr
End Sub

Public Function DrawString(ByVal hDC As Long, ByVal Font As stdole.IFont, ByVal S As String, X, Y, dx, dy, Optional ByVal DTFlags As eDrawText = DT_SINGLELINE Or DT_VCENTER, Optional ByVal TextColor&, Optional ByVal Alpha! = 1) As Long
  Dim LFW&(0 To 22), gdipFont As Long, gdipFormat As Long, R!(0 To 3), M!(0 To 3), CPs As Long
  Prepare hDC, TextColor, Alpha, 0
  GetObjectW Font.hFont, 92, LFW(0): GdipCreateFontFromLogfontW hDC, LFW(0), gdipFont
  If gdipFont = 0 Then  'let's make one more attempt, falling back to Arial
    Font.name = "Arial" 'because the passed font was probably not a TrueType-Font
    GetObjectW Font.hFont, 92, LFW(0): GdipCreateFontFromLogfontW hDC, LFW(0), gdipFont
  End If
  If gdipFont = 0 Then Err.Raise vbObjectError, , "Couldn't create GDIP-FontObject" 'we dont' try anymore
 
  GdipStringFormatGetGenericDefault gdipFormat
  If gdipFormat = 0 Then Err.Raise vbObjectError, , "Couldn't create GDIP-FormatObject"
  
  If DTFlags And DT_CENTER Then GdipSetStringFormatAlign gdipFormat, 1
  If DTFlags And DT_RIGHT Then GdipSetStringFormatAlign gdipFormat, 2
  If DTFlags And DT_VCENTER Then GdipSetStringFormatLineAlign gdipFormat, 1
  If DTFlags And DT_BOTTOM Then GdipSetStringFormatLineAlign gdipFormat, 2
  
  R(0) = X: R(1) = Y: R(2) = X + dx: R(3) = Y + dy
  GdipMeasureString mCtx, StrPtr(S), Len(S), gdipFont, R(0), gdipFormat, M(0), 0, DrawString
  If DTFlags And DT_CALCRECT Then
     X = M(0): Y = M(1): dx = M(2): dy = M(3)
  Else
     GdipDrawString mCtx, StrPtr(S), Len(S), gdipFont, R(0), gdipFormat, mBrush
  End If
  
  GdipDeleteStringFormat gdipFormat
  GdipDeleteFont gdipFont
End Function

Private Sub Prepare(hDC As Long, Color As Long, Alpha As Single, PenWidth As Single)
  ContextBindTo hDC
  If PenWidth Then GdipSetPenWidth mPen, PenWidth
  If PenWidth Then GdipSetPenColor mPen, MakeARGB(Color, Alpha) Else GdipSetSolidFillColor mBrush, MakeARGB(Color, Alpha)
End Sub

Public Function MakeARGB(ByVal Color As Long, ByVal Alpha As Single) As Long
  If Color = -1 Then Exit Function
  If (Color And &HFF000000) = &H80000000 Then Color = GetSysColor(Color And &HFFFF&)
  If Alpha < 0 Then Alpha = 0 Else If Alpha > 1 Then Alpha = 1
  Static B(0 To 3) As Byte: B(0) = Color \ 65536: B(1) = Color \ 256 And 255: B(2) = Color And 255: B(3) = Alpha * 255:
  MemCopy MakeARGB, B(0), 4
End Function

'another internally used Function, but defined Public - because it might be usefuly on the outside as well for other stuff
Public Function vtblCall(ByVal pObj As Long, ByVal vtblIdx As Long, ParamArray P() As Variant) As Long
  Const MaxArgs& = 32, CC_STDCALL& = 4
  Static vType(0 To MaxArgs - 1) As Integer, VPtr(0 To MaxArgs - 1) As Long
  Dim i As Long, V(), HRes As Long
  
  If pObj Then V = P Else Exit Function
  For i = 0 To UBound(V)
    vType(i) = VarType(V(i))
    VPtr(i) = VarPtr(V(i))
  Next
  
  HRes = DispCallFunc(pObj, vtblIdx * 4, CC_STDCALL, vbLong, i, vType(0), VPtr(0), vtblCall)
  If HRes Then Err.Raise HRes
End Function

'**** two pairs of Codec-Functions (for Base64 and RLE), mainly to support easy decoding of Image-Resources from a Base64-String-Const)
Public Function Base64Dec(sText) As Byte()
  Dim E As Object
  Set E = CreateObject("MSXML2.DOMDocument").CreateElement("E")
      E.DataType = "bin.base64": E.Text = sText
  Base64Dec = E.NodeTypedValue
End Function
Public Function Base64Enc(Bytes) As String
  Dim E As Object
  Set E = CreateObject("MSXML2.DOMDocument").CreateElement("E")
      E.DataType = "bin.base64": E.NodeTypedValue = Bytes
  Base64Enc = Replace(Replace(E.Text, vbLf, ""), vbCr, "")
End Function

Public Function RLEEnc(B() As Byte) As Byte()
  ReDim E(0 To UBound(B) * 2 + 1) As Byte 'twice the input-size
  Dim i As Long, j As Long, k As Long, n As Long
  For i = 0 To UBound(B)
    E(k) = B(i) \ 16: E(k + 1) = B(i) And 15: k = k + 2
  Next
  
  For i = 0 To UBound(E)
    For k = i + 1 To UBound(E)
      If n = 15 Or E(i) <> E(k) Then Exit For Else n = n + 1
    Next
    E(j) = n + 16 * E(i): j = j + 1: i = i + n: n = 0
  Next
  ReDim Preserve E(0 To j - 1): RLEEnc = E 'return the result
End Function
Public Function RLEDec(B() As Byte) As Byte()
  Dim i As Long, j As Long, k As Long, sz As Long
  For i = 0 To UBound(B): sz = sz + (B(i) And 15) + 1: Next
  ReDim D(0 To sz - 1) As Byte
  For i = 0 To UBound(B)
    For k = 0 To B(i) And 15: D(j) = B(i) \ 16: j = j + 1: Next
  Next
  
  For i = 0 To sz \ 2 - 1
    D(i) = D(i + i) * 16 + D(i + i + 1)
  Next
  ReDim Preserve D(0 To i - 1): RLEDec = D 'return the result
End Function
