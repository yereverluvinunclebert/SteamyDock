Attribute VB_Name = "GDIPlusAPI"
Option Explicit
' Translated by Avery P. - 7/29/2002
' NOTES:
'   - All GDI+ Strings expect and return ONLY Unicode - you'll need to use StrConv when using them or convert the APIs to use StrPtr.
'     As always, there are a few exceptions to a rule, and ImageTitles are one string that uses only ANSI strings.
'   - Functions with an I (i) at the end are non-floating point declarations.
'   - If a function without the I (i) at the end doesn't work, try the one with (if any).
'     If neither version worked, then you did something wrong, or the API *may* be misdeclared.
'   - The word (ALL) next to an API set mean all of the functions are declared. (ALL GDI+ functions are now declared 8/12/2002)
'   - Search for "TODO:" (no quotes) to see what still needs done within the file. If there is no TODO, there is nothing to do!
'   - If you want to get all of the encoder or decoder file extensions, MIME type, or other values, try to use the Get__Clsid functions as a base.
'   - If you don't like the idea of converting strings to and from Unicode, change all As String occurances in the API declarations
'     to As Long, and pass the StrPtr() result there instead. I opted to use the As String for clarity, especially since the GDI+ docs are
'     geared toward how to use the C++ classes.
'   - I may have misdeclared the IStream functions as I'm not too familiar with them. Do a "TODO:" or "IStream" search (no quotes) to see the
'     IStream functions. All parameters except one where declared as 'IStream* stream' in C++. The exception has a comment above it. The possible
'     problem is that the IStream parameters should be passed ByRef instead of ByVal. If they are wrong, please tell me!
'   - APIs are in ordered groups, just like the C++ header, but the groups themselves are not in the same order as in the C++ headers.
'
' WARNINGS:
'   - Some of the structs may not work, though I didn't test them all fully.
'   - If a function causes a GPF or performs unexpectedly, make sure you are passing correct arguments.
'     It also couldn't hurt to double-check the declarations as there is a chance they could be a bit off and looking in the MSDN can't hurt either.
'   - Some APIs that have a ByRef parameter may expect an array; check the MSDN to find out if unsure.
'
'-----------------------------------------------
' 2/6/2003
' - I suppose I should put change notifications here in case you missed them on PSC.
' - Altered the ColorPalette to have 256 color palette entries regardless and the flags member is mapped to the PaletteFlags enum.
' - ImageCodecInfo now has the flags member mapped to the ImageCodecFlags enum.
' - GdipBitmapLockBits flags parameter changed to the ImageLockMode enum instead of a Long.
' - Altered the LOGFONTW lfFaceName member to be a String, which is twice as long as the ANSI to adjust for unicode.
'   You'll want to use a StrConv on it to get the ANSI text. Also introduced a new constant to ease maintainance: LF_FACESIZEW
'-----------------------------------------------


'-----------------------------------------------
' String Pointer Related APIs (For the String Utilities)
'-----------------------------------------------
Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long


'-----------------------------------------------
' CLSID Generation Related APIs
'-----------------------------------------------
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long


'-----------------------------------------------
' GDI+ Constants
'-----------------------------------------------
Public Const LF_FACESIZE As Long = 32
Public Const LF_FACESIZEW As Long = LF_FACESIZE * 2

Public Const FlatnessDefault As Single = 1# / 4#

'Shift count and bit mask for A, R, G, B components
Public Const AlphaShift = 24
Public Const RedShift = 16
Public Const GreenShift = 8
Public Const BlueShift = 0

Public Const AlphaMask = &HFF000000
Public Const RedMask = &HFF0000
Public Const GreenMask = &HFF00
Public Const BlueMask = &HFF


' In-memory pixel data formats:
' bits 0-7 = format index
' bits 8-15 = pixel size (in bits)
' bits 16-23 = flags
' bits 24-31 = reserved
Public Const PixelFormatIndexed = &H10000           ' Indexes into a palette
Public Const PixelFormatGDI = &H20000               ' Is a GDI-supported format
Public Const PixelFormatAlpha = &H40000             ' Has an alpha component
Public Const PixelFormatPAlpha = &H80000            ' Pre-multiplied alpha
Public Const PixelFormatExtended = &H100000         ' Extended color 16 bits/channel
Public Const PixelFormatCanonical = &H200000

Public Const PixelFormatUndefined = 0
Public Const PixelFormatDontCare = 0

Public Const PixelFormat1bppIndexed = &H30101
Public Const PixelFormat4bppIndexed = &H30402
Public Const PixelFormat8bppIndexed = &H30803
Public Const PixelFormat16bppGreyScale = &H101004
Public Const PixelFormat16bppRGB555 = &H21005
Public Const PixelFormat16bppRGB565 = &H21006
Public Const PixelFormat16bppARGB1555 = &H61007
Public Const PixelFormat24bppRGB = &H21808
Public Const PixelFormat32bppRGB = &H22009
Public Const PixelFormat32bppARGB = &H26200A
Public Const PixelFormat32bppPARGB = &HE200B
Public Const PixelFormat48bppRGB = &H10300C
Public Const PixelFormat64bppARGB = &H34400D
Public Const PixelFormat64bppPARGB = &H1C400E
Public Const PixelFormatMax = 15 '&HF



' Image property types
Public Const PropertyTagTypeByte = 1
Public Const PropertyTagTypeASCII = 2
Public Const PropertyTagTypeShort = 3
Public Const PropertyTagTypeLong = 4
Public Const PropertyTagTypeRational = 5
Public Const PropertyTagTypeUndefined = 7
Public Const PropertyTagTypeSLONG = 9
Public Const PropertyTagTypeSRational = 10


' Image property ID tags
Public Const PropertyTagExifIFD = &H8769
Public Const PropertyTagGpsIFD = &H8825

Public Const PropertyTagNewSubfileType = &HFE
Public Const PropertyTagSubfileType = &HFF
Public Const PropertyTagImageWidth = &H100
Public Const PropertyTagImageHeight = &H101
Public Const PropertyTagBitsPerSample = &H102
Public Const PropertyTagCompression = &H103
Public Const PropertyTagPhotometricInterp = &H106
Public Const PropertyTagThreshHolding = &H107
Public Const PropertyTagCellWidth = &H108
Public Const PropertyTagCellHeight = &H109
Public Const PropertyTagFillOrder = &H10A
Public Const PropertyTagDocumentName = &H10D
Public Const PropertyTagImageDescription = &H10E
Public Const PropertyTagEquipMake = &H10F
Public Const PropertyTagEquipModel = &H110
Public Const PropertyTagStripOffsets = &H111
Public Const PropertyTagOrientation = &H112
Public Const PropertyTagSamplesPerPixel = &H115
Public Const PropertyTagRowsPerStrip = &H116
Public Const PropertyTagStripBytesCount = &H117
Public Const PropertyTagMinSampleValue = &H118
Public Const PropertyTagMaxSampleValue = &H119
Public Const PropertyTagXResolution = &H11A            ' Image resolution in width direction
Public Const PropertyTagYResolution = &H11B            ' Image resolution in height direction
Public Const PropertyTagPlanarConfig = &H11C           ' Image data arrangement
Public Const PropertyTagPageName = &H11D
Public Const PropertyTagXPosition = &H11E
Public Const PropertyTagYPosition = &H11F
Public Const PropertyTagFreeOffset = &H120
Public Const PropertyTagFreeByteCounts = &H121
Public Const PropertyTagGrayResponseUnit = &H122
Public Const PropertyTagGrayResponseCurve = &H123
Public Const PropertyTagT4Option = &H124
Public Const PropertyTagT6Option = &H125
Public Const PropertyTagResolutionUnit = &H128         ' Unit of X and Y resolution
Public Const PropertyTagPageNumber = &H129
Public Const PropertyTagTransferFuncition = &H12D
Public Const PropertyTagSoftwareUsed = &H131
Public Const PropertyTagDateTime = &H132
Public Const PropertyTagArtist = &H13B
Public Const PropertyTagHostComputer = &H13C
Public Const PropertyTagPredictor = &H13D
Public Const PropertyTagWhitePoint = &H13E
Public Const PropertyTagPrimaryChromaticities = &H13F
Public Const PropertyTagColorMap = &H140
Public Const PropertyTagHalftoneHints = &H141
Public Const PropertyTagTileWidth = &H142
Public Const PropertyTagTileLength = &H143
Public Const PropertyTagTileOffset = &H144
Public Const PropertyTagTileByteCounts = &H145
Public Const PropertyTagInkSet = &H14C
Public Const PropertyTagInkNames = &H14D
Public Const PropertyTagNumberOfInks = &H14E
Public Const PropertyTagDotRange = &H150
Public Const PropertyTagTargetPrinter = &H151
Public Const PropertyTagExtraSamples = &H152
Public Const PropertyTagSampleFormat = &H153
Public Const PropertyTagSMinSampleValue = &H154
Public Const PropertyTagSMaxSampleValue = &H155
Public Const PropertyTagTransferRange = &H156

Public Const PropertyTagJPEGProc = &H200
Public Const PropertyTagJPEGInterFormat = &H201
Public Const PropertyTagJPEGInterLength = &H202
Public Const PropertyTagJPEGRestartInterval = &H203
Public Const PropertyTagJPEGLosslessPredictors = &H205
Public Const PropertyTagJPEGPointTransforms = &H206
Public Const PropertyTagJPEGQTables = &H207
Public Const PropertyTagJPEGDCTables = &H208
Public Const PropertyTagJPEGACTables = &H209

Public Const PropertyTagYCbCrCoefficients = &H211
Public Const PropertyTagYCbCrSubsampling = &H212
Public Const PropertyTagYCbCrPositioning = &H213
Public Const PropertyTagREFBlackWhite = &H214

Public Const PropertyTagICCProfile = &H8773            ' This TAG is defined by ICC
                                                ' for embedded ICC in TIFF
Public Const PropertyTagGamma = &H301
Public Const PropertyTagICCProfileDescriptor = &H302
Public Const PropertyTagSRGBRenderingIntent = &H303

Public Const PropertyTagImageTitle = &H320
Public Const PropertyTagCopyright = &H8298

' Extra TAGs (Like Adobe Image Information tags etc.)

Public Const PropertyTagResolutionXUnit = &H5001
Public Const PropertyTagResolutionYUnit = &H5002
Public Const PropertyTagResolutionXLengthUnit = &H5003
Public Const PropertyTagResolutionYLengthUnit = &H5004
Public Const PropertyTagPrintFlags = &H5005
Public Const PropertyTagPrintFlagsVersion = &H5006
Public Const PropertyTagPrintFlagsCrop = &H5007
Public Const PropertyTagPrintFlagsBleedWidth = &H5008
Public Const PropertyTagPrintFlagsBleedWidthScale = &H5009
Public Const PropertyTagHalftoneLPI = &H500A
Public Const PropertyTagHalftoneLPIUnit = &H500B
Public Const PropertyTagHalftoneDegree = &H500C
Public Const PropertyTagHalftoneShape = &H500D
Public Const PropertyTagHalftoneMisc = &H500E
Public Const PropertyTagHalftoneScreen = &H500F
Public Const PropertyTagJPEGQuality = &H5010
Public Const PropertyTagGridSize = &H5011
Public Const PropertyTagThumbnailFormat = &H5012            ' 1 = JPEG, 0 = RAW RGB
Public Const PropertyTagThumbnailWidth = &H5013
Public Const PropertyTagThumbnailHeight = &H5014
Public Const PropertyTagThumbnailColorDepth = &H5015
Public Const PropertyTagThumbnailPlanes = &H5016
Public Const PropertyTagThumbnailRawBytes = &H5017
Public Const PropertyTagThumbnailSize = &H5018
Public Const PropertyTagThumbnailCompressedSize = &H5019
Public Const PropertyTagColorTransferFunction = &H501A
Public Const PropertyTagThumbnailData = &H501B            ' RAW thumbnail bits in
                                                   ' JPEG format or RGB format
                                                   ' depends on
                                                   ' PropertyTagThumbnailFormat

' Thumbnail related TAGs
Public Const PropertyTagThumbnailImageWidth = &H5020        ' Thumbnail width
Public Const PropertyTagThumbnailImageHeight = &H5021       ' Thumbnail height
Public Const PropertyTagThumbnailBitsPerSample = &H5022     ' Number of bits per
                                                     ' component
Public Const PropertyTagThumbnailCompression = &H5023       ' Compression Scheme
Public Const PropertyTagThumbnailPhotometricInterp = &H5024 ' Pixel composition
Public Const PropertyTagThumbnailImageDescription = &H5025  ' Image Tile
Public Const PropertyTagThumbnailEquipMake = &H5026         ' Manufacturer of Image
                                                     ' Input equipment
Public Const PropertyTagThumbnailEquipModel = &H5027        ' Model of Image input
                                                     ' equipment
Public Const PropertyTagThumbnailStripOffsets = &H5028      ' Image data location
Public Const PropertyTagThumbnailOrientation = &H5029       ' Orientation of image
Public Const PropertyTagThumbnailSamplesPerPixel = &H502A   ' Number of components
Public Const PropertyTagThumbnailRowsPerStrip = &H502B      ' Number of rows per strip
Public Const PropertyTagThumbnailStripBytesCount = &H502C   ' Bytes per compressed
                                                     ' strip
Public Const PropertyTagThumbnailResolutionX = &H502D       ' Resolution in width
                                                     ' direction
Public Const PropertyTagThumbnailResolutionY = &H502E       ' Resolution in height
                                                     ' direction
Public Const PropertyTagThumbnailPlanarConfig = &H502F      ' Image data arrangement
Public Const PropertyTagThumbnailResolutionUnit = &H5030    ' Unit of X and Y
                                                     ' Resolution
Public Const PropertyTagThumbnailTransferFunction = &H5031  ' Transfer function
Public Const PropertyTagThumbnailSoftwareUsed = &H5032      ' Software used
Public Const PropertyTagThumbnailDateTime = &H5033          ' File change date and
                                                     ' time
Public Const PropertyTagThumbnailArtist = &H5034            ' Person who created the
                                                     ' image
Public Const PropertyTagThumbnailWhitePoint = &H5035        ' White point chromaticity
Public Const PropertyTagThumbnailPrimaryChromaticities = &H5036
                                                     ' Chromaticities of
                                                     ' primaries
Public Const PropertyTagThumbnailYCbCrCoefficients = &H5037 ' Color space transforma-
                                                     ' tion coefficients
Public Const PropertyTagThumbnailYCbCrSubsampling = &H5038  ' Subsampling ratio of Y
                                                     ' to C
Public Const PropertyTagThumbnailYCbCrPositioning = &H5039  ' Y and C position
Public Const PropertyTagThumbnailRefBlackWhite = &H503A     ' Pair of black and white
                                                     ' reference values
Public Const PropertyTagThumbnailCopyRight = &H503B         ' CopyRight holder

Public Const PropertyTagLuminanceTable = &H5090
Public Const PropertyTagChrominanceTable = &H5091

Public Const PropertyTagFrameDelay = &H5100
Public Const PropertyTagLoopCount = &H5101

Public Const PropertyTagPixelUnit = &H5110          ' Unit specifier for pixel/unit
Public Const PropertyTagPixelPerUnitX = &H5111      ' Pixels per unit in X
Public Const PropertyTagPixelPerUnitY = &H5112      ' Pixels per unit in Y
Public Const PropertyTagPaletteHistogram = &H5113   ' Palette histogram

' EXIF specific tag

Public Const PropertyTagExifExposureTime = &H829A
Public Const PropertyTagExifFNumber = &H829D

Public Const PropertyTagExifExposureProg = &H8822
Public Const PropertyTagExifSpectralSense = &H8824
Public Const PropertyTagExifISOSpeed = &H8827
Public Const PropertyTagExifOECF = &H8828

Public Const PropertyTagExifVer = &H9000
Public Const PropertyTagExifDTOrig = &H9003         ' Date & time of original
Public Const PropertyTagExifDTDigitized = &H9004    ' Date & time of digital data generation

Public Const PropertyTagExifCompConfig = &H9101
Public Const PropertyTagExifCompBPP = &H9102

Public Const PropertyTagExifShutterSpeed = &H9201
Public Const PropertyTagExifAperture = &H9202
Public Const PropertyTagExifBrightness = &H9203
Public Const PropertyTagExifExposureBias = &H9204
Public Const PropertyTagExifMaxAperture = &H9205
Public Const PropertyTagExifSubjectDist = &H9206
Public Const PropertyTagExifMeteringMode = &H9207
Public Const PropertyTagExifLightSource = &H9208
Public Const PropertyTagExifFlash = &H9209
Public Const PropertyTagExifFocalLength = &H920A
Public Const PropertyTagExifMakerNote = &H927C
Public Const PropertyTagExifUserComment = &H9286
Public Const PropertyTagExifDTSubsec = &H9290        ' Date & Time subseconds
Public Const PropertyTagExifDTOrigSS = &H9291        ' Date & Time original subseconds
Public Const PropertyTagExifDTDigSS = &H9292         ' Date & TIme digitized subseconds

Public Const PropertyTagExifFPXVer = &HA000
Public Const PropertyTagExifColorSpace = &HA001
Public Const PropertyTagExifPixXDim = &HA002
Public Const PropertyTagExifPixYDim = &HA003
Public Const PropertyTagExifRelatedWav = &HA004      ' related sound file
Public Const PropertyTagExifInterop = &HA005
Public Const PropertyTagExifFlashEnergy = &HA20B
Public Const PropertyTagExifSpatialFR = &HA20C       ' Spatial Frequency Response
Public Const PropertyTagExifFocalXRes = &HA20E       ' Focal Plane X Resolution
Public Const PropertyTagExifFocalYRes = &HA20F       ' Focal Plane Y Resolution
Public Const PropertyTagExifFocalResUnit = &HA210    ' Focal Plane Resolution Unit
Public Const PropertyTagExifSubjectLoc = &HA214
Public Const PropertyTagExifExposureIndex = &HA215
Public Const PropertyTagExifSensingMethod = &HA217
Public Const PropertyTagExifFileSource = &HA300
Public Const PropertyTagExifSceneType = &HA301
Public Const PropertyTagExifCfaPattern = &HA302

Public Const PropertyTagGpsVer = &H0
Public Const PropertyTagGpsLatitudeRef = &H1
Public Const PropertyTagGpsLatitude = &H2
Public Const PropertyTagGpsLongitudeRef = &H3
Public Const PropertyTagGpsLongitude = &H4
Public Const PropertyTagGpsAltitudeRef = &H5
Public Const PropertyTagGpsAltitude = &H6
Public Const PropertyTagGpsGpsTime = &H7
Public Const PropertyTagGpsGpsSatellites = &H8
Public Const PropertyTagGpsGpsStatus = &H9
Public Const PropertyTagGpsGpsMeasureMode = &HA
Public Const PropertyTagGpsGpsDop = &HB              ' Measurement precision
Public Const PropertyTagGpsSpeedRef = &HC
Public Const PropertyTagGpsSpeed = &HD
Public Const PropertyTagGpsTrackRef = &HE
Public Const PropertyTagGpsTrack = &HF
Public Const PropertyTagGpsImgDirRef = &H10
Public Const PropertyTagGpsImgDir = &H11
Public Const PropertyTagGpsMapDatum = &H12
Public Const PropertyTagGpsDestLatRef = &H13
Public Const PropertyTagGpsDestLat = &H14
Public Const PropertyTagGpsDestLongRef = &H15
Public Const PropertyTagGpsDestLong = &H16
Public Const PropertyTagGpsDestBearRef = &H17
Public Const PropertyTagGpsDestBear = &H18
Public Const PropertyTagGpsDestDistRef = &H19
Public Const PropertyTagGpsDestDist = &H1A


'//---------------------------------------------------------------------------
'// Image file format identifiers
'//---------------------------------------------------------------------------
Public Const ImageFormatSuffix        As String = "-0728-11D3-9D7B-0000F81EF32E}"
Public Const ImageFormatUndefined     As String = "{B96B3CA9" & ImageFormatSuffix
Public Const ImageFormatMemoryBMP     As String = "{B96B3CAA" & ImageFormatSuffix
Public Const ImageFormatBMP           As String = "{B96B3CAB" & ImageFormatSuffix
Public Const ImageFormatEMF           As String = "{B96B3CAC" & ImageFormatSuffix
Public Const ImageFormatWMF           As String = "{B96B3CAD" & ImageFormatSuffix
Public Const ImageFormatJPEG          As String = "{B96B3CAE" & ImageFormatSuffix
Public Const ImageFormatPNG           As String = "{B96B3CAF" & ImageFormatSuffix
Public Const ImageFormatGIF           As String = "{B96B3CB0" & ImageFormatSuffix
Public Const ImageFormatTIFF          As String = "{B96B3CB1" & ImageFormatSuffix
Public Const ImageFormatEXIF          As String = "{B96B3CB2" & ImageFormatSuffix
Public Const ImageFormatIcon          As String = "{B96B3CB5" & ImageFormatSuffix
'//---------------------------------------------------------------------------
'// Predefined multi-frame dimension IDs
'//---------------------------------------------------------------------------
Public Const FrameDimensionTime       As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
Public Const FrameDimensionResolution As String = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"
Public Const FrameDimensionPage       As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"
'//---------------------------------------------------------------------------
'// Property sets
'//---------------------------------------------------------------------------
Public Const FormatIDImageInformation As String = "{E5836CBE-5EEF-0F1D-ACDE-AE4C43B608CE}"
Public Const FormatIDJpegAppHeaders   As String = "{1C4AFDCD-6177-43CF-ABC7-5F51AF39EE85}"
'//---------------------------------------------------------------------------
'// Encoder parameter sets
'//---------------------------------------------------------------------------
Public Const EncoderCompression       As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EncoderColorDepth        As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EncoderScanMethod        As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EncoderVersion           As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EncoderRenderMethod      As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EncoderQuality           As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EncoderTransformation    As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EncoderLuminanceTable    As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EncoderChrominanceTable  As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EncoderSaveFlag          As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Public Const CodecIImageBytes         As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"


'-----------------------------------------------
' The following types are NOT in the GDI+ docs, per se
'-----------------------------------------------
Public Type POINTL    ' aka Point
   x As Long
   y As Long
End Type

Public Type POINTF   ' aka PointF
   x As Single
   y As Single
End Type

Public Type RECTL     ' aka Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RECTF    ' aka RectF
   Left As Single
   Top As Single
   Right As Single
   Bottom As Single
End Type

Public Type SIZEL    ' aka Size
   cx As Long
   cy As Long
End Type

Public Type SIZEF    ' aka SizeF
   cx As Single
   cy As Single
End Type

' Custom types
Public Type COLORBYTES
   BlueByte As Byte
   GreenByte As Byte
   RedByte As Byte
   AlphaByte As Byte
End Type
Public Type COLORLONG
   longval As Long
End Type



'-----------------------------------------------
' GDI+ Structs/Types
'-----------------------------------------------

Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
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

' Encoder Parameter structure
Public Type EncoderParameter
   GUID As CLSID                          ' GUID of the parameter
   NumberOfValues As Long                 ' Number of the parameter values; usually 1
   type As EncoderParameterValueType      ' Value type, like ValueTypeLONG  etc.
   value As Long                          ' A pointer to the parameter values
End Type

' Encoder Parameters structure
Public Type EncoderParameters
   count As Long                          ' Number of parameters in this structure; Should be 1
   Parameter As EncoderParameter          ' Parameter values; this CAN be an array!!!! (Use CopyMemory and a string or byte array as workaround)
End Type

Public Type ColorPalette
   flags As PaletteFlags      ' Palette flags; should be a Long equivalent
   count As Long              ' Number of color entries used
   Entries(0 To 255) As Long  ' Palette color entries. WARNING: SDK defines as 1 entry, but I made it
                              ' contain 256 (a reasonable amount) for use with any palette count 256 or less
                              ' since VB can't malloc and do type casting like C/C++.
End Type

Public Type ColorMap
   oldColor As Long       ' Color oldColor;
   newColor As Long       ' Color newColor;
End Type

Public Type ColorMatrix
   m(0 To 4, 0 To 4) As Single
End Type

' Information about image pixel data
Public Type BitmapData
   Width As Long
   Height As Long
   stride As Long
   PixelFormat As Long
   scan0 As Long
   Reserved As Long
End Type

Public Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors As RGBQUAD
End Type

Public Type PathData
   count As Long
   Points As Long    ' Pointer to POINTF array
   types As Long     ' Pointer to BYTE array
End Type

Public Type PropertyItem
   propId As Long              ' ID of this property
   length As Long              ' Length of the property value, in bytes
   type As Integer             ' Type of the value, as one of TAG_TYPE_XXX
                               ' defined above
   value As Long               ' property value
End Type

Public Type LOGFONTA
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
   lfFaceName As String * LF_FACESIZE
End Type

Public Type LOGFONTW
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
   lfFaceName As String * LF_FACESIZEW
End Type

Public Type CharacterRange
   First As Long
   length As Long
End Type

Public Type PWMFRect16
   Left As Integer
   Top As Integer
   Right As Integer
   Bottom As Integer
End Type

Public Type WmfPlaceableFileHeader
   Key As Long                        ' GDIP_WMF_PLACEABLEKEY
   Hmf As Integer                     ' Metafile HANDLE number (always 0)
   boundingBox As PWMFRect16          ' Coordinates in metafile units
   Inch As Integer                    ' Number of metafile units per inch
   Reserved As Long                   ' Reserved (always 0)
   Checksum As Integer                ' Checksum value for previous 10 WORDs
End Type

Public Type ENHMETAHEADER3
   itype As Long               ' Record type EMR_HEADER
   nSize As Long               ' Record size in bytes.  This may be greater
                               ' than the sizeof(ENHMETAHEADER).
   rclBounds As RECTL        ' Inclusive-inclusive bounds in device units
   rclFrame As RECTL         ' Inclusive-inclusive Picture Frame .01mm unit
   dSignature As Long          ' Signature.  Must be ENHMETA_SIGNATURE.
   nVersion As Long            ' Version number
   nBytes As Long              ' Size of the metafile in bytes
   nRecords As Long            ' Number of records in the metafile
   nHandles As Integer         ' Number of handles in the handle table
                               ' Handle index zero is reserved.
   sReserved As Integer        ' Reserved.  Must be zero.
   nDescription As Long        ' Number of chars in the unicode desc string
                               ' This is 0 if there is no description string
   offDescription As Long      ' Offset to the metafile description record.
                               ' This is 0 if there is no description string
   nPalEntries As Long         ' Number of entries in the metafile palette.
   szlDevice As SIZEL           ' Size of the reference device in pels
   szlMillimeters As SIZEL      ' Size of the reference device in millimeters
End Type

Public Type METAHEADER
   mtType As Integer
   mtHeaderSize As Integer
   mtVersion As Integer
   mtSize As Long
   mtNoObjects As Integer
   mtMaxRecord As Long
   mtNoParameters As Integer
End Type

Public Type MetafileHeader
   mType As MetafileType
   size As Long                ' Size of the metafile (in bytes)
   Version As Long             ' EMF+, EMF, or WMF version
   EmfPlusFlags As Long
   DpiX As Single
   DpiY As Single
   x As Long                   ' Bounds in device units
   y As Long
   Width As Long
   Height As Long
   'Union
   '{
   '    METAHEADER WmfHeader
   '    ENHMETAHEADER3 EmfHeader
   '}
   EmfHeader As ENHMETAHEADER3 ' NOTE: You'll have to use CopyMemory to view the METAHEADER type
   EmfPlusHeaderSize As Long   ' size of the EMF+ header in file
   LogicalDpiX As Long         ' Logical Dpi of reference Hdc
   LogicalDpiY As Long         ' usually valid only for EMF+
End Type



'-----------------------------------------------
' GDI+ Enums
'-----------------------------------------------

Public Enum PaletteFlags
   PaletteFlagsHasAlpha = &H1
   PaletteFlagsGrayScale = &H2
   PaletteFlagsHalftone = &H4
End Enum

Public Enum GpUnit  ' aka Unit
   UnitWorld      ' 0 -- World coordinate (non-physical unit)
   UnitDisplay    ' 1 -- Variable -- for PageTransform only
   UnitPixel      ' 2 -- Each unit is one device pixel.
   UnitPoint      ' 3 -- Each unit is a printer's point, or 1/72 inch.
   UnitInch       ' 4 -- Each unit is 1 inch.
   UnitDocument   ' 5 -- Each unit is 1/300 inch.
   UnitMillimeter ' 6 -- Each unit is 1 millimeter.
End Enum

' Common color constants
' NOTE: Oringinal enum was unnamed
Public Enum Colors
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
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
End Enum

' Quality mode constants
Public Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       ' Best performance
   QualityModeHigh = 2       ' Best rendering quality
End Enum

' Alpha Compositing mode constants
Public Enum CompositingMode
   CompositingModeSourceOver    ' 0
   CompositingModeSourceCopy    ' 1
End Enum

' Alpha Compositing quality constants
Public Enum CompositingQuality
   CompositingQualityInvalid = QualityModeInvalid
   CompositingQualityDefault = QualityModeDefault
   CompositingQualityHighSpeed = QualityModeLow
   CompositingQualityHighQuality = QualityModeHigh
   CompositingQualityGammaCorrected
   CompositingQualityAssumeLinear
End Enum

' Generic font families
Public Enum GenericFontFamily
   GenericFontFamilySerif
   GenericFontFamilySansSerif
   GenericFontFamilyMonospace
End Enum

' FontStyle: face types and common styles
Public Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum

Public Enum SmoothingMode
   SmoothingModeInvalid = QualityModeInvalid
   SmoothingModeDefault = QualityModeDefault
   SmoothingModeHighSpeed = QualityModeLow
   SmoothingModeHighQuality = QualityModeHigh
   SmoothingModeNone
   SmoothingModeAntiAlias
End Enum

Public Enum FillMode
   FillModeAlternate        ' 0
   FillModeWinding           ' 1
End Enum

Public Enum InterpolationMode
   InterpolationModeInvalid = QualityModeInvalid
   InterpolationModeDefault = QualityModeDefault
   InterpolationModeLowQuality = QualityModeLow
   InterpolationModeHighQuality = QualityModeHigh
   InterpolationModeBilinear
   InterpolationModeBicubic
   InterpolationModeNearestNeighbor
   InterpolationModeHighQualityBilinear
   InterpolationModeHighQualityBicubic
End Enum

' Various wrap modes for brushes
Public Enum WrapMode
   WrapModeTile         ' 0
   WrapModeTileFlipX    ' 1
   WrapModeTileFlipY    ' 2
   WrapModeTileFlipXY   ' 3
   WrapModeClamp        ' 4
End Enum

Public Enum LinearGradientMode
   LinearGradientModeHorizontal          ' 0
   LinearGradientModeVertical            ' 1
   LinearGradientModeForwardDiagonal     ' 2
   LinearGradientModeBackwardDiagonal    ' 3
End Enum

Public Enum ImageType
   ImageTypeUnknown    ' 0
   ImageTypeBitmap     ' 1
   ImageTypeMetafile   ' 2
End Enum

' Various Hatch Styles
Public Enum HatchStyle
   HatchStyleHorizontal                   ' 0
   HatchStyleVertical                     ' 1
   HatchStyleForwardDiagonal              ' 2
   HatchStyleBackwardDiagonal             ' 3
   HatchStyleCross                        ' 4
   HatchStyleDiagonalCross                ' 5
   HatchStyle05Percent                    ' 6
   HatchStyle10Percent                    ' 7
   HatchStyle20Percent                    ' 8
   HatchStyle25Percent                    ' 9
   HatchStyle30Percent                    ' 10
   HatchStyle40Percent                    ' 11
   HatchStyle50Percent                    ' 12
   HatchStyle60Percent                    ' 13
   HatchStyle70Percent                    ' 14
   HatchStyle75Percent                    ' 15
   HatchStyle80Percent                    ' 16
   HatchStyle90Percent                    ' 17
   HatchStyleLightDownwardDiagonal        ' 18
   HatchStyleLightUpwardDiagonal          ' 19
   HatchStyleDarkDownwardDiagonal         ' 20
   HatchStyleDarkUpwardDiagonal           ' 21
   HatchStyleWideDownwardDiagonal         ' 22
   HatchStyleWideUpwardDiagonal           ' 23
   HatchStyleLightVertical                ' 24
   HatchStyleLightHorizontal              ' 25
   HatchStyleNarrowVertical               ' 26
   HatchStyleNarrowHorizontal             ' 27
   HatchStyleDarkVertical                 ' 28
   HatchStyleDarkHorizontal               ' 29
   HatchStyleDashedDownwardDiagonal       ' 30
   HatchStyleDashedUpwardDiagonal         ' 31
   HatchStyleDashedHorizontal             ' 32
   HatchStyleDashedVertical               ' 33
   HatchStyleSmallConfetti                ' 34
   HatchStyleLargeConfetti                ' 35
   HatchStyleZigZag                       ' 36
   HatchStyleWave                         ' 37
   HatchStyleDiagonalBrick                ' 38
   HatchStyleHorizontalBrick              ' 39
   HatchStyleWeave                        ' 40
   HatchStylePlaid                        ' 41
   HatchStyleDivot                        ' 42
   HatchStyleDottedGrid                   ' 43
   HatchStyleDottedDiamond                ' 44
   HatchStyleShingle                      ' 45
   HatchStyleTrellis                      ' 46
   HatchStyleSphere                       ' 47
   HatchStyleSmallGrid                    ' 48
   HatchStyleSmallCheckerBoard            ' 49
   HatchStyleLargeCheckerBoard            ' 50
   HatchStyleOutlinedDiamond              ' 51
   HatchStyleSolidDiamond                 ' 52

   HatchStyleTotal
   HatchStyleLargeGrid = HatchStyleCross  ' 4

   HatchStyleMin = HatchStyleHorizontal
   HatchStyleMax = HatchStyleTotal - 1
End Enum

Public Enum MatrixOrder
   MatrixOrderPrepend = 0
   MatrixOrderAppend = 1
End Enum

Public Enum ColorAdjustType
   ColorAdjustTypeDefault
   ColorAdjustTypeBitmap
   ColorAdjustTypeBrush
   ColorAdjustTypePen
   ColorAdjustTypeText
   ColorAdjustTypeCount
   ColorAdjustTypeAny      ' Reserved
End Enum

Public Enum ColorChannelFlags
   ColorChannelFlagsC = 0
   ColorChannelFlagsM
   ColorChannelFlagsY
   ColorChannelFlagsK
   ColorChannelFlagsLast
End Enum

Public Enum ColorMatrixFlags
   ColorMatrixFlagsDefault = 0
   ColorMatrixFlagsSkipGrays = 1
   ColorMatrixFlagsAltGray = 2
End Enum

Public Enum PenAlignment
    PenAlignmentCenter = 0
    PenAlignmentInset = 1
End Enum

Public Enum BrushType
   BrushTypeSolidColor = 0
   BrushTypeHatchFill = 1
   BrushTypeTextureFill = 2
   BrushTypePathGradient = 3
   BrushTypeLinearGradient = 4
End Enum

Public Enum DashStyle
   DashStyleSolid          ' 0
   DashStyleDash           ' 1
   DashStyleDot            ' 2
   DashStyleDashDot        ' 3
   DashStyleDashDotDot     ' 4
   DashStyleCustom         ' 5
End Enum

' Dash cap constants
Public Enum DashCap
   DashCapFlat = 0
   DashCapRound = 2
   DashCapTriangle = 3
End Enum

' Line cap constants (only the lowest 8 bits are used).
Public Enum LineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3

   LineCapNoAnchor = &H10         ' corresponds to flat cap
   LineCapSquareAnchor = &H11     ' corresponds to square cap
   LineCapRoundAnchor = &H12      ' corresponds to round cap
   LineCapDiamondAnchor = &H13    ' corresponds to triangle cap
   LineCapArrowAnchor = &H14      ' no correspondence

   LineCapCustom = &HFF           ' custom cap

   LineCapAnchorMask = &HF0        ' mask to check for anchor or not.
End Enum

' Custom Line cap type constants
Public Enum CustomLineCapType
   CustomLineCapTypeDefault = 0
   CustomLineCapTypeAdjustableArrow = 1
End Enum

' Line join constants
Public Enum LineJoin
   LineJoinMiter = 0
   LineJoinBevel = 1
   LineJoinRound = 2
   LineJoinMiterClipped = 3
End Enum

' Pen's Fill types
Public Enum PenType
   PenTypeSolidColor = BrushTypeSolidColor
   PenTypeHatchFill = BrushTypeHatchFill
   PenTypeTextureFill = BrushTypeTextureFill
   PenTypePathGradient = BrushTypePathGradient
   PenTypeLinearGradient = BrushTypeLinearGradient
   PenTypeUnknown = -1
End Enum

Public Enum WarpMode
   WarpModePerspective     ' 0
   WarpModeBilinear        ' 1
End Enum

' Region Comine Modes
Public Enum CombineMode
   CombineModeReplace      ' 0
   CombineModeIntersect    ' 1
   CombineModeUnion        ' 2
   CombineModeXor          ' 3
   CombineModeExclude      ' 4
   CombineModeComplement   ' 5 (Exclude From)
End Enum

Public Enum RotateFlipType
   RotateNoneFlipNone = 0
   Rotate90FlipNone = 1
   Rotate180FlipNone = 2
   Rotate270FlipNone = 3

   RotateNoneFlipX = 4
   Rotate90FlipX = 5
   Rotate180FlipX = 6
   Rotate270FlipX = 7

   RotateNoneFlipY = Rotate180FlipX
   Rotate90FlipY = Rotate270FlipX
   Rotate180FlipY = RotateNoneFlipX
   Rotate270FlipY = Rotate90FlipX

   RotateNoneFlipXY = Rotate180FlipNone
   Rotate90FlipXY = Rotate270FlipNone
   Rotate180FlipXY = RotateNoneFlipNone
   Rotate270FlipXY = Rotate90FlipNone
End Enum


' String format flags
'
'  DirectionRightToLeft          - For horizontal text, the reading order is
'                                  right to left. This value is called
'                                  the base embedding level by the Unicode
'                                  bidirectional engine.
'                                  For vertical text, columns are read from
'                                  right to left.
'                                  By default, horizontal or vertical text is
'                                  read from left to right.
'
'  DirectionVertical             - Individual lines of text are vertical. In
'                                  each line, characters progress from top to
'                                  bottom.
'                                  By default, lines of text are horizontal,
'                                  each new line below the previous line.
'
'  NoFitBlackBox                 - Allows parts of glyphs to overhang the
'                                  bounding rectangle.
'                                  By default glyphs are first aligned
'                                  inside the margines, then any glyphs which
'                                  still overhang the bounding box are
'                                  repositioned to avoid any overhang.
'                                  For example when an italic
'                                  lower case letter f in a font such as
'                                  Garamond is aligned at the far left of a
'                                  rectangle, the lower part of the f will
'                                  reach slightly further left than the left
'                                  edge of the rectangle. Setting this flag
'                                  will ensure the character aligns visually
'                                  with the lines above and below, but may
'                                  cause some pixels outside the formatting
'                                  rectangle to be clipped or painted.
'
'  DisplayFormatControl          - Causes control characters such as the
'                                  left-to-right mark to be shown in the
'                                  output with a representative glyph.
'
'  NoFontFallback                - Disables fallback to alternate fonts for
'                                  characters not supported in the requested
'                                  font. Any missing characters will be
'                                  be displayed with the fonts missing glyph,
'                                  usually an open square.
'
'  NoWrap                        - Disables wrapping of text between lines
'                                  when formatting within a rectangle.
'                                  NoWrap is implied when a point is passed
'                                  instead of a rectangle, or when the
'                                  specified rectangle has a zero line length.
'
'  NoClip                        - By default text is clipped to the
'                                  formatting rectangle. Setting NoClip
'                                  allows overhanging pixels to affect the
'                                  device outside the formatting rectangle.
'                                  Pixels at the end of the line may be
'                                  affected if the glyphs overhang their
'                                  cells, and either the NoFitBlackBox flag
'                                  has been set, or the glyph extends to far
'                                  to be fitted.
'                                  Pixels above/before the first line or
'                                  below/after the last line may be affected
'                                  if the glyphs extend beyond their cell
'                                  ascent / descent. This can occur rarely
'                                  with unusual diacritic mark combinations.
Public Enum StringFormatFlags
   StringFormatFlagsDirectionRightToLeft = &H1
   StringFormatFlagsDirectionVertical = &H2
   StringFormatFlagsNoFitBlackBox = &H4
   StringFormatFlagsDisplayFormatControl = &H20
   StringFormatFlagsNoFontFallback = &H400
   StringFormatFlagsMeasureTrailingSpaces = &H800
   StringFormatFlagsNoWrap = &H1000
   StringFormatFlagsLineLimit = &H2000

   StringFormatFlagsNoClip = &H4000
End Enum

Public Enum StringTrimming
   StringTrimmingNone = 0
   StringTrimmingCharacter = 1
   StringTrimmingWord = 2
   StringTrimmingEllipsisCharacter = 3
   StringTrimmingEllipsisWord = 4
   StringTrimmingEllipsisPath = 5
End Enum

' National language digit substitution
Public Enum StringDigitSubstitute
   StringDigitSubstituteUser = 0         ' As NLS setting
   StringDigitSubstituteNone = 1
   StringDigitSubstituteNational = 2
   StringDigitSubstituteTraditional = 3
End Enum

' Hotkey prefix interpretation
Public Enum HotkeyPrefix
   HotkeyPrefixNone = 0
   HotkeyPrefixShow = 1
   HotkeyPrefixHide = 2
End Enum

Public Enum StringAlignment
   ' Left edge for left-to-right text,
   ' right for right-to-left text,
   ' and top for vertical
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Public Enum FlushIntention
   FlushIntentionFlush = 0         ' Flush all batched rendering operations
   FlushIntentionSync = 1          ' Flush all batched rendering operations
                                   ' and wait for them to complete
End Enum

' Image encoder parameter related types
Public Enum EncoderParameterValueType
   EncoderParameterValueTypeByte = 1              ' 8-bit unsigned int
   EncoderParameterValueTypeASCII = 2             ' 8-bit byte containing one 7-bit ASCII
                                                   ' code. NULL terminated.
   EncoderParameterValueTypeShort = 3             ' 16-bit unsigned int
   EncoderParameterValueTypeLong = 4              ' 32-bit unsigned int
   EncoderParameterValueTypeRational = 5          ' Two Longs. The first Long is the
                                                   ' numerator the second Long expresses the
                                                   ' denomintor.
   EncoderParameterValueTypeLongRange = 6         ' Two longs which specify a range of
                                                   ' integer values. The first Long specifies
                                                   ' the lower end and the second one
                                                   ' specifies the higher end. All values
                                                   ' are inclusive at both ends
   EncoderParameterValueTypeUndefined = 7         ' 8-bit byte that can take any value
                                                   ' depending on field definition
   EncoderParameterValueTypeRationalRange = 8      ' Two Rationals. The first Rational
                                                   ' specifies the lower end and the second
                                                   ' specifies the higher end. All values
                                                   ' are inclusive at both ends
End Enum

' Image encoder value types
Public Enum EncoderValue
   EncoderValueColorTypeCMYK
   EncoderValueColorTypeYCCK
   EncoderValueCompressionLZW
   EncoderValueCompressionCCITT3
   EncoderValueCompressionCCITT4
   EncoderValueCompressionRle
   EncoderValueCompressionNone
   EncoderValueScanMethodInterlaced
   EncoderValueScanMethodNonInterlaced
   EncoderValueVersionGif87
   EncoderValueVersionGif89
   EncoderValueRenderProgressive
   EncoderValueRenderNonProgressive
   EncoderValueTransformRotate90
   EncoderValueTransformRotate180
   EncoderValueTransformRotate270
   EncoderValueTransformFlipHorizontal
   EncoderValueTransformFlipVertical
   EncoderValueMultiFrame
   EncoderValueLastFrame
   EncoderValueFlush
   EncoderValueFrameDimensionTime
   EncoderValueFrameDimensionResolution
   EncoderValueFrameDimensionPage
End Enum

Public Enum PixelOffsetMode
   PixelOffsetModeInvalid = QualityModeInvalid
   PixelOffsetModeDefault = QualityModeDefault
   PixelOffsetModeHighSpeed = QualityModeLow
   PixelOffsetModeHighQuality = QualityModeHigh
   PixelOffsetModeNone    ' No pixel offset
   PixelOffsetModeHalf     ' Offset by -0.5 -0.5 for fast anti-alias perf
End Enum

Public Enum TextRenderingHint
   TextRenderingHintSystemDefault = 0            ' Glyph with system default rendering hint
   TextRenderingHintSingleBitPerPixelGridFit     ' Glyph bitmap with hinting
   TextRenderingHintSingleBitPerPixel            ' Glyph bitmap without hinting
   TextRenderingHintAntiAliasGridFit             ' Glyph anti-alias bitmap with hinting
   TextRenderingHintAntiAlias                    ' Glyph anti-alias bitmap without hinting
   TextRenderingHintClearTypeGridFit              ' Glyph CT bitmap with hinting
End Enum

Public Enum MetafileType
   MetafileTypeInvalid            ' Invalid metafile
   MetafileTypeWmf                ' Standard WMF
   MetafileTypeWmfPlaceable       ' Placeable WMF
   MetafileTypeEmf                ' EMF (not EMF+)
   MetafileTypeEmfPlusOnly        ' EMF+ without dual down-level records
   MetafileTypeEmfPlusDual         ' EMF+ with dual down-level records
End Enum

' Specifies the type of EMF to record
Public Enum EmfType
    EmfTypeEmfOnly = MetafileTypeEmf               ' no EMF+  only EMF
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly   ' no EMF  only EMF+
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual   ' both EMF+ and EMF
End Enum

' EMF+ Persistent object types
Public Enum ObjectType
    ObjectTypeInvalid
    ObjectTypeBrush
    ObjectTypePen
    ObjectTypePath
    ObjectTypeRegion
    ObjectTypeImage
    ObjectTypeFont
    ObjectTypeStringFormat
    ObjectTypeImageAttributes
    ObjectTypeCustomLineCap

    ObjectTypeMax = ObjectTypeCustomLineCap
    ObjectTypeMin = ObjectTypeBrush
End Enum

' The frameRect for creating a metafile can be specified in any of these
' units.  There is an extra frame unit value (MetafileFrameUnitGdi) so
' that units can be supplied in the same units that GDI expects for
' frame rects -- these units are in .01 (1/100ths) millimeter units
' as defined by GDI.
Public Enum MetafileFrameUnit
   MetafileFrameUnitPixel = UnitPixel
   MetafileFrameUnitPoint = UnitPoint
   MetafileFrameUnitInch = UnitInch
   MetafileFrameUnitDocument = UnitDocument
   MetafileFrameUnitMillimeter = UnitMillimeter
   MetafileFrameUnitGdi                        ' GDI compatible .01 MM units
End Enum

' Coordinate space identifiers
Public Enum CoordinateSpace
   CoordinateSpaceWorld     ' 0
   CoordinateSpacePage      ' 1
   CoordinateSpaceDevice     ' 2
End Enum

' Added 12/4/2002
' This enum was translated by: Dana Seaman
Public Enum EmfPlusRecordType
   '//Since we have to enumerate GDI records right along with GDI+ records
   '//We list all the GDI records here so that they can be part of the
   '//same enumeration type which is used in the enumeration callback.
   WmfRecordTypeSetBkColor = &H10201
   WmfRecordTypeSetBkMode = &H10102
   WmfRecordTypeSetMapMode = &H10103
   WmfRecordTypeSetROP2 = &H10104
   WmfRecordTypeSetRelAbs = &H10105
   WmfRecordTypeSetPolyFillMode = &H10106
   WmfRecordTypeSetStretchBltMode = &H10107
   WmfRecordTypeSetTextCharExtra = &H10108
   WmfRecordTypeSetTextColor = &H10209
   WmfRecordTypeSetTextJustification = &H1020A
   WmfRecordTypeSetWindowOrg = &H1020B
   WmfRecordTypeSetWindowExt = &H1020C
   WmfRecordTypeSetViewportOrg = &H1020D
   WmfRecordTypeSetViewportExt = &H1020E
   WmfRecordTypeOffsetWindowOrg = &H1020F
   WmfRecordTypeScaleWindowExt = &H10410
   WmfRecordTypeOffsetViewportOrg = &H10211
   WmfRecordTypeScaleViewportExt = &H10412
   WmfRecordTypeLineTo = &H10213
   WmfRecordTypeMoveTo = &H10214
   WmfRecordTypeExcludeClipRect = &H10415
   WmfRecordTypeIntersectClipRect = &H10416
   WmfRecordTypeArc = &H10817
   WmfRecordTypeEllipse = &H10418
   WmfRecordTypeFloodFill = &H10419
   WmfRecordTypePie = &H1081A
   WmfRecordTypeRectangle = &H1041B
   WmfRecordTypeRoundRect = &H1061C
   WmfRecordTypePatBlt = &H1061D
   WmfRecordTypeSaveDC = &H1001E
   WmfRecordTypeSetPixel = &H1041F
   WmfRecordTypeOffsetClipRgn = &H10220
   WmfRecordTypeTextOut = &H10521
   WmfRecordTypeBitBlt = &H10922
   WmfRecordTypeStretchBlt = &H10B23
   WmfRecordTypePolygon = &H10324
   WmfRecordTypePolyline = &H10325
   WmfRecordTypeEscape = &H10626
   WmfRecordTypeRestoreDC = &H10127
   WmfRecordTypeFillRegion = &H10228
   WmfRecordTypeFrameRegion = &H10429
   WmfRecordTypeInvertRegion = &H1012A
   WmfRecordTypePaintRegion = &H1012B
   WmfRecordTypeSelectClipRegion = &H1012C
   WmfRecordTypeSelectObject = &H1012D
   WmfRecordTypeSetTextAlign = &H1012E
   WmfRecordTypeDrawText = &H1062F
   WmfRecordTypeChord = &H10830
   WmfRecordTypeSetMapperFlags = &H10231
   WmfRecordTypeExtTextOut = &H10A32
   WmfRecordTypeSetDIBToDev = &H10D33
   WmfRecordTypeSelectPalette = &H10234
   WmfRecordTypeRealizePalette = &H10035
   WmfRecordTypeAnimatePalette = &H10436
   WmfRecordTypeSetPalEntries = &H10037
   WmfRecordTypePolyPolygon = &H10538
   WmfRecordTypeResizePalette = &H10139
   WmfRecordTypeDIBBitBlt = &H10940
   WmfRecordTypeDIBStretchBlt = &H10B41
   WmfRecordTypeDIBCreatePatternBrush = &H10142
   WmfRecordTypeStretchDIB = &H10F43
   WmfRecordTypeExtFloodFill = &H10548
   WmfRecordTypeSetLayout = &H10149
   WmfRecordTypeResetDC = &H1014C
   WmfRecordTypeStartDoc = &H1014D
   WmfRecordTypeStartPage = &H1004F
   WmfRecordTypeEndPage = &H10050
   WmfRecordTypeAbortDoc = &H10052
   WmfRecordTypeEndDoc = &H1005E
   WmfRecordTypeDeleteObject = &H101F0
   WmfRecordTypeCreatePalette = &H100F7
   WmfRecordTypeCreateBrush = &H100F8
   WmfRecordTypeCreatePatternBrush = &H101F9
   WmfRecordTypeCreatePenIndirect = &H102FA
   WmfRecordTypeCreateFontIndirect = &H102FB
   WmfRecordTypeCreateBrushIndirect = &H102FC
   WmfRecordTypeCreateBitmapIndirect = &H102FD
   WmfRecordTypeCreateBitmap = &H106FE
   WmfRecordTypeCreateRegion = &H106FF
   EmfRecordTypeHeader = 1
   EmfRecordTypePolyBezier = 2
   EmfRecordTypePolygon = 3
   EmfRecordTypePolyline = 4
   EmfRecordTypePolyBezierTo = 5
   EmfRecordTypePolyLineTo = 6
   EmfRecordTypePolyPolyline = 7
   EmfRecordTypePolyPolygon = 8
   EmfRecordTypeSetWindowExtEx = 9
   EmfRecordTypeSetWindowOrgEx = 10
   EmfRecordTypeSetViewportExtEx = 11
   EmfRecordTypeSetViewportOrgEx = 12
   EmfRecordTypeSetBrushOrgEx = 13
   EmfRecordTypeEOF = 14
   EmfRecordTypeSetPixelV = 15
   EmfRecordTypeSetMapperFlags = 16
   EmfRecordTypeSetMapMode = 17
   EmfRecordTypeSetBkMode = 18
   EmfRecordTypeSetPolyFillMode = 19
   EmfRecordTypeSetROP2 = 20
   EmfRecordTypeSetStretchBltMode = 21
   EmfRecordTypeSetTextAlign = 22
   EmfRecordTypeSetColorAdjustment = 23
   EmfRecordTypeSetTextColor = 24
   EmfRecordTypeSetBkColor = 25
   EmfRecordTypeOffsetClipRgn = 26
   EmfRecordTypeMoveToEx = 27
   EmfRecordTypeSetMetaRgn = 28
   EmfRecordTypeExcludeClipRect = 29
   EmfRecordTypeIntersectClipRect = 30
   EmfRecordTypeScaleViewportExtEx = 31
   EmfRecordTypeScaleWindowExtEx = 32
   EmfRecordTypeSaveDC = 33
   EmfRecordTypeRestoreDC = 34
   EmfRecordTypeSetWorldTransform = 35
   EmfRecordTypeModifyWorldTransform = 36
   EmfRecordTypeSelectObject = 37
   EmfRecordTypeCreatePen = 38
   EmfRecordTypeCreateBrushIndirect = 39
   EmfRecordTypeDeleteObject = 40
   EmfRecordTypeAngleArc = 41
   EmfRecordTypeEllipse = 42
   EmfRecordTypeRectangle = 43
   EmfRecordTypeRoundRect = 44
   EmfRecordTypeArc = 45
   EmfRecordTypeChord = 46
   EmfRecordTypePie = 47
   EmfRecordTypeSelectPalette = 48
   EmfRecordTypeCreatePalette = 49
   EmfRecordTypeSetPaletteEntries = 50
   EmfRecordTypeResizePalette = 51
   EmfRecordTypeRealizePalette = 52
   EmfRecordTypeExtFloodFill = 53
   EmfRecordTypeLineTo = 54
   EmfRecordTypeArcTo = 55
   EmfRecordTypePolyDraw = 56
   EmfRecordTypeSetArcDirection = 57
   EmfRecordTypeSetMiterLimit = 58
   EmfRecordTypeBeginPath = 59
   EmfRecordTypeEndPath = 60
   EmfRecordTypeCloseFigure = 61
   EmfRecordTypeFillPath = 62
   EmfRecordTypeStrokeAndFillPath = 63
   EmfRecordTypeStrokePath = 64
   EmfRecordTypeFlattenPath = 65
   EmfRecordTypeWidenPath = 66
   EmfRecordTypeSelectClipPath = 67
   EmfRecordTypeAbortPath = 68
   EmfRecordTypeReserved_069 = 69
   EmfRecordTypeGdiComment = 70
   EmfRecordTypeFillRgn = 71
   EmfRecordTypeFrameRgn = 72
   EmfRecordTypeInvertRgn = 73
   EmfRecordTypePaintRgn = 74
   EmfRecordTypeExtSelectClipRgn = 75
   EmfRecordTypeBitBlt = 76
   EmfRecordTypeStretchBlt = 77
   EmfRecordTypeMaskBlt = 78
   EmfRecordTypePlgBlt = 79
   EmfRecordTypeSetDIBitsToDevice = 80
   EmfRecordTypeStretchDIBits = 81
   EmfRecordTypeExtCreateFontIndirect = 82
   EmfRecordTypeExtTextOutA = 83
   EmfRecordTypeExtTextOutW = 84
   EmfRecordTypePolyBezier16 = 85
   EmfRecordTypePolygon16 = 86
   EmfRecordTypePolyline16 = 87
   EmfRecordTypePolyBezierTo16 = 88
   EmfRecordTypePolylineTo16 = 89
   EmfRecordTypePolyPolyline16 = 90
   EmfRecordTypePolyPolygon16 = 91
   EmfRecordTypePolyDraw16 = 92
   EmfRecordTypeCreateMonoBrush = 93
   EmfRecordTypeCreateDIBPatternBrushPt = 94
   EmfRecordTypeExtCreatePen = 95
   EmfRecordTypePolyTextOutA = 96
   EmfRecordTypePolyTextOutW = 97
   EmfRecordTypeSetICMMode = 98
   EmfRecordTypeCreateColorSpace = 99
   EmfRecordTypeSetColorSpace = 100
   EmfRecordTypeDeleteColorSpace = 101
   EmfRecordTypeGLSRecord = 102
   EmfRecordTypeGLSBoundedRecord = 103
   EmfRecordTypePixelFormat = 104
   EmfRecordTypeDrawEscape = 105
   EmfRecordTypeExtEscape = 106
   EmfRecordTypeStartDoc = 107
   EmfRecordTypeSmallTextOut = 108
   EmfRecordTypeForceUFIMapping = 109
   EmfRecordTypeNamedEscape = 110
   EmfRecordTypeColorCorrectPalette = 111
   EmfRecordTypeSetICMProfileA = 112
   EmfRecordTypeSetICMProfileW = 113
   EmfRecordTypeAlphaBlend = 114
   EmfRecordTypeSetLayout = 115
   EmfRecordTypeTransparentBlt = 116
   EmfRecordTypeReserved_117 = 117
   EmfRecordTypeGradientFill = 118
   EmfRecordTypeSetLinkedUFIs = 119
   EmfRecordTypeSetTextJustification = 120
   EmfRecordTypeColorMatchToTargetW = 121
   EmfRecordTypeCreateColorSpaceW = 122
   EmfRecordTypeMax = 122
   EmfRecordTypeMin = 1
   '//That is the END of the GDI EMF records.
   '//Now we start the list of EMF+ records.  We leave quite
   '//a bit of room here for the addition of any new GDI
   '//records that may be added later.
   EmfPlusRecordTypeInvalid = 16384 '//GDIP_EMFPLUS_RECORD_BASE
   EmfPlusRecordTypeHeader = 16385
   EmfPlusRecordTypeEndOfFile = 16386
   EmfPlusRecordTypeComment = 16387
   EmfPlusRecordTypeGetDC = 16388
   EmfPlusRecordTypeMultiFormatStart = 16389
   EmfPlusRecordTypeMultiFormatSection = 16390
   EmfPlusRecordTypeMultiFormatEnd = 16391
   '//For all persistent objects
   EmfPlusRecordTypeObject = 16392
   '//Drawing Records
   EmfPlusRecordTypeClear = 16393
   EmfPlusRecordTypeFillRects = 16394
   EmfPlusRecordTypeDrawRects = 16395
   EmfPlusRecordTypeFillPolygon = 16396
   EmfPlusRecordTypeDrawLines = 16397
   EmfPlusRecordTypeFillEllipse = 16398
   EmfPlusRecordTypeDrawEllipse = 16399
   EmfPlusRecordTypeFillPie = 16400
   EmfPlusRecordTypeDrawPie = 16401
   EmfPlusRecordTypeDrawArc = 16402
   EmfPlusRecordTypeFillRegion = 16403
   EmfPlusRecordTypeFillPath = 16404
   EmfPlusRecordTypeDrawPath = 16405
   EmfPlusRecordTypeFillClosedCurve = 16406
   EmfPlusRecordTypeDrawClosedCurve = 16407
   EmfPlusRecordTypeDrawCurve = 16408
   EmfPlusRecordTypeDrawBeziers = 16409
   EmfPlusRecordTypeDrawImage = 16410
   EmfPlusRecordTypeDrawImagePoints = 16411
   EmfPlusRecordTypeDrawString = 16412
   '//Graphics State Records
   EmfPlusRecordTypeSetRenderingOrigin = 16413
   EmfPlusRecordTypeSetAntiAliasMode = 16414
   EmfPlusRecordTypeSetTextRenderingHint = 16415
   EmfPlusRecordTypeSetTextContrast = 16416
   EmfPlusRecordTypeSetInterpolationMode = 16417
   EmfPlusRecordTypeSetPixelOffsetMode = 16418
   EmfPlusRecordTypeSetCompositingMode = 16419
   EmfPlusRecordTypeSetCompositingQuality = 16420
   EmfPlusRecordTypeSave = 16421
   EmfPlusRecordTypeRestore = 16422
   EmfPlusRecordTypeBeginContainer = 16423
   EmfPlusRecordTypeBeginContainerNoParams = 16424
   EmfPlusRecordTypeEndContainer = 16425
   EmfPlusRecordTypeSetWorldTransform = 16426
   EmfPlusRecordTypeResetWorldTransform = 16427
   EmfPlusRecordTypeMultiplyWorldTransform = 16428
   EmfPlusRecordTypeTranslateWorldTransform = 16429
   EmfPlusRecordTypeScaleWorldTransform = 16430
   EmfPlusRecordTypeRotateWorldTransform = 16431
   EmfPlusRecordTypeSetPageTransform = 16432
   EmfPlusRecordTypeResetClip = 16433
   EmfPlusRecordTypeSetClipRect = 16434
   EmfPlusRecordTypeSetClipPath = 16435
   EmfPlusRecordTypeSetClipRegion = 16436
   EmfPlusRecordTypeOffsetClip = 16437
   EmfPlusRecordTypeDrawDriverString = 16438
   EmfPlusRecordTotal = 16439
   EmfPlusRecordTypeMax = 16438
   EmfPlusRecordTypeMin = 16385
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

' Access modes used when calling Image::LockBits (GdipBitmapLockBits API)
Public Enum ImageLockMode
   ImageLockModeRead = &H1
   ImageLockModeWrite = &H2
   ImageLockModeUserInputBuf = &H4
End Enum

Public Enum DebugEventLevel
    DebugEventLevelFatal
    DebugEventLevelWarning
End Enum




'-----------------------------------------------
' APIs
'-----------------------------------------------

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

' Graphics Functions (ALL)
Public Declare Function GdipFlush Lib "gdiplus" (ByVal graphics As Long, ByVal intention As FlushIntention) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal graphics As Long, ByVal CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal graphics As Long, CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal graphics As Long, ByVal x As Long, ByVal y As Long) As GpStatus
Public Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal graphics As Long, x As Long, y As Long) As GpStatus
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal graphics As Long, ByVal CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal graphics As Long, CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal graphics As Long, ByVal mode As TextRenderingHint) As GpStatus
Public Declare Function GdipGetTextRenderingHint Lib "gdiplus" (ByVal graphics As Long, mode As TextRenderingHint) As GpStatus
Public Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal graphics As Long, ByVal contrast As Long) As GpStatus
Public Declare Function GdipGetTextContrast Lib "gdiplus" (ByVal graphics As Long, contrast As Long) As GpStatus
Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipMultiplyWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipGetWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPageTransform Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetPageUnit Lib "gdiplus" (ByVal graphics As Long, unit As GpUnit) As GpStatus
Public Declare Function GdipGetPageScale Lib "gdiplus" (ByVal graphics As Long, sscale As Single) As GpStatus
Public Declare Function GdipSetPageUnit Lib "gdiplus" (ByVal graphics As Long, ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipSetPageScale Lib "gdiplus" (ByVal graphics As Long, ByVal sscale As Single) As GpStatus
Public Declare Function GdipGetDpiX Lib "gdiplus" (ByVal graphics As Long, dpi As Single) As GpStatus
Public Declare Function GdipGetDpiY Lib "gdiplus" (ByVal graphics As Long, dpi As Single) As GpStatus
Public Declare Function GdipTransformPoints Lib "gdiplus" (ByVal graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipTransformPointsI Lib "gdiplus" (ByVal graphics As Long, ByVal destSpace As CoordinateSpace, ByVal srcSpace As CoordinateSpace, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipGetNearestColor Lib "gdiplus" (ByVal graphics As Long, argb As Long) As GpStatus
' Creates the Win9x Halftone Palette (even on NT) with correct Desktop colors
Public Declare Function GdipCreateHalftonePalette Lib "gdiplus" () As Long
Public Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GpStatus
Public Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GpStatus
Public Declare Function GdipDrawLines Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawArc Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawBezier Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GpStatus
Public Declare Function GdipDrawBezierI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As GpStatus
Public Declare Function GdipDrawBeziers Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangles Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, rects As RECTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, rects As RECTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawPie Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPieI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipDrawCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3 Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTL, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As Long, ByVal lColor As Long) As GpStatus
Public Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangles Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, rects As RECTF, ByVal count As Long) As GpStatus
Public Declare Function GdipFillRectanglesI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, rects As RECTL, ByVal count As Long) As GpStatus
Public Declare Function GdipFillPolygon Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTF, ByVal count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTL, ByVal count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2 Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPieI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipFillClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTL, ByVal count As Long, ByVal tension As Single, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillRegion Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal region As Long) As GpStatus
Public Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single) As GpStatus
Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePoints Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, dstpoints As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, dstpoints As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, _
                     ByVal y As Single, ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, _
                     ByVal y As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal dstx As Single, _
                     ByVal dsty As Single, ByVal dstwidth As Long, ByVal dstheight As Single, _
                     ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, _
                     ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, _
                     Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
' Callback declaration: Public Function DrawImageAbort(ByVal lpData as Long) as Long
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal dstx As Long, _
                     ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                     ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                     ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, _
                     Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, _
                     Points As POINTF, ByVal count As Long, _
                     ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, _
                     ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, _
                     Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, _
                     Points As POINTL, ByVal count As Long, _
                     ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                     ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, _
                     Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoint Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoint As POINTF, lpEnumerateMetafileProc As Long, ByVal callbackData As Long, imageAttributes As Long) As GpStatus
' Callback declaration: Public Function EnumMetafilesProc(Byval rtype as EmfPlusRecordType, byval _ as Long, byval _ as Long, bytes as Any, byval callbackData as long) as long
Public Declare Function GdipEnumerateMetafileDestPointI Lib "gdiplus" (graphics As Long, ByVal metafile As Long, destPoint As POINTL, ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRect Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destRect As RECTF, lpEnumerateMetafileProc As Long, ByVal callbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRectI Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destRect As RECTL, lpEnumerateMetafileProc As Long, ByVal callbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoints Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoint As POINTF, ByVal count As Long, lpEnumerateMetafileProc As Long, ByVal callbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointsI Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoint As POINTL, ByVal count As Long, lpEnumerateMetafileProc As Long, ByVal callbackData As Long, imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoint Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoint As POINTF, srcRect As RECTF, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointI Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoint As POINTL, srcRect As RECTL, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRect Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destRect As RECTF, srcRect As RECTF, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRectI Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destRect As RECTL, srcRect As RECTL, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoints As POINTF, ByVal count As Long, srcRect As RECTF, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI Lib "gdiplus" (ByVal graphics As Long, ByVal metafile As Long, destPoints As POINTL, ByVal count As Long, srcRect As RECTL, ByVal srcUnit As GpUnit, _
                     ByVal lpEnumerateMetafileProc As Long, ByVal callbackData As Long, ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipPlayMetafileRecord Lib "gdiplus" (ByVal metafile As Long, ByVal recordType As EmfPlusRecordType, ByVal flags As Long, ByVal dataSize As Long, byteData As Any) As GpStatus
Public Declare Function GdipSetClipGraphics Lib "gdiplus" (ByVal graphics As Long, ByVal srcgraphics As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRect Lib "gdiplus" (ByVal graphics As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal graphics As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipPath Lib "gdiplus" (ByVal graphics As Long, ByVal path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal graphics As Long, ByVal region As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipHrgn Lib "gdiplus" (ByVal graphics As Long, ByVal hRgn As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipTranslateClip Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateClipI Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Long, ByVal dy As Long) As GpStatus
Public Declare Function GdipGetClip Lib "gdiplus" (ByVal graphics As Long, ByVal region As Long) As GpStatus
Public Declare Function GdipGetClipBounds Lib "gdiplus" (ByVal graphics As Long, rect As RECTF) As GpStatus
Public Declare Function GdipGetClipBoundsI Lib "gdiplus" (ByVal graphics As Long, rect As RECTL) As GpStatus
Public Declare Function GdipIsClipEmpty Lib "gdiplus" (ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipGetVisibleClipBounds Lib "gdiplus" (ByVal graphics As Long, rect As RECTF) As GpStatus
Public Declare Function GdipGetVisibleClipBoundsI Lib "gdiplus" (ByVal graphics As Long, rect As RECTL) As GpStatus
Public Declare Function GdipIsVisibleClipEmpty Lib "gdiplus" (ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisiblePoint Lib "gdiplus" (ByVal graphics As Long, ByVal x As Single, ByVal y As Single, result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI Lib "gdiplus" (ByVal graphics As Long, ByVal x As Long, ByVal y As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisibleRect Lib "gdiplus" (ByVal graphics As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, result As Long) As GpStatus
Public Declare Function GdipIsVisibleRectI Lib "gdiplus" (ByVal graphics As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, result As Long) As GpStatus
Public Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal graphics As Long, state As Long) As GpStatus
Public Declare Function GdipRestoreGraphics Lib "gdiplus" (ByVal graphics As Long, ByVal state As Long) As GpStatus
Public Declare Function GdipBeginContainer Lib "gdiplus" (ByVal graphics As Long, dstrect As RECTF, srcRect As RECTF, ByVal unit As GpUnit, state As Long) As GpStatus
Public Declare Function GdipBeginContainerI Lib "gdiplus" (ByVal graphics As Long, dstrect As RECTL, srcRect As RECTL, ByVal unit As GpUnit, state As Long) As GpStatus
Public Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal graphics As Long, state As Long) As GpStatus
Public Declare Function GdipEndContainer Lib "gdiplus" (ByVal graphics As Long, ByVal state As Long) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromWmf Lib "gdiplus" (ByVal hWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromEmf Lib "gdiplus" (ByVal hEmf As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromFile Lib "gdiplus" (ByVal filename As String, header As MetafileHeader) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipGetMetafileHeaderFromStream Lib "gdiplus" (Byval stream as IStream, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal metafile As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As Long, hEmf As Long) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
' NOTE: The C++ stream parameter was declared as IStream** stream
'Public Declare Function GdipCreateStreamOnFile Lib "gdiplus" (ByVal filename As String, ByVal access As Long, stream As IStream) As GpStatus
Public Declare Function GdipCreateMetafileFromWmf Lib "gdiplus" (ByVal hWmf As Long, ByVal bDeleteWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, ByVal metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromEmf Lib "gdiplus" (ByVal hEmf As Long, ByVal bDeleteEmf As Long, metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromFile Lib "gdiplus" (byvalfile As String, metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromWmfFile Lib "gdiplus" (ByVal file As String, WmfPlaceableFileHdr As WmfPlaceableFileHeader, metafile As Long) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipCreateMetafileFromStream Lib "gdiplus" (Byval stream as IStream, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafile Lib "gdiplus" (ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileName Lib "gdiplus" (ByVal filename As String, ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileNameI Lib "gdiplus" (ByVal filename As String, ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipRecordMetafileStream Lib "gdiplus" (Byval stream as IStream, ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTF, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
'Public Declare Function GdipRecordMetafileStreamI Lib "gdiplus" (Byval stream as IStream, ByVal referenceHdc As Long, etype As EmfType, frameRect As RECTL, ByVal frameUnit As MetafileFrameUnit, ByVal description As String, metafile As Long) As GpStatus
Public Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal metafile As Long, ByVal metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal metafile As Long, metafileRasterizationLimitDpi As Long) As GpStatus
' NOTE: These encoders/decoders functions expect an ImageCodecInfo array
Public Declare Function GdipGetImageDecodersSize Lib "gdiplus" (numDecoders As Long, size As Long) As GpStatus
Public Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numDecoders As Long, ByVal size As Long, decoders As Any) As GpStatus
Public Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, size As Long) As GpStatus
Public Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal size As Long, encoders As Any) As GpStatus
Public Declare Function GdipComment Lib "gdiplus" (ByVal graphics As Long, ByVal sizeData As Long, data As Any) As GpStatus

' Image Functions (ALL)
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, image As Long) As GpStatus
Public Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal filename As String, image As Long) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (Byval stream as IStream, image As Long) As GpStatus
'Public Declare Function GdipLoadImageFromStreamICM Lib "gdiplus" (Byval stream as IStream, image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
' NOTE: The encoderParams parameter expects a EncoderParameters struct or a NULL
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As String, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal image As Long, ByVal stream As IStream, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
Public Declare Function GdipSaveAdd Lib "gdiplus" (ByVal image As Long, encoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipSaveAddImage Lib "gdiplus" (ByVal image As Long, ByVal newImage As Long, encoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As GpStatus
Public Declare Function GdipGetImageBounds Lib "gdiplus" (ByVal image As Long, srcRect As RECTF, srcUnit As GpUnit) As GpStatus
Public Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal image As Long, Width As Single, Height As Single) As GpStatus
Public Declare Function GdipGetImageType Lib "gdiplus" (ByVal image As Long, itype As ImageType) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal image As Long, resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal image As Long, flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal image As Long, format As CLSID) As GpStatus
Public Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal image As Long, PixelFormat As Long) As GpStatus
Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, _
                        Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipGetEncoderParameterListSize Lib "gdiplus" (ByVal image As Long, clsidEncoder As CLSID, size As Long) As GpStatus
Public Declare Function GdipGetEncoderParameterList Lib "gdiplus" (ByVal image As Long, clsidEncoder As CLSID, ByVal size As Long, buffer As EncoderParameters) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsCount Lib "gdiplus" (ByVal image As Long, count As Long) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsList Lib "gdiplus" (ByVal image As Long, dimensionIDs As CLSID, ByVal count As Long) As GpStatus
Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal image As Long, dimensionID As CLSID, count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal image As Long, dimensionID As CLSID, ByVal frameIndex As Long) As GpStatus
Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal image As Long, ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipGetImagePalette Lib "gdiplus" (ByVal image As Long, palette As ColorPalette, ByVal size As Long) As GpStatus
Public Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal image As Long, palette As ColorPalette) As GpStatus
Public Declare Function GdipGetImagePaletteSize Lib "gdiplus" (ByVal image As Long, size As Long) As GpStatus
Public Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal image As Long, numOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList Lib "gdiplus" (ByVal image As Long, ByVal numOfProperty As Long, list As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal image As Long, ByVal propId As Long, size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal image As Long, ByVal propId As Long, ByVal propSize As Long, buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal image As Long, totalBufferSize As Long, numProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal image As Long, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal image As Long, ByVal propId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal image As Long, item As PropertyItem) As GpStatus
Public Declare Function GdipImageForceValidation Lib "gdiplus" (ByVal image As Long) As GpStatus

' Pen Functions (ALL)
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal brush As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipClonePen Lib "gdiplus" (ByVal pen As Long, clonepen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal pen As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal pen As Long, Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit Lib "gdiplus" (ByVal pen As Long, ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit Lib "gdiplus" (ByVal pen As Long, unit As GpUnit) As GpStatus
Public Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal pen As Long, ByVal startCap As LineCap, ByVal endCap As LineCap, ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal pen As Long, ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal pen As Long, startCap As LineCap) As GpStatus
Public Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal pen As Long, endCap As LineCap) As GpStatus
Public Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal pen As Long, dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal LnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal pen As Long, LnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomStartCap Lib "gdiplus" (ByVal pen As Long, customCap As Long) As GpStatus
Public Declare Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal pen As Long, ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomEndCap Lib "gdiplus" (ByVal pen As Long, customCap As Long) As GpStatus
Public Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, ByVal miterLimit As Single) As GpStatus
Public Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal pen As Long, miterLimit As Single) As GpStatus
Public Declare Function GdipSetPenMode Lib "gdiplus" (ByVal pen As Long, ByVal penMode As PenAlignment) As GpStatus
Public Declare Function GdipGetPenMode Lib "gdiplus" (ByVal pen As Long, penMode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPenTransform Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipMultiplyPenTransform Lib "gdiplus" (ByVal pen As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePenTransform Lib "gdiplus" (ByVal pen As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetPenColor Lib "gdiplus" (ByVal pen As Long, ByVal argb As Long) As GpStatus
Public Declare Function GdipGetPenColor Lib "gdiplus" (ByVal pen As Long, argb As Long) As GpStatus
Public Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal brush As Long) As GpStatus
Public Declare Function GdipGetPenBrushFill Lib "gdiplus" (ByVal pen As Long, brush As Long) As GpStatus
Public Declare Function GdipGetPenFillType Lib "gdiplus" (ByVal pen As Long, ptype As PenType) As GpStatus
Public Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal pen As Long, dStyle As DashStyle) As GpStatus
Public Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As DashStyle) As GpStatus
Public Declare Function GdipGetPenDashOffset Lib "gdiplus" (ByVal pen As Long, Offset As Single) As GpStatus
Public Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal pen As Long, ByVal Offset As Single) As GpStatus
Public Declare Function GdipGetPenDashCount Lib "gdiplus" (ByVal pen As Long, count As Long) As GpStatus
Public Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPenDashArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundCount Lib "gdiplus" (ByVal pen As Long, count As Long) As GpStatus
Public Declare Function GdipSetPenCompoundArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundArray Lib "gdiplus" (ByVal pen As Long, dash As Single, ByVal count As Long) As GpStatus

' CustomLineCap Functions (ALL)
Public Declare Function GdipCreateCustomLineCap Lib "gdiplus" (ByVal fillPath As Long, ByVal strokePath As Long, ByVal baseCap As LineCap, ByVal baseInset As Single, customCap As Long) As GpStatus
Public Declare Function GdipDeleteCustomLineCap Lib "gdiplus" (ByVal customCap As Long) As GpStatus
Public Declare Function GdipCloneCustomLineCap Lib "gdiplus" (ByVal customCap As Long, clonedCap As Long) As GpStatus
Public Declare Function GdipGetCustomLineCapType Lib "gdiplus" (ByVal customCap As Long, capType As CustomLineCapType) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal customCap As Long, ByVal startCap As LineCap, ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal customCap As Long, startCap As LineCap, endCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal customCap As Long, ByVal LnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal customCap As Long, LnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseCap Lib "gdiplus" (ByVal customCap As Long, ByVal baseCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseCap Lib "gdiplus" (ByVal customCap As Long, baseCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseInset Lib "gdiplus" (ByVal customCap As Long, ByVal inset As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseInset Lib "gdiplus" (ByVal customCap As Long, inset As Single) As GpStatus
Public Declare Function GdipSetCustomLineCapWidthScale Lib "gdiplus" (ByVal customCap As Long, ByVal widthScale As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapWidthScale Lib "gdiplus" (ByVal customCap As Long, widthScale As Single) As GpStatus

' AdjustableArrowCap Functions (ALL)
Public Declare Function GdipCreateAdjustableArrowCap Lib "gdiplus" (ByVal Height As Single, ByVal Width As Single, ByVal isFilled As Long, cap As Long) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapHeight Lib "gdiplus" (ByVal cap As Long, ByVal Height As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapHeight Lib "gdiplus" (ByVal cap As Long, Height As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapWidth Lib "gdiplus" (ByVal cap As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapWidth Lib "gdiplus" (ByVal cap As Long, Width As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal cap As Long, ByVal middleInset As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal cap As Long, middleInset As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapFillState Lib "gdiplus" (ByVal cap As Long, ByVal bFillState As Long) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapFillState Lib "gdiplus" (ByVal cap As Long, bFillState As Long) As GpStatus

' Bitmap Functions (ALL)
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFileICM Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As GpStatus
' TODO: Uncomment if you have the IStream object declared, or equivalent
'Public Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (Byval stream as IStream, bitmap As Long) As GpStatus
'Public Declare Function GdipCreateBitmapFromStreamICM Lib "gdiplus" (Byval stream as IStream, bitmap As Long) As GpStatus
' NOTE: The scan0 parameter is treated as a byte array
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal graphics As Long, bitmap As Long) As GpStatus
' TODO: Uncomment if DirectX is in your program
'Public Declare Function GdipCreateBitmapFromDirectDrawSurface Lib "gdiplus" (surface As DirectDrawSurface7, bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, ByVal gdiBitmapData As Long, bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hicon As Long, bitmap As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource Lib "gdiplus" (ByVal hInstance As Long, ByVal lpBitmapName As String, bitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapArea Lib "gdiplus" (ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal PixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus
Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, rect As RECTL, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As GpStatus
Public Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As GpStatus

' CachedBitmap Functions (ALL)
Public Declare Function GdipCreateCachedBitmap Lib "gdiplus" (ByVal bitmap As Long, ByVal graphics As Long, cachedBitmap As Long) As GpStatus
Public Declare Function GdipDeleteCachedBitmap Lib "gdiplus" (ByVal cachedBitmap As Long) As GpStatus
Public Declare Function GdipDrawCachedBitmap Lib "gdiplus" (ByVal graphics As Long, ByVal cachedBitmap As Long, ByVal x As Long, ByVal y As Long) As GpStatus

' Brush Functions (ALL)
Public Declare Function GdipCloneBrush Lib "gdiplus" (ByVal brush As Long, cloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipGetBrushType Lib "gdiplus" (ByVal brush As Long, brshType As BrushType) As GpStatus

' HatchBrush Functions (ALL)
Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As HatchStyle, ByVal forecolr As Long, ByVal backcolr As Long, brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle Lib "gdiplus" (ByVal brush As Long, style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor Lib "gdiplus" (ByVal brush As Long, forecolr As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor Lib "gdiplus" (ByVal brush As Long, backcolr As Long) As GpStatus

' SolidBrush Functions (ALL)
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal brush As Long, ByVal argb As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal brush As Long, argb As Long) As GpStatus

' LineBrush Functions (ALL)
Public Declare Function GdipCreateLineBrush Lib "gdiplus" (point1 As POINTF, point2 As POINTF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (point1 As POINTL, point2 As POINTL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (rect As RECTL, ByVal color1 As Long, ByVal color2 As Long, ByVal angle As Single, ByVal isAngleScalable As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipSetLineColors Lib "gdiplus" (ByVal brush As Long, ByVal color1 As Long, ByVal color2 As Long) As GpStatus
Public Declare Function GdipGetLineColors Lib "gdiplus" (ByVal brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipGetLineRect Lib "gdiplus" (ByVal brush As Long, rect As RECTF) As GpStatus
Public Declare Function GdipGetLineRectI Lib "gdiplus" (ByVal brush As Long, rect As RECTL) As GpStatus
Public Declare Function GdipSetLineGammaCorrection Lib "gdiplus" (ByVal brush As Long, ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineGammaCorrection Lib "gdiplus" (ByVal brush As Long, useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineBlendCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipGetLineBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetLineBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlendCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetLineSigmaBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineLinearBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineTransform Lib "gdiplus" (ByVal brush As Long, matrix As Long) As GpStatus
Public Declare Function GdipSetLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetLineTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateLineTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus

' TextureBrush Functions (ALL)
Public Declare Function GdipCreateTexture Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2 Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIA Lib "gdiplus" (ByVal image As Long, ByVal imageAttributes As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2I Lib "gdiplus" (ByVal image As Long, ByVal WrapMd As WrapMode, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIAI Lib "gdiplus" (ByVal image As Long, ByVal imageAttributes As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, texture As Long) As GpStatus
Public Declare Function GdipGetTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal brush As Long, image As Long) As GpStatus

' PathGradientBrush Functions (ALL)
Public Declare Function GdipCreatePathGradient Lib "gdiplus" (Points As POINTF, ByVal count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As POINTL, ByVal count As Long, ByVal WrapMd As WrapMode, polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal path As Long, polyGradient As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterColor Lib "gdiplus" (ByVal brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipSetPathGradientCenterColor Lib "gdiplus" (ByVal brush As Long, ByVal lColors As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal brush As Long, argb As Long, count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal brush As Long, argb As Long, count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPath Lib "gdiplus" (ByVal brush As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipSetPathGradientPath Lib "gdiplus" (ByVal brush As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterPoint Lib "gdiplus" (ByVal brush As Long, Points As POINTF) As GpStatus
Public Declare Function GdipGetPathGradientCenterPointI Lib "gdiplus" (ByVal brush As Long, Points As POINTL) As GpStatus
Public Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal brush As Long, Points As POINTF) As GpStatus
Public Declare Function GdipSetPathGradientCenterPointI Lib "gdiplus" (ByVal brush As Long, Points As POINTL) As GpStatus
Public Declare Function GdipGetPathGradientRect Lib "gdiplus" (ByVal brush As Long, rect As RECTF) As GpStatus
Public Declare Function GdipGetPathGradientRectI Lib "gdiplus" (ByVal brush As Long, rect As RECTL) As GpStatus
Public Declare Function GdipGetPathGradientPointCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipSetPathGradientGammaCorrection Lib "gdiplus" (ByVal brush As Long, ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientGammaCorrection Lib "gdiplus" (ByVal brush As Long, useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlendCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend Lib "gdiplus" (ByVal brush As Long, blend As Single, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlendCount Lib "gdiplus" (ByVal brush As Long, count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal brush As Long, blend As Long, positions As Single, ByVal count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSigmaBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal sscale As Single) As GpStatus
Public Declare Function GdipSetPathGradientLinearBlend Lib "gdiplus" (ByVal brush As Long, ByVal focus As Single, ByVal sscale As Single) As GpStatus
Public Declare Function GdipGetPathGradientWrapMode Lib "gdiplus" (ByVal brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPathGradientTransform Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyPathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal matrix As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePathGradientTransform Lib "gdiplus" (ByVal brush As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipGetPathGradientFocusScales Lib "gdiplus" (ByVal brush As Long, xScale As Single, yScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientFocusScales Lib "gdiplus" (ByVal brush As Long, ByVal xScale As Single, ByVal yScale As Single) As GpStatus

' GraphicsPath Functions (ALL)
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As FillMode, path As Long) As GpStatus
' NOTE: The types parameter is treated as a byte array
Public Declare Function GdipCreatePath2 Lib "gdiplus" (Points As POINTF, types As Any, ByVal count As Long, brushmode As FillMode, path As Long) As GpStatus
Public Declare Function GdipCreatePath2I Lib "gdiplus" (Points As POINTL, types As Any, ByVal count As Long, brushmode As FillMode, path As Long) As GpStatus
Public Declare Function GdipClonePath Lib "gdiplus" (ByVal path As Long, clonePath As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipResetPath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipGetPointCount Lib "gdiplus" (ByVal path As Long, count As Long) As GpStatus
' NOTE: The types parameter is treated as a byte array
Public Declare Function GdipGetPathTypes Lib "gdiplus" (ByVal path As Long, types As Any, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPathPoints Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal path As Long, ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal path As Long, ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipGetPathData Lib "gdiplus" (ByVal path As Long, pdata As PathData) As GpStatus
Public Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClosePathFigures Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipSetPathMarker Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClearPathMarkers Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipReversePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipGetPathLastPoint Lib "gdiplus" (ByVal path As Long, lastPoint As POINTF) As GpStatus
Public Declare Function GdipAddPathLine Lib "gdiplus" (ByVal path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GpStatus
Public Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathArc Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GpStatus
Public Declare Function GdipAddPathBeziers Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathRectangles Lib "gdiplus" (ByVal path As Long, rect As RECTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathPie Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal path As Long, Points As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathPath Lib "gdiplus" (ByVal path As Long, ByVal addingPath As Long, ByVal bConnect As Long) As GpStatus
Public Declare Function GdipAddPathString Lib "gdiplus" (ByVal path As Long, ByVal str As String, ByVal length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTF, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal path As Long, ByVal str As String, ByVal length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTL, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal path As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GpStatus
Public Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezierI Lib "gdiplus" (ByVal path As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As GpStatus
Public Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long, ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long, ByVal Offset As Long, ByVal numberOfSegments As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long, ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathRectanglesI Lib "gdiplus" (ByVal path As Long, rects As RECTL, ByVal count As Long) As GpStatus
Public Declare Function GdipAddPathEllipseI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathPieI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal path As Long, Points As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipFlattenPath Lib "gdiplus" (ByVal path As Long, Optional ByVal matrix As Long = 0, Optional ByVal flatness As Single = FlatnessDefault) As GpStatus
Public Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal path As Long, ByVal matrix As Long, ByVal flatness As Single) As GpStatus
Public Declare Function GdipWidenPath Lib "gdiplus" (ByVal nativePath As Long, ByVal pen As Long, ByVal matrix As Long, ByVal flatness As Single) As GpStatus
Public Declare Function GdipWarpPath Lib "gdiplus" (ByVal path As Long, ByVal matrix As Long, Points As POINTF, ByVal count As Long, ByVal srcx As Single, ByVal srcy As Single, ByVal srcwidth As Single, ByVal srcheight As Single, ByVal WarpMd As WarpMode, ByVal flatness As Single) As GpStatus
Public Declare Function GdipTransformPath Lib "gdiplus" (ByVal path As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal path As Long, bounds As RECTF, ByVal matrix As Long, ByVal pen As Long) As GpStatus
Public Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal path As Long, bounds As RECTL, ByVal matrix As Long, ByVal pen As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal pen As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal path As Long, ByVal x As Long, ByVal y As Long, ByVal pen As Long, ByVal graphics As Long, result As Long) As GpStatus

' PathIterator Functions (ALL)
Public Declare Function GdipCreatePathIter Lib "gdiplus" (iterator As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipDeletePathIter Lib "gdiplus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, startIndex As Long, endIndex As Long, isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpathPath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, ByVal path As Long, isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextPathType Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, pathType As Any, startIndex As Long, endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarker Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, startIndex As Long, endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarkerPath Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, ByVal path As Long) As GpStatus
Public Declare Function GdipPathIterGetCount Lib "gdiplus" (ByVal iterator As Long, count As Long) As GpStatus
Public Declare Function GdipPathIterGetSubpathCount Lib "gdiplus" (ByVal iterator As Long, count As Long) As GpStatus
Public Declare Function GdipPathIterIsValid Lib "gdiplus" (ByVal iterator As Long, valid As Long) As GpStatus
Public Declare Function GdipPathIterHasCurve Lib "gdiplus" (ByVal iterator As Long, hasCurve As Long) As GpStatus
Public Declare Function GdipPathIterRewind Lib "gdiplus" (ByVal iterator As Long) As GpStatus
' NOTE: The types parameter is treated as a byte array
Public Declare Function GdipPathIterEnumerate Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, Points As POINTF, types As Any, ByVal count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData Lib "gdiplus" (ByVal iterator As Long, resultCount As Long, Points As POINTF, types As Any, ByVal startIndex As Long, ByVal endIndex As Long) As GpStatus

' Matrix Functions (ALL)
Public Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single, matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3 Lib "gdiplus" (rect As RECTF, dstplg As POINTF, matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3I Lib "gdiplus" (rect As RECTL, dstplg As POINTL, matrix As Long) As GpStatus
Public Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal matrix As Long, cloneMatrix As Long) As GpStatus
Public Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements Lib "gdiplus" (ByVal matrix As Long, ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipMultiplyMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal matrix2 As Long, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal angle As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipShearMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI Lib "gdiplus" (ByVal matrix As Long, pts As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints Lib "gdiplus" (ByVal matrix As Long, pts As POINTF, ByVal count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI Lib "gdiplus" (ByVal matrix As Long, pts As POINTL, ByVal count As Long) As GpStatus
Public Declare Function GdipGetMatrixElements Lib "gdiplus" (ByVal matrix As Long, matrixOut As Single) As GpStatus
Public Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal matrix As Long, result As Long) As GpStatus
Public Declare Function GdipIsMatrixIdentity Lib "gdiplus" (ByVal matrix As Long, result As Long) As GpStatus
Public Declare Function GdipIsMatrixEqual Lib "gdiplus" (ByVal matrix As Long, ByVal matrix2 As Long, result As Long) As GpStatus

' Region Functions (ALL)
Public Declare Function GdipCreateRegion Lib "gdiplus" (region As Long) As GpStatus
Public Declare Function GdipCreateRegionRect Lib "gdiplus" (rect As RECTF, region As Long) As GpStatus
Public Declare Function GdipCreateRegionRectI Lib "gdiplus" (rect As RECTL, region As Long) As GpStatus
Public Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal path As Long, region As Long) As GpStatus
' NOTE: The regionData parameter is treated as a byte array
Public Declare Function GdipCreateRegionRgnData Lib "gdiplus" (regionData As Any, ByVal size As Long, region As Long) As GpStatus
Public Declare Function GdipCreateRegionHrgn Lib "gdiplus" (ByVal hRgn As Long, region As Long) As GpStatus
Public Declare Function GdipCloneRegion Lib "gdiplus" (ByVal region As Long, cloneRegion As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetInfinite Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetEmpty Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal region As Long, rect As RECTF, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal region As Long, rect As RECTF, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal region As Long, ByVal path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal region As Long, ByVal region2 As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipTranslateRegion Lib "gdiplus" (ByVal region As Long, ByVal dx As Single, ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateRegionI Lib "gdiplus" (ByVal region As Long, ByVal dx As Long, ByVal dy As Long) As GpStatus
Public Declare Function GdipTransformRegion Lib "gdiplus" (ByVal region As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, rect As RECTF) As GpStatus
Public Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, rect As RECTL) As GpStatus
Public Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, hRgn As Long) As GpStatus
Public Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal region As Long, ByVal region2 As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipGetRegionDataSize Lib "gdiplus" (ByVal region As Long, bufferSize As Long) As GpStatus
' NOTE: The buffer parameter is treated as a byte array
Public Declare Function GdipGetRegionData Lib "gdiplus" (ByVal region As Long, buffer As Any, ByVal bufferSize As Long, sizeFilled As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal region As Long, ByVal x As Single, ByVal y As Single, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal region As Long, ByVal x As Long, ByVal y As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRect Lib "gdiplus" (ByVal region As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRectI Lib "gdiplus" (ByVal region As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal graphics As Long, result As Long) As GpStatus
Public Declare Function GdipGetRegionScansCount Lib "gdiplus" (ByVal region As Long, Ucount As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScans Lib "gdiplus" (ByVal region As Long, rects As RECTF, count As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScansI Lib "gdiplus" (ByVal region As Long, rects As RECTL, count As Long, ByVal matrix As Long) As GpStatus

' ImageAttributes APIs (ALL)
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As GpStatus
Public Declare Function GdipCloneImageAttributes Lib "gdiplus" (ByVal imageattr As Long, cloneImageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesToIdentity Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipResetImageAttributes Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
' NOTE: The colourMatrix and grayMatrix parameters expect a ColorMatrix structure or NULL
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As ColorMatrix, grayMatrix As Any, ByVal flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipSetImageAttributesThreshold Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal threshold As Single) As GpStatus
Public Declare Function GdipSetImageAttributesGamma Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal gamma As Single) As GpStatus
Public Declare Function GdipSetImageAttributesNoOp Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal colorLow As Long, ByVal colorHigh As Long) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannel Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjstType As ColorAdjustType, ByVal enableFlag As Long, ByVal channelFlags As ColorChannelFlags) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannelColorProfile Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal colorProfileFilename As String) As GpStatus
Public Declare Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, ByVal mapSize As Long, map As ColorMap) As GpStatus
Public Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As Long, ByVal wrap As WrapMode, ByVal argb As Long, ByVal bClamp As Long) As GpStatus
Public Declare Function GdipSetImageAttributesICMMode Lib "gdiplus" (ByVal imageattr As Long, ByVal bOn As Long) As GpStatus
Public Declare Function GdipGetImageAttributesAdjustedPalette Lib "gdiplus" (ByVal imageattr As Long, colorPal As ColorPalette, ByVal ClrAdjType As ColorAdjustType) As GpStatus

' FontFamily Functions (ALL)
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As String, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily Lib "gdiplus" (ByVal fontFamily As Long, clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace Lib "gdiplus" (nativeFamily As Long) As GpStatus
' NOTE: name must be LF_FACESIZE in length or less
Public Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal family As Long, ByVal name As String, ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal family As Long, ByVal style As Long, IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable Lib "gdiplus" (ByVal fontCollection As Long, ByVal graphics As Long, numFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, ByVal numFound As Long, ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight Lib "gdiplus" (ByVal family As Long, ByVal style As Long, EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent Lib "gdiplus" (ByVal family As Long, ByVal style As Long, CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent Lib "gdiplus" (ByVal family As Long, ByVal style As Long, CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing Lib "gdiplus" (ByVal family As Long, ByVal style As Long, LineSpacing As Integer) As GpStatus

' Font Functions (ALL)
Public Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hdc As Long, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontA Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTA, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontW Lib "gdiplus" (ByVal hdc As Long, logfont As LOGFONTW, createdfont As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As FontStyle, ByVal unit As GpUnit, createdfont As Long) As GpStatus
Public Declare Function GdipCloneFont Lib "gdiplus" (ByVal curFont As Long, cloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily Lib "gdiplus" (ByVal curFont As Long, family As Long) As GpStatus
Public Declare Function GdipGetFontStyle Lib "gdiplus" (ByVal curFont As Long, style As Long) As GpStatus
Public Declare Function GdipGetFontSize Lib "gdiplus" (ByVal curFont As Long, size As Single) As GpStatus
Public Declare Function GdipGetFontUnit Lib "gdiplus" (ByVal curFont As Long, unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI Lib "gdiplus" (ByVal curFont As Long, ByVal dpi As Single, Height As Single) As GpStatus
Public Declare Function GdipGetLogFontA Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, logfont As LOGFONTA) As GpStatus
Public Declare Function GdipGetLogFontW Lib "gdiplus" (ByVal curFont As Long, ByVal graphics As Long, logfont As LOGFONTW) As GpStatus
Public Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As Long, numFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, numFound As Long) As GpStatus
Public Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As Long, ByVal filename As String) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As Long, ByVal memory As Long, ByVal length As Long) As GpStatus

' Text Functions (ALL)
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal regionCount As Long, regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, ByVal brush As Long, positions As POINTF, ByVal flags As Long, ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString Lib "gdiplus" (ByVal graphics As Long, ByVal str As String, ByVal length As Long, ByVal thefont As Long, positions As POINTF, ByVal flags As Long, ByVal matrix As Long, boundingBox As RECTF) As GpStatus

' String format Functions (ALL)
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat Lib "gdiplus" (ByVal StringFormat As Long, newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, ByVal flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, trimming As Long) As GpStatus
Public Declare Function GdipSetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, ByVal hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipGetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal firstTabOffset As Single, ByVal count As Long, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal count As Long, firstTabOffset As Single, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount Lib "gdiplus" (ByVal StringFormat As Long, count As Long) As GpStatus
Public Declare Function GdipSetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, ByVal language As Integer, ByVal substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, language As Integer, substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount Lib "gdiplus" (ByVal StringFormat As Long, count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges Lib "gdiplus" (ByVal StringFormat As Long, ByVal rangeCount As Long, ranges As CharacterRange) As GpStatus

' GDI+ Memory Management Functions (ALL)
Public Declare Function GdipAlloc Lib "gdiplus" (ByVal size As Long) As Long
Public Declare Sub GdipFree Lib "gdiplus" (ByVal ptr As Long)



'-----------------------------------------------
' Helper Functions
'-----------------------------------------------


' Use this in lieu of the Color class constructor
' Thanks to Richard Mason for help with this
Public Function ColorARGB(ByVal alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   
   With bytestruct
      .AlphaByte = alpha
      .RedByte = Red
      .GreenByte = Green
      .BlueByte = Blue
   End With
   
   LSet result = bytestruct
   ColorARGB = result.longval
End Function

Public Function ColorSetAlpha(ByVal lColor As Long, ByVal alpha As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   
   result.longval = lColor
   LSet bytestruct = result

   bytestruct.AlphaByte = alpha

   LSet result = bytestruct
   ColorSetAlpha = result.longval
End Function

' Pass a GDI+ color to this function and get the VB compatible color
Public Function GetRGB_GDIP2VB(ByVal lColor As Long) As Long
   Dim argb As COLORBYTES
   CopyMemory argb, lColor, 4
   GetRGB_GDIP2VB = RGB(argb.RedByte, argb.GreenByte, argb.BlueByte)
End Function

' Pass a VB/standard color to this function and get the GDI+ compatible color
Public Function GetRGB_VB2GDIP(ByVal lColor As Long, Optional ByVal alpha As Byte = 255) As Long
   Dim rgbq As RGBQUAD
   CopyMemory rgbq, lColor, 4
   ' I must have done something wrong, but swapping Red and Blue works...
   GetRGB_VB2GDIP = ColorARGB(alpha, rgbq.rgbBlue, rgbq.rgbGreen, rgbq.rgbRed)
End Function


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
   Dim num As Long, size As Long, I As Long
   Dim ICI() As ImageCodecInfo
   Dim buffer() As Byte
   
   GetEncoderClsid = -1 'Failure flag

   ' Get the encoder array size
   Call GdipGetImageEncodersSize(num, size)
   If size = 0 Then Exit Function ' Failed!

   ' Allocate room for the arrays dynamically
   ReDim ICI(1 To num) As ImageCodecInfo
   ReDim buffer(1 To size) As Byte

   ' Get the array and string data
   Call GdipGetImageEncoders(num, size, buffer(1))
   ' Copy the class headers
   Call CopyMemory(ICI(1), buffer(1), (Len(ICI(1)) * num))

   ' Loop through all the codecs
   For I = 1 To num
      ' Must convert the pointer into a usable string
      If StrComp(PtrToStrW(ICI(I).MimeType), strMimeType, vbTextCompare) = 0 Then
         ClassID = ICI(I).ClassID   ' Save the class id
         GetEncoderClsid = I        ' return the index number for success
         Exit For
      End If
   Next
   ' Free the memory
   Erase ICI
   Erase buffer
End Function

' Same as above, but for decoders
' Built in decoders: (You can *try* to get other types also)
'   image/bmp
'   image/jpeg
'   image/gif
'   image/x-emf
'   image/x-wmf
'   image/tiff
'   image/png
'   image/x-icon
Public Function GetDecoderClsid(strMimeType As String, ClassID As CLSID)
   Dim num As Long, size As Long, I As Long
   Dim ICI() As ImageCodecInfo
   Dim buffer() As Byte

   GetDecoderClsid = -1 'Failure flag

   ' Get the encoder array size
   Call GdipGetImageDecodersSize(num, size)
   If size = 0 Then Exit Function ' Failed!

   ' Allocate room for the arrays dynamically
   ReDim ICI(1 To num) As ImageCodecInfo
   ReDim buffer(1 To size) As Byte

   ' Get the array and string data
   Call GdipGetImageDecoders(num, size, buffer(1))
   ' Copy the class headers
   Call CopyMemory(ICI(1), buffer(1), (Len(ICI(1)) * num))

   ' Loop through all the codecs
   For I = 1 To num
      ' Must convert the pointer into a usable string
      If StrComp(PtrToStrW(ICI(I).MimeType), strMimeType, vbTextCompare) = 0 Then
         ClassID = ICI(I).ClassID   ' Save the class id
         GetDecoderClsid = I        ' return the index number for success
         Exit For
      End If
   Next
   ' Free the memory
   Erase ICI
   Erase buffer
End Function


' Courtesy of: Dana Seaman
' Helper routine to convert a CLSID(aka GUID) string to a structure
Public Function DEFINE_GUID(ByVal sGuid As String) As CLSID
   ' Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
   Call CLSIDFromString(StrPtr(sGuid), DEFINE_GUID)
End Function

' From www.mvps.org/vbnet...i think
'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
Public Function PtrToStrW(ByVal lpsz As Long) As String
    Dim sOut As String
    Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function
Public Function PtrToStrA(ByVal lpsz As Long) As String
    Dim sOut As String
    Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        PtrToStrA = sOut
    End If
End Function

' This should hopefully simplify property item value retrieval
' NOTE: We are raising errors in this function; ensure the caller has error handing code.
'       The resulting arrays are using a base of one.
Public Function GetPropValue(item As PropertyItem) As Variant
   ' We need a valid pointer and length
   If item.value = 0 Or item.length = 0 Then Err.Raise 5, "GetPropValue"

   Select Case item.type
      ' We'll make Undefined types a Btye array as it seems the safest choice...
      Case PropertyTagTypeByte, PropertyTagTypeUndefined:
         Dim bte() As Byte: ReDim bte(1 To item.length)
         CopyMemory bte(1), ByVal item.value, item.length
         GetPropValue = bte
         Erase bte

      Case PropertyTagTypeASCII:
         GetPropValue = PtrToStrA(item.value)
         
      Case PropertyTagTypeShort:
         Dim short() As Integer: ReDim short(1 To (item.length / 2))
         CopyMemory short(1), ByVal item.value, item.length
         GetPropValue = short
         Erase short
         
      Case PropertyTagTypeLong, PropertyTagTypeSLONG:
         Dim lng() As Long: ReDim lng(1 To (item.length / 4))
         CopyMemory lng(1), ByVal item.value, item.length
         GetPropValue = lng
         Erase lng
         
      Case PropertyTagTypeRational, PropertyTagTypeSRational:
         Dim lngpair() As Long: ReDim lngpair(1 To (item.length / 8), 1 To 2)
         CopyMemory lngpair(1, 1), ByVal item.value, item.length
         GetPropValue = lngpair
         Erase lngpair

      Case Else: Err.Raise 461, "GetPropValue"
   End Select
End Function
