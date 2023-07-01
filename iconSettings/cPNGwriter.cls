VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPNGwriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Write PNG image format when use of GDI+ not available ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' This class is only used on non-GDI+ systems


' This class is a partial version of a full-blown PNG creation class.
' It creates a PNG with many options, but not nearly all PNG options avaialble
' and is specifically modified to creating PNGs from pre-multiplied 32bpp DIBs.
' There still may be references to Interlacing. Interlacing options do not
' exist in this class and any such references are results of extracting the
' routines from the full-version PNG creation class.
' Note that GDI+ does not offer any PNG options when creating PNGs, this class
' exposes several options and can be modified to support all PNG options.

' CUSTOM TAILORED FOR PRE-MULTIPLIED 32bpp DIBS. Routines not portable for normal DIBs.

' Required is a version of the zLIB DLL which can be found at www.zlib.net.
' zLIB comes in at least two varieties: C calling convention (_cdecl) and
' VB/PASCAL calling convention (_stdcall). This routine can use either of those
' conventions, but the zLIB file must be named one of the two following,
' not case sensitive, both are original filenames:  zLib.dll or zLib1.dll

' Key highlights:
' 1. PNGs can be created without GDI+ as long as zlib or zlib1 is present
' 2. Using bit reduction algorithms, a 32bpp DIB can be converted to one of the
'       the following:  8 bpp paletted, 24 bpp or 32 bpp PNGs; supporting full alpha
' 3. The PNG compression filtering mechanism in this routine is user-selected.
'       Filters assist in reorganizing byte information to make it compress better
'       Speed vs Size tradeoffs: filter type None is fastest while type Paeth is smallest (generally)
'       See notes in FilterImage routine, set filter in c32bppDIB.PngPropertySet routine
' 4. Over a dozen options available when creating PNGs, see Me.AddProperty
' 5. This class almost always creates smaller PNG files than GDI+ when default filtering is used
' 6. PNGs can be saved to file or saved to an array



' array mapping structures
Private Type SafeArray
    cDims As Integer                ' numer of dimensions
    fFeatures As Integer            ' not used
    cbElements As Long              ' byte size of each element (byte=1,Int=2,Long=4)
    cLocks As Long                  ' not used
    pvData As Long                  ' pointer to memory space containing array
    rgSABound(0 To 3) As Long
End Type

'Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Const INVALID_HANDLE_VALUE = -1

Private Enum eColorTypes    ' internal use only
    clrGrayScale = 0
    clrTrueColor = 2
    clrPalette = 3
    clrGrayAlpha = 4
    clrTrueAlpha = 6
End Enum

'Private Const zlibMaxCompression = 9 ' when zlib compress2 is available

'----------------------------------------------------------------------------
' following are optional PNG properties only
'Private m_Filter As eFilterMethods
'Private m_bKGD As Long          ' default PNG background color if a view opts to render against solid bkg
'Private m_Captions() As String  ' see c32bppDIB.ePngProperties
'Private m_PNGprops As Long      ' indicates which, if any, m_Captions are used
'----------------------------------------------------------------------------

'Private cCfunction As cCDECL        ' class to allow using C calling convention
'Private m_ZLIBadv As Boolean        ' which version of zLIB? true=use compress2 else use compress
'Private m_zLIBname As String        ' either zlib1 or zlib
'Private m_Palette() As Byte         ' PNG palette if image can be palettized
'Private m_transPal() As Byte        ' alpha values for PNG palettes as needed
'Private m_Uncompressed() As Byte    ' initialized, contains uncompressed DIB bytes in 8,24,32 bit formats
'Private m_Stream() As Byte          ' never initialized, overlay to host 32bpp DIB
'Private m_Trans As Long             ' flag indicating whether or not transparency is used in DIB
'Private m_ColorType As eColorTypes  ' the color type the PNG will be created in








'
'Private Function PropertyIndex(ByRef PropertyID As ePngProperties, Optional ByRef Caption_Misc As String) As Long
'
'    ' Helper function. Returns the m_Captions() array index for the passed PropertyID
'
'    Dim x As Long
'Dim CaptionID As Long
'
'
'    If PropertyID = pngProp_Miscellaneous Then
'        For CaptionID = 11 To UBound(m_Captions)
'            x = InStr(m_Captions(CaptionID), Chr$(0))
'            If StrComp(Left$(m_Captions(CaptionID), x - 1&), Caption_Misc) = 0& Then Exit For
'        Next
'        If CaptionID > UBound(m_Captions) Then CaptionID = -1&
'    Else
'        x = PropertyID
'        Do Until x = 1&
'            x = x \ 2&
'            CaptionID = CaptionID + 1&
'        Loop
'    End If
'
'    PropertyIndex = CaptionID
'
'End Function

'Private Function PalettizeImage(ByRef isAlpha As AlphaTypeEnum) As Boolean
'
'    ' Function determines if image can be palettized vs 24/32 bpp true color
'    ' Once determined it can be paletted, it will optimize to include the following:
'    ' 1. Convert to PNG grayscale palette if possible, saves at least 768 bytes vs color palette
'    ' 2. Rearrange palette to reduce alpha/palette entries, saves up to 200+ bytes if alpha is used
'    ' 3. Converts per-color grayscale to a modified color palette, reducing size at least 50%
'    ' This modified version does not reduce to 1,2,or 4 bits per pixel
'    '   -- Any paletted image is 256 colors, but only the needed palette entries are cached in PNG
'
'    Dim x As Long
'Dim y As Long
'Dim scanWidth As Long
'
'    Dim palCount As Long
'Dim Index As Long
'
'    Dim Color As Long
'Dim newColor As Boolean
'
'    Dim palXRef() As Byte
'Dim palAlpha() As Byte
'
'    Dim tSortIndex() As Long
'Dim tPalette() As Long
'
'
'    On Error GoTo ExitRoutine
'
'    ' count unique colors (maximum of 256 if we are to palettize)
'    ' Note that alphas are included in the tSortIndex. This is because any color
'    ' using more than one alpha value would require separate palette entries:
'    ' Example: Red @ Alpha 255 & Red @ Alpha 128 requires two palette entries
'    ReDim m_transPal(1 To 256)          ' array to hold alpha values only
'    ReDim tSortIndex(1 To 256)          ' sort indexes
'    ReDim tPalette(1 To 256) As Long    ' palette
'    For y = 0& To UBound(m_Stream, 2)
'        For x = 0& To UBound(m_Stream, 1) Step 4&
'
'            CopyMemory Color, m_Stream(x, y), 4&
'            Index = iparseFindColor(tSortIndex, Color, palCount, newColor)   ' use binary search routine
'            If newColor = True Then
'                If palCount = 256& Then Exit Function       ' exceeded palette entries limit
'                palCount = palCount + 1&                    ' increment entry count & shift palette to maintain asc sort
'                If Index < palCount Then
'                    CopyMemory tSortIndex(Index + 1&), tSortIndex(Index), (palCount - Index) * 4&
'                    CopyMemory tPalette(Index + 1&), tPalette(Index), (palCount - Index) * 4&
'                End If
'                tSortIndex(Index) = Color                    ' add new color to the palette
'                CopyMemory tPalette(Index), Color, 3&
'            End If
'
'        Next
'    Next
'
'    ' if we got here, then image can be palettized, but to which of the following?
'    ' 1. Palette - no transparency? like non-transparent GIFs (isAlpha=False)
'    ' 2. Palette - simple transparency? like transparent GIFs (grayscale handled differently in PNGs)
'    ' 3. Palette - per-color index transparency?
'
'    y = 0&
'    If isAlpha Then
'        ' separate alpha from color and count how many non-opaque alpha values
'        For x = 1& To palCount
'            If (tSortIndex(x) And &H7FFFFFFF) = tSortIndex(x) Then ' high bit not set
'                m_transPal(x) = tSortIndex(x) \ &H1000000
'            Else                                                 ' high bit is set
'                m_transPal(x) = ((tSortIndex(x) And &H7FFFFFFF) \ &H1000000) Or &H80
'            End If
'            If Not m_transPal(x) = 255 Then
'                y = y + 1&   ' count different levels of transparency
'                Index = x    ' track last palette entry with alpha value <> 255
'            End If
'        Next
'    Else
'        FillMemory m_transPal(1), 256&, 255 ' all alphas are opaque
'    End If
'
'    Select Case y
'    Case 0&
'    ' 1. Palette - no transparency? like non-transparent GIFs (isAlpha=False)
'        m_Trans = -1& ' no transparency
'    Case 1&
'    ' 2. Palette - simple transparency? like transparent GIFs
'        m_Trans = Index ' flag & may be changed later in this routine
'    Case Else
'    ' 3. Palette - per-color transparency?
'        m_Trans = 0& ' > -1 means we have transparency at some level
'        ' alphas are kept in the m_transPal() array
'    End Select
'
''     Now for the last optimization attempt: check for grayscale but only for non per-color
''     alpha images. Why restrict grayscale to non per-color alpha when PNG can support grayscale
''     per-color alpha? Here's why: per-color grayscale alpha is ColorType 4. ColorType 4
''     always requires 16 bits per pixel (bpp), regardless of grayscale bit depth, but
''     ColorType 3 requires 8 bpp (max) + 768 palette bytes (max) + 256 bytes (max) for alpha info:
''       ColorType 4, 256x256 image: 256*256*2=131072 bytes for color information (grayscale has no palette in PNGs)
''       ColorType 3 (8bpp), 256x256 image: 256*256*1+768+256=66560 bytes for color information
''           note: ColorType 4 is always 16bpp, but ColorType 3 can be 1,2,4,8 bpp
''           and palette/alpha arrays can be reduced too
'
'    m_ColorType = clrPalette    ' Color Type 3 (color palette)
'    If Not m_Trans = 0& Then
'        ' check each palette entry to see if grayscale or not. When not, abort loop
'        For Index = 1& To palCount
'            If Not (tPalette(Index) And &HFF) = ((tPalette(Index) \ &H100&) And &HFF) Then ' compare B to G
'                Exit For
'            ElseIf Not (tPalette(Index) And &HFF) = ((tPalette(Index) \ &H10000) And &HFF) Then ' compare B to R
'                Exit For
'            End If
'        Next
'
'        If Index > palCount Then    ' need to tweak transparency possibly
'
'            m_ColorType = clrGrayScale ' Color Type 0
'
'            If isAlpha > AlphaNone Then
'                ' we only got here because just 1 color was transparent & with a pre-multiplied DIB
'                ' that color is always black. But if non-transparent black was used elsewhere in the
'                ' grayscale then we need to change the transparency. Non-transparent black is very
'                ' common in grayscales
'                ReDim palXRef(1 To 256) ' track which grayscales are used
'                For x = 1& To palCount
'                    If tPalette(x) = 0& Then             ' this is black
'                        If Not m_transPal(x) = 0& Then   ' and not our transparent black
'                            palXRef(1) = 1     ' mark black as used
'                        End If
'                    Else
'                        palXRef(Index + 1&) = 1 ' non-black, mark as used
'                    End If
'                Next
'                If palXRef(1) = 1 Then
'                    ' non-transparent black is used in the grayscale, so we must change our
'                    ' tranparent black - Locate a grayscale not in use
'                    For Index = 2& To palCount
'                        If palXRef(Index) = 0 Then
'                            ' bingo, we'll use this one
'                            m_Trans = Index - 1&
'                            Exit For
'                        End If
'                    Next
'                Else    ' black was not in the image, we can use black as transparency
'                    m_Trans = 0&
'                End If
'                Erase palXRef
'            End If
'        End If
'    End If
'
'    scanWidth = UBound(m_Stream, 1) \ 4& + 1&         ' width of image
'    x = (UBound(m_Stream, 2) + 1) * scanWidth - 1&   ' calculate size of total image bytes
'    ReDim m_Uncompressed(0 To x)    ' the Filter function expects 1D arrays
'
'    If m_ColorType = clrGrayScale Then
'        ' grayscale is easy enough, transfer 32bpp info to 8bpp info
'        ' Remember, PNG grayscale color types do not use palettes.
'        ' Grayscale palettes are PNG decoders responsibility
'        For y = 0& To UBound(m_Stream, 2)
'            Index = y * scanWidth
'            For x = 0& To UBound(m_Stream, 1) Step 4&
'                If m_Stream(x + 3&, y) = 0& Then     ' transparency index needed
'                    m_Uncompressed(Index) = m_Trans ' use modified transparency index as necessary
'                Else
'                    m_Uncompressed(Index) = m_Stream(x, y)  ' use grayscale index
'                End If
'                Index = Index + 1&
'            Next
'        Next
'
'    Else
'        ' for color palettes, we want to re-order entries when per-color alpha is used.
'        ' Why the hassle? Shrink PNG a bit more. When color palettes have transparency,
'        ' you must have a 1 byte Alpha value for each palette entry. But, that 1 byte
'        ' alpha value, when = 255, is optional and implied. Therefore, if we move all
'        ' palette entries with transparency to top of array, then all those 255s at the
'        ' bottom of the array don't need to be cached in the PNG; not being there, PNG
'        ' decoders must assume value is 255. We can save anywhere up to 200+ bytes
'        ' depending on the image.
'        ReDim palXRef(0 To 1, 0 To palCount - 1)
'
'        If m_Trans = -1& Then   ' no transparencies and not grayscale
'            For x = 0& To palCount - 1& ' all entries are opaque, no cross-referencing needed
'                palXRef(0, x) = x
'                palXRef(1, x) = x
'            Next
'        Else                    ' per-color alpha being used
'            ' since we are re-ordering, we also need to build a cross-reference so
'            ' we can reference palette locations, old to new and vice versa
'            y = 0&: x = palCount - 1&  ' starting points for top & bottom of array
'            For Index = 0& To palCount - 1&
'                If m_transPal(Index + 1&) = 255 Then
'                    palXRef(1, x) = Index ' keep full opaque entries at bottom of array
'                    palXRef(0, Index) = x ' double link reference
'                    x = x - 1&
'                Else
'                    palXRef(0, Index) = y ' move non-opaque entries near top of array
'                    palXRef(1, y) = Index ' double link reference
'                    y = y + 1&
'                End If
'            Next
'        End If
'
'        ' now we build our 8 bpp paletted image, referencing the re-sorted palette entires
'        For y = 0& To UBound(m_Stream, 2)
'            Index = y * scanWidth
'            For x = 0& To UBound(m_Stream, 1) Step 4&
'                ' get 32bit color from DIB
'                CopyMemory Color, m_Stream(x, y), 4&
'                ' locate it in our temp palette using binary search algorithm
'                Color = iparseFindColor(tSortIndex, Color, palCount, False)
'                ' now cache its re-sorted reference
'                m_Uncompressed(Index) = palXRef(0, Color - 1&)
'                Index = Index + 1&
'            Next
'        Next
'
'        ' good, now we need to build the palette the PNG will use,
'        ' but we will be using 3 byte values, not 4 byte values & colors need to be RGB vs BGR
'        ReDim m_Palette(1 To palCount * 3& + 4&) ' extra 4 bytes are used during Write_PLTE
'        For x = 1& To palCount
'            ' calculate new index for this palette entry
'            Index = palXRef(0, x - 1&) * 3& + 5&   ' offset that extra 4 bytes too
'            ' simultaneously remove pre-multiplication
'            Select Case m_transPal(x)
'            Case 255    ' full opaque
'                m_Palette(Index) = (tPalette(x) \ &H10000) And &HFF&
'                m_Palette(Index + 1&) = (tPalette(x) \ &H100&) And &HFF&
'                m_Palette(Index + 2&) = tPalette(x) And &HFF&
'            Case 0: ' do nothing, color is always 0,0,0
'            Case Else
'                m_Palette(Index) = (((tPalette(x) \ &H10000) And &HFF&) * m_transPal(x) \ 255)
'                m_Palette(Index + 1&) = (((tPalette(x) \ &H100&) And &HFF&) * m_transPal(x) \ 255)
'                m_Palette(Index + 2&) = ((tPalette(x) And &HFF&) * m_transPal(x) \ 255)
'            End Select
'        Next
'        Erase tPalette()
'
'        If m_Trans = -1& Then
'            Erase m_transPal()
'        Else
'            ' now we are going to double check how many non-opaque palette entries we have
'            For x = 0& To palCount - 1&
'                If m_transPal(palXRef(1, x) + 1&) = 255 Then
'                    palCount = x    ' we saved 256-X bytes at least
'                    Exit For
'                End If
'            Next
'            ReDim palAlpha(1 To palCount + 4&)   ' extra 4 bytes used in Write_tRNS
'            ' rewrite the m_transPal array, only caching non-opaque palette entries
'            For x = 0& To palCount - 1&
'                palAlpha(x + 5&) = m_transPal(palXRef(1, x) + 1&)
'            Next
'            m_transPal = palAlpha
'
'        End If
'    End If
'
'    PalettizeImage = True
'
'ExitRoutine:
'End Function

'Private Function OptimizeTrueColor(ByVal isAlpha As AlphaTypeEnum) As Boolean
'
'    ' Function attempts to reduce 32bpp DIB to 24bpp DIB.  Check is not brute force
'    ' Reduction to Palette already tried before this routine was called
'    ' Reduction can occur when:
'    '   1. No transparency is used (all Alpha values are 255), or
'    '   2. Only simple transparency is used (if alpha between 1 & 254 then no reduction)
'
'    Dim x As Long
'Dim y As Long
'
'    Dim bAbort As Boolean
'Dim tOffset As Long
'
'    Dim palAlpha(0 To 255) As Byte
'Dim palCount As Long
'
'    Dim scanWidth As Long
'Dim Color As Long
'
'
'    m_Trans = -1&       ' flag indicating no simple transparency. ColorType 6 (clrTrueAlpha) implies transparency
'    If isAlpha Then
'        m_ColorType = clrTrueAlpha  ' default color type for this DIB
'
'        ' can reduction to 24bpp be done?
'
'        ' This routine will run quickly. We will not make the effort to check every possible
'        ' color in the 16 million color range, rather, we will be looking at a max of 768
'        ' colors: 256 Reds, 256 Greens, & 256 Blues (grayscales only). If we find one we
'        ' can use, bingo, else write as 32bpp
'
'        ' Now here's the catch. When full transparency is in play, black is always the transparent
'        ' color in premultiplied DIBs, but if non-transparent black (i.e., RGB=0,0,0:Alpha>0) is
'        ' used anywhere else in the image, then we can't leave black as the transparent color;
'        ' we'll need to change it.
'
'        For Color = 0& To 2&  ' Color = B, G, R
'
'            tOffset = 3& - Color ' location of the alpha byte relative to "Color"
'            palCount = 0&       ' number of "Color" shades used; from 1 to 256
'
'            For y = 0& To UBound(m_Stream, 2)
'                For x = Color To UBound(m_Stream, 1) Step 4&
'
'                    Select Case m_Stream(x + tOffset, y)
'                    Case 255    ' fully opaque
'                        If palAlpha(m_Stream(x, y)) = 0 Then    ' has it been counted?
'                            palCount = palCount + 1&            ' up the count & abort if we maxed out
'                            If palCount = 256& Then             ' all 256 shades of "Color" used
'                                y = UBound(m_Stream, 2)         ' force early termination of loop Y
'                                Exit For
'                            End If
'                            palAlpha(m_Stream(x, y)) = 1        ' flag it
'                        End If
'                    Case 0      ' fully transparent
'                    Case Else   ' partial transparency              ' can't reduce to 24bpp
'                        y = UBound(m_Stream, 2)                     ' force early termination of loop Y
'                        palCount = 256&                             ' prevent next IF statement from executing
'                        Color = 3&                                  ' force early termination of loop Color
'                        Exit For
'                    End Select
'                Next
'            Next
'
'            If Not palCount = 256& Then ' did we find a color we can use? It could be black
'                For x = 0& To 255&  ' lets find out which it is
'                    If palAlpha(x) = 0 Then
'                        ' since the X-shade of the R, G, or B isn't used in the image,
'                        ' we can safely state that RGB(X,X,X) is also not in the image
'                        m_Trans = x Or (x * &H100&) Or (x * &H10000)
'                        Exit For
'                    End If
'                Next
'                m_ColorType = clrTrueColor  ' reduce to 24bpp vs 32bpp
'                Exit For
'            End If
'
'            Erase palAlpha()    ' reset to zeros
'        Next
'    Else
'        m_ColorType = clrTrueColor      ' no transparency, reduction to 24bpp
'    End If
'
'    ' Remove premultiplication & convert to 24bpp if needed
'    ' Use separate loops vs adding an IF statement for every pixel to test for color type
'
'    If m_ColorType = clrTrueAlpha Then  ' 32bpp (ColorType 6)
'        scanWidth = UBound(m_Stream, 1) + 1&
'        ReDim m_Uncompressed(0 To scanWidth * (UBound(m_Stream, 2) + 1&) - 1&)
'        For y = 0& To UBound(m_Stream, 2)
'            tOffset = y * scanWidth
'            For x = 0& To UBound(m_Stream, 1) Step 4&
'                ' remove pre-multiplication
'                Select Case m_Stream(x + 3&, y)
'                Case 255
'                    m_Uncompressed(tOffset) = m_Stream(x + 2&, y)
'                    m_Uncompressed(tOffset + 1&) = m_Stream(x + 1&, y)
'                    m_Uncompressed(tOffset + 2&) = m_Stream(x, y)
'                    m_Uncompressed(tOffset + 3&) = 255
'                Case 0 ' do nothing
'                Case Else
'                    Color = m_Stream(x + 3&, y)
'                    m_Uncompressed(tOffset) = (255& * m_Stream(x + 2&, y) \ Color)
'                    m_Uncompressed(tOffset + 1&) = (255& * m_Stream(x + 1&, y) \ Color)
'                    m_Uncompressed(tOffset + 2&) = (255& * m_Stream(x, y) \ Color)
'                    m_Uncompressed(tOffset + 3&) = Color
'                End Select
'                tOffset = tOffset + 4&
'            Next
'        Next
'    Else            ' 24bpp (Color Type 2) with or without simple transparency
'        scanWidth = iparseByteAlignOnWord(24, UBound(m_Stream, 1) \ 4 + 1&)
'        ' convert BGR to RGB, the Filter function expects 1D arrays
'        ReDim m_Uncompressed(0 To scanWidth * (UBound(m_Stream, 2) + 1&) - 1&)
'        For y = 0& To UBound(m_Stream, 2)
'            tOffset = y * scanWidth
'            For x = 0& To UBound(m_Stream, 1) Step 4&
'                ' simultaneously remove pre-multiplication. Don't carry over any alpha values
'                Select Case m_Stream(x + 3&, y)
'                Case 255
'                    m_Uncompressed(tOffset) = m_Stream(x + 2&, y)
'                    m_Uncompressed(tOffset + 1&) = m_Stream(x + 1&, y)
'                    m_Uncompressed(tOffset + 2&) = m_Stream(x, y)
'                Case 0 ' uses simple transparency (1 color is transparent)
'                    CopyMemory m_Uncompressed(tOffset), m_Trans, 3&
'                Case Else
'                    Color = m_Stream(x + 3&, y)
'                    m_Uncompressed(tOffset) = (255& * m_Stream(x + 2&, y) \ Color)
'                    m_Uncompressed(tOffset + 1&) = (255& * m_Stream(x + 1&, y) \ Color)
'                    m_Uncompressed(tOffset + 2&) = (255& * m_Stream(x, y) \ Color)
'                End Select
'                tOffset = tOffset + 3&
'            Next
'        Next
'    End If
'End Function



'Private Function Write_IHDR(ByRef FileNum As Long, ByRef Stream() As Byte, ByRef Host As c32bppDIB, ByRef isInterlaced As Boolean) As Boolean
'
'    Const png_Signature1 As Long = 1196314761
'    Const png_Signature2 As Long = 169478669
'    Const chnk_IHDR As Long = &H52444849 'Image header
'
'    On Error GoTo eh
'    Dim pngData(0 To 16) As Byte ' 13 byte header + 4 byte chunk name
'    Dim gpLong As Long           ' general purpose variable
'    Dim rwLen As Long
'
'    ' build header
'    CopyMemory pngData(0), chnk_IHDR, 4&    ' chunk name
'    gpLong = iparseReverseLong(Host.Width)   ' png width
'    CopyMemory pngData(4), gpLong, 4&
'    gpLong = iparseReverseLong(Host.Height)  ' png height
'    CopyMemory pngData(8), gpLong, 4&
'
'    ' bit depth, 16bit (PNG 16 bytes per R,G,B element or 48 bytes per pixel)
'    ' not supported via this class
'    pngData(12) = 8 ' only 1,2,4,48 bpp are different, 8,24,32 bpp is 8
'
'    pngData(13) = m_ColorType
'    ' pngData(14) & (15) will always be zero (compression/filter methods)
'    ' next byte is 1 if interlaced
'    pngData(16) = Abs(isInterlaced)
'
'    If FileNum = 0& Then ' writing to array vs file
'        ReDim Stream(0 To 32) ' png signature, header len, header, crc value (33 bytes)
'        CopyMemory Stream(0), png_Signature1, 4&
'        CopyMemory Stream(4), png_Signature2, 4&
'        gpLong = iparseReverseLong(13&) ' len of header
'        CopyMemory Stream(8), gpLong, 4&
'        CopyMemory Stream(12), pngData(0), 17&
'        gpLong = zCreateCRC(VarPtr(pngData(0)), 17&)
'        CopyMemory Stream(29), gpLong, 4&
'        Write_IHDR = True
'    Else
'
'        WriteFile FileNum, png_Signature1, 4&, rwLen, ByVal 0&
'        If rwLen = 4& Then
'            WriteFile FileNum, png_Signature2, rwLen, rwLen, ByVal 0&
'            If rwLen = 4& Then
'                WriteFile FileNum, iparseReverseLong(13&), rwLen, rwLen, ByVal 0&
'                If rwLen = 4& Then
'                    WriteFile FileNum, pngData(0), 17&, rwLen, ByVal 0&
'                    If rwLen = 17& Then
'                        WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), rwLen), 4&, rwLen, ByVal 0&
'                        Write_IHDR = (rwLen = 4&)
'                    End If
'                End If
'            End If
'        End If
'    End If
'eh:
'    If Err Then Err.Clear
'End Function
'
'
'Private Function Write_PLTE(ByRef FileNum As Long, ByRef Stream() As Byte, ByRef Invalid_bKGD As Boolean) As Boolean
'
'    ' Note: the palette is preprocessed before it arrives here: BGR>RGB
'    On Error GoTo eh
'
'    If m_ColorType = clrPalette Then ' paletted images only
'
'        Const chnk_PLTE As Long = &H45544C50 'Palette
'
'        Dim gpLong As Long          ' general purpose variable
'        Dim Index As Long
'        Dim rwLen As Long
'
'        ' when paletted, the bKGD chunk comes after the palette, but for palettes the
'        ' bkgd chunk must be one of the palette entries, therefore, we will attempt to
'        ' find the color in the palette, add it to the palette if possible, or skip
'        ' the optional chunk if color is not in the palette
'        If (m_PNGprops And ePngProperties.pngProp_DefaultBkgColor) = ePngProperties.pngProp_DefaultBkgColor Then
'            Dim bkg(0 To 2) As Byte
'            CopyMemory bkg(0), m_bKGD, 3&
'            For Index = 5& To UBound(m_Palette) Step 3&
'                If bkg(0) = m_Palette(Index) Then
'                    If bkg(1) = m_Palette(Index + 1&) Then
'                        If bkg(2) = m_Palette(Index + 2&) Then Exit For
'                    End If
'                End If
'            Next
'            If Index < UBound(m_Palette) Then   ' found it, ref the index
'                m_bKGD = (Index - 5&) \ 3&
'            ElseIf UBound(m_Palette) < 772& Then ' we can add it, let's do that
'                ' ^^ 772 is 256*3+4
'                ReDim Preserve m_Palette(1 To UBound(m_Palette) + 3&)
'                m_bKGD = (UBound(m_Palette) - 5&) \ 3&
'                CopyMemory m_Palette(UBound(m_Palette) - 2&), bkg(0), 3&
'            Else
'                Invalid_bKGD = True ' do not write the bkgd chunk
'            End If
'        End If
'
'        CopyMemory m_Palette(1), chnk_PLTE, 4&
'        gpLong = UBound(m_Palette)
'
'        If FileNum = 0& Then 'writing to array vs file
'            Index = UBound(Stream) + 1&
'            ReDim Preserve Stream(0 To Index + gpLong + 7&)
'            rwLen = iparseReverseLong(gpLong - 4&)
'            CopyMemory Stream(Index), rwLen, 4&                 ' size of chunk
'            CopyMemory Stream(Index + 4&), m_Palette(1), gpLong ' palette
'            rwLen = zCreateCRC(VarPtr(m_Palette(1)), gpLong)
'            CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&   ' crc
'            Write_PLTE = True
'        Else
'            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'            If rwLen = 4& Then
'                WriteFile FileNum, m_Palette(1), gpLong, rwLen, ByVal 0&
'                If rwLen = gpLong Then
'                    WriteFile FileNum, zCreateCRC(VarPtr(m_Palette(1)), gpLong), 4&, rwLen, ByVal 0&
'                    Write_PLTE = (rwLen = 4&)
'                End If
'            End If
'        End If
'        Erase m_Palette()   ' no longer needed
'    Else
'        Write_PLTE = True
'    End If
'eh:
'    If Err Then Err.Clear
'
'End Function
'
'Private Function Write_tEXt(ByRef FileNum As Long, ByRef Stream() As Byte, ByVal bTitleAuthorOnly As Boolean) As Boolean
'
'    ' Function writes uncompressed standard Keywords & text to the PNG
'
'    ' Note. Per PNG specs, some text should be written near top of the file while others
'    ' should be written near the end. There is no requirement for text to appear in
'    ' any specific location. The logic for writing some near the top is for search
'    ' engines only. It would be faster to find that text if nearer the top.
'    ' Therefore, this routine is called twice, once near the top of the PNG and
'    ' and again just before the IEND chunk is written
'
'    Const chnk_tEXt As Long = &H74584574 'Text - uncompressed
'
'    On Error GoTo ExitRoutine
'
'    Dim pngData() As Byte   ' data to be written to PNG file
'    Dim txtData() As Byte
'    Dim gpLong As Long
'    Dim lenKeyword As Long
'    Dim lenText As Long
'
'    Dim Index As Long
'    Dim CaptionID As Long
'    Dim tProps As Long
'    Dim lastCaption As Long
'    Dim keyWord As String
'    Dim rwLen As Long
'
'    If bTitleAuthorOnly Then    ' called after writing IHDR
'        CaptionID = ePngProperties.pngProp_Title
'        lastCaption = ePngProperties.pngProp_Description
'    Else                        ' called before writing IEND
'        CaptionID = ePngProperties.pngProp_Description
'        lastCaption = ePngProperties.pngProp_Miscellaneous
'    End If
'    tProps = m_PNGprops
'    Do Until CaptionID = lastCaption
'        If (tProps And CaptionID) = CaptionID Then
'            tProps = tProps And Not CaptionID
'            Select Case CaptionID
'            Case pngProp_Title: Index = 0
'                keyWord = "Title" & Chr$(0)
'            Case pngProp_Author: Index = 1&
'                keyWord = "Author" & Chr$(0)
'            Case pngProp_Comment: Index = 9&
'                keyWord = "Comment" & Chr$(0)
'            Case pngProp_Copyright: Index = 3&
'                keyWord = "Copyright" & Chr$(0)
'            Case pngProp_CreationTime: Index = 4&
'                keyWord = "Creation Time" & Chr$(0)
'            Case pngProp_Description: Index = 2&
'                keyWord = "Description" & Chr$(0)
'            Case pngProp_Disclaimer: Index = 6&
'                keyWord = "Disclaimer" & Chr$(0)
'            Case pngProp_Software: Index = 5&
'                keyWord = "Software" & Chr$(0)
'            Case pngProp_Source: Index = 8&
'                keyWord = "Source" & Chr$(0)
'            Case pngProp_Warning: Index = 7&
'                keyWord = "Warning" & Chr$(0)
'            End Select
'
'            ' tXTt chunk format::
'            'Keyword 1-79 bytes (character string)
'            'Null separator 1 byte (null character)
'            'Text string 0 or more bytes (character string)
'
'            lenKeyword = Len(keyWord)
'            txtData() = StrConv(keyWord, vbFromUnicode)
'            If Len(m_Captions(Index)) > 0& Then
'                lenText = Len(m_Captions(Index))
'                ReDim pngData(1 To lenKeyword + lenText + 4&)
'                CopyMemory pngData(5), txtData(0), lenKeyword
'                txtData() = StrConv(m_Captions(Index), vbFromUnicode)
'                CopyMemory pngData(5& + lenKeyword), txtData(0), lenText
'
'            Else ' handle zero-length chunks.
'                ' Note: I would prefer to just skip these, but maybe you might
'                ' decide to use one as a flag for something else?
'                ReDim pngData(1 To lenKeyword + 4&)
'                CopyMemory pngData(5), txtData(0), lenKeyword
'            End If
'            CopyMemory pngData(1), chnk_tEXt, 4&
'            gpLong = lenKeyword + lenText + 4&
'
'            If FileNum = 0& Then ' writing to stream
'                Index = UBound(Stream) + 1&
'                ReDim Preserve Stream(0 To Index + gpLong + 7&)
'                rwLen = iparseReverseLong(gpLong - 4&)
'                CopyMemory Stream(Index), rwLen, 4&
'                CopyMemory Stream(Index + 4), pngData(1), gpLong
'                rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
'                CopyMemory Stream(Index + 4& + gpLong), rwLen, 4&
'                Write_tEXt = True
'            Else
'                WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'                If rwLen = 4& Then
'                    WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
'                    If rwLen = gpLong Then
'                        WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
'                        Write_tEXt = (rwLen = 4&)
'                    End If
'                End If
'            End If
'        End If
'        CaptionID = CaptionID * 2&
'    Loop
'
'ExitRoutine:
'    If Err Then
'        Err.Clear
'    Else
'        If lenKeyword = 0& Then Write_tEXt = True
'    End If
'
'End Function
'
'Private Function Write_tIMe(ByRef FileNum As Long, ByRef Stream() As Byte) As Boolean
'
'    ' Note: the time stamp should be Universal Time, not local area
'
'    If (m_PNGprops And ePngProperties.pngProp_DateTimeModified) = ePngProperties.pngProp_DateTimeModified Then
'
'        Const chnk_tIME As Long = &H454D4974 'Timestamp
'
'        On Error GoTo eh
'        Dim pngData(0 To 10) As Byte ' 7 byte date/time + 4 byte chunk name
'        Dim gpLong As Long
'        Dim gpInt As Integer
'        Dim dtStamp As Date
'        Dim rwLen As Long
'
'        dtStamp = CDate(m_Captions(10))
'
'        CopyMemory pngData(0), chnk_tIME, 4&
'            gpInt = Year(dtStamp)
'        CopyMemory pngData(5), gpInt, 2&
'        pngData(4) = pngData(6)             ' swap endian of integer
'            gpInt = Month(dtStamp)
'        CopyMemory pngData(6), gpInt, 1&
'            gpInt = Day(dtStamp)
'        CopyMemory pngData(7), gpInt, 1&
'            gpInt = Hour(dtStamp)
'        CopyMemory pngData(8), gpInt, 1&
'            gpInt = Minute(dtStamp)
'        CopyMemory pngData(9), gpInt, 1&
'            gpInt = Second(dtStamp)
'        CopyMemory pngData(10), gpInt, 1&
'
'        If FileNum = 0& Then ' writing to stream
'            gpLong = UBound(Stream) + 1&
'            ReDim Preserve Stream(0 To gpLong + 18&)
'            rwLen = iparseReverseLong(7)
'            CopyMemory Stream(gpLong), rwLen, 4&
'            CopyMemory Stream(gpLong + 4&), pngData(0), 11&
'            rwLen = zCreateCRC(VarPtr(pngData(0)), 11&)
'            CopyMemory Stream(gpLong + 15&), rwLen, 4&
'            Write_tIMe = True
'        Else
'            WriteFile FileNum, iparseReverseLong(7), 4&, rwLen, ByVal 0&
'            If rwLen = 4& Then
'                WriteFile FileNum, pngData(0), 11&, rwLen, ByVal 0&
'                If rwLen = 11& Then
'                    WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), rwLen), 4&, rwLen, ByVal 0&
'                    Write_tIMe = (rwLen = 4&)
'                End If
'            End If
'        End If
'    Else
'        Write_tIMe = True
'    End If
'eh:
'    If Err Then Err.Clear
'
'End Function
'
'Private Function Write_tRNS(ByRef FileNum As Long, ByRef Stream() As Byte) As Boolean
'
'    ' For paletted/grayscale images, tRNS is the palette index, otherwise RGB value
'    On Error GoTo eh
'
'    If m_Trans = -1& Then
'        Write_tRNS = True
'
'    Else ' transparency not used
'
'        Const chnk_tRNS As Long = &H534E5274 'Simple Transparency & palette transparency
'        Dim Index As Long
'        Dim gpLong As Long
'        Dim rwLen As Long
'
'        Select Case m_ColorType
'
'            Case clrPalette ' Paletted (palette count * 3 + 4 byte chunk name)
'                ' nothing to do; done during PalettizeImage
'
'            Case clrGrayScale ' grayscale
'                ReDim m_transPal(1 To 6)   ' 2 bytes + 4 byte chunk name
'                m_transPal(6) = m_Trans
'                ' Note: m_transPal(5) used with 48bit per pixel images (not supported)
'
'            Case clrTrueColor ' we have simple transparency for true color
'                ReDim m_transPal(1 To 10)   ' 6 bytes + 4 byte chunk name
'                m_transPal(6) = m_Trans And &HFF
'                m_transPal(8) = (m_Trans \ &H100&) And &HFF
'                m_transPal(10) = (m_Trans \ &H10000) And &HFF
'                ' Note: m_transPal(5,7,9) used with 48bit per pixel images (not supported)
'
'            Case Else
'                ' Color Types 4 & 6 are prohibited from having a tRNS chunk
'                Write_tRNS = True
'                Exit Function
'        End Select
'
'        CopyMemory m_transPal(1), chnk_tRNS, 4&
'        gpLong = UBound(m_transPal)
'
'        ' write the chunk
'        If FileNum = 0& Then ' writing to array vs file
'            Index = UBound(Stream) + 1&
'            ReDim Preserve Stream(0 To Index + gpLong + 7&)
'            rwLen = iparseReverseLong(gpLong - 4&)
'            CopyMemory Stream(Index), rwLen, 4&                 ' chunk size
'            CopyMemory Stream(Index + 4&), m_transPal(1), gpLong ' palette
'            gpLong = zCreateCRC(VarPtr(m_transPal(1)), gpLong)
'            CopyMemory Stream(Index + UBound(m_transPal) + 4&), gpLong, 4&       ' crc value
'            Write_tRNS = True
'        Else
'            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'            If rwLen = 4& Then
'                WriteFile FileNum, m_transPal(1), gpLong, rwLen, ByVal 0&
'                If rwLen = gpLong Then
'                    WriteFile FileNum, zCreateCRC(VarPtr(m_transPal(1)), gpLong), 4&, rwLen, ByVal 0&
'                    Write_tRNS = (rwLen = 4&)
'                End If
'            End If
'        End If
'        Erase m_transPal()
'    End If
'eh:
'    If Err Then Err.Clear
'
'End Function
'
'Private Function Write_zTXt(ByRef FileNum As Long, ByRef Stream() As Byte) As Boolean
'
'    ' Function writes non-reserved keyword compressed/uncompressed text to the PNG
'
'    If (m_PNGprops And ePngProperties.pngProp_Miscellaneous) = ePngProperties.pngProp_Miscellaneous Then
'
'        On Error GoTo eh
'
'        Const chnk_tEXt As Long = &H74584574 'Text - uncompressed
'        Const chnk_zTXt As Long = &H7458547A 'Text - compressed
'
'        Dim txtData() As Byte   ' comments/text in bytes
'        Dim pngData() As Byte   ' data to be written to PNG file
'        Dim sText As String
'        Dim gpLong As Long
'        Dim Index As Long
'        Dim rwLen As Long
'        Dim lenKeyword As Long
'        Dim lenText As Long
'        Dim bWritten As Boolean
'
'        For Index = 11& To UBound(m_Captions)
'            ' convert keyword to bytes
'            lenKeyword = InStr(m_Captions(Index), Chr$(0))
'            lenText = Len(m_Captions(Index)) - lenKeyword
'            txtData() = StrConv(m_Captions(Index), vbFromUnicode)
'
'            ' per PNG specs....
'            ' It is recommended that text items less than 1K (1024 bytes)
'            ' in size should be output using uncompressed text chunks
'            If lenText > 1024& Then
'
'                ' IMPORTANT: This portion of the routine is not equipped to write
'                ' zero-length text block. That is only handled below where the
'                ' .Text length is < 1025... DO NOT modify that IF statement to
'                ' allow zero-length chunks to fall thru to this portion of IF
'
'                ' zTXt chunk format::
'                'Keyword 1-79 bytes (character string)
'                'Null separator 1 byte (null character)
'                'Compression method 1 byte
'                'Compressed text datastream n bytes
'
'                ' Note that the compression byte of zero needs to be included too,
'                ' but we don't add it to the txtData conversion above cause zero
'                ' would be converted to 48 -- Asc("0").
'
'                gpLong = lenText * 0.01 + 12& + lenText
'                '^^ Text won't always compress smaller; it should, but may not
'                ' That is why it is recommended to allow 1024 bytes as uncompressed
'                ReDim pngData(1 To gpLong + (lenKeyword + 5&))
'                ' ^^ include 4 bytes for chunk name + keyword length + 1 byte compression method
'
'                If zDeflate(VarPtr(pngData(6& + lenKeyword)), gpLong, VarPtr(txtData(lenKeyword)), lenText) = True Then
'                    ' ^^ store compression after chunk name, after keyword and after compression method
'                    ' ^^ begin compression on 1st byte of the text, not the caption or compression method
'
'                    CopyMemory pngData(1), chnk_zTXt, 4&
'                    CopyMemory pngData(5), txtData(0), lenKeyword
'                    gpLong = gpLong + lenKeyword + 5&
'
'                    If FileNum = 0& Then    ' writing to array
'                        Index = UBound(Stream) + 1&
'                        ReDim Preserve Stream(0 To Index + gpLong + 7&)
'                        rwLen = iparseReverseLong(gpLong - 4&)
'                        CopyMemory Stream(Index), rwLen, 4&
'                        CopyMemory Stream(Index + 4&), pngData(1), gpLong
'                        rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
'                        CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
'                        Write_zTXt = True
'                    Else
'                        WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'                        If rwLen = 4& Then
'                            WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
'                            If rwLen = gpLong Then
'                                WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
'                                Write_zTXt = (rwLen = 4&)
'                            End If
'                        End If
'                    End If
'                    bWritten = True
'                Else    ' failed to compress. Which means our buffer was too small
'                        ' Therefore we will add it as uncompressed instead of
'                        ' making the buffer even bigger
'                End If
'            End If
'
'            If bWritten Then  'either len<1025 or compression failed
'                bWritten = False    ' reset
'            Else
'                ' tXTt chunk format::
'                'Keyword 1-79 bytes (character string)
'                'Null separator 1 byte (null character)
'                'Text string 0 or more bytes (character string)
'
'                gpLong = lenText + lenKeyword + 4&  ' size of chunk
'                ReDim pngData(1 To gpLong)
'                CopyMemory pngData(1), chnk_tEXt, 4&
'                CopyMemory pngData(5), txtData(0), lenKeyword
'
'                If Not lenText = 0& Then ' zero-length text; not prohibited by PNG specs
'                    CopyMemory pngData(5 + lenKeyword), txtData(lenKeyword), lenText
'                End If
'
'                If FileNum = 0& Then ' writing to array
'                    Index = UBound(Stream) + 1&
'                    ReDim Preserve Stream(0 To Index + gpLong + 7&)
'                    rwLen = iparseReverseLong(gpLong - 4&)
'                    CopyMemory Stream(Index), rwLen, 4&
'                    CopyMemory Stream(Index + 4&), pngData(1), gpLong
'                    rwLen = zCreateCRC(VarPtr(pngData(1)), gpLong)
'                    CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
'                    Write_zTXt = True
'                Else
'                    WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'                    If rwLen = 4& Then
'                        WriteFile FileNum, pngData(1), gpLong, rwLen, ByVal 0&
'                        If rwLen = gpLong Then
'                            WriteFile FileNum, zCreateCRC(VarPtr(pngData(1)), gpLong), 4&, rwLen, ByVal 0&
'                            Write_zTXt = (rwLen = 4&)
'                        End If
'                    End If
'                End If
'            End If
'        Next
'    Else
'        Write_zTXt = True
'    End If
'
'eh:
'    If Err Then Err.Clear
'
'End Function

'Private Function Write_bKGD(ByRef FileNum As Long, ByRef Stream() As Byte) As Boolean
'
'    ' For paletted/grayscale images, this is the palette index, otherwise RGB value
'    On Error GoTo eh
'    Const chnk_bKGD As Long = &H44474B62 'Window Background Color
'
'    If (m_PNGprops And ePngProperties.pngProp_DefaultBkgColor) = ePngProperties.pngProp_DefaultBkgColor Then
'
'        Dim pngData() As Byte
'        Dim gpLong As Long
'        Dim rwLen As Long
'        Dim Index As Long
'
'        ' Per PNG specs, bKGD chunk must come before IDAT and after PLTE
'        Select Case m_ColorType
'        Case clrPalette  ' 1 byte + 4 byte chunk name
'            ReDim pngData(0 To 4)
'            pngData(4) = CByte(m_bKGD)
'
'        Case clrGrayScale, clrGrayAlpha   ' grayscales, 2 bytes + 4 byte chunk name
'            ReDim pngData(0 To 6)
'            pngData(5) = (m_bKGD And &HFF)
'            ' pngData(4) used with 48bit per pixel images (not supported)
'
'        Case Else ' true color, RGB format
'            ReDim pngData(0 To 9)   ' 6 bytes + 4 byte chunk name
'            pngData(5) = m_bKGD And &HFF
'            pngData(7) = (m_bKGD \ &H100&) And &HFF
'            pngData(9) = (m_bKGD \ &H10000) And &HFF
'            ' Note: pngData(4,6,8) used with 48bit per pixel images (not supported)
'        End Select
'
'        CopyMemory pngData(0), chnk_bKGD, 4&
'        gpLong = UBound(pngData) + 1&
'
'        If FileNum = 0& Then 'writing to array
'            Index = UBound(Stream) + 1&
'            ReDim Preserve Stream(0 To Index + gpLong + 7&)
'            rwLen = iparseReverseLong(gpLong - 4&)
'            CopyMemory Stream(Index), rwLen, 4&
'            CopyMemory Stream(Index + 4&), pngData(0), gpLong
'            rwLen = zCreateCRC(VarPtr(pngData(0)), gpLong)
'            CopyMemory Stream(Index + gpLong + 4&), rwLen, 4&
'            Write_bKGD = True
'        Else
'            WriteFile FileNum, iparseReverseLong(gpLong - 4&), 4&, rwLen, ByVal 0&
'            If rwLen = 4& Then
'                WriteFile FileNum, pngData(0), gpLong, rwLen, ByVal 0&
'                If rwLen = gpLong Then
'                    WriteFile FileNum, zCreateCRC(VarPtr(pngData(0)), gpLong), 4&, rwLen, ByVal 0&
'                    Write_bKGD = (rwLen = 4&)
'                End If
'            End If
'        End If
'    Else
'        Write_bKGD = True
'    End If
'eh:
'    If Err Then Err.Clear
'
'End Function




























' =======================================
' FOLLOWING 3 FUNCTIONS ARE ZLIB RELATED
' =======================================

























