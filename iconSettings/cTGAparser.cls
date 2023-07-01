VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTGAparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Read/Write TGA image format & Convert to Bitmap ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' TGA Format specifications retrieved from wotsit.org
' additional info from: http://netghost.narod.ru/gff/graphics/summary/tga.htm

    ' TGA is loosely formatted with hardly any restrictions to force conformity
    ' to its structure. Basically, documentation recommends a user identify TGA by
    ' its extension or v2 footer. But v1 of the structure had no footer and those
    ' files are widely available today, still.
    
    ' FYI. The only real difference btwn a normal bitmap and a true color uncompressed tga is
    ' that bitmap scan lines are word aligned & tga are byte aligned and also the the headers:
    ' a bitmap has a 40 byte header & tga has a 44 byte header/footer.
    ' But if tga not true color, then that is where the similarities pretty much end, which is
    ' probably why so many routines out on the net only support true color tgas: image type of 2.
    

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Const INVALID_HANDLE_VALUE = -1

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND ' reusable UDT for 1 & 2 dim arrays
End Type

' TGA header follows (18 byte structure)
' Offset + 0 : 1 byte     ID Length (max of 255 characters that follow the header; can contain anything)
' Offset + 1 : 1 byte     Color Map Type (1=palette included, 0=no palette, other values are reserved)
' Offset + 2 : 1 byte     Image Type (0,1,2,3,9,10,11,32,33). 0=no img, 1&9=paletted, 2&10=true color, 3&11=Grayscale, 32&33 huffman compressed, 9&10&11 are compressed)
' Offset + 3 : 2 bytes    Color Map Index (offset into the palette where 1st palette entry is found)
' Offset + 5 : 2 bytes    Number of Map Entries (number of palette entries)
' Offset + 7 : 1 byte     Entry bit depth (bit depth of each palette entry: 8,15,16,24,32)
' Offset + 8 : 2 bytes    X origin where image should be displayed on screen (rarely used)
' Offset + 10: 2 bytes    Y origin where image should be displayed on screen (rarely used)
' Offset + 12: 2 bytes    Image Width (unsigned)
' Offset + 14: 2 bytes    Image Height (unsigned)
' Offset + 16: 1 byte     Pixel bit depth (any value is valid except zero; this class does 8,15,16,24,32)
' Offset + 17: 1 byte     Image Descriptor (image orientation & alpha usage)
' Total 18 bytes

' See CreateTGAextensionBlock for description of the Extension Block for V2 TGA format (495 blocks)
' The other optional block for TGAs, v2, is the developer block which is not used (variable size & unlimited)

Private pal32LUT(0 To 31) As Byte  ' 5-bit, 32 entry, palette lookup table

Public Function LoadStream(ByRef inStream() As Byte, ByRef cHost As c32bppDIB, _
                    Optional ByVal streamOffset As Long, Optional ByVal streamLength As Long) As Boolean

    ' Function processes an array of bytes suspected of containing a TGA file.
    ' Note: If desired, process filenames by simply opening the file and caching the bytes in an array
    '   and then pass the array to this function, which is what c32bppDIB.LoadPicture_File does.

    ' Parameters:
    ' insSream() :: a byte array containing a TGA
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim bOk As Boolean
    If isTGA(inStream, streamOffset, streamLength, bOk) = False Then Exit Function
    
    ' definitely a tga (v2 footer was included); else potentially a tga (no footer to guarantee)
    
    If bOk Then ' passed key validation checks; assume it is a tga & process it
    
        Select Case inStream(2 + streamOffset) ' check image type
        
        Case 1, 9  ' uncompressed/compressed color map (paletted)
            ProcessColorMap inStream(), streamOffset, streamLength, cHost
        
        Case 2, 3, 10, 11 ' uncompressed/compressed true color / grayscale (easy)
            ProcessTrueColor inStream(), streamOffset, streamLength, cHost
        
        Case 0, 32, 33 ' won't get here because isTGA already ruled it out; just FYI
            ' 0 unsupported. No image contained; probably not a tga
            ' 32 & 33 unsupported. Uses Huffman Delta compression & couldn't find
            '         documentation to support those types and no code exists within to process them
        End Select
        
        LoadStream = (cHost.Handle <> 0&)
        
    End If
    
End Function




Private Function isTGA(ByRef inStream() As Byte, ByRef Offset As Long, ByVal Length As Long, ByRef isValid As Boolean) As Boolean

    ' Function attempts to rule out whether stream contains a TGA file
    ' There are some limitations when the image is paletted:
    ' -- will only support palette indexes of 8 bits, 1 byte
    '    This is not a TGA restriction; rather my own. Until I find or am provided
    '    with sample images other than 1 byte indexes; image is denied

    Dim tData(0 To 23) As Byte
    Dim expectedSize As Long
    Dim Cx As Long
Dim Cy As Long

    Dim tVal As Long
Dim pixBpp As Long

    
    isValid = False
    If Length > 23 Then ' check for v2 TGA footer first
        CopyMemory tData(0), inStream(Offset + Length - 22), 23&
        If InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare) > 0 Then isTGA = True
        ' if the above is true,
        '   :: we will still do some validation before we return.
        '   :: Parser must return True even if image is invalid; this prevents other parsers
        '           from attempting to process the data, when we know 100% it is a tga
    End If
    
    If Length > 18 Then ' possible to have a one pixel B&W image = 19 bytes
        
        ' This will attempt to validate both version 1 & 2 tgas before we actually parse image data
    
        ' v1 has no magic number we can use. We will try to exclude & process if we cannot
    
        ' we will look for unusual values or known values that cannot exist. If any are found,
        ' our routines won't process the file, so other parsers can have a chance if applicable
        
        ' Text 1: Validate image type
        Select Case inStream(Offset + 2)
            Case 1, 2, 3, 9, 10, 11 ' supported
            Case Else ' 0=no image, 32 & 33 may be valid, but no documentation on those types available to me
                Exit Function
        End Select
    
        ' Test 2: Image size
        CopyMemory Cx, inStream(Offset + 12), 2&
        CopyMemory Cy, inStream(Offset + 14), 2&
        ' width/height are unsigned; any negative values would indicate absolutely huge image which
        ' we are not prepared to process or simply that this is not a tga file
        If Cx < 1& Or Cy < 1& Then Exit Function ' invalid tga or not a tga
        
        ' Test 3. Ignore unsupported pixel bit depths
        Select Case inStream(Offset + 16)
            Case 8: pixBpp = 1 ' supported
            Case 15, 16 ' supported if not paletted
                If (inStream(Offset + 2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = 2
            Case 24, 32 ' supported if not paletted
                If (inStream(Offset + 2) And Not 8) = 1 Then Exit Function ' paletted
                pixBpp = (inStream(Offset + 16) \ 8)
            Case Else: Exit Function ' per documentation; other bit depths may be valid, but unsupported here
        End Select
        
        ' Test 4: paletted image validations
        If (inStream(Offset + 2) And Not 8) = 1 Then
            If Not inStream(Offset + 1) = 1 Then Exit Function  ' gotta have a palette if paletted image
        End If
        
        ' Test 5. Validate palette attributes. These can exist in image even if image is not paletted
        If inStream(Offset + 1) = 1 Then
            CopyMemory expectedSize, inStream(Offset + 5), 2&   ' are nr entries valid?
            If expectedSize < 1& Then Exit Function
            CopyMemory tVal, inStream(Offset + 3), 2&           ' is palette offset valid?
            If tVal < 0 Then Exit Function
            expectedSize = expectedSize + tVal                  ' how many indexes are expected
            
            Select Case inStream(Offset + 7)                    ' ignore unsupported palette bit depths
                Case 8, 24, 32 ' supported
                    expectedSize = expectedSize * (inStream(Offset + 7) \ 8)
                Case 15, 16 ' supported
                    expectedSize = expectedSize * 2&
                Case Else: Exit Function ' per documentation; nothing else is valid
            End Select
        End If
        
        ' Test 6: Minimal file size expected
        If (inStream(Offset + 2) And 8) = 0& Then ' uncompressed, else compressed
            expectedSize = expectedSize + (pixBpp * Cx * Cy)
        End If
        If isTGA = True Then expectedSize = expectedSize + 26 ' has footer
        If expectedSize + inStream(Offset) + 18& <= Length Then
            ' file may be much larger, but should be at least the size of expectedSize
            isValid = True
            isTGA = True
        End If
        
    End If

End Function

Private Sub ProcessTrueColor(ByRef inStream() As Byte, ByRef Offset As Long, ByRef streamLength As Long, ByRef cHost As c32bppDIB)

    ' Processes compressed/uncompressed, TrueColor & GrayScale TGA files
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim aUncompressed() As Byte
    Dim x As Long
    Dim y As Long
    Dim dBytes() As Byte
    Dim dSA As SafeArray
    Dim lScanWidth As Long
    Dim rowOffset As Long
    Dim bAlpha As AlphaTypeEnum
    Dim lType As Long
    Dim bPP As Long
    Dim int16 As Integer
    Dim maxWidth As Long
    
    Dim dstStepX As Long
    Dim dstStartX As Long
    Dim dstX As Long
    Dim dstStepY As Long
    Dim dstStartY As Long
    Dim dstStopY As Long

    Select Case inStream(16 + Offset) ' pixel bit/byte depth
        Case 15: bPP = 2
            Create16bitLUT
        Case 16: bPP = 2
            Create16bitLUT
            ' 16bpp can have alpha too; either fully transparent or fully opaque; 1bit value
            ' See if flag is set & verify usage from version 2 tga data if available
            If (inStream(17 + Offset) And 31) > 0& Then
                bAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
            End If
        Case 8, 24, 32: bPP = inStream(16 + Offset) \ 8
    End Select
    
    If inStream(Offset + 1) = 1 Then   ' color map exists before true color data (unusual but not prohibited)
        CopyMemory int16, inStream(Offset + 3), 2&          ' number of palette entries to skip
        If int16 < 0& Then Exit Sub                         ' invalid entry, invlaid TGA
        CopyMemory rowOffset, inStream(Offset + 5), 2&      ' number of color map entries
        Select Case inStream(Offset + 7) ' bit depth per entry
        Case 8, 24, 32
            rowOffset = (rowOffset + int16) * (inStream(Offset + 7) \ 8)
        Case 15, 16
            rowOffset = (rowOffset + int16) * 2
        End Select
        
    End If
    
    CopyMemory x, inStream(12 + Offset), 2&            ' get image width
    CopyMemory y, inStream(14 + Offset), 2&            ' and image height
    
    lScanWidth = bPP * x                               ' source/raw scan width
    rowOffset = 18& + inStream(Offset) + rowOffset ' identify where first image data byte exists
    
    cHost.InitializeDIB x, y                          ' create blank dib to write to
    
    If (inStream(Offset + 17) \ 32 And 1) = 1 Then  ' top down image
        dstStepY = -1&: dstStartY = y + dstStepY: dstStopY = 0&
    Else                                            ' bottom up image (like our dib)
        dstStepY = 1&: dstStartY = 0&: dstStopY = y - dstStepY
    End If
    If (inStream(Offset + 17) \ 64 And 1) = 1 Then  ' right to left
        dstStepX = -4&: dstStartX = x * 4& - 5&
    Else
        dstStepX = 4&: dstStartX = 0&               ' left to right image
    End If
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2, y, cHost.scanWidth, cHost.BitsPointer
    
    If (inStream(Offset + 2) And 8) = 8 Then        ' compressed
        ReDim aUncompressed(0 To x * 4& - 1)       ' 32bpp array
        For y = dstStartY To dstStopY Step dstStepY
            
            If UncompressTrueColor(inStream(), rowOffset, aUncompressed(), bPP, cHost.Width, bAlpha) = False Then
                cHost.DestroyDIB ' corrupt data or not a tga file
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            
            If dstStepX = bPP Then  ' left to right source & 32bpp
                CopyMemory dBytes(dstX, y), aUncompressed(0), lScanWidth
            Else                    ' right to left and/or 24bpp source
                dstX = dstStartX
                For x = 0 To UBound(aUncompressed) Step 4
                    CopyMemory dBytes(dstX, y), aUncompressed(x), 4&
                    dstX = dstX + dstStepX
                Next
            End If
        Next
        Erase aUncompressed()
        
    Else
    
        maxWidth = Offset + streamLength ' prevent writing past allocated dib memory
        ' If isTGA function is working correctly, any uncompressed image that would not
        ' contain enough bytes to fill up the image would be rejected; therefore we do
        ' not slow down the loops checking for potential scanwidth violations.
        
        Select Case bPP
        Case 3, 4 ' 32-24 bits per pixel / 4-3 bytes per pixel
            ' simple & straightforward copy
            If dstStepX = bPP Then      ' faster updating when left to right & 32bpp
                If dstStartY = 0& Then  ' even faster if image is bottom up
                    CopyMemory dBytes(0, 0), inStream(rowOffset), lScanWidth * (dstStopY + 1)
                Else
                    For y = dstStartY To dstStopY Step dstStepY
                        CopyMemory dBytes(dstX, y), inStream(rowOffset), lScanWidth
                        rowOffset = rowOffset + lScanWidth
                    Next
                End If
            Else            ' either top-down or right-left source or 24bpp
                For y = dstStartY To dstStopY Step dstStepY
                    dstX = dstStartX
                    For x = 0 To lScanWidth - 1 Step bPP
                        CopyMemory dBytes(dstX, y), inStream(rowOffset + x), bPP
                        dstX = dstX + dstStepX
                    Next
                    rowOffset = rowOffset + lScanWidth
                Next
            End If
            
        Case 2 ' 15 or 16 bits per pixel / 2 bytes per pixel
            
            ' some bit shifting required
            For y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For x = 0 To lScanWidth - 1 Step bPP
                    ' use a 5,5,5 pattern. Build 15 bits into a 16 bit/2 byte Integer
                    int16 = (inStream(rowOffset + x + 1) And &H7F) * &H100 Or inStream(rowOffset + x)
                    ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                    dBytes(dstX, y) = pal32LUT((int16 And &H1F))
                    dBytes(dstX + 1, y) = pal32LUT(((int16 \ &H20) And &H1F))
                    dBytes(dstX + 2, y) = pal32LUT((int16 \ &H400) And &H1F)
                    If bAlpha Then
                        If (inStream(rowOffset + x + 1) And &H80) = 0 Then dBytes(dstX + 3, y) = &HFF
                    End If
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + lScanWidth
            Next
        
        Case 1    ' 8 bits per pixel / 1 byte per pixel (grayscale)
            ' copy 1 byte at a time, skip alpha (overwrite it to 255 at end)
            For y = dstStartY To dstStopY Step dstStepY
                dstX = dstStartX
                For x = 0 To lScanWidth - 1
                    dBytes(dstX, y) = inStream(rowOffset + x)
                    dBytes(dstX + 1, y) = dBytes(dstX, y)
                    dBytes(dstX + 2, y) = dBytes(dstX, y)
                    dstX = dstX + dstStepX
                Next
                rowOffset = rowOffset + lScanWidth
            Next
        End Select
    
    End If
    
    iparseValidateAlphaChannel dBytes, True, bAlpha, lType
    iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0

    cHost.Alpha = bAlpha
    cHost.ImageType = imgTGA

End Sub

Private Sub ProcessColorMap(ByRef inStream() As Byte, ByRef Offset As Long, ByRef streamLength As Long, ByRef cHost As c32bppDIB)

    ' Processes compressed/uncompressed, paletted TGA images
    ' Note: the isTGA function already validated key portions of the header; so this routine does no
    ' further validation, other than to ensure compressed images uncompress correctly

    Dim pal() As Byte
    Dim aUncompressed() As Byte
    
    Dim x As Long
    Dim y As Long
    
    Dim lScanWidth As Long
    Dim rowOffset As Long
    
    Dim bAlpha As AlphaTypeEnum
    Dim bPalAlpha As Boolean
    Dim lType As Long
    
    Dim bPP As Long
    Dim palBpp As Long
    
    Dim mapSize As Long
    Dim int16 As Integer
    
    Dim dBytes() As Byte
    Dim dSA As SafeArray
    
    
    Dim dstStepX As Long
    Dim dstStartX As Long
    Dim dstX As Long
    
    Dim dstStepY As Long
    Dim dstStartY As Long
    Dim dstStopY As Long
    
    Dim maxWidth As Long
    
    'Select Case inStream(16 + Offset) ' source pixel bit/byte depth
    '    Case 8: bPP = 1
    'End Select
    ' currently, only 8bit, 1byte palette indexes are supported
    bPP = 1
    
    CopyMemory mapSize, inStream(Offset + 5), 2&       ' number of palette entries
    Select Case inStream(Offset + 7)                   ' bit depth per entry
    Case 8: palBpp = 1
    Case 15: palBpp = 2
        Create16bitLUT
    Case 16: palBpp = 2
        Create16bitLUT
        ' It is possible for the palette to have alpha and the pixels to also have alpha
        ' If pixels have alpha, then their alpha overrides any palette alpha
        bPalAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
    Case 24: palBpp = 3
    Case 32: palBpp = 4
        bPalAlpha = AllowAlphaChannel(inStream(), Offset, streamLength)
    End Select
    
    CopyMemory int16, inStream(Offset + 3), 2& ' number of palette entries to skip
    rowOffset = (mapSize + int16) * palBpp + 18& + inStream(Offset) ' this is where the palette indexes begin
    
    CopyMemory x, inStream(12 + Offset), 2&            ' get image width
    CopyMemory y, inStream(14 + Offset), 2&            ' and image height
    
    If mapSize < 256 Then
        ReDim pal(0 To 1023)    ' oversize palette to prevent any errors if palette index exceeds palette
    Else
        ReDim pal(0 To mapSize * 4 - 1)           ' use actual size of palette
    End If
    cHost.InitializeDIB x, y                      ' create blank dib to write to
    
    ' determine where in the source palette to start reading entries from
    x = 0&: y = rowOffset - (mapSize * palBpp)
    For y = y To rowOffset - 1 Step palBpp
        Select Case palBpp
        Case 1:
            pal(x) = inStream(y)
            pal(x + 1) = inStream(y)
            pal(x + 2) = inStream(y)
        Case 2
            int16 = (inStream(y + 1) And &H7F) * &H100 Or inStream(y)
            ' shift the 5,5,5 and ref a 5 bit/32 entry palette
            pal(x) = pal32LUT(int16 And &H1F)
            pal(x + 1) = pal32LUT((int16 \ &H20) And &H1F)
            pal(x + 2) = pal32LUT((int16 \ &H400) And &H1F)
            If bPalAlpha Then
                If (inStream(y + 1) And &H80) = 0& Then pal(x + 3) = &HFF
            End If
        Case 3, 4
            CopyMemory pal(x), inStream(y), 3
            If bPalAlpha Then pal(x + 3) = inStream(y + palBpp - 1) ' 32bpp palettes
        End Select
        x = x + 4
    Next
    
    If (inStream(Offset + 17) \ 32 And 1) = 1 Then  ' top down image
        dstStepY = -1&: dstStartY = cHost.Height - 1&: dstStopY = 0&
    Else                                            ' bottom up (like our dib)
        dstStepY = 1&: dstStartY = 0&: dstStopY = cHost.Height - 1&
    End If
    If (inStream(Offset + 17) \ 64 And 1) = 1 Then  ' right to left
        dstStepX = -4&: dstStartX = cHost.Width * 4& - 5&
    Else
        dstStepX = 4&: dstStartX = 0&               ' left to right
    End If
    
    iparseOverlayHost_Byte dBytes(), VarPtr(dSA), 2, cHost.Height, cHost.scanWidth, cHost.BitsPointer
    
    If (inStream(Offset + 2) And 8) = 8 Then        ' compressed
    
        lScanWidth = cHost.scanWidth                ' dib scan width
        ReDim aUncompressed(0 To lScanWidth - 1) ' 32bpp format
        For y = dstStartY To dstStopY Step dstStepY
            
            If UncompressColorMap(inStream(), rowOffset, aUncompressed(), bPP, cHost.Width, pal()) = False Then
                cHost.DestroyDIB    ' corrupt data or not a tga
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            If dstStepX = 4& Then  ' left to right image, copy entire scan line in one call
                CopyMemory dBytes(dstX, y), aUncompressed(0), lScanWidth
            Else                    ' right to left
                dstX = dstStartX
                For x = 0 To UBound(aUncompressed) Step 4
                    CopyMemory dBytes(dstX, y), aUncompressed(x), 4
                    dstX = dstX + dstStepX
                Next
            End If
        Next
        Erase aUncompressed()
    
    Else
    
        maxWidth = Offset + streamLength ' prevent writing past allocated dib memory
        lScanWidth = bPP * cHost.Width  ' source/raw scan width
    
        ' 8 bits per pixel / 1 byte per pixel
        ' copy 1 byte at a time, skip alpha (overwrite it to 255 at end)
        For y = dstStartY To dstStopY Step dstStepY
            If lScanWidth + rowOffset > maxWidth Then
                ' safety check that should never fire if isTGA function is working correctly
                cHost.DestroyDIB ' corrupt data or not a tga file
                iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0
                Exit Sub
            End If
            dstX = dstStartX
            For x = 0 To lScanWidth - 1
                CopyMemory dBytes(dstX, y), pal(inStream(rowOffset + x) * 4&), 4&
                dstX = dstX + dstStepX
            Next
            rowOffset = rowOffset + lScanWidth
        Next
        
    End If
    
    iparseValidateAlphaChannel dBytes, True, bAlpha, lType
    iparseOverlayHost_Byte dBytes(), 0, 0, 0, 0, 0

    cHost.Alpha = bAlpha
    cHost.ImageType = imgTGA

End Sub

Private Function UncompressTrueColor(ByRef inStream() As Byte, ByRef Offset As Long, ByRef bScan() As Byte, ByRef bPP As Long, _
                                    ByVal Width As Long, ByRef AlphaType As AlphaTypeEnum) As Boolean

    ' RLE uncompression of TGA images: TrueColor only

    Dim pixCount As Long
Dim pixLoc As Long

    Dim int16 As Integer
Dim bReturn As Boolean

    Dim tileStop As Long
Dim tileCx As Long

    Dim x As Long
Dim lScanWidth As Long

    
    bReturn = True
    lScanWidth = Width * 4&
    Do Until pixCount = Width
        If (inStream(Offset) And 128) = 0 Then  ' this set of pixels not RLE compressed
            ' when not compressed, the control byte says how many non-compressed pixels to process
            
            If bPP * inStream(Offset) + pixLoc > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            End If
            Offset = Offset + 1                 ' move to next byte
            
            Select Case bPP
            Case 4                  ' 32bpp; straight copy
                x = inStream(Offset - 1) + 1
                CopyMemory bScan(pixLoc), inStream(Offset), x * bPP
                pixLoc = pixLoc + x * bPP
                Offset = Offset + x * bPP
            Case 3                  ' 24 bpp
                For x = 0 To inStream(Offset - 1)
                    CopyMemory bScan(pixLoc), inStream(Offset), bPP ' copy 3 bytes
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move 3 bytes
                Next                            ' already positioned on next RLE control byte
            Case 2                  ' 16 bpp
                For x = 0 To (inStream(Offset - 1)) ' create 15bit integer & shift into palette
                    int16 = (inStream(Offset + 1) And &H7F) * &H100 Or inStream(Offset)
                    bScan(pixLoc) = pal32LUT((int16 And &H1F))
                    bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                    bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                    If AlphaType > AlphaNone Then
                        If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                    End If
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move 2 bytes
                Next                            ' already positioned on next RLE control byte
            Case 1                  ' 8 bpp/grayscale
                For x = 0 To (inStream(Offset - 1))
                    bScan(pixLoc) = inStream(Offset)    ' add grayscale values
                    bScan(pixLoc + 1) = bScan(pixLoc)
                    bScan(pixLoc + 2) = bScan(pixLoc)
                    pixLoc = pixLoc + 4
                    Offset = Offset + bPP       ' move to next byte
                Next                            ' already positioned on next RLE control byte
            End Select
            
            pixCount = pixCount + x
            
        Else                                    ' RLE compressed. TGA does RLE runs on pixels; not bits/bytes
            
            Select Case bPP
            Case 3, 4                           ' 24/32bpp
                CopyMemory bScan(pixLoc), inStream(Offset + 1), bPP ' get color from next bytes
            Case 2
                int16 = (inStream(Offset + 2) And &H7F) * &H100 Or inStream(Offset + 1)
                ' shift the 5,5,5 and ref a 5 bit/32 entry palette
                bScan(pixLoc) = pal32LUT((int16 And &H1F))
                bScan(pixLoc + 1) = pal32LUT(((int16 \ &H20) And &H1F))
                bScan(pixLoc + 2) = pal32LUT((int16 \ &H400) And &H1F)
                If AlphaType > AlphaNone Then
                    If (inStream(Offset + 1) And &H80) = 0 Then bScan(pixLoc + 3) = &HFF
                End If
            Case 1
                bScan(pixLoc) = inStream(Offset + 1)
                bScan(pixLoc + 1) = bScan(pixLoc)
                bScan(pixLoc + 2) = bScan(pixLoc)
            End Select
            
            ' here we use tiling techniques to quickly fill a scan line with repeated pixels
            tileCx = 4                              ' how many pixels copied so far
            tileStop = ((inStream(Offset) And Not 128) + 1) * tileCx ' how many pixels will be duplicated
            If tileStop + pixLoc > lScanWidth Then
                bReturn = False             ' corrupt data; copying will exceed scanwidth
                Exit Do
            End If
            
            Do Until tileCx + tileCx > tileStop     ' copy another group of pixels
                CopyMemory bScan(pixLoc + tileCx), bScan(pixLoc), tileCx
                tileCx = tileCx * 2                 ' increment group by x2
            Loop
            tileCx = tileStop - tileCx              ' see if any partial tiling is needed
            If tileCx Then                          ' and copy required remaining pixels
                CopyMemory bScan(pixLoc + tileStop - tileCx), bScan(pixLoc), tileCx
            End If
            pixLoc = pixLoc + tileStop      ' move bScan() pointer to next byte to be filled
            Offset = Offset + bPP + 1       ' move to next RLE control byte
            pixCount = pixCount + tileStop \ 4 ' number of pixels processed during decompression
        End If
        
    Loop
    UncompressTrueColor = bReturn

End Function

Private Function UncompressColorMap(ByRef inStream() As Byte, ByRef Offset As Long, ByRef bScan() As Byte, ByVal bPP As Long, _
                                    ByVal Width As Long, ByRef Palette() As Byte) As Boolean

    ' RLE uncompression of TGA images: Paletted TGA only
    
    Dim pixCount As Long
Dim pixLoc As Long

    'Dim int16 As Integer
Dim bReturn As Boolean

    Dim tileStop As Long
Dim tileCx As Long

    Dim x As Long
Dim lScanWidth As Long

    Dim maxSourceCount As Long
    
    bReturn = True
    lScanWidth = Width * 4&
    maxSourceCount = UBound(bScan)
    Do Until pixCount = Width
        If (inStream(Offset) And 128) = 0 Then
            ' not RLE encoded
            If bPP * inStream(Offset) + pixLoc > lScanWidth Then
                bReturn = False ' corrupt data or not tga; will exceed scanwidth
                Exit Do
            ElseIf maxSourceCount < pixLoc + inStream(Offset) * 4& Then
                bReturn = False ' corrupt data or not tga; will exceed destination size
                Exit Do
            End If
            
            Offset = Offset + 1
            For x = 0 To inStream(Offset - 1)
                CopyMemory bScan(pixLoc), Palette(inStream(Offset) * 4), 4
                pixLoc = pixLoc + 4
                Offset = Offset + bPP
            Next
            pixCount = pixCount + x
            
        Else
            
            CopyMemory bScan(pixLoc), Palette(inStream(Offset + 1) * 4), 4
            
            tileCx = 4&                              ' how many pixels copied so far
            tileStop = ((inStream(Offset) And Not 128) + 1) * tileCx ' how many bytes will be duplicated
            If tileStop + pixLoc > lScanWidth Then
                bReturn = False     ' corrupt data or not a tga; will exceed scanwidth
                Exit Do
            ElseIf maxSourceCount < tileStop + pixLoc Then
                bReturn = False     ' corrupt data or not tga; will exceed destination size
                Exit Do
            End If
            
            Do Until tileCx + tileCx > tileStop     ' copy another group of pixels
                CopyMemory bScan(pixLoc + tileCx), bScan(pixLoc), tileCx
                tileCx = tileCx + tileCx            ' increment group by x2
            Loop
            tileCx = tileStop - tileCx              ' see if any partial tiling is needed
            If tileCx Then                          ' and copy required remaining pixels
                CopyMemory bScan(pixLoc + tileStop - tileCx), bScan(pixLoc), tileCx
            End If
            pixLoc = pixLoc + tileStop      ' move bScan() pointer to next byte to be filled
            Offset = Offset + bPP + 1       ' move to next RLE control byte
            pixCount = pixCount + tileStop \ 4 ' number of pixels processed during decompression
        
        End If
    
    Loop
    UncompressColorMap = bReturn
    
End Function

Private Sub Create16bitLUT()
    ' Initialize a 32 entry palette
    Dim x As Byte
    If pal32LUT(31) = 0 Then ' not yet created
        For x = 0 To 31
            pal32LUT(x) = x * 8 + (x Mod 8)
        Next
    End If
    
End Sub









Private Function AllowAlphaChannel(ByRef theStream() As Byte, ByVal Offset As Long, ByVal Length As Long) As Boolean

    ' Helper Function. Called when 16bpp images/palettes contain an alpha bit and for 32bpp palettes.
    
    ' This function will check if version 2 data exists and if so whether or not v2 says
    ' to ignore the alpha channel. Only if v2 exists and ignores the alpha will this
    ' function return False. 32bpp true color images are handled via modParsers.iparseValidateAlphaChannel
    
    Dim tData(0 To 29) As Byte
    Dim tVal As Long
Dim appBlockLoc As Long

    Dim bReturn As Boolean
    
    Const v2BlockLength As Long = 495 ' v2 application/extension block length
    
    bReturn = True      ' default - don't override & allow any alpha
    If Length > 43& Then ' check for v2 TGA footer first
        CopyMemory tData(0), theStream(Offset + Length - 29&), 30&
        tVal = InStr(1, StrConv(tData(), vbUnicode), "TRUEVISION-XFILE.", vbTextCompare)
        If tVal > 8& Then
            ' ok, does it have an application/extension block? 8 bytes before TRUEVISION
            CopyMemory appBlockLoc, tData(tVal - 9&), 4&
            ' ensure it is not corrupt; must be a valid pointer & must be within the stream/file
            If appBlockLoc > 0& And appBlockLoc + Offset + v2BlockLength < Length Then
                ' ok, it appears valid, double check
                CopyMemory tVal, theStream(Offset + appBlockLoc), 2&
                If tVal = v2BlockLength Then ' we have v2 app data; expected
                    ' fine, let's get the overriding value
                    bReturn = (theStream(Offset + appBlockLoc + v2BlockLength - 1&) > 2)
                    ' If the byte = 0, 1 or 2; then alpha is not to be used
                End If
            End If
        End If
    Else
        
    End If
    AllowAlphaChannel = bReturn

End Function








