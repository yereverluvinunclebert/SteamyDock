Attribute VB_Name = "iconExtract"
'To test this, start a new project, and pass an hIcon handle & target filename to the function.
' Either pass a Picture.Handle (if picture is an icon), or use LoadImage or other APIs to create an hIcon handle.

Option Explicit

Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
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
    bmiColors(1 To 256) As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Function SaveHICONtoArray(ByVal hIcon As Long, outArray() As Byte) As Boolean

    ' Function takes an HICON handle and converts it to 1,4,8,24, or 32 bit icon file format
    ' If return value is False, outArray() contents are undefined
    ' Note: Bit reduction is in play. Example: If original source for HICON was 24 bit
    '   and it can be reduced/saved as 8 bit or lower without color loss, it will.
    ' Note: The end result's quality should be identical to HICON
    '       XP and above required to show/save 32bpp alphablended icons correctly
    '       This routine not coded to save icons in PNG format (Vista and above)

    Dim bits() As Long, pow2(0 To 8) As Long
    Dim tDC As Long, maskScan As Long, clrScan As Long
    Dim X As Long, Y As Long, clrOffset As Long, bNewColor As Boolean
    Dim palIndex As Long, palShift As Long, palPtr As Long, lPrevPal As Long
    Dim ICI As ICONINFO, BHI As BITMAPINFO

    If hIcon = 0& Then Exit Function
    If GetIconInfo(hIcon, ICI) = 0& Then Exit Function

    ' A properly formatted icon file will contain this information:
    ' :: 6 byte ICONDIRECTORY structure
    ' :: 16 byte ICONDIRECTORYENTRY structure
    ' If stored in PNG format then
    '    :: The entire PNG
    ' Else
    '    :: 40 byte BITMAPINFOHEADER structure
    '    If paletted then: Palette entries, each in BGRA format
    '    :: Color data packed & word-aligned per Bitcount of 1,4,8,24,32 bits per pixel
    '    :: 1-bit word-aligned Mask data, even if mask not used (i.e., 32bpp)
    '    Size of any single icon's file can be calculated as:
    '    FileSize = 62 + NrPaletteEntries*4 + (ByteAlignOnWord(BitCount,Width) + ByteAlignOnWord(1,Width))*Height)
    ' Icon sizes are limited to maximum dimensions of 256x256

    On Error GoTo Catch_Exception
    tDC = GetDC(0&)
    With BHI.bmiHeader
        .biSize = 40&
        If ICI.hbmColor = 0& Then  ' black and white icon (rare, but so easy)
            If GetDIBits(tDC, ICI.hbmMask, 0, 0&, ByVal 0&, BHI, 0&) Then
                .biClrUsed = 2&                 ' should be filled in, but ensure it is so
                .biClrImportant = .biClrUsed    ' should be filled in, but ensure it is so
                .biCompression = 0&
                .biSizeImage = 0&
                BHI.bmiColors(2) = vbWhite      ' set 2nd palette entry to white
                ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
                ReDim outArray(0 To ByteAlignOnWord(1, .biWidth) * .biHeight + 69&)
                ' this next call gets the entire icon data; just need to fill in the directory a bit further down this routine
                If GetDIBits(tDC, ICI.hbmMask, 0, BHI.bmiHeader.biHeight, outArray(70), BHI, 0&) = 0& Then
                    .biBitCount = 0
                Else
                    .biClrUsed = 2&                 ' fill in; last GetDIBits call erased it
                    .biClrImportant = .biClrUsed    ' fill in; last GetDIBits call erased it
                    .biHeight = .biHeight \ 2&      ' set to real height, not height*2 as is now
                End If
            End If
            DeleteObject ICI.hbmMask: ICI.hbmMask = 0& ' destroy; no longer needed

        Else    ' color icon vs black & white
            If GetDIBits(tDC, ICI.hbmColor, 0, 0&, ByVal 0&, BHI, 0&) Then
                .biBitCount = 32
                .biCompression = 0&
                .biClrImportant = 0&
                .biClrUsed = 0&
                .biSizeImage = 0&
                ReDim bits(0 To .biWidth * .biHeight - 1&) ' number colors we will process
                If GetDIBits(tDC, ICI.hbmColor, 0, .biHeight, bits(0), BHI, 0&) = 0 Then
                    .biBitCount = 0
                Else
                    ' determine if this icon can be paletted or not; fast routine for small images (256x256 or less)
                    lPrevPal = bits(X) Xor 1&                       ' forces mismatch in loop start
                    For Y = X To .biWidth * .biHeight - 1&          ' process each color
                        If bits(Y) <> lPrevPal Then
                            If (bits(Y) And &HFF000000) Then        ' uses alpha channel; 32bpp
                                .biClrImportant = 0&                ' we can abort loop; won't be paletted
                                .biBitCount = 32
                                Exit For

                            ElseIf .biBitCount = 32 Then                ' continue processing else identified as potential 24bpp icon
                                palIndex = FindColor(BHI.bmiColors(), bits(Y), .biClrImportant, bNewColor) ' have we seen this color?
                                If bNewColor Then                       ' if not, add to our palette
                                    If .biClrImportant = 256& Then      ' max'd out on palette entries; treat as 24bpp
                                        .biBitCount = 24                ' but don't exit loop cause we don't know now
                                        .biClrImportant = 0&            ' if it is not a 32bpp icon

                                    Else                                ' prepare to add to our palette if new
                                        .biClrImportant = .biClrImportant + 1&
                                        If palIndex < .biClrImportant Then  ' keep our palette in ascending order for binary search
                                            CopyMemory BHI.bmiColors(palIndex + 1&), BHI.bmiColors(palIndex), (.biClrImportant - palIndex) * 4&
                                        End If
                                        BHI.bmiColors(palIndex) = bits(Y) ' add color now
                                    End If
                                End If
                            End If
                            lPrevPal = bits(Y) ' track for faster looping
                        End If
                    Next
                    maskScan = ByteAlignOnWord(1, .biWidth) ' scan width for the mask portion of this icon

                    If .biClrImportant Then                ' then can be paletted
                        Select Case .biClrImportant        ' set destination bit count
                            Case Is < 3:    .biBitCount = 1
                            Case Is < 17:   .biBitCount = 4
                            Case Else:      .biBitCount = 8
                        End Select
                        pow2(0) = 1&                                ' setup a power of two lookup table
                        For Y = pow2(0) To .biBitCount
                            pow2(Y) = pow2(Y - 1&) * 2&
                        Next
                        clrScan = ByteAlignOnWord(.biBitCount, .biWidth)    ' scan width of destination's color data
                        .biClrUsed = pow2(.biBitCount)                      ' how many palette entries we will provide
                        .biSizeImage = clrScan * .biHeight                  ' new size of color data
                        clrOffset = .biClrUsed * 4& + 62&                   ' where color data starts
                        ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
                        ReDim outArray(0 To .biSizeImage + maskScan * .biHeight + clrOffset - 1&)

                        lPrevPal = bits(X) Xor 1&                   ' forces mismatch when loop starts
                        For Y = X To .biHeight - 1&                 ' start packing the palette indexes into bytes
                            palShift = 8& - .biBitCount             ' 1st position of byte where palette index will be written
                            palPtr = clrOffset + Y * clrScan        ' position where that byte will start for current row
                            For X = X To X + .biWidth - 1&          ' process each row of the source bitmap
                                ' locate the color in our palette & subtract one (palette is 1-based, indexes are 0-based)
                                If lPrevPal <> bits(X) Then
                                    palIndex = FindColor(BHI.bmiColors(), bits(X), .biClrImportant, bNewColor) - 1&
                                    lPrevPal = bits(X) ' track for faster looping
                                End If
                                outArray(palPtr) = outArray(palPtr) Or (palIndex * pow2(palShift)) ' pack the index
                                If palShift = 0& Then               ' done with this byte
                                    palPtr = palPtr + 1&            ' move destination to next byte
                                    palShift = 8& - .biBitCount     ' reset the position where next index will be written
                                Else
                                    palShift = palShift - .biBitCount ' adjust position where next index will be written
                                End If
                            Next
                        Next

                    Else ' 24 or 32 bit color

                        .biSizeImage = ByteAlignOnWord(.biBitCount, .biWidth) * .biHeight ' size of color data
                        ' size array to the entire icon file format, includes Icon Directory structure, bitmap header
                        ReDim outArray(0 To .biSizeImage + maskScan * .biHeight + 61&)
                        If .biBitCount = 32 Then    ' just copy the entire bitmap to our array
                            CopyMemory outArray(62), bits(X), .biSizeImage
                        Else
                            ' we can loop & transfer 3 of 4 bytes for each pixel or just call the API one more time
                            Call GetDIBits(tDC, ICI.hbmColor, 0&, .biHeight, outArray(62), BHI, 0&)
                        End If
                    End If
                    Erase bits()
                End If
            End If
        End If
    End With

    If BHI.bmiHeader.biBitCount Then
        With BHI.bmiHeader
            ' let's build the icon structure (22 bytes for single icon)
            ' 6 byte ICONDIRECTORY
            '   Integer: Reserved; must be zero
            '   Integer: Type. 1=Icon, 2=Cursor
            '   Integer: Count. Number ico/cur in this resource
            ' 16 BYTE ICONDIRECTORYENTRY
            ' -------- 1 of these for each ico/cur in resource. ICO entry differs from CUR entry
            '   Byte: Width; 256=0
            '   Byte: Height; 256=0
            '   Byte: Color Count; 256=0 & 16-32bit = 0
            '   Byte: Reserved; must be 0
            '   Integer: Planes; must be 1
            '   Integer: Bitcount
            '   Long: Number of bytes for this entry's ico/cur data
            '   Long: Offset into resource where ico/cur data starts
            outArray(2) = 1                                      ' type: icon
            outArray(4) = 1                                      ' count
            If .biWidth < 256& Then outArray(6) = .biWidth       ' width
            If .biHeight < 256& Then outArray(7) = .biHeight     ' height
            If .biClrUsed < 256& Then outArray(8) = .biClrUsed   ' color count
            outArray(10) = 1                                     ' planes
            outArray(12) = .biBitCount                           ' bitcount
            CopyMemory outArray(14), CLng(UBound(outArray) - 21&), 4& ' bytes in resource
            outArray(18) = 22                                    ' offset into directory where BHI starts
            .biHeight = .biHeight + .biHeight                    ' icon's store height*2 in bitmap header
        End With
        ' copy the bitmap header & palette, if used
        CopyMemory outArray(outArray(18)), BHI, BHI.bmiHeader.biClrUsed * 4& + BHI.bmiHeader.biSize

        ' done with the icon directory, now to the mask portion
        If ICI.hbmMask Then
            BHI.bmiColors(1) = vbBlack: BHI.bmiColors(2) = vbWhite      ' set up black/white palette
            With BHI.bmiHeader                                          ' set up bitmapinfo header
                .biBitCount = 1
                .biClrUsed = 2&
                .biClrImportant = .biClrUsed
                .biHeight = .biHeight \ 2&
                .biSizeImage = 0&
                palPtr = UBound(outArray) - maskScan * .biHeight + 1&    ' location where mask will be written
            End With
            GetDIBits tDC, ICI.hbmMask, 0&, BHI.bmiHeader.biHeight, outArray(palPtr), BHI, 0& ' get the mask
        End If
        SaveHICONtoArray = True
    End If

Catch_Exception:
    ReleaseDC 0&, tDC
    If ICI.hbmColor Then DeleteObject ICI.hbmColor
    If ICI.hbmMask Then DeleteObject ICI.hbmMask

End Function

Private Function FindColor(ByRef PaletteItems() As Long, ByVal Color As Long, ByVal Count As Long, ByRef isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations

    ' [in] PaletteItems(). Long Array to search within. Array must be 1-bound
    ' [in] Color. A value to search for. Order is always ascending
    ' [in] Count. Number of items in PaletteItems() to compare against
    ' [out] isNew. If Color not found, isNew is True else False
    ' [out] Return value: The Index where Color was found or where the new Color should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long

    If Count = 0& Then
        FindColor = 1&
        isNew = True
        Exit Function
    End If

    UB = Count
    LB = 1&

    Do Until LB > UB
        newIndex = LB + ((UB - LB) \ 2&)
        Select Case PaletteItems(newIndex) - Color
        Case 0& ' match found
            Exit Do
        Case Is > 0& ' new color is lower in sort order
            UB = newIndex - 1&
        Case Else ' new color is higher in sort order
            LB = newIndex + 1&
        End Select
    Loop

    If LB > UB Then  ' color was not found

        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True

    Else
        isNew = False
    End If

    FindColor = newIndex

End Function

Private Function ByteAlignOnWord(ByVal bitDepth As Long, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

