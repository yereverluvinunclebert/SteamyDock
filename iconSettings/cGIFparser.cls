VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGIFparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  -----======== PURPOSE: Convert GIF to Bitmap and Bitmap to GIF ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs, or different versions of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

' used to extract data from a converted GIF
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long


'Private Declare Function CreatePalette Lib "gdi32.dll" (ByRef lpLogPalette As Any) As Long
'Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

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
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type tNode                              'Octree node struct
    bIsLeaf As Boolean                          'Leaf flag
    bAddedReduce As Boolean                     'Linked list flag
    vRed As Long                                'Red Value
    vGreen As Long                              'Green Value
    vBlue As Long                               'Blue Value
    cClrs As Long                               'Pixel count
    iChildren(0 To 1, 0 To 1, 0 To 1) As Long   'Child pointers
    iNext As Long                               'Next reducable node
End Type
'Private OctreeNodes() As tNode
'Private OctNodeCount As Long

'Private OctDepth As Byte
'Private OctTopGarbage As Long

'Private OctColor As Long
'Private OctReduction() As Long

'Private c_Normalized32bpp() As Long

' variables used during GIF compression
'Private c_aMask() As Long           ' Power of 2 look up table
'Private c_aBuff() As Byte           ' general use byte array
'Private c_BitBucketCount As Long    ' bit bucket bit counter
'Private c_BitBucketBuff As Long     ' bit bucket contents
'Private c_lCurrentBits As Long      ' current lzw compression size (variable)
'Private c_lSubBlockSize As Long     ' flag to track last byte in a data subBlock
'Private c_Palette() As Long         ' GIF palette working array
'Private c_hPalette As Long          ' windows palette

Private c_GIFdata() As Byte     ' source bytes (mapped array, never initialized)
Private c_GIFframe() As Byte    ' 1st GIF frame from source bytes

Public Function LoadStream(ByRef inStream() As Byte, ByRef cHost As c32bppDIB, _
                    Optional ByVal streamOffset As Long, Optional ByVal streamLength As Long) As Boolean

    ' Parameters:
    ' insSream() :: a byte array containing a GIF
    ' cHost :: an initialized c32bppDIB
    ' streamOffset :: array position for 1st byte in the stream
    ' streamLength :: size of stream that contains the image
    
    ' IMPORTANT: the array offset & length are not checked in this class.
    '   They were checked before this class was called. If this class is to
    '   be pulled out and put in another project, ensure you include the
    '   validation shown in c32bppDIB.LoadPicture_Stream
    
    Dim tTSA As SafeArray
    ' overlay our module level array onto the passed array
    With tTSA
        .cbElements = 1         ' byte array
        .cDims = 1              ' 1 dimensional
        .pvData = VarPtr(inStream(streamOffset))
        .rgSABound(0).cElements = streamLength
    End With
    CopyMemory ByVal VarPtrArray(c_GIFdata), VarPtr(tTSA), 4& ' apply overlay
    
    ' call routine to parse the GIF & convert it to 32bpp
    LoadStream = ParseGIF(cHost)
    CopyMemory ByVal VarPtrArray(c_GIFdata), 0&, 4&    ' remove overlay
    
End Function

Private Function ParseGIF(ByRef cHost As c32bppDIB) As Boolean

    On Error Resume Next
    ' a modified routine from some of my other GIF postings
    ' This version is scaled back and only extracts first frame
    
    ' This routine has one limitation. Some rare GIFs do not follow the
    ' standards and when those are encountered, the routine will return
    ' True to prevent GIF from being sent to other parsers.  This is
    ' important because the BMP, the last parser, sends the stream to an API
    ' to convert an unknown image to a stdPicture. If the GIF stream
    ' isn't formatted within standards that API hangs the application.
    
    Dim gLong As Long
    Dim aPointer As Long
    Dim gHeaderLen As Long
    Dim g87aStart As Long
Dim g87aStop As Long

    Dim g89aStart As Long
'Dim g89aStop As Long

    
    ' transparency flags and variables use to tweak GIF
    Dim transUsed As Byte
Dim TransIndex As Long

    Dim aLocalTbl As Long
Dim gColorsUsed As Long

    Dim uniquePalette(0 To 767) As Byte
    Dim p As Long
    
    On Error GoTo ExitReadRoutine
    
    ' read signature
    ReDim c_GIFframe(0 To 5)
    CopyMemory c_GIFframe(0), c_GIFdata(0), 6&
    Select Case LCase$(StrConv(c_GIFframe, vbUnicode))
        Case "gif89a", "gif87a"
        Case Else
            Exit Function
    End Select
        
    ' skip to the global color table information
    If (c_GIFdata(10) And 128) = 128 Then ' color table used? If so, skip it
        gColorsUsed = 2& ^ (1& + (c_GIFdata(10) And &H7)) ' count colors
        gHeaderLen = gColorsUsed * 3& + 13&
    Else 'no global color table; probably uses local color tables
        gHeaderLen = 13&
    End If
    aPointer = gHeaderLen
    
    Do
        Select Case c_GIFdata(aPointer)    ' read a single byte
        Case 0  ' block terminators
            aPointer = aPointer + 1&
            
        Case 33 'Extension Introducer
            aPointer = aPointer + 1&
            
            Select Case c_GIFdata(aPointer) ' read the extension type
            
            Case 255    ' application extension
                ' Get the length of extension: will always be 11
                aPointer = aPointer + c_GIFdata(aPointer + 1&) + 2&
                Call SkipGifBlock(aPointer)
                
            Case 249    ' Graphic Control Label
                        ' (description of frame & is an optional block) 8 bytes
                transUsed = (c_GIFdata(aPointer + 2&) And 1&)
                If transUsed = 1& Then ' has transparency?
                    TransIndex = c_GIFdata(aPointer + 5&) ' cache transparency index
                End If
                g89aStart = aPointer - 1&    ' location where 89a block starts
                aPointer = aPointer + 7&     ' move to end of block
                
            Case Else   ' Comment block, plain text extension, or Unknown extension
                aPointer = aPointer + 1&
                Call SkipGifBlock(aPointer)
            End Select
                
        Case 44 ' Image Descriptor (image dimensions & color table)
                ' mark position where image description starts
            g87aStart = aPointer
            aPointer = aPointer + 9& ' image data starts 10 bytes after header
            ' next byte indicates if local color table used
            If (c_GIFdata(aPointer) And 128) = 128 Then   ' local color table used?
                gColorsUsed = 2& ^ (1& + (c_GIFdata(aPointer) And &H7)) ' count colors
                aPointer = aPointer + gColorsUsed * 3&
                aLocalTbl = 1&  ' flag indicating colors from local table vs global table
            End If
            aPointer = aPointer + 2& ' move to position of first data block
            Call SkipGifBlock(aPointer)
                
            g87aStop = aPointer - 1&    ' this is where the data ends
            If g87aStop - g87aStart < 3& Then Exit Function ' invalid frame
            Exit Do
            
        Case Else
            ' shouldn't happen; abort with what we have
            Exit Function
        End Select
    Loop
    
    If Not (g87aStart = 0& Or gColorsUsed = 0&) Then ' we have a valid gif frame
    
        ' rebuild the GIF file to include only the 1st frame read
        If g89aStart > 0 Then   ' gif is 89a format
            ' resize array, copy header info & gif89a info
            ReDim c_GIFframe(0 To gHeaderLen + (g87aStop - g87aStart + 1&) + 8&)
            CopyMemory c_GIFframe(0), c_GIFdata(0), gHeaderLen
            CopyMemory c_GIFframe(gHeaderLen), c_GIFdata(g89aStart), 8&
            aPointer = gHeaderLen + 8&  ' adjust pointer for gif87a info
        Else
            ' resize array and copy header info only
            ReDim c_GIFframe(0 To gHeaderLen + (g87aStop - g87aStart + 1&))
            CopyMemory c_GIFframe(0), c_GIFdata(0), gHeaderLen
            aPointer = gHeaderLen
        End If
        ' now copy the gif 87a info
        CopyMemory c_GIFframe(aPointer), c_GIFdata(g87aStart), g87aStop - g87aStart + 1&
        c_GIFframe(UBound(c_GIFframe)) = 59 ' trailer/end of file
    
        ' fix up the left/top & width/height of overall frame
        CopyMemory c_GIFframe(aPointer + 1), 0&, 4& ' make frame left/top zero
        CopyMemory c_GIFframe(6), c_GIFframe(aPointer + 5), 4& ' make window & frame size same
    
        If transUsed = 1& Then
            ' Fix up the color table/indexes for images with transparency
            ' Why? Instead of parsing/decompressing the GIF manually, we will allow an API to do it for us.
            ' But that API can re-index the GIF which means we may lose track of the transparency
            ' color/index.  This happens whenever any color in the GIF's palette is duplicated.
            ' To prevent this from occuring, we simply replace the GIF's palette with another
            ' palette of non-duplicated entries.
            ' BTW: This way of creating GIF is still much faster than parsing the GIF by hand
            If aLocalTbl = 1& Then ' local color table else global
                ' local color table starts 10 bytes after the gif87a block
                aPointer = gHeaderLen + 10&  ' location of table within single frame array
                aLocalTbl = g87aStart + 10&  ' location of table within souce array
                ' offset single frame array when gif89a structure is used
                If Not g89aStart = 0& Then aPointer = aPointer + 8&
            Else
                aPointer = 13&        ' global table location
                aLocalTbl = 13&       ' same in both arrays
            End If
            For p = 1& To gColorsUsed - 1&  ' create non-duplicating color palette
                gLong = p * 3&
                uniquePalette(gLong) = p
                uniquePalette(gLong + 1) = p
                uniquePalette(gLong + 2) = p
            Next
            ' replace the old palette with the new one
            CopyMemory c_GIFframe(aPointer), uniquePalette(0), gColorsUsed * 3&
            Erase uniquePalette()
        Else
            TransIndex = -1&
        End If
        
        ' all done parsing the GIF file, send it to routine to convert it to a 32bpp
        ParseGIF = ConvertGIFto32bpp(TransIndex, aLocalTbl, cHost)
        
    End If
    
ExitReadRoutine:
If Err Then
    Err.Clear           ' this is a GIF format, but the format is invalid
    cHost.DestroyDIB    ' something is wrong; don't allow it to continue
    ParseGIF = True     ' to other parsers
End If
End Function

Private Sub SkipGifBlock(ByRef Ptr As Long)
    ' Routine skips a block of data within the GIF file
    Dim curByte As Byte
    curByte = c_GIFdata(Ptr)
    Do While Not curByte = 0
        Ptr = Ptr + 1& + curByte
        curByte = c_GIFdata(Ptr)
    Loop
    Ptr = Ptr + 1&
End Sub

Private Function ConvertGIFto32bpp(ByVal TransIndex As Long, ByVal tblOffset As Long, ByRef cHost As c32bppDIB) As Boolean

    ' Function converts GIF to a standard picture and then premultiplies RGB values based on the
    ' GIFs transparent index, if applicable.
    ' Note: The c_GIFframe array was already processed/filled in the ParseGIF function
    
    Dim tPic As StdPicture
Dim tBMP As BITMAP

    
    ' used for parsing a transparent gif
    Dim X As Long
Dim Y As Long
Dim m As Long
Dim dX As Long
Dim Index As Long

    Dim gSA As SafeArray
Dim dSA As SafeArray

    Dim Pow2(0 To 8) As Long
Dim dibBytes() As Byte

    Dim maskShift As Long
Dim maskAND As Long

    Dim hostDC As Long
    
    ' first: have API create a stdPicture for us
    Set tPic = iparseArrayToPicture(c_GIFframe, 0&, 1& + UBound(c_GIFframe))
    Erase c_GIFframe
    If Not tPic Is Nothing Then
        
        ' a VB stdPicture is a DIB, therefore it has a handle to the DIB bits; get it
        GetGDIObject tPic.Handle, Len(tBMP), tBMP
        If Not tBMP.bmBits = 0& Then
        
            ' have host create application's 32bpp DIB
            cHost.InitializeDIB tBMP.bmWidth, tBMP.bmHeight
            
            ' we only need to parse the palette & indexes if transparency is used
            If TransIndex = -1& Then                ' opaque GIF
                ' render GIF to our DIB DC, then ensure all alpha bytes are 255
                hostDC = cHost.LoadDIBinDC(True)
                tPic.Render hostDC + 0&, 0&, 0&, tBMP.bmWidth + 0&, tBMP.bmHeight + 0&, _
                    0&, tPic.Height, tPic.Width, -tPic.Height, ByVal 0&
                cHost.LoadDIBinDC False
                With dSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = cHost.BitsPointer
                    .rgSABound(0).cElements = cHost.Height
                    .rgSABound(1).cElements = cHost.scanWidth
                End With
                CopyMemory ByVal VarPtrArray(dibBytes), VarPtr(dSA), 4&
                iparseValidateAlphaChannel dibBytes(), True, 0&, -1&
                CopyMemory ByVal VarPtrArray(dibBytes), 0&, 4&
                cHost.Alpha = AlphaNone
                
            Else
                ' next: getting ready to parse the paletted stdPic
                Pow2(0) = 1&
                For X = 1& To tBMP.bmBitsPixel   ' power of 2 array
                    Pow2(X) = Pow2(X - 1&) * 2&
                Next
                maskAND = Pow2(tBMP.bmBitsPixel) - 1& ' AND mask for stdPic indexes
                ' we need to overlay arrays onto the GIF and the host's DIB pointers
                With gSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = tBMP.bmBits
                    .rgSABound(0).cElements = tBMP.bmHeight
                    .rgSABound(1).cElements = iparseByteAlignOnWord(tBMP.bmBitsPixel, tBMP.bmWidth)
                End With
                With dSA
                    .cbElements = 1
                    .cDims = 2
                    .pvData = cHost.BitsPointer
                    .rgSABound(0).cElements = cHost.Height
                    .rgSABound(1).cElements = cHost.scanWidth
                End With
                CopyMemory ByVal VarPtrArray(c_GIFframe), VarPtr(gSA), 4&
                CopyMemory ByVal VarPtrArray(dibBytes), VarPtr(dSA), 4&
                
                ' last: start parsing stdPic's paletted DIB
                For Y = 0& To tBMP.bmHeight - 1&
                    dX = 0&: m = 0&   ' reset dX=host DIB's X & M=stdPic DIB's X
                    maskShift = 8& - tBMP.bmBitsPixel    ' 1st bit to process
                    
                    ' note: do not loop thru using gif ScanWidth. If the GIF
                    ' width is not DWORD ligned , you will overflow the target
                    ' DIB width and eventually write to uninitialized memory
                    For X = 1& To tBMP.bmWidth&
                        ' get the palette index by shifting bits
                        Index = ((c_GIFframe(m, Y) \ Pow2(maskShift)) And maskAND)
                        
                        If Not Index = TransIndex Then  ' 100% opaque else 100% transparent
                            Index = Index * 3& + tblOffset
                            dibBytes(dX, Y) = c_GIFdata(Index + 2&)     ' make BGR vs RGB
                            dibBytes(dX + 1, Y) = c_GIFdata(Index + 1&)
                            dibBytes(dX + 2, Y) = c_GIFdata(Index)
                            dibBytes(dX + 3, Y) = 255
                        End If
                        
                        ' adjust for parsing/shifting the next index
                        If maskShift = 0& Then
                            maskShift = 8& - tBMP.bmBitsPixel ' start new byte
                            m = m + 1&                        ' next GIF byte
                        Else
                            maskShift = maskShift - tBMP.bmBitsPixel ' adjust
                        End If
                        dX = dX + 4&                          ' next Host pixel
                    Next
                Next
                ' done, remove overlays
                CopyMemory ByVal VarPtrArray(c_GIFframe), 0&, 4&
                CopyMemory ByVal VarPtrArray(dibBytes), 0&, 4&
                cHost.Alpha = AlphaSimple
            End If
            cHost.ImageType = imgGIF
            ConvertGIFto32bpp = True
        End If
    End If

End Function








' THE PALETTE ROUTINES BELOW AND THIS ROUTINE ARE FROM THE FOLLOWING SOURCE.
' I have modified each slightly, so get the original routines from that source...
' source: http://www.vbaccelerator.com/home/VB/Code/vbMedia/Image_Processing/Colour_Depth_Reduction/VB6_Colour_Depth_Sample.asp



























