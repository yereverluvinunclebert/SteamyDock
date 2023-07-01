Attribute VB_Name = "modParsers"
Option Explicit

' Module contains functions that are required by two or more classes.

' No APIs are declared public. This is to prevent possibly, differently
' declared APIs or different versions, of the same API, from conflicting
' with any APIs you declared in your project. Same rule for UDTs.

' Though many of these routines are made Public so that the classes can use them,
' you should not call these routines from your own project. Those that you may
' wish to call anyway, ensure you pass valid, expected parameters. Within the
' classes, parameters are validated and these routines may not have additional
' validation checks which could result in crashes or memory leaks if used incorrectly.

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND ' 32 bytes as used. Can be used for 1D and/or 2D arrays
End Type
Private Type PictDesc
    Size As Long
    TypeIt As Long
    hHandle As Long
    hPal As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (ByRef pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As Any, ByRef ppvObj As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (ByRef lpPictDesc As PictDesc, ByRef riid As Any, ByVal fPictureOwnsHandle As Long, ByRef iPic As IPicture) As Long

' used to see if DLL exported function exists
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hmodule As Long, ByVal lpProcName As String) As Long

' GDI32 APIs
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (ByRef lpXform As Any, ByVal nCount As Long, ByRef lpRgnData As Any) As Long
'Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
'Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' User32 APIs
Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' Kernel32/User32 APIs for Unicode Filename Support
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "Kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function DeleteFile Lib "Kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributesW Lib "Kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributes Lib "Kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesW Lib "Kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "Kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Const FILE_ATTRIBUTE_NORMAL = &H80&

Public Function iparseCreateShapedRegion(ByRef cHost As c32bppDIB, ByRef regionStyle As eRegionStyles) As Long

    '*******************************************************
    ' FUNCTION RETURNS A HANDLE TO A REGION IF SUCCESSFUL.
    ' If unsuccessful, function retuns zero.
    ' The fastest region from bitmap routines around, custom
    ' designed by LaVolpe. This version modified to create
    ' regions from alpha masks.
    '*******************************************************
    ' Note: See c32bppDIB.CreateRegion for description of the regionStyle parameter
    
    ' declare bunch of variables...
    Dim rgnRects() As RECT ' array of rectangles comprising region
    Dim rectCount As Long ' number of rectangles & used to increment above array
    Dim rStart As Long ' pixel that begins a new regional rectangle
    
    Dim x As Long
    Dim y As Long
    Dim Z As Long
    ' loop counters
    
    Dim bDib() As Byte  ' the DIB bit array
    Dim tSA As SafeArray ' array overlay
    Dim rtnRegion As Long ' region handle returned by this function
    Dim Width As Long
    Dim Height As Long

    Dim lScanWidth As Long ' used to size the DIB bit array
    
    ' Simple sanity checks
    If cHost.Alpha = AlphaNone Then
        iparseCreateShapedRegion = CreateRectRgn(0&, 0&, cHost.Width, cHost.Height)
        Exit Function
    End If
    
    Width = cHost.Width
    If Width < 1& Then Exit Function
    Height = cHost.Height
    If Height < 1& Then Exit Function
    
    On Error GoTo CleanUp
      
    lScanWidth = Width * 4& ' how many bytes per bitmap line?
    With tSA                ' prepare array overlay
        .cbElements = 1     ' byte elements
        .cDims = 2          ' two dim array
        .pvData = cHost.BitsPointer  ' data location
        .rgSABound(0).cElements = Height
        .rgSABound(1).cElements = lScanWidth
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
    
    If regionStyle = regionShaped Then
        
        ReDim rgnRects(0 To Width * 3&) ' start with an arbritray number of rectangles
        
        ' begin pixel by pixel comparisons
        For y = Height - 1 To 0& Step -1&
            ' the alpha byte is every 4th byte
            For x = 3& To lScanWidth - 1& Step 4&
            
                ' test to see if next pixel is 100% transparent
                If bDib(x, y) = 0 Then
                    If Not rStart = 0& Then ' we're currently tracking a rectangle,
                        ' so let's close it, but see if array needs to be resized
                        If rectCount + 1& = UBound(rgnRects) Then _
                            ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                         
                         ' add the rectangle to our array
                         SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - y - 1&, x \ 4 + 1&, Height - y
                         rStart = 0&                    ' reset flag
                         rectCount = rectCount + 1&     ' keep track of nr in use
                    End If
                
                Else
                    ' non-transparent, ensure start value set
                    If rStart = 0& Then rStart = x  ' set start point
                End If
            Next x
            If Not rStart = 0& Then
                ' got to end of bitmap without hitting another transparent pixel
                ' but we're tracking so we'll close rectangle now
               
               ' see if array needs to be resized
               If rectCount + 1& = UBound(rgnRects) Then _
                   ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                   
                ' add the rectangle to our array
                SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - y - 1&, Width, Height - y
                rStart = 0&                     ' reset flag
                rectCount = rectCount + 1&      ' keep track of nr in use
            End If
        Next y

    ElseIf regionStyle = regionEnclosed Then
        
        ReDim rgnRects(0 To Width * 3&) ' start with an arbritray number of rectangles
        
        ' begin pixel by pixel comparisons
        For y = Height - 1 To 0& Step -1&
            ' the alpha byte is every 4th byte
            For x = 3& To lScanWidth - 1& Step 4&
            
                ' test to see if next pixel has any opaqueness
                If Not bDib(x, y) = 0 Then
                    ' we got the left side of the scan line, check the right side
                    For Z = lScanWidth - 1 To x + 4& Step -4&
                        ' when we hit a non-transparent pixel, exit loop
                        If Not bDib(Z, y) = 0 Then Exit For
                    Next
                    ' see if array needs to be resized
                    If rectCount + 1& = UBound(rgnRects) Then _
                        ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                     
                     ' add the rectangle to our array
                     SetRect rgnRects(rectCount + 2&), x \ 4, Height - y - 1&, Z \ 4 + 1&, Height - y
                     rectCount = rectCount + 1&     ' keep track of nr in use
                     Exit For
                End If
            Next x
        Next y
        
    ElseIf regionStyle = regionBounds Then
        
        ReDim rgnRects(0 To 0) ' we will only have 1 regional rectangle
        
        ' set the min,max bounding parameters
        SetRect rgnRects(0), Width * 4, Height, 0, 0
        With rgnRects(0)
            ' begin pixel by pixel comparisons
            For y = Height - 1 To 0& Step -1&
                ' the alpha byte is every 4th byte
                For x = 3& To lScanWidth - 1& Step 4&
                
                    ' test to see if next pixel has any opaqueness
                    If Not bDib(x, y) = 0 Then
                        ' we got the left side of the scan line, check the right side
                        For Z = lScanWidth - 1 To x + 4& Step -4&
                            ' when we hit a non-transparent pixel, exit loop
                            If Not bDib(Z, y) = 0 Then Exit For
                        Next
                        rStart = 1& ' flag indicating we have opaqueness on this line
                        ' resize our bounding rectangle's left/right as needed
                        If x < .Left Then .Left = x
                        If Z > .Right Then .Right = Z
                        Exit For
                    End If
                Next x
                If rStart = 1& Then
                    ' resize our bounding rectangle's top/bottom as needed
                    If y < .Top Then .Top = y
                    If y > .Bottom Then .Bottom = y
                    rStart = 0& ' reset flag indicating we do not have any opaque pixels
                End If
            Next y
        End With
        If rgnRects(0).Right > rgnRects(0).Left Then
            rtnRegion = CreateRectRgn(rgnRects(0).Left \ 4, Height - rgnRects(0).Bottom - 1&, rgnRects(0).Right \ 4 + 1&, _
                                     (rgnRects(0).Bottom - rgnRects(0).Top) + (Height - rgnRects(0).Bottom))
        End If
    End If

    ' remove the array overlay
    CopyMemory ByVal VarPtrArray(bDib()), 0&, 4&
        
    On Error Resume Next
    ' check for failure & engage backup plan if needed
    If Not rectCount = 0 Then
        ' there were rectangles identified, try to create the region in one step
        rtnRegion = local_CreatePartialRegion(rgnRects(), 2&, rectCount + 1&, 0&, Width)
        
        ' ok, now to test whether or not we are good to go...
        ' if less than 2000 rectangles, region should have been created & if it didn't
        ' it wasn't due to O/S restrictions -- failure
        If rtnRegion = 0& Then
            If rectCount > 2000& Then
                ' Win98 has limitation of approximately 4000 regional rectangles
                ' In cases of failure, we will create the region in steps of
                ' 2000 vs trying to create the region in one step
                rtnRegion = local_CreateWin98Region(rgnRects, rectCount + 1&, 0&, Width)
            End If
        End If
    End If

CleanUp:
    Erase rgnRects()
    
    If Err Then ' failure; probably low on resources
        If Not rtnRegion = 0& Then DeleteObject rtnRegion
        Err.Clear
    Else
        iparseCreateShapedRegion = rtnRegion
    End If


End Function

Private Function local_CreatePartialRegion(ByRef rgnRects() As RECT, ByRef lIndex As Long, ByRef uIndex As Long, ByVal leftOffset As Long, ByVal Cx As Long) As Long
    ' Helper function for CreateShapedRegion & local_CreateWin98Region
    ' Called to create a region in its entirety or stepped (see local_CreateWin98Region)

    On Error Resume Next
    ' Note: Ideally contiguous rectangles of equal height & width should be combined
    ' into one larger rectangle. However, thru trial & error I found that Windows
    ' does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region is the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2) ' bytes 0-15
        .Left = 32&                     ' length of region header in bytes
        .Top = 1&                       ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
    End With
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .Top = rgnRects(lIndex).Top         ' top
        .Right = leftOffset + Cx            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    local_CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Private Function local_CreateWin98Region(ByRef rgnRects() As RECT, ByVal rectCount As Long, ByRef leftOffset As Long, ByRef Cx As Long) As Long
    ' Fall-back routine when a very large region fails to be created.
    ' Win98 has problems with regional rectangles over 4000
    ' So, we'll try again in case this is the prob with other systems too.
    ' We'll step it at 2000 at a time which is stil very quick

    Dim x As Long
Dim y As Long
 ' loop counters
    Dim win98Rgn As Long     ' partial region
    Dim rtnRegion As Long    ' combined region & return value of this function
    Const RGN_OR As Long = 2&
    Const scanSize As Long = 2000&

    ' we start with 2 'cause first 2 RECTs are the header
    For x = 2& To rectCount Step scanSize
    
        If x + scanSize > rectCount Then
            y = rectCount
        Else
            y = x + scanSize
        End If
        
        ' attempt to create partial region, scanSize rects at a time
        win98Rgn = local_CreatePartialRegion(rgnRects(), x, y, leftOffset, Cx)
        
        If win98Rgn = 0& Then    ' failure
            ' cleaup combined region if needed
            If Not rtnRegion = 0& Then DeleteObject rtnRegion
            Exit For ' abort; system won't allow us to create the region
        Else
            If rtnRegion = 0& Then ' first time thru
                rtnRegion = win98Rgn
            Else ' already started
                ' use combineRgn, but only every scanSize times
                CombineRgn rtnRegion, rtnRegion, win98Rgn, RGN_OR
                DeleteObject win98Rgn
            End If
        End If
    Next
    ' done; return result
    local_CreateWin98Region = rtnRegion
    
End Function

Public Function iparseIsArrayEmpty(ByVal FarPointer As Long) As Long
  ' test to see if an array has been initialized
  CopyMemory iparseIsArrayEmpty, ByVal FarPointer, 4&
End Function

Public Function iparseByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    iparseByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Public Function iparseArrayToPicture(ByRef inArray() As Byte, ByRef Offset As Long, ByRef Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Note: The array was already validated as not empty when calling class' LoadStream was called
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), iparseArrayToPicture)
            End If
        End If
    End If

End Function

Public Function iparseHandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)
    
    Dim lpPictDesc As PictDesc
    Dim aGUID(0 To 3) As Long

    With lpPictDesc
        .Size = Len(lpPictDesc)
        .TypeIt = imgType
        .hHandle = hImage
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, iparseHandleToStdPicture)
    
End Function

Public Function iparseReverseLong(ByVal inLong As Long) As Long

    ' fast function to reverse a long value from big endian to little endian
    ' PNG files contain reversed longs
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long
    Dim lHighBit As Long
    
    lHighBit = inLong And &H80000000
    If lHighBit Then
      inLong = inLong And Not &H80000000
    End If
    
    b1 = inLong And &HFF
    b2 = (inLong And &HFF00&) \ &H100&
    b3 = (inLong And &HFF0000) \ &H10000
    If lHighBit Then
      b4 = inLong \ &H1000000 Or &H80&
    Else
      b4 = inLong \ &H1000000
    End If
    
    If b1 And &H80& Then
      iparseReverseLong = ((b1 And &H7F&) * &H1000000 Or &H80000000) Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    Else
      iparseReverseLong = b1 * &H1000000 Or _
          b2 * &H10000 Or b3 * &H100& Or b4
    End If

End Function

Public Function iparseValidateDLL(ByVal DllName As String, ByVal dllProc As String) As Boolean
    
    ' PURPOSE: Test a DLL for a specific function.
    
    Dim LB As Long
Dim pa As Long

    
    'attempt to open the DLL to be checked
    LB = LoadLibrary(DllName)
    If LB Then
        'if so, retrieve the address of one of the function calls
        pa = GetProcAddress(LB, dllProc)
        ' free references
        FreeLibrary LB
    End If
    iparseValidateDLL = (Not (LB = 0 Or pa = 0))
    
End Function

Public Sub iparseValidateAlphaChannel(ByRef inStream() As Byte, ByVal bPreMultiply As Boolean, ByRef bIsAlpha As AlphaTypeEnum, ByRef imgType As Long)

    ' Purpose: Modify 32bpp DIB's alpha bytes depending on whether or not they are used
    
    ' Parameters
    ' inStream(). 2D array overlaying the DIB to be checked
    ' bPreMultiply. If true, image will be premultiplied if not already
    ' bIsAlpha. Returns whether or not the image contains transparency
    ' imgType. If passed as -1 then image is known to be not alpha, but will have its alpha values set to 255
    '          When routine returns, imgType is either imgBmpARGB, imgBmpPARGB or imgBitmap

    Dim x As Long
Dim y As Long

    Dim lPARGB As Long
Dim zeroCount As Long
Dim opaqueCount As Long

    Dim bPARGB As Boolean
Dim bAlpha As AlphaTypeEnum


    ' see if the 32bpp is premultiplied or not and if it is alpha or not
    If Not imgType = -1 Then
        For y = 0 To UBound(inStream, 2)
            For x = 3 To UBound(inStream, 1) Step 4
                Select Case inStream(x, y)
                Case 0
                    If lPARGB = 0 Then
                        ' zero alpha, if any of the RGB bytes are non-zero, then this is not pre-multiplied
                        If Not inStream(x - 1, y) = 0 Then
                            lPARGB = 1 ' not premultiplied
                        ElseIf Not inStream(x - 2, y) = 0 Then
                            lPARGB = 1
                        ElseIf Not inStream(x - 3, y) = 0 Then
                            lPARGB = 1
                        End If
                        ' but don't exit loop until we know if any alphas are non-zero
                    End If
                    zeroCount = zeroCount + 1 ' helps in decision factor at end of loop
                Case 255
                    ' no way to indicate if premultiplied or not, unless...
                    If lPARGB = 1 Then
                        lPARGB = 2    ' not pre-multiplied because of the zero check above
                        Exit For
                    End If
                    opaqueCount = opaqueCount + 1
                Case Else
                    ' if any Exit For's below get triggered, not pre-multiplied
                    If lPARGB = 1 Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 3, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 2, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 1, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    Else
                        bAlpha = AlphaComplex
                    End If
                End Select
            Next
            If lPARGB = 2 Then Exit For
        Next
        
        ' if we got all the way thru the image without hitting Exit:For then
        ' the image is not alpha unless the bAlpha flag was set in the loop
        
        If zeroCount = (x \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha value was zero
            bPARGB = False: bAlpha = AlphaNone ' assume RGB, else 100% transparent ARGB
            ' also if lPARGB=0, then image is completely black
        ElseIf opaqueCount = (x \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha is 255
            bPARGB = False: bAlpha = AlphaNone
        Else
            Select Case lPARGB
                Case 2: bPARGB = False ' 100% positive ARGB
                Case 1: bPARGB = False ' now 100% positive ARGB
                Case 0: bPARGB = True
            End Select
            If bAlpha = AlphaNone Then bAlpha = AlphaSimple
        End If
    End If
    
    ' see if caller wants the non-premultiplied alpha channel premultiplied
    If bAlpha Then
        If bPARGB Then ' else force premultiplied
            imgType = imgBmpPARGB
        Else
            imgType = imgBmpARGB
            If bPreMultiply = True Then
                bAlpha = AlphaSimple
                For y = 0 To UBound(inStream, 2)
                    For x = 3 To UBound(inStream, 1) Step 4
                        If inStream(x, y) = 0 Then
                            CopyMemory inStream(x - 3, y), 0&, 4&
                        ElseIf Not inStream(x, y) = 255 Then
                            bAlpha = AlphaComplex
                            For lPARGB = x - 3 To x - 1
                                inStream(lPARGB, y) = ((0& + inStream(lPARGB, y)) * inStream(x, y)) \ &HFF
                            Next
                        End If
                    Next
                Next
            End If
        End If
    Else
        imgType = imgBitmap
        If bPreMultiply = True Then
            For y = 0 To UBound(inStream, 2)
                For x = 3 To UBound(inStream, 1) Step 4
                    inStream(x, y) = 255
                Next
            Next
        End If
    End If
    bIsAlpha = bAlpha

End Sub

Public Sub iparseGrayScaleRatios(ByRef Formula As eGrayScaleFormulas, ByRef R As Single, ByRef g As Single, ByRef b As Single)

        ' note: when adding your own formulas, ensure they add up to 1.0 or less;
        ' else unexpected colors may be calculated. Exception: non-grayscale are always: 1,1,1
        Select Case Formula
        Case gsclNone   ' no grayscale
            R = 1: g = 1: b = 1
        Case gsclNTSCPAL
            R = 0.299: g = 0.587: b = 0.114 ' standard weighted average
        Case gsclSimpleAvg
            R = 0.333: g = 0.334: b = R     ' pure average
        Case gsclCCIR709
            R = 0.213: g = 0.715: b = 0.072 ' Formula.CCIR 709, Default
        Case gsclRedMask
            R = 0.8: g = 0.1: b = g     ' personal preferences: could be r=1:g=0:b=0 or other weights
        Case gsclGreenMask
            R = 0.1: g = 0.8: b = R     ' personal preferences: could be r=0:g=1:b=0 or other weights
        Case gsclBlueMask
            R = 0.1: g = R: b = 0.8     ' personal preferences: could be r=0:g=0:b=1 or other weights
        Case gsclBlueGreenMask
            R = 0.1: g = 0.45: b = g    ' personal preferences: could be r=0:g=.5:b=.5 or other weights
        Case gsclRedGreenMask
            R = 0.45: g = R: b = 0.1    ' personal preferences: could be r=.5:g=.5:b=0 or other weights
        Case Else
            R = 0.299: g = 0.587: b = 0.114 ' use gsclNTSCPAL
        End Select

End Sub

Public Function iparseGetFileHandle(ByVal filename As String, ByVal bOpen As Boolean, Optional ByVal useUnicode As Boolean = False) As Long

    ' Function uses APIs to read/create files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    'Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    
    Dim flags As Long
Dim Access As Long

    Dim Disposition As Long
Dim Share As Long

    
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If bOpen Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If useUnicode Then
            flags = GetFileAttributesW(StrPtr(filename))
        Else
            flags = GetFileAttributes(filename)
        End If
        If flags < 0& Then flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If
    
    If useUnicode Then
        iparseGetFileHandle = CreateFileW(StrPtr(filename), Access, Share, ByVal 0&, Disposition, flags, 0&)
    Else
        iparseGetFileHandle = CreateFile(filename, Access, Share, ByVal 0&, Disposition, flags, 0&)
    End If

End Function

Public Function iparseDeleteFile(ByRef filename As String, Optional ByVal useUnicode As Boolean = False) As Boolean

    ' Function uses APIs to delete files :: unicode supported

    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        If Not (SetFileAttributesW(StrPtr(filename), FILE_ATTRIBUTE_NORMAL) = 0&) Then
            iparseDeleteFile = Not (DeleteFileW(StrPtr(filename)) = 0&)
        End If
    Else
        If Not (SetFileAttributes(filename, FILE_ATTRIBUTE_NORMAL) = 0&) Then
            iparseDeleteFile = Not (DeleteFile(filename) = 0&)
        End If
    End If

End Function

Public Function iparseFileExists(ByRef filename As String, Optional ByVal useUnicode As Boolean) As Boolean
    ' test to see if a file exists
    Const INVALID_HANDLE_VALUE = -1&
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        iparseFileExists = Not (GetFileAttributesW(StrPtr(filename)) = INVALID_HANDLE_VALUE)
    Else
        iparseFileExists = Not (GetFileAttributes(filename) = INVALID_HANDLE_VALUE)
    End If
End Function

Public Sub iparseOverlayHost_Byte(ByRef aOverlay() As Byte, ByRef ptrSafeArray As Long, ByVal nrDims As Long, ByVal ElemCount_Dim1 As Long, ByVal ElemCount_Dim2 As Long, ByVal memPtr As Long)

    ' Routine overlays a BYTE array on top of some memory address. Passing incorrect values will crash the app and maybe the system
    ' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10
    
    ' aOverlay() is an uninitialized, dynamic Byte array. If initialized, call Erase on the array before passing it
    ' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
    ' nrDims must be 1 or 2. Not used if memPtr=0
    ' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
    ' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
    ' memPtr is the memory address to overlay the array onto
    
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
        Dim tSA As SafeArray
        With tSA
            .cbElements = 1     '1=byte
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSABound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If

End Sub

Public Sub iparseOverlayHost_Long(ByRef aOverlay() As Long, ByRef ptrSafeArray As Long, ByVal nrDims As Long, ByVal ElemCount_Dim1 As Long, ByVal ElemCount_Dim2 As Long, ByVal memPtr As Long)

    ' Routine overlays a LONG array on top of some memory address. Passing incorrect values will crash the app and maybe the system
    ' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10
    
    ' aOverlay() is an uninitialized, dynamic Long array. If initialized, call Erase on the array before passing it
    ' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
    ' nrDims must be 1 or 2. Not used if memPtr=0
    ' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
    ' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
    ' memPtr is the memory address to overlay the array onto
    
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
        Dim tSA As SafeArray
        With tSA
            .cbElements = 4     '4=long
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSABound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSABound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If

End Sub

Public Function iparseArrayProps(ByVal arrayPtr As Long, _
                                Optional ByRef Dimensions As Long, _
                                Optional ByRef FirstElementPtr As Long) As Long

    ' Function returns the overall size of the array in bytes or returns zero
    ' if the array is uninitialized or contains no elements
    
    ' Parameters
    '   ArrayPtr :: result of call from iparseGetArrayPointer
    '   Dimensions [out] :: number of dimensions for the array
    '   FirstElementPtr [out] :: basically VarPtr(first element of array)
    
    Dim tSA As SafeArray
    Dim lBounds() As Long
    Dim x As Long
Dim totalSize As Long

    
    If arrayPtr = 0& Then Exit Function
    CopyMemory arrayPtr, ByVal arrayPtr, 4&
    If arrayPtr = 0& Then Exit Function             ' uninitialized array
    
    CopyMemory ByVal VarPtr(tSA), ByVal arrayPtr, 16&     ' safe array structure minus bounds info
    Dimensions = tSA.cDims
    FirstElementPtr = tSA.pvData
    ReDim lBounds(1 To 2, 1 To Dimensions)
    CopyMemory lBounds(1, 1), ByVal arrayPtr + 16&, Dimensions * 8&
    
    totalSize = 1
    For x = 1 To Dimensions
        totalSize = totalSize * lBounds(1, x)
    Next
    
    iparseArrayProps = totalSize * tSA.cbElements
    
End Function

Public Sub iparseOverlayHost_2DbyHost(ByRef aOverlay() As Byte, ByRef ptrSafeArray As Long, ByRef Host As c32bppDIB)

    ' Routine overlays a BYTE array on top of some memory address. Passing incorrect values will crash the app and maybe the system
    ' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10
    
    ' aOverlay() is an uninitialized, dynamic Byte array. If initialized, call Erase on the array before passing it
    ' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
    ' nrDims must be 1 or 2. Not used if memPtr=0
    ' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
    ' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
    ' memPtr is the memory address to overlay the array onto
    
    If ptrSafeArray = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&      ' remove overlay
    ElseIf Not Host Is Nothing Then
        If Host.Handle Then
            Dim tSA As SafeArray
            With tSA
                .cbElements = 1               '1=byte
                .pvData = Host.BitsPointer    'memory address
                .cDims = 2                    'nr of dimensions
                .rgSABound(0).cElements = Host.Height  'number array items (1st dimension)
                .rgSABound(1).cElements = Host.scanWidth 'number array items (2nd dimension)
            End With
            CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
            CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
        End If
    End If

End Sub

Public Function iparseBlendImageToColor(ByRef cHost As c32bppDIB, ByVal FillColor As Long, ByRef outStream() As Byte, Optional ByVal bitDepth24 As Boolean = False) As Boolean

    ' Function basically renders an image over a solid bkg color
    ' Function called from SaveToFlle/Stream_JPG & BMP when the image to be saved
    ' has premultiplied pixels.
    
    Dim x As Long
Dim y As Long

    Dim R As Byte
Dim g As Byte
Dim b As Byte

    Dim pAlpha As Byte
Dim dAlpha As Long

    Dim tSA As SafeArray
Dim srcBytes() As Byte

    Dim scanWidth As Long
Dim destX As Long


    scanWidth = cHost.scanWidth ' cache vs recalculating each scan line
    iparseOverlayHost_2DbyHost srcBytes(), VarPtr(tSA), cHost
    ' extract individual RGB values & convert FillColor
    R = (FillColor And &HFF)
    g = ((FillColor \ &H100) And &HFF)
    b = ((FillColor \ &H10000) And &HFF)
    FillColor = R * &H10000 Or (FillColor And &HFF00&) Or b
    
    If bitDepth24 = False Then ' requesting to blend & save as 32bpp (GDI+ JPG routine calls this)
        ReDim outStream(0 To scanWidth - 1, 0 To cHost.Height - 1)
        For y = 0 To cHost.Height - 1
            For x = 0 To scanWidth - 1& Step 4&
                pAlpha = srcBytes(x + 3&, y)
                If pAlpha = 255 Then
                    CopyMemory outStream(x, y), srcBytes(x, y), 4&
                ElseIf pAlpha = 0 Then
                    CopyMemory outStream(x, y), FillColor, 4&
                Else ' blend to backcolor
                    dAlpha = &HFF& - pAlpha
                    outStream(x, y) = (dAlpha * b) \ &HFF + srcBytes(x, y)
                    outStream(x + 1&, y) = (dAlpha * g) \ &HFF + srcBytes(x + 1&, y)
                    outStream(x + 2&, y) = (dAlpha * R) \ &HFF + srcBytes(x + 2&, y)
                    outStream(x + 3&, y) = 255 ' indicate pixel is fully opaque now
                End If
            Next
        Next
    Else ' requesting to blend as a 24bpp image
        ReDim outStream(0 To iparseByteAlignOnWord(24, cHost.Width) - 1, 0 To cHost.Height - 1)
        For y = 0 To cHost.Height - 1
            destX = 0&
            For x = 0 To scanWidth - 1& Step 4&
                pAlpha = srcBytes(x + 3&, y)
                If pAlpha = 255 Then
                    CopyMemory outStream(destX, y), srcBytes(x, y), 3&
                ElseIf pAlpha = 0 Then
                    CopyMemory outStream(destX, y), FillColor, 3&
                Else ' blend to backcolor
                    dAlpha = &HFF& - pAlpha
                    outStream(destX, y) = (dAlpha * b) \ &HFF + srcBytes(x, y)
                    outStream(destX + 1&, y) = (dAlpha * g) \ &HFF + srcBytes(x + 1&, y)
                    outStream(destX + 2&, y) = (dAlpha * R) \ &HFF + srcBytes(x + 2&, y)
                End If
                destX = destX + 3&
            Next
        Next
    End If
    iparseOverlayHost_2DbyHost srcBytes(), 0&, Nothing

End Function

Public Function iparseFindColor(ByRef PaletteItems() As Long, ByVal Color As Long, ByVal count As Long, ByRef isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    
    ' [in] PaletteItems(). Long Array to search within. Array must be 1-bound, sorted ascending
    ' [in] Color. A value to search for.
    ' [in] Count. Number of items in PaletteItems() to compare against
    ' [out] isNew. If Color not found, isNew is True else False
    ' [out] Return value: The Index where Color was found or where the new Color should be inserted

    Dim UB As Long
Dim LB As Long

    Dim NewIndex As Long
    
    If count = 0& Then
        iparseFindColor = 1&
        isNew = True
        Exit Function
    End If
    
    UB = count
    LB = 1&
    
    Do Until LB > UB
        NewIndex = LB + ((UB - LB) \ 2&)
        If PaletteItems(NewIndex) = Color Then
            Exit Do
        ElseIf PaletteItems(NewIndex) > Color Then ' new color is lower in sort order
            UB = NewIndex - 1&
        Else ' new color is higher in sort order
            LB = NewIndex + 1&
        End If
    Loop

    If LB > UB Then  ' color was not found
            
        If Color > PaletteItems(NewIndex) Then NewIndex = NewIndex + 1&
        isNew = True
        
    Else
        isNew = False
    End If
    
    iparseFindColor = NewIndex

End Function




