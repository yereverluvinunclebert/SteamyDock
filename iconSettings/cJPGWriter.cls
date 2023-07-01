VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJPGWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'  -----======== PURPOSE: Write JPG image format when use of GDI+ not available ========-----
' ._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._._.

' This class is only used on non-GDI+ systems

' Credits to John Korejwa for his excellent VB version of the JPG encoding algorithm
' However, I have made modifications to his class so it can be used within this
' c32bppDIB Suite project. John's original version did not directly support 32bpp
' DIBs and required a secondary bitmap. Other than that, I removed some unused
' constants, removed some properties, and changed many Public routines to Private.
' I have also removed the JPG Comment that was written to each JPG identifying
' John as the encoder. This was done simply to protect his respect should the
' changes I made adversely effect the quality of the encoding.  Therefore,
' if something breaks, it is my fault and not the original author's.

' ////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Class Name:   cJpeg.cls  "JPEG Encoder Class"
'Author:       John Korejwa  <korejwa@tiac.net>
'Version:      0.9 beta  [26 / November / 2003]
'Legal:
'        This class is intended for and was uploaded to www.planetsourcecode.com
'        This product includes JPEG compression code developed by John Korejwa.  <korejwa@tiac.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'Credits:
'        Special thanks to Barry G., a government research scientist who took an interest in my
'        steganography software and research in late 1999.  I never met Barry in person, but he
'        was kind enough to buy and mail me a book with the ISO DIS 10918-1 JPEG standard.
'Description:  This class contains code for compressing pictures, sampled via hDC, into
'              baseline .JPG files.  Please report any errors or unusual behavior to the email
'              address above.
'Dependencies: None
' ////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Private Const SOF0    As Long = &HC0& 'Baseline DCT
                                      'Other Markers
'Define Huffman tables
'Private Const SOI     As Long = &HD8& 'Start of image
'Private Const EOI     As Long = &HD9& 'End of image
'Private Const SOS     As Long = &HDA& 'Start of scan
'Define quantization table(s)

'Consider these arrays of constants.
'They are initialized with the class and do not change.
Private QLumin(0 To 63)    As Integer 'Standard Luminance   Quantum (for 50% quality)
Private QChrom(0 To 63)    As Integer 'Standard Chrominance Quantum (for 50% quality)
Private FDCTScale(0 To 7)  As Double  'Constants for scaling FDCT Coefficients
'Private IDCTScale(0 To 7)  As Double  'Constants for scaling IDCT Coefficients
Private ZigZag(0 To 7, 0 To 7)  As Long    'Zig Zag order of 8X8 block of samples

Private Type SAFEARRAYBOUND
    cElements         As Long
    lLbound           As Long
End Type
Private Type SAFEARRAY2D
    cDims             As Integer
    fFeatures         As Integer
    cbElements        As Long
    cLocks            As Long
    pvData            As Long
    Bounds(0 To 1)    As SAFEARRAYBOUND
End Type

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSource As Any, ByVal cbCopy As Long)
'Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Const INVALID_HANDLE_VALUE = -1

'Custom variable types used for this JPEG encoding implementation
Private Type QUANTIZATIONTABLE
    Qk(0 To 63)            As Integer 'Quantization Values
    FScale(0 To 63)        As Single  'Multiplication values to scale and Quantize   FDCT output
    IScale(0 To 63)        As Single  'Multiplication values to scale and DeQuantize IDCT input
End Type
Private Type HUFFMANTABLE
    BITS(0 To 15)          As Byte    'Number of huffman codes of length i+1
    HUFFVAL(0 To 255)      As Byte    'Huffman symbol values
    EHUFSI(0 To 255)       As Long    'Huffman code size for symbol i
    EHUFCO(0 To 255)       As Long    'Huffman code      for symbol i
    MINCODE(0 To 15)       As Long    '
    MAXCODE(0 To 15)       As Long    'Largest code value for length i+1
End Type
Private Type COMPONENT
    Ci                As Long    'Component ID                       [0-255]
    Hi                As Long    'Horizontal Sampling Factor         [1-4]
    Vi                As Long    'Vertical   Sampling Factor         [1-4]
    Tqi               As Long    'Quantization Table Select          [0-3]
    data()            As Integer 'DCT Coefficients
End Type

'Private PP            As Long    'Sample Precision [8, 12]
''Private YY            As Long    'Number of lines             [Image Height]
'Private XX            As Long    'Number of samples per line  [Image Width]
Private Nf            As Long    'Number of components in Frame

Private HMax          As Long    'Maximum horizontal sampling frequency
Private VMax          As Long    'Maximum vertical   sampling frequency

'Private m_Data()      As Byte    'JPEG File Data
'Private m_Chr         As Long    'Current Character in m_Data
'Private m_Ptr         As Long    'Byte index in m_Data
'Private m_Bit         As Long    'Bit  index in m_Chr

Private m_Block(0 To 7, 0 To 7) As Single  'Buffer for calculating DCT

'4 Quantization Tables
'4 DC Huffman Tables
'4 AC Huffman Tables
Private Comp()             As COMPONENT          'Scan Components


'========================================================================================
'              D I S C R E T E   C O S I N E   T R A N S F O R M A T I O N
'========================================================================================
'@Ignore ProcedureNotUsed
Private Sub FDCT()
    Static t0   As Single 'Given an 8X8 block of discretely sampled values [m_Block(0-7, 0-7)],
    Static t1   As Single 'replace them with their (scaled) Forward Discrete Cosine Transformation values.
    Static t2   As Single '80 (+64) multiplications and 464 additions are needed.
    Static t3   As Single 'Values are scaled on output, meaning that each of the 64 elements must be
    Static t4   As Single 'multiplied by constants for a final FDCT.  These final constants are combined
    Static t5   As Single 'with Quantization constants, so a final 64 multiplications combine the
    Static t6   As Single 'completion of the FDCT and Quantization in one step.
    Static t7   As Single
    Static t8   As Single
    Static I    As Long

    For I = 0 To 7                  'Process 1D FDCT on each row
        t0 = m_Block(I, 0) + m_Block(I, 7)
        t1 = m_Block(I, 0) - m_Block(I, 7)
        t2 = m_Block(I, 1) + m_Block(I, 6)
        t3 = m_Block(I, 1) - m_Block(I, 6)
        t4 = m_Block(I, 2) + m_Block(I, 5)
        t5 = m_Block(I, 2) - m_Block(I, 5)
        t6 = m_Block(I, 3) + m_Block(I, 4)
        t7 = m_Block(I, 3) - m_Block(I, 4)

        t7 = t7 + t5
        t8 = t0 - t6
        t6 = t6 + t0
        t0 = t2 + t4
        t2 = (t2 - t4 + t8) * 0.707106781186548   'Cos(2# * PI / 8#)
        t4 = t1 + t3
        t3 = (t3 + t5) * 0.707106781186548        'Cos(2# * PI / 8#)
        t5 = (t4 - t7) * 0.382683432365091        'Cos(3# * PI / 8#)
        t7 = t7 * 0.541196100146196 - t5          'Cos(PI / 8#) - Cos(3# * PI / 8#)
        t4 = t4 * 1.30656296487638 - t5           'Cos(PI / 8#) + Cos(3# * PI / 8#)
        t5 = t1 + t3
        t1 = t1 - t3

        m_Block(I, 0) = t6 + t0
        m_Block(I, 4) = t6 - t0
        m_Block(I, 1) = t5 + t4
        m_Block(I, 7) = t5 - t4
        m_Block(I, 2) = t8 + t2
        m_Block(I, 6) = t8 - t2
        m_Block(I, 5) = t1 + t7
        m_Block(I, 3) = t1 - t7
    Next I

    For I = 0 To 7                   'Process 1D FDCT on each column
        t0 = m_Block(0, I) + m_Block(7, I)
        t1 = m_Block(0, I) - m_Block(7, I)
        t2 = m_Block(1, I) + m_Block(6, I)
        t3 = m_Block(1, I) - m_Block(6, I)
        t4 = m_Block(2, I) + m_Block(5, I)
        t5 = m_Block(2, I) - m_Block(5, I)
        t6 = m_Block(3, I) + m_Block(4, I)
        t7 = m_Block(3, I) - m_Block(4, I)

        t7 = t7 + t5
        t8 = t0 - t6
        t6 = t6 + t0
        t0 = t2 + t4
        t2 = (t2 - t4 + t8) * 0.707106781186548   'Cos(2# * PI / 8#)
        t4 = t1 + t3
        t3 = (t3 + t5) * 0.707106781186548        'Cos(2# * PI / 8#)
        t5 = (t4 - t7) * 0.382683432365091        'Cos(3# * PI / 8#)
        t7 = t7 * 0.541196100146196 - t5          'Cos(PI / 8#) - Cos(3# * PI / 8#)
        t4 = t4 * 1.30656296487638 - t5           'Cos(PI / 8#) + Cos(3# * PI / 8#)
        t5 = t1 + t3
        t1 = t1 - t3

        m_Block(0, I) = t6 + t0
        m_Block(4, I) = t6 - t0
        m_Block(1, I) = t5 + t4
        m_Block(7, I) = t5 - t4
        m_Block(2, I) = t8 + t2
        m_Block(6, I) = t8 - t2
        m_Block(5, I) = t1 + t7
        m_Block(3, I) = t1 - t7
    Next I
End Sub




'================================================================================
'                 H U F F M A N   T A B L E   G E N E R A T I O N
'================================================================================






'================================================================================
'                           E N T R O P Y   C O D I N G
'================================================================================









'========================================================================================
'                      C O L L E C T I N G   S T A T I S T I C S
'========================================================================================
'These procedures collect statistics of run-length and size categories of DCT coefficients
'so optimized Huffman tables can be generated to compress them.







'========================================================================================
'                                Q U A N T I Z A T I O N
'========================================================================================






'================================================================================
'                           I M A G E   S A M P L I N G
'================================================================================
Private Sub SetSamplingFrequencies(ByVal H1 As Long, ByVal V1 As Long, ByVal H2 As Long, ByVal V2 As Long, ByVal H3 As Long, ByVal V3 As Long)

' Note: the modifications of the code by LaVolpe only support 1,1,1,1,1,1 sampling. Therefore
' this routine changed from Public to Private, othewise, this routine left unmodified


'This class always samples and compresses pictures in YCbCr colorspace.  The first component, Y,
'represents the Luminance of the pixels.  This is "how bright" a pixel is.  The Cb and Cr
'components are Chrominance, which is a measure of how far from neutral-white (toward a color)
'a pixel is.  The human visual sensory system can discriminate Luminance differences about
'twice as well as it can discriminate Chrominance differences.
'
'Virtually all JPEG files are in YCbCr colorspace.  Other JPEG compliant colorspaces exist, but
'they are used in specialty equipment.  For example, people in the astronomy or medical fields
'choose colorspaces that best record the information they are interested in, and don't care about
'how pretty the picture looks to a person when displayed on a computer monitor.
'[Apple/Machintosh sometimes uses a four component colorspace, but that colorspace is rare and
'not widely supported]
'
'Sampling frequencies define how often each component is sampled.  Higher frequencies store more
'information, while lower frequencies store less.  Typically, sampling frequencies are set at
'2,2, 1,1, 1,1.  This corresponds to the human visual sensory system.  The first component,
'Luminance, is sampled twice as much because our eyes notice differences in Luminance quite easily.
'The two Chrominance components are sampled half as much as because our eyes can't distinguish
'the difference in color changes as well.  One Luminance value is sampled for every pixel, and
'one Chrominance value is sampled for each 2X2 block of pixels.
'
'Digital cameras typically record at sampling frequencies of 1,1, 1,1, 1,1.  This samples every
'pixel for all three components.  The quality of the picture is a little better when viewed by
'a person, but the compression benefits drop significantly.  If the picture to be compressed
'is from a Scanner or Digital camera, and you plan on printing it in the future, and storage
'space is not a problem, then sampling at these frequencies makes sense.  Otherwise, if you only
'plan on using the picture to display on a monitor or a web page, [2,2, 1,1, 1,1] makes the
'most sense.
'
'The JPEG standard specifies that sampling frequencies may range from 1-4 for each component
'in both directions.  However, if any component has a sampling frequency of '3', and another
'component has a coresponding sampling frequency of '2' or '4', the downsampling process
'will map fractional pixels to sample values.  This is leagal in the JPEG standard, and this
'class will compress fractional pixel samplings, but this is not widely supported.  It is
'highly recommended to AVOID SAMPLING FACTORS OF 3 for maximum compatability with JPEG decoders.
'
'Some JPEG encoders avoid the fractional pixel problem by only allowing the end user to pick
'a "sub-sampling" value.  In such "Sub Sampling" schemes, all Chrominance frequencies are set
'to one, and the (one or two) sub-sampling value(s) specify Luminance frequencies.
'
'There should *never* be an error raised if you are using this class correctly.  It should
'not be possible for the end user to specify illegal sampling frequency values!
'[For tinkerers - If you delete the error raising code and specify illegal sampling
'frequencies, this class will procede to create a non-JPEG compliant file with the values
'specified]

    Dim I As Long

    If H1 < 1 Or H1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
    If V1 < 1 Or V1 > 4 Then Err.Raise 1, , "Invalid Sampling Value"

    If (H2 Or H3 Or V2 Or V3) = 0 Then  'if H2,H3,V2,V3 are all zero ...
        Nf = 1         'Luminance only.
        ReDim Comp(0 To 0)
        Comp(0).Hi = 1 'Set up for sampling Greyscale
        Comp(0).Vi = 1 '(Black and White picture)
    Else
        If H2 < 1 Or H2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If H3 < 1 Or H3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If V2 < 1 Or V2 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        If V3 < 1 Or V3 > 4 Then Err.Raise 1, , "Invalid Sampling Value"
        Nf = 3         'YCbCr
        ReDim Comp(0 To 2)
        Comp(0).Hi = H1
        Comp(0).Vi = V1
        Comp(0).Tqi = 0
        Comp(1).Hi = H2
        Comp(1).Vi = V2
        Comp(1).Tqi = 1
        Comp(2).Hi = H3
        Comp(2).Vi = V3
        Comp(2).Tqi = 1
    End If

    HMax = -1
    VMax = -1
    For I = 0 To Nf - 1 'determine max h, v sampling factors
        If HMax < Comp(I).Hi Then HMax = Comp(I).Hi
        If VMax < Comp(I).Vi Then VMax = Comp(I).Vi
    Next I

End Sub





'================================================================================
'                         E M I T I N G   M A R K E R S
'================================================================================









'================================================================================
'                           E M I T I N G   S C A N S
'================================================================================










'========================================================================================
'                               W R I T I N G   F I L E
'========================================================================================







'========================================================================================
'                         C L A S S   I N I T I A L I Z A T I O N
'========================================================================================
Private Sub Class_Initialize()
    Dim I As Long
    Dim J As Long
    Dim dX As Long
    Dim zz As Long

    I = 0                   'Initialize the ZigZag() array, which maps out the
    J = 0                   '  zig-zag sequence of quantized DCT coefficients
    dX = 1                  '  in approximately low to high spatial frequencies
    For zz = 0 To 63
        ZigZag(I, J) = zz
        I = I + dX
        J = J - dX
        If I > 7 Then              '  0   1   5   6  14  15  27  28
            I = 7                  '  2   4   7  13  16  26  29  42
            J = J + 2              '  3   8  12  17  25  30  41  43
            dX = -1                '  9  11  18  24  31  40  44  53
        ElseIf J > 7 Then          ' 10  19  23  32  39  45  52  54
            J = 7                  ' 20  22  33  38  46  51  55  60
            I = I + 2              ' 21  34  37  47  50  56  59  61
            dX = 1                 ' 35  36  48  49  57  58  62  63
        ElseIf I < 0 Then
            I = 0 'check (j>7) first
            dX = 1
        ElseIf J < 0 Then
            J = 0
            dX = -1
        End If
    Next zz

   'Luminance Quantization table for Quality = 50
    QLumin(0) = 16:   QLumin(1) = 11:   QLumin(2) = 12:   QLumin(3) = 14
    QLumin(4) = 12:   QLumin(5) = 10:   QLumin(6) = 16:   QLumin(7) = 14
    QLumin(8) = 13:   QLumin(9) = 14:   QLumin(10) = 18:  QLumin(11) = 17
    QLumin(12) = 16:  QLumin(13) = 19:  QLumin(14) = 24:  QLumin(15) = 40
    QLumin(16) = 26:  QLumin(17) = 24:  QLumin(18) = 22:  QLumin(19) = 22
    QLumin(20) = 24:  QLumin(21) = 49:  QLumin(22) = 35:  QLumin(23) = 37
    QLumin(24) = 29:  QLumin(25) = 40:  QLumin(26) = 58:  QLumin(27) = 51
    QLumin(28) = 61:  QLumin(29) = 60:  QLumin(30) = 57:  QLumin(31) = 51
    QLumin(32) = 56:  QLumin(33) = 55:  QLumin(34) = 64:  QLumin(35) = 72
    QLumin(36) = 92:  QLumin(37) = 78:  QLumin(38) = 64:  QLumin(39) = 68
    QLumin(40) = 87:  QLumin(41) = 69:  QLumin(42) = 55:  QLumin(43) = 56
    QLumin(44) = 80:  QLumin(45) = 109: QLumin(46) = 81:  QLumin(47) = 87
    QLumin(48) = 95:  QLumin(49) = 98:  QLumin(50) = 103: QLumin(51) = 104
    QLumin(52) = 103: QLumin(53) = 62:  QLumin(54) = 77:  QLumin(55) = 113
    QLumin(56) = 121: QLumin(57) = 112: QLumin(58) = 100: QLumin(59) = 120
    QLumin(60) = 92:  QLumin(61) = 101: QLumin(62) = 103: QLumin(63) = 99

   'Chrominance Quantization table for Quality = 50
    QChrom(0) = 17:   QChrom(1) = 18:   QChrom(2) = 18:   QChrom(3) = 24
    QChrom(4) = 21:   QChrom(5) = 24:   QChrom(6) = 47:   QChrom(7) = 26
    QChrom(8) = 26:   QChrom(9) = 47:   QChrom(10) = 99:  QChrom(11) = 66
    QChrom(12) = 56:  QChrom(13) = 66:  QChrom(14) = 99:  QChrom(15) = 99
    QChrom(16) = 99:  QChrom(17) = 99:  QChrom(18) = 99:  QChrom(19) = 99
    QChrom(20) = 99:  QChrom(21) = 99:  QChrom(22) = 99:  QChrom(23) = 99
    QChrom(24) = 99:  QChrom(25) = 99:  QChrom(26) = 99:  QChrom(27) = 99
    QChrom(28) = 99:  QChrom(29) = 99:  QChrom(30) = 99:  QChrom(31) = 99
    QChrom(32) = 99:  QChrom(33) = 99:  QChrom(34) = 99:  QChrom(35) = 99
    QChrom(36) = 99:  QChrom(37) = 99:  QChrom(38) = 99:  QChrom(39) = 99
    QChrom(40) = 99:  QChrom(41) = 99:  QChrom(42) = 99:  QChrom(43) = 99
    QChrom(44) = 99:  QChrom(45) = 99:  QChrom(46) = 99:  QChrom(47) = 99
    QChrom(48) = 99:  QChrom(49) = 99:  QChrom(50) = 99:  QChrom(51) = 99
    QChrom(52) = 99:  QChrom(53) = 99:  QChrom(54) = 99:  QChrom(55) = 99
    QChrom(56) = 99:  QChrom(57) = 99:  QChrom(58) = 99:  QChrom(59) = 99
    QChrom(60) = 99:  QChrom(61) = 99:  QChrom(62) = 99:  QChrom(63) = 99


    FDCTScale(0) = 0.353553390593273     '0.25 / Cos(4 / 16 * PI)
    FDCTScale(1) = 0.25489778955208      '0.25 / Cos(1 / 16 * PI)
    FDCTScale(2) = 0.270598050073098     '0.25 / Cos(2 / 16 * PI)
    FDCTScale(3) = 0.300672443467523     '0.25 / Cos(3 / 16 * PI)
    FDCTScale(4) = 0.353553390593273     '0.25 / Cos(4 / 16 * PI)
    FDCTScale(5) = 0.449988111568207     '0.25 / Cos(5 / 16 * PI)
    FDCTScale(6) = 0.653281482438186     '0.25 / Cos(6 / 16 * PI)
    FDCTScale(7) = 1.28145772387074      '0.25 / Cos(7 / 16 * PI)

    'SetSamplingFrequencies 2, 2, 1, 1, 1, 1
    ' Modified by LaVolpe: modifications to the SampleHDC routine now only supports highest quality sampling
    SetSamplingFrequencies 1, 1, 1, 1, 1, 1
End Sub






