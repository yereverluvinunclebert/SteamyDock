VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPlusDockRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cGDIPlusDockRenderer.cls
' Author    : beededea
' Date      : 05/11/2026
' Purpose   : Implements the contract using GDI+, in Visual Basic 6.0 (VB6), interfaces are created using regular Class Modules (.cls) that contain only declarations without any implementation code.
'             Because VB6 does not feature an explicit Interface keyword like modern VB.NET, a normal class acts as an abstract contract or blueprint. Other classes
'             then adopt this contract using the Implements keyword. The subroutines, functions, and properties inside an interface class contain absolutely no executable code. They are essentially just empty shell definitions.
'             The background is that VB6 does not support traditional class inheritance (inheriting code from a base class), interface classes are the primary tool used to achieve polymorphism—allowing different objects to be
'             treated interchangeably through a common interface
'---------------------------------------------------------------------------------------

Option Explicit

'Any class implementing IDockRenderer must provide all three methods with matching signatures.
Implements IDockRenderer

Private mIconCount As Long

Private mFrameActive As Boolean

Private Sub IDockRenderer_BeginFrame()
    ' GDI+ initialisation
    
    If mFrameActive Then Exit Sub
    
    mIconCount = 0

    Debug.Print "----- BEGIN FRAME -----"

    ' 1. Obtain the backbuffer.
    ' 2. Create or obtain the GDI+ graphics context.
    ' 3. Clear the backbuffer to transparent.
    
    'creates a bitmap section in memory that applications can write to directly
    'Call createNewGDIPBitmap

    mFrameActive = True
    
End Sub

Private Sub IDockRenderer_DrawIcon(ByVal Icon As cwMainIcon)

    ' Actual GDI+ drawing

    If Not mFrameActive Then Exit Sub
    
    mIconCount = mIconCount + 1

    Debug.Print "Icon: " & Icon.FileName
    Debug.Print "Position: "; Icon.X; ","; Icon.Y
    Debug.Print "Size: "; Icon.Width; "x"; Icon.Height

    ' 1. Retrieve the appropriate cached image.
    ' 2. Obtain its GDI+ image handle.
    ' 3. Draw using the icon's current properties.
    
    'Call updateDisplayFromIconClass(vbNullString, Icon.X, Icon.Y, Icon.Width, Icon.Height, 0, Icon.IconIndex)

    ' Destination:
    ' Icon.X
    ' Icon.Y
    ' Icon.Width
    ' Icon.Height
    
End Sub

Private Sub IDockRenderer_EndFrame()

    ' Complete GDI+ rendering

    If Not mFrameActive Then Exit Sub

    Debug.Print "Icons drawn: "; mIconCount
    Debug.Print "----- END FRAME -----"

    ' 1. Complete drawing.
    ' 2. Present the backbuffer.
    ' 3. Release temporary resources.
    
    ' calling a public routine within a form is fine, a form is a class
    'Call dock.updateScreenUsingGDIPBitmap

    mFrameActive = False

End Sub
