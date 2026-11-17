VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDebugDockRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Any class implementing IDockRenderer must provide all three methods with matching signatures.
Implements IDockRenderer

Private mIconCount As Long

Public Sub IDockRenderer_BeginFrame()

    mIconCount = 0

    Debug.Print "----- BEGIN FRAME -----"

End Sub

Public Sub IDockRenderer_DrawIcon(ByVal Icon As cwMainIcon)

    mIconCount = mIconCount + 1

    Debug.Print "Icon: " & Icon.FileName
    Debug.Print "Position: "; Icon.X; ","; Icon.Y
    Debug.Print "Size: "; Icon.Width; "x"; Icon.Height

End Sub

Public Sub IDockRenderer_EndFrame()

    Debug.Print "Icons drawn: "; mIconCount
    Debug.Print "----- END FRAME -----"

End Sub

