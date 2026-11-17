VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDockRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IDockRenderer.cls
' Author    : beededea
' Date      : 05/11/2026
' Purpose   : Implements an interface to a graphics renderer be it GDI+ or Cairo or another. Using Visual Basic 6.0 (VB6), interfaces are created using regular Class
'             Modules (.cls) that contain only declarations without any implementation code.
'             Because VB6 does not feature an explicit Interface keyword like modern TwinBasic, a normal class acts as an abstract contract or blueprint for an interface.
'             Other classes then adopt this contract using the Implements keyword. The subroutines, functions, and properties inside an interface class contain
'             absolutely no executable code. They are essentially just empty shell definitions.
'
'             The background is that VB6 does not support traditional class inheritance (inheriting code from a base class), interface classes are the primary tool used
'             to achieve polymorphism—allowing different objects to be treated interchangeably through a common interface
'---------------------------------------------------------------------------------------
'
' currently the only graphics method implemented is cGDIPlusDockRenderer.cls
' but this leaves room for using an alternative such as Cairo or even GDI later
'
Option Explicit

Public Sub BeginFrame()
End Sub

Public Sub DrawIcon(ByVal Icon As cwMainIcon)
End Sub

Public Sub EndFrame()
End Sub

