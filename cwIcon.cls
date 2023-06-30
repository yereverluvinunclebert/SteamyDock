VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMainIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''Private WithEvents W As cWidgetBase '<- this is required in each cwImplementation...
'
'Private mainIconxo As Integer: Private mainIconyo As Integer
'
'
'Private Sub Class_Initialize()
'
'    'Call initialiseVars
'
'    mainIconxo = 0: mainIconyo = 0
'
''    Set W = Cairo.WidgetBase '<- this is required in each cwImplementation...
''        W.Alpha = 0.5
'End Sub
'
''the following two Public Props also need to be exposed in each Widget-Implementation (but that was it)
''Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
''Public Property Get Widgets() As cWidgets:  Set Widgets = W.Widgets: End Property
'Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    If Button = vbRightButton Then
'         Call menuForm.PopupMenu(menuForm.mnuMainMenu)
'    End If
'End Sub
'
'Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
'    Draw CC, dx_Aligned, dy_Aligned
'End Sub
'
'Private Sub Draw(CC As cCairoContext, dx, dy)
'
'    CC.RenderSurfaceContent "mainIcon", mainIconxo, mainIconyo

'End Sub
