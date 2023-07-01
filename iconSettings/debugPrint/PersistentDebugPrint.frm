VERSION 5.00
Begin VB.Form frmDebugPrint 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Persistent Debug Print Window"
   ClientHeight    =   8025
   ClientLeft      =   1005
   ClientTop       =   3015
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "PersistentDebugPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5355
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   540
      Width           =   4695
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu mnuSeparate 
      Caption         =   "Separate"
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Font"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuBackColor 
      Caption         =   "BackColor"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuForeColor 
      Caption         =   "ForeColor"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuReset 
      Caption         =   "Reset"
   End
End
Attribute VB_Name = "frmDebugPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long ' This is +1 (right - left = width)
    Bottom As Long ' This is +1 (bottom - top = height)
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
'
Private Const WM_SETREDRAW      As Long = &HB&
Private Const EM_SETSEL         As Long = &HB1&
Private Const EM_REPLACESEL     As Long = &HC2&
'
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
'

Private Sub Form_Load()
    On Error Resume Next
        Me.Left = GetSetting(App.Title, "Settings", "Left", 0&)
        Me.Top = GetSetting(App.Title, "Settings", "Top", 0&)
        Me.Width = GetSetting(App.Title, "Settings", "Width", 6600&)
        Me.Height = GetSetting(App.Title, "Settings", "Height", 6600&)
        If Not FormIsFullyOnMonitor(Me) Then
            Me.Left = 0&
            Me.Top = 0&
        End If
        '
        txt.FontName = GetSetting(App.Title, "Settings", "FontName", "Courier New")
        txt.FontBold = GetSetting(App.Title, "Settings", "FontBold", False)
        txt.FontItalic = GetSetting(App.Title, "Settings", "FontItalic", False)
        txt.FontSize = GetSetting(App.Title, "Settings", "FontSize", 8)
        txt.FontStrikethru = GetSetting(App.Title, "Settings", "FontStrikethru", False)
        txt.FontUnderline = GetSetting(App.Title, "Settings", "FontUnderline", False)
        '
        txt.BackColor = GetSetting(App.Title, "Settings", "BackColor", vbWhite)
        txt.ForeColor = GetSetting(App.Title, "Settings", "ForeCOlor", vbBlack)
    On Error GoTo 0
    '
    SubclassFormToReceiveStringMsg Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "Left", Me.Left
    SaveSetting App.Title, "Settings", "Top", Me.Top
    SaveSetting App.Title, "Settings", "Width", Me.Width
    SaveSetting App.Title, "Settings", "Height", Me.Height
    '
    SaveSetting App.Title, "Settings", "FontName", txt.FontName
    SaveSetting App.Title, "Settings", "FontBold", txt.FontBold
    SaveSetting App.Title, "Settings", "FontItalic", txt.FontItalic
    SaveSetting App.Title, "Settings", "FontSize", txt.FontSize
    SaveSetting App.Title, "Settings", "FontStrikethru", txt.FontStrikethru
    SaveSetting App.Title, "Settings", "FontUnderline", txt.FontUnderline
    '
    SaveSetting App.Title, "Settings", "BackColor", txt.BackColor
    SaveSetting App.Title, "Settings", "ForeCOlor", txt.ForeColor
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        txt.Move 0&, 0&, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub





Private Sub mnuClear_Click()
    txt.Text = vbNullString
End Sub

Private Sub mnuSeparate_Click()
    Out vbCrLf
    Out "-----------------------"
    Out vbCrLf
End Sub

Private Sub mnuFont_Click()
'    cdl.flags = cdlCFScreenFonts Or cdlCFForceFontExist
'    '
'    cdl.FontName = txt.FontName
'    cdl.FontBold = txt.FontBold
'    cdl.FontItalic = txt.FontItalic
'    cdl.FontSize = txt.FontSize
'    cdl.FontStrikethru = txt.FontStrikethru
'    cdl.FontUnderline = txt.FontUnderline
'    '
'    cdl.ShowFont
    '
    txt.FontName = "courier new"
    txt.FontBold = False
    txt.FontItalic = False
    txt.FontSize = 8
    txt.FontStrikethru = False
    txt.FontUnderline = False
End Sub

Private Sub mnuBackColor_Click()
    ShowColorDialog Me.hWnd, txt.BackColor, , "BackColor"
    If ColorDialogSuccessful Then txt.BackColor = ColorDialogColor
End Sub

Private Sub mnuForeColor_Click()
    ShowColorDialog Me.hWnd, txt.BackColor, , "ForeColor"
    If ColorDialogSuccessful Then txt.ForeColor = ColorDialogColor
End Sub

Private Sub mnuReset_Click()
        Me.Left = 0&
        Me.Top = 0&
        Me.Width = 6600&
        Me.Height = 6600&
        '
        txt.FontName = "Fixedsys"
        txt.FontBold = False
        txt.FontItalic = False
        txt.FontSize = 9
        txt.FontStrikethru = False
        txt.FontUnderline = False
        '
        txt.BackColor = vbBlack
        txt.ForeColor = vbYellow
End Sub






Public Sub Out(s As String, Optional bHoldLine As Boolean)
    SendMessageW txt.hWnd, EM_SETSEL, &H7FFFFFFF, ByVal &H7FFFFFFF          ' txt.SelStart = &H7FFFFFFF
    If bHoldLine Then
        SendMessageW txt.hWnd, EM_REPLACESEL, 0, ByVal StrPtr(s)            ' txt.SelText = s
    Else
        SendMessageW txt.hWnd, EM_REPLACESEL, 0, ByVal StrPtr(s & vbCrLf)   ' txt.SelText = s & vbCrLf
    End If
End Sub






Private Function FormIsFullyOnMonitor(frm As Form) As Boolean
    ' This tells us whether or not a form is FULLY visible on its monitor.
    '
    Dim hMonitor As Long
    Dim r1 As RECT
    Dim r2 As RECT
    Dim uMonInfo As MONITORINFO
    '
    hMonitor = hMonitorForForm(frm)
    GetWindowRect frm.hWnd, r1
    uMonInfo.cbSize = LenB(uMonInfo)
    GetMonitorInfo hMonitor, uMonInfo
    r2 = uMonInfo.rcWork
    '
    FormIsFullyOnMonitor = (r1.Top >= r2.Top) And (r1.Left >= r2.Left) And (r1.Bottom <= r2.Bottom) And (r1.Right <= r2.Right)
End Function

Public Function hMonitorForForm(frm As Form) As Long
    ' The monitor that the window is MOSTLY on.
    Const MONITOR_DEFAULTTONULL = &H0
    hMonitorForForm = MonitorFromWindow(frm.hWnd, MONITOR_DEFAULTTONULL)
End Function




