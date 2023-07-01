VERSION 5.00
Begin VB.Form frmConfirmDock 
   Caption         =   "Confirm Generation"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2925
   Icon            =   "frmConfirmDock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConfirmation 
      Height          =   720
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmConfirmDock.frx":058A
      Top             =   1815
      Width           =   2835
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   435
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1305
      Width           =   975
   End
   Begin VB.Frame fraRadioButtons 
      Height          =   1770
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   1800
      Begin VB.OptionButton rdbCurrent 
         Caption         =   "At Current Icon"
         Height          =   270
         Left            =   165
         TabIndex        =   4
         Top             =   1290
         Width           =   1425
      End
      Begin VB.OptionButton rdbPrepend 
         Caption         =   "Prepend"
         Height          =   270
         Left            =   165
         TabIndex        =   3
         Top             =   945
         Width           =   1335
      End
      Begin VB.OptionButton rdbAppend 
         Caption         =   "Append"
         Height          =   270
         Left            =   165
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton rdbOverwrite 
         Caption         =   "Overwrite"
         Height          =   270
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConfirmDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' .01 DAEB 29/05/2022 frmConfirmDock.frm Add the ability to turn the tooltips off in the generate dock utility as per ico. sett.
' .01 DAEB 29/05/2022 frmConfirmDock.frm Add balloon tooltips to the generate dock utility


Private Sub btnCancel_Click()
    frmConfirmDock.Hide
End Sub


Private Sub btnGenerate_Click()
    'msgBoxA "Note: Dock generation is not yet fully implemented", vbExclamation + vbOKOnly, "Dock Generation Tool"
    
    ' backup the current dock, use the backup code
    Call backupDockSettings
    
    ' call the generation tool code in the SoftwareList form
    Call formSoftwareList.generateDockInformation
    
End Sub


Private Sub Form_Activate()
    ' .01 DAEB 29/05/2022 frmConfirmDock.frm Add the ability to turn the tooltips off in the generate dock utility as per ico. sett.
    If rDIconConfigForm.chkToggleDialogs.Value = 0 Then
        rdbCurrent.ToolTipText = "This will insert the new items at the current dock location."
        rdbAppend.ToolTipText = "This will append the new items to the end of the existing dock."
        rdbOverwrite.ToolTipText = "This will completely overwrite the existing dock!"
        rdbPrepend.ToolTipText = "This will prepend the new items to the beginning of the existing dock."
    Else
        rdbCurrent.ToolTipText = ""
        rdbAppend.ToolTipText = ""
        rdbOverwrite.ToolTipText = ""
        rdbPrepend.ToolTipText = ""
    End If
End Sub

Private Sub Form_Load()
    Dim storedFont As String: storedFont = vbNullString
    
    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False

    'storedFont = txtTextFont.Text 'TBD
    
    fntFont = SDSuppliedFont
    fntSize = SDSuppliedFontSize
    fntItalics = CBool(SDSuppliedFontItalics)
    fntColour = CLng(SDSuppliedFontColour)
    
    ' .TBD DAEB 26/05/2022 rdIconConfig.frm Call the font tool for this form
    Call changeFont(Me, False, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)

    ' set the theme colour on startup
    Call setThemeSkin(Me)
    
    rdbCurrent.Value = True
    

End Sub

Private Sub rdbAppend_Click()
    txtConfirmation.Text = "This will append the new items to the end of the existing dock."

End Sub

Private Sub rdbCurrent_Click()
    txtConfirmation.Text = "This will insert the new items at the current dock location."
End Sub

Private Sub rdbOverwrite_Click()
    txtConfirmation.Text = "This will completely overwrite the existing dock!"
End Sub


' .02 DAEB 29/05/2022 frmConfirmDock.frm Add balloon tooltips to the generate dock utility STARTS

Private Sub btnGenerate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If rDEnableBalloonTooltips = "1" Then CreateToolTip btnGenerate.hwnd, "This button generates the new dock. Take care, this is the final step!", _
                  TTIconInfo, "Help on the Final Generate Dock button", , , , True
End Sub

Private Sub rdbAppend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbAppend.hwnd, "This radio box selects the append option when generating the new dock. Your current dock will added to on the right...", _
                  TTIconInfo, "Help on the append button", , , , True
End Sub

Private Sub rdbCurrent_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbCurrent.hwnd, "This radio box the new dock items to be added to the currently selected dock map position - go check the icon settings tool and see which has been selected...", _
                  TTIconInfo, "Help on the current position button", , , , True
End Sub


Private Sub rdbOverwrite_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbOverwrite.hwnd, "This radio box selects the overwrite option when generating the new dock. Your current dock will be lost!", _
                  TTIconInfo, "Help on the overwrite button", , , , True
End Sub

Private Sub rdbPrepend_Click()
    txtConfirmation.Text = "This will prepend the new items to the beginning of the existing dock."

End Sub

Private Sub rdbPrepend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip rdbPrepend.hwnd, "This radio box selects the prepend option when generating the new dock. Your current dock will be added to from the left...", _
                  TTIconInfo, "Help on the prepend button", , , , True
End Sub

Private Sub btnCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If rDEnableBalloonTooltips = "1" Then CreateToolTip btnCancel.hwnd, "This button cancels the generation of the new dock.", _
                  TTIconInfo, "Help on the Cancel button", , , , True

End Sub
' .02 DAEB 29/05/2022 frmConfirmDock.frm Add balloon tooltips to the generate dock utility ENDS

