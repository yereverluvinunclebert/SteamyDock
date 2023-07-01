VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "SteamyDock Enhanced Icon Settings Tool"
   ClientHeight    =   2100
   ClientLeft      =   4845
   ClientTop       =   4800
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "message.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMessage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   -30
      TabIndex        =   2
      Top             =   0
      Width           =   5970
      Begin VB.Frame fraPicVB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   195
         TabIndex        =   4
         Top             =   270
         Width           =   735
         Begin VB.Image picVBInformation 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":030A
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBCritical 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":14F4
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image picVBExclamation 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":26DC
            Top             =   0
            Width           =   720
         End
         Begin VB.Image picVBQuestion 
            Height          =   720
            Left            =   0
            Picture         =   "message.frx":3914
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "*"
         Height          =   195
         Left            =   1110
         TabIndex        =   3
         Top             =   570
         Width           =   4455
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton btnButtonTwo 
      Caption         =   "&No"
      Height          =   372
      Left            =   4980
      TabIndex        =   1
      Top             =   1620
      Width           =   972
   End
   Begin VB.CommandButton btnButtonOne 
      Caption         =   "&Yes"
      Height          =   372
      Left            =   3885
      TabIndex        =   0
      Top             =   1620
      Width           =   972
   End
   Begin VB.CheckBox chkShowAgain 
      Caption         =   "&Hide this message for the rest of this session"
      Height          =   420
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3435
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen STARTS
Option Explicit
Private mintLabelHeight As Integer
Private yesNoReturnValue As Integer
Private formMsgContext As String
Private formShowAgainChkBox As Boolean

Private Sub btnButtonTwo_Click()
    If formShowAgainChkBox = True Then SaveSetting App.EXEName, "Options", "Show message" & formMsgContext, chkShowAgain.Value
    yesNoReturnValue = 7
    Unload Me
End Sub

Private Sub btnButtonOne_Click()
    Me.Visible = False
    If formShowAgainChkBox = True Then SaveSetting App.EXEName, "Options", "Show message" & formMsgContext, chkShowAgain.Value
    yesNoReturnValue = 6
    Unload Me
End Sub

Public Sub Display()

    Dim intShow As Integer
    
    If formShowAgainChkBox = True Then
    
        chkShowAgain.Visible = True
        ' Returns a key setting value from an application's entry in the Windows registry
        intShow = GetSetting(App.EXEName, "Options", "Show message" & formMsgContext, vbUnchecked)
        
        If intShow = vbUnchecked Then
            Me.Show vbModal
        End If
    Else
        Me.Show vbModal
    End If

End Sub
' property to allow a message to be passed to the form
Public Property Let propMessage(ByVal strMessage As String)

    Dim intDiff As Integer
    
    lblMessage.Caption = strMessage
    
    ' Expand the form and move the other controls if the message is too long to show.
    intDiff = lblMessage.Height - mintLabelHeight
    Me.Height = Me.Height + intDiff
    
    fraMessage.Height = fraMessage.Height + intDiff
    fraPicVB.Top = fraPicVB.Top + (intDiff / 2)
        
    chkShowAgain.Top = chkShowAgain.Top + intDiff
    btnButtonOne.Top = btnButtonOne.Top + intDiff
    btnButtonTwo.Top = btnButtonTwo.Top + intDiff

End Property

Public Property Let propTitle(ByVal strTitle As String)
    If strTitle = "" Then
        frmMessage.Caption = "SteamyDock Icon Enhanced Settings"
    Else
        frmMessage.Caption = strTitle
    End If
End Property

Public Property Let propMsgContext(ByVal thisContext As String)
    formMsgContext = thisContext
End Property

Public Property Let propShowAgainChkBox(ByVal showAgainVis As Boolean)
    formShowAgainChkBox = showAgainVis
End Property

Public Property Let propButtonVal(ByVal buttonVal As Integer)
    
    Dim fileToPlay As String: fileToPlay = vbNullString

    btnButtonOne.Visible = False
    btnButtonTwo.Visible = False
    'btnButtonThree.Visible = false

    picVBInformation.Visible = False
    picVBCritical.Visible = False
    picVBExclamation.Visible = False
    picVBQuestion.Visible = False

    btnButtonOne.Left = 3885

    If buttonVal >= 64 Then ' vbInformation
       buttonVal = buttonVal - 64
       picVBInformation.Visible = True
    ElseIf buttonVal >= 48 Then '    vbExclamation
        buttonVal = buttonVal - 48
        picVBExclamation.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "ting.wav"
        If FExists(App.Path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    ElseIf buttonVal >= 32 Then '    vbQuestion
        buttonVal = buttonVal - 32
        picVBQuestion.Visible = True
    ElseIf buttonVal >= 20 Then '    vbCritical
        buttonVal = buttonVal - 20
        picVBCritical.Visible = True
        
        ' .86 DAEB 06/06/2022 rDIConConfig.frm Add a sound to the msgbox for critical and exclamations? ting and belltoll.wav files
        fileToPlay = "belltoll01.wav"
        If FExists(App.Path & "\resources\sounds\" & fileToPlay) Then
            PlaySound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    End If

    If buttonVal = 2 Then 'vbAbortRetryIgnore 2
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Abort"
        btnButtonOne.Caption = "Retry"
        'btnButtonThree.Caption = "Ignore"
    End If
    If buttonVal = 0 Then '    vbOKOnly 0
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = False
        btnButtonOne.Caption = "OK"
        btnButtonOne.Left = 4620
    End If
    If buttonVal = 1 Then '    vbOKCancel 1
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "OK"
        btnButtonTwo.Caption = "Cancel"
    End If
    If buttonVal = 2 Then '    vbCancel 2
        btnButtonOne.Visible = False
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = ""
        btnButtonTwo.Caption = "Cancel"
    End If
    If buttonVal = 3 Then '    vbYesNoCancel 3
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        'btnButtonThree.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
        'btnButtonThree.Caption = "Cancel"
    End If
    If buttonVal = 4 Then '    vbYesNo 4
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Yes"
        btnButtonTwo.Caption = "No"
    End If
    If buttonVal = 5 Then '    vbRetryCancel 5
        btnButtonOne.Visible = True
        btnButtonTwo.Visible = True
        btnButtonOne.Caption = "Retry"
        btnButtonTwo.Caption = "Cancel"
    End If

        
End Property

Public Property Get propReturnedValue()

    propReturnedValue = yesNoReturnValue
    
End Property


Private Sub Form_Load()

    Dim Ctrl As Control

    mintLabelHeight = lblMessage.Height
        
    ' .TBD DAEB 05/05/2021 frmMessage.frm Added the font mod. here instead of within the changeFont tool
    '                       as each instance of the form is new, the font modification must be here.
    For Each Ctrl In Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
           If SDSuppliedFont <> "" Then Ctrl.Font.Name = SDSuppliedFont
           If Val(SDSuppliedFontSize) > 0 Then Ctrl.Font.Size = Val(SDSuppliedFontSize)
                       'Ctrl.Font.Italic = CBool(SDSuppliedFontItalics) TBD
           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
        End If
    Next

    chkShowAgain.Visible = False
    
End Sub
' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen ENDS
