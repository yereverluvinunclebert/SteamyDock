VERSION 5.00
Begin VB.Form hiddenForm 
   Caption         =   "do not delete me as I am a temporary structure used to hold a picbox"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton btnUpdateSingle 
      Caption         =   "Update single record"
      Height          =   705
      Left            =   2280
      TabIndex        =   17
      Top             =   5790
      Width           =   1455
   End
   Begin VB.CommandButton btnWriteRandom 
      Caption         =   "Write Random File"
      Enabled         =   0   'False
      Height          =   705
      Left            =   540
      TabIndex        =   16
      Top             =   5790
      Width           =   1455
   End
   Begin VB.TextBox txtSingleField 
      Enabled         =   0   'False
      Height          =   345
      Left            =   2760
      TabIndex        =   15
      Top             =   5040
      Width           =   3735
   End
   Begin VB.ComboBox cmbSingleFieldRecordNumber 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "hiddenForm.frx":0000
      Left            =   2070
      List            =   "hiddenForm.frx":00B5
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5070
      Width           =   675
   End
   Begin VB.CommandButton lblGetField 
      Caption         =   "Get Single Record, One Field"
      Enabled         =   0   'False
      Height          =   645
      Left            =   510
      TabIndex        =   13
      Top             =   4950
      Width           =   1485
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   675
      Left            =   7830
      TabIndex        =   12
      Top             =   5820
      Width           =   1785
   End
   Begin VB.ComboBox cmbRecordNumber 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "hiddenForm.frx":01A5
      Left            =   2070
      List            =   "hiddenForm.frx":025A
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4350
      Width           =   675
   End
   Begin VB.TextBox txtSingleRecord 
      Enabled         =   0   'False
      Height          =   345
      Left            =   2760
      TabIndex        =   10
      Top             =   4350
      Width           =   3735
   End
   Begin VB.CommandButton lblGetRecord 
      Caption         =   "Get Single Record All Fields"
      Enabled         =   0   'False
      Height          =   645
      Left            =   510
      TabIndex        =   9
      Top             =   4200
      Width           =   1485
   End
   Begin VB.CommandButton Command 
      Caption         =   "Kill .db "
      Height          =   615
      Left            =   3570
      TabIndex        =   8
      Top             =   3300
      Width           =   1425
   End
   Begin VB.CommandButton CommandClose 
      Caption         =   "Close.db"
      Height          =   615
      Left            =   5010
      TabIndex        =   3
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert fresh data into .db"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2070
      TabIndex        =   2
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton CommandConnect 
      Caption         =   "Connect.db && get multiple records"
      Height          =   615
      Left            =   540
      TabIndex        =   1
      Top             =   3300
      Width           =   1485
   End
   Begin VB.PictureBox hiddenPicbox 
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   480
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "The buttons below will connect tot he db, clear it down and reload fresh from schema and close the db."
      Height          =   795
      Left            =   3240
      TabIndex        =   7
      Top             =   2100
      Width           =   6345
   End
   Begin VB.Label lblRecordNum 
      Caption         =   "0"
      Height          =   225
      Left            =   540
      TabIndex        =   6
      Top             =   2940
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   $"hiddenForm.frx":034A
      Height          =   795
      Left            =   3180
      TabIndex        =   5
      Top             =   1110
      Width           =   6345
   End
   Begin VB.Label Label 
      Caption         =   $"hiddenForm.frx":03D4
      Height          =   795
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   6345
   End
End
Attribute VB_Name = "hiddenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : hiddenForm
' Author    : beededea
' Date      : 24/11/2025
' Purpose   : Temporary structure used to hold a picbox used when extracting and converting an icon to a PNG
'             as well as temporarily holding some code for doing trial database operations
'---------------------------------------------------------------------------------------

Option Explicit

Private mIsLoaded As Boolean ' property

'#If (VBA7 = 0) Then
'    Private Enum LongPtr
'        [_]
'    End Enum
'#End If
'#If Win64 Then
'    Private Const NULL_PTR As LongPtr = 0 ' this may glow red but is NOT an error, suitable for 64bit TwinBasic
'    Private Const PTR_SIZE As Long = 8
'#Else
'    Private Const NULL_PTR As Long = 0
'    Private Const PTR_SIZE As Long = 4
'#End If


' The next line implements an Interface from an External COM DLL VBSQLite12.dll,
' accepting COM QueryInterface calls for the specified interface ISQLiteProgressHandler
' which is a COM/ActiveX DLL, referenced in project references that refers itself to a raw C DLL, winsqlite3.dll registered in sysWow64 using regsvr32
' The COM object exposes a dispatch interface in its type library.

'Implements ISQLiteProgressHandler ' only allowed in classes and forms (classes)

'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 04/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()

    On Error GoTo btnClose_Click_Error

    Unload hiddenForm

    On Error GoTo 0
    Exit Sub

btnClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdateSingle_Click
' Author    : beededea
' Date      : 07/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdateSingle_Click()

    On Error GoTo btnUpdateSingle_Click_Error

    Call UPDATEFieldInSingleRecord(3, "fIconTitle", "arseburgers")

    On Error GoTo 0
    Exit Sub

btnUpdateSingle_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdateSingle_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnWriteRandom_Click
' Author    : beededea
' Date      : 05/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnWriteRandom_Click()

    Dim srcFile As String
    Dim trgtFile As String

    On Error GoTo btnWriteRandom_Click_Error

    srcFile = SpecialFolder(SpecialFolder_AppData) & "\steamyDock\iconSettings.dat"
    trgtFile = SpecialFolder(SpecialFolder_AppData) & "\steamyDock\iconSettings.bkp"
    
    'List1.Clear
    
    'FileCopy srcFile, trgtFile
    
    lblRecordNum.Caption = "Reading from Database, inserting into file."

    Call insertAllFieldsIntoRandomDataFile

    lblRecordNum.Caption = "Inserting into file complete."

    On Error GoTo 0
    Exit Sub

btnWriteRandom_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnWriteRandom_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Command_Click
' Author    : beededea
' Date      : 04/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Command_Click()

    Dim ans As VbMsgBoxResult
    
    On Error GoTo Command_Click_Error
    
    ans = MsgBox("This will close and then remove the database completely by deleting it, are you sure you wish to do this?")
    If ans = vbNo Then Exit Sub

    Call closeDatabase
    
    Kill SpecialFolder(SpecialFolder_AppData) & "\steamyDock\iconSettings.db"
    
    lblRecordNum.Caption = "Database Deleted"
    
    CommandInsert.Enabled = False
    'List1.Clear
    'List1.Enabled = False

    On Error GoTo 0
    Exit Sub

Command_Click_Error:

     MsgBox "Error " & Err.Number & " Database could not be deleted, probably due to it being open by the DB browser or other instance of this program. (" & Err.Description & ") in procedure Command_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblGetField_Click
' Author    : beededea
' Date      : 04/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGetField_Click()

    Dim recordToFind As Integer: recordToFind = 0
    
    On Error GoTo lblGetField_Click_Error

    recordToFind = CInt(cmbSingleFieldRecordNumber.List(cmbSingleFieldRecordNumber.ListIndex))

    txtSingleField.Text = getSingleFieldFromSingleRecord("fIconCommand", recordToFind)

    On Error GoTo 0
    Exit Sub

lblGetField_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGetField_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 04/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    On Error GoTo Form_Load_Error

    IsLoaded = True
    cmbRecordNumber.ListIndex = 0
    cmbSingleFieldRecordNumber.ListIndex = 0
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form hiddenForm"
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : ISQLiteProgressHandler_Callback
'' Author    : jbPro
'' Date      : 24/11/2025
'' Purpose   : The SetProgressHandler method (which registers this callback) has a default value of 100 for the
''             number of virtual machine instructions that are evaluated between successive invocations of this callback.
''             This means that this callback is never invoked for very short running SQL statements.
''
''             Any running SQL operation will be interrupted if the cancel parameter is set to true.
''             This can be used to implement a "cancel" button on a GUI progress dialog box.
''
''---------------------------------------------------------------------------------------
''
'Public Sub ISQLiteProgressHandler_Callback(Cancel As Boolean)
'
'    On Error GoTo ISQLiteProgressHandler_Callback_Error
'
'    DoEvents
'
'    On Error GoTo 0
'    Exit Sub
'
'ISQLiteProgressHandler_Callback_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ISQLiteProgressHandler_Callback of Form hiddenForm"
'End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Error

    'If Not DBConnection Is Nothing Then DBConnection.CloseDB
    
    'Call closeDatabase

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form hiddenForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : CommandConnect_Click
' Author    : beededea
' Date      : 24/11/2025
' Purpose   : Returns a shared connection to the local SQLite iconDataTable.
'                If the database file doesn't exist, it is created and initialised
'                with the iconDataTable table and triggers.
'---------------------------------------------------------------------------------------
'
Private Sub CommandConnect_Click()
    Dim PathName As String: PathName = vbNullString
    Dim retResult As String: retResult = vbNullString
    
    On Error GoTo CommandConnect_Click_Error
    
    ' test connection to DB exists, if not then connect or create.
    'retResult = connectDatabase
   
'    If DBConnection Is Nothing Then
'
'        PathName = App.Path
'        If Not Right$(PathName, 1) = "\" Then PathName = PathName & "\"
'        PathName = "C:\Users\beededea\AppData\Roaming\steamyDock\iconSettings.db"
'
'        ' check database file exists on the system
'        If fFExists(PathName) = True Then
'            With New SQLiteConnection
'                ' connect to SQLite db
'                .OpenDB PathName, SQLiteReadWrite
'
'                ' connection is good?
'                If .hDB <> NULL_PTR Then
'                    Set DBConnection = .object
'                End If
'            End With
'            hiddenForm.lblRecordNum.Caption = "Database Connected."
'
'        Else ' if db not exists then create it and set up the new database with hard coded schema
'            If MsgBox(PathName & " does not exist. Create new?", vbExclamation + vbOKCancel) <> vbCancel Then
'                Call createUnpopulatedDBFromSchema(PathName)
'
'            Else
'                Exit Sub
'            End If
'        End If
'    End If
    
  
        
    If DBConnection Is Nothing Then
        ' do nothing
    Else
    
'        With DBConnection
'            .SetProgressHandler Me ' Registers the progress handler callback
'        End With

        hiddenForm.lblRecordNum.Caption = "Database Connected."
    
        CommandInsert.Enabled = True
'        List1.Enabled = True
'        List1.Clear

        lblGetRecord.Enabled = True
        cmbRecordNumber.Enabled = True
        txtSingleRecord.Enabled = True
        
        lblGetField.Enabled = True
        cmbSingleFieldRecordNumber.Enabled = True
        txtSingleField.Enabled = True
        
        btnWriteRandom.Enabled = True
        
        Call getSingleFieldFromMultipleRecords("fIconTitle")
        
'    Else
'        MsgBox "Already connected.", vbExclamation

    End If
    


    On Error GoTo 0
    Exit Sub

CommandConnect_Click_Error:

     MsgBox "Error " & PathName & " has a problem." & Err.Number & " (" & Err.Description & ") in procedure CommandConnect_Click of Form hiddenForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : CommandInsert_Click
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CommandInsert_Click()
    
    Dim ans As VbMsgBoxResult
    
'    List1.Clear
    CommandInsert.Enabled = False
'    List1.Enabled = False
        
    ans = MsgBox("This will remove all the icon elements from the database, are you sure you wish to do this?")
    If ans = vbNo Then Exit Sub

    Call insertRecordsFromRandomDataFileIntoDatabase
    hiddenForm.lblRecordNum.Caption = "Data Inserted."

    On Error GoTo 0
    Exit Sub

End Sub






'---------------------------------------------------------------------------------------
' Procedure : CommandClose_Click
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CommandClose_Click()
    On Error GoTo CommandClose_Click_Error
    
    Dim ans As VbMsgBoxResult
        
    ans = MsgBox("This will close the database which will have a negative effect on the dock and any current dock operations, are you sure you wish to do this?")
    If ans = vbNo Then Exit Sub
    
    If DBConnection Is Nothing Then
        MsgBox "Not connected.", vbExclamation
    Else
        Call closeDatabase
        CommandInsert.Enabled = False
'        List1.Clear
'        List1.Enabled = False
    End If
    
    hiddenForm.lblRecordNum.Caption = "Database closed."

    On Error GoTo 0
    Exit Sub

CommandClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CommandClose_Click of Form hiddenForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblGetRecord_Click
' Author    : beededea
' Date      : 04/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGetRecord_Click()

    Dim recordToFind As Integer: recordToFind = 0
    
    On Error GoTo lblGetRecord_Click_Error
    
    recordToFind = CInt(cmbRecordNumber.List(cmbRecordNumber.ListIndex))

    txtSingleRecord.Text = getAllFieldsFromSingleRecord(recordToFind)

    On Error GoTo 0
    Exit Sub

lblGetRecord_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGetRecord_Click of Form hiddenForm"
    
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : List1_KeyDown
'' Author    : beededea
'' Date      : 02/12/2025
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Dim keyToDelete As String: keyToDelete = vbNullString
'
'    On Error GoTo List1_KeyDown_Error
'
'    If List1.ListCount > 0 Then
'        If KeyCode = vbKeyDelete Then
'            keyToDelete = Left$(List1.Text, InStr(List1.Text, "_") - 1)
'
'            If MsgBox("Delete record number " & keyToDelete & " " & List1.Text & "?", vbQuestion + vbYesNo) <> vbNo Then
'                Call deleteSpecificKey(keyToDelete)
'                List1.RemoveItem List1.ListIndex
'            End If
'        End If
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'List1_KeyDown_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure List1_KeyDown of Form hiddenForm"
'
'End Sub




'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
 
   On Error GoTo IsLoaded_Error

    IsLoaded = mIsLoaded
    

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Let IsLoaded(ByVal newValue As Boolean)
 
   On Error GoTo IsLoaded_Error

   If mIsLoaded <> newValue Then mIsLoaded = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property
