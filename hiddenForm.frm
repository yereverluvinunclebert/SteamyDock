VERSION 5.00
Begin VB.Form hiddenForm 
   Caption         =   "do not delete me as I am a temporary structure used to hold a picbox"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   510
      TabIndex        =   4
      Top             =   4230
      Width           =   4695
   End
   Begin VB.CommandButton CommandClose 
      Caption         =   "Close Test.db"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   3330
      Width           =   1455
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert into test_table"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2220
      TabIndex        =   2
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton CommandConnect 
      Caption         =   "Connect Test.db"
      Height          =   615
      Left            =   540
      TabIndex        =   1
      Top             =   3300
      Width           =   1455
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
   Begin VB.Label Label 
      Caption         =   $"hiddenForm.frx":0000
      Height          =   795
      Left            =   3240
      TabIndex        =   5
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
#If (VBA7 = 0) Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
#If Win64 Then
    Private Const NULL_PTR As LongPtr = 0 ' this may glow red but is NOT an error, suitable for 64bit TwinBasic
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const NULL_PTR As Long = 0
    Private Const PTR_SIZE As Long = 4
#End If

Implements ISQLiteProgressHandler

Private DBConnection As SQLiteConnection  ' requires the SQLLite project reference VBSQLLite12.DLL

'---------------------------------------------------------------------------------------
' Procedure : ISQLiteProgressHandler_Callback
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ISQLiteProgressHandler_Callback(Cancel As Boolean)
' The SetProgressHandler method (which registers this callback) has a default value of 100 for the
' number of virtual machine instructions that are evaluated between successive invocations of this callback.
' This means that this callback is never invoked for very short running SQL statements.
    On Error GoTo ISQLiteProgressHandler_Callback_Error

DoEvents
' The operation will be interrupted if the cancel parameter is set to true.
' This can be used to implement a "cancel" button on a GUI progress dialog box.

    On Error GoTo 0
    Exit Sub

ISQLiteProgressHandler_Callback_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ISQLiteProgressHandler_Callback of Form hiddenForm"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Error

    If Not DBConnection Is Nothing Then DBConnection.CloseDB

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form hiddenForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : CommandConnect_Click
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
' Returns a shared connection to the local SQLite iconData.
' If the database file doesn't exist, it is created and initialized
' with the iconData table and triggers.
'---------------------------------------------------------------------------------------
'
Private Sub CommandConnect_Click()
    On Error GoTo CommandConnect_Click_Error

    If DBConnection Is Nothing Then
        Dim PathName As String
        PathName = App.Path
        If Not Right$(PathName, 1) = "\" Then PathName = PathName & "\"
    
        PathName = "C:\Users\beededea\AppData\Roaming\steamyDock\iconSettings.db"
        If fFExists(PathName) Then
            With New SQLiteConnection
                On Error Resume Next
                .OpenDB PathName, SQLiteReadWrite
                If Err.Number <> 0 Then
                    Err.Clear
                    If MsgBox("iconSettings.db does not exist. Create new?", vbExclamation + vbOKCancel) <> vbCancel Then
                        .OpenDB PathName & "iconSettings.db", SQLiteReadWriteCreate
    
                        ' Create main iconData table:
                        '  - key: logical identifier (case-insensitive primary key)
                        '  - data: payload stored as BLOB
                        '  - update_counter: monotonically increasing integer, used for change tracking
                        .Execute _
                            "CREATE TABLE iconData (" & _
                            " key TEXT NOT NULL COLLATE NOCASE," & _
                            " data BLOB NOT NULL," & _
                            " update_counter INTEGER NOT NULL DEFAULT 0," & _
                            " PRIMARY KEY(key))"
    
                        ' Trigger to bump update_counter on UPDATE of data:
                        '  - AFTER UPDATE OF data: only fires when the data column changes
                        '  - Sets update_counter to current max(update_counter)+1 across the table
                        '  - WHERE key = NEW.key ensures only the updated row is changed
                        .Execute _
                              "CREATE TRIGGER iconData_update_counter " & _
                              "AFTER UPDATE OF data ON iconData " & _
                              "FOR EACH ROW " & _
                              "BEGIN " & _
                              "  UPDATE iconData " & _
                              "  SET update_counter = (SELECT COALESCE(MAX(update_counter), 0) + 1 FROM iconData) " & _
                              "  WHERE key = NEW.key; " & _
                              "END;"
    
                        ' Trigger to bump update_counter on INSERT:
                        '  - AFTER INSERT: runs after the row is inserted
                        '  - Sets update_counter for just the new row (NEW.key)
                        '  - Uses same global max(update_counter)+1 logic
                        .Execute _
                              "CREATE TRIGGER iconData_insert_counter " & _
                              "AFTER INSERT ON iconData " & _
                              "BEGIN " & _
                              "  UPDATE iconData " & _
                              "  SET update_counter = (SELECT COALESCE(MAX(update_counter), 0) + 1 FROM iconData) " & _
                              "  WHERE key = NEW.key; " & _
                              "END;"
    
                    End If
                End If
                On Error GoTo 0
                If .hDB <> NULL_PTR Then
                    Set DBConnection = .Object
                    .SetProgressHandler Me ' Registers the progress handler callback
                    CommandInsert.Enabled = True
                    List1.Enabled = True
                    Call Requery
                End If
            End With
        End If
    Else
        MsgBox "Already connected.", vbExclamation
    End If

    On Error GoTo 0
    Exit Sub

CommandConnect_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CommandConnect_Click of Form hiddenForm"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : CommandInsert_Click
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CommandInsert_Click()
    Dim Text As String
    
    Text = VBA.InputBox("szText")
    If StrPtr(Text) = NULL_PTR Then Exit Sub
    On Error GoTo CATCH_EXCEPTION
    With DBConnection
    .Execute "INSERT INTO test_table (szText) VALUES ('" & Text & "')"
    End With
    Call Requery
    Exit Sub
CATCH_EXCEPTION:
    MsgBox Err.Description, vbCritical + vbOKOnly

    On Error GoTo 0
    Exit Sub

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Requery
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Requery()

    On Error GoTo CATCH_EXCEPTION
    List1.Clear
    Dim DataSet As SQLiteDataSet
    Set DataSet = DBConnection.OpenDataSet("SELECT ID, szText FROM icondata")
    DataSet.MoveFirst
    Do Until DataSet.EOF
        List1.AddItem DataSet!id & "_" & DataSet!szText
        DataSet.MoveNext
    Loop
    Exit Sub
CATCH_EXCEPTION:
    MsgBox Err.Description, vbCritical + vbOKOnly

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

    If DBConnection Is Nothing Then
        MsgBox "Not connected.", vbExclamation
    Else
        DBConnection.SetProgressHandler Nothing ' Unregisters the progress handler callback
        DBConnection.CloseDB
        Set DBConnection = Nothing
        CommandInsert.Enabled = False
        List1.Clear
        List1.Enabled = False
    End If

    On Error GoTo 0
    Exit Sub

CommandClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CommandClose_Click of Form hiddenForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : List1_KeyDown
' Author    : Krool
' Date      : 24/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo CATCH_EXCEPTION
    If List1.ListCount > 0 Then
        If KeyCode = vbKeyDelete Then
            If MsgBox("Delete?", vbQuestion + vbYesNo) <> vbNo Then
                Dim Command As SQLiteCommand
                Set Command = DBConnection.CreateCommand("DELETE FROM iconData WHERE ID = @oid")
                Command.SetParameterValue Command![@oid], Left$(List1.Text, InStr(List1.Text, "_") - 1)
                Command.Execute
                List1.RemoveItem List1.ListIndex
            End If
        End If
    End If
    Exit Sub
CATCH_EXCEPTION:
    MsgBox Err.Description, vbCritical + vbOKOnly

    On Error GoTo 0
    Exit Sub

End Sub




