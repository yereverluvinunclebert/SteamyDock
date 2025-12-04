VERSION 5.00
Begin VB.Form hiddenForm 
   Caption         =   "do not delete me as I am a temporary structure used to hold a picbox"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
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
   Begin VB.Label Label1 
      Caption         =   $"hiddenForm.frx":0000
      Height          =   795
      Left            =   3180
      TabIndex        =   6
      Top             =   1110
      Width           =   6345
   End
   Begin VB.Label Label 
      Caption         =   $"hiddenForm.frx":008A
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


' The next line implements an Interface from an External COM DLL VBSQLite12.dll,
' accepting COM QueryInterface calls for the specified interface ISQLiteProgressHandler
' which is a COM/ActiveX DLL, referenced in project references that refers itself to a raw C DLL, winsqlite3.dll registered in sysWow64 using regsvr32
' The COM object exposes a dispatch interface in its type library.

Implements ISQLiteProgressHandler ' only allowed in classes and forms (classes)

'---------------------------------------------------------------------------------------
' Procedure : ISQLiteProgressHandler_Callback
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : The SetProgressHandler method (which registers this callback) has a default value of 100 for the
'             number of virtual machine instructions that are evaluated between successive invocations of this callback.
'             This means that this callback is never invoked for very short running SQL statements.
'
'             Any running SQL operation will be interrupted if the cancel parameter is set to true.
'             This can be used to implement a "cancel" button on a GUI progress dialog box.
'
'---------------------------------------------------------------------------------------
'
Public Sub ISQLiteProgressHandler_Callback(Cancel As Boolean)

    On Error GoTo ISQLiteProgressHandler_Callback_Error

    DoEvents

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
' Author    : beededea
' Date      : 24/11/2025
' Purpose   : Returns a shared connection to the local SQLite iconDataTable.
'                If the database file doesn't exist, it is created and initialised
'                with the iconDataTable table and triggers.
'---------------------------------------------------------------------------------------
'
Private Sub CommandConnect_Click()
    Dim PathName As String: PathName = vbNullString
    
    On Error GoTo CommandConnect_Click_Error

    ' test connection to DB exists, if not then connect or create.
    If DBConnection Is Nothing Then
            
        PathName = App.Path
        If Not Right$(PathName, 1) = "\" Then PathName = PathName & "\"
        PathName = "C:\Users\beededea\AppData\Roaming\steamyDock\iconSettings.db"
        
        ' check database file exists on the system
        If fFExists(PathName) = True Then
            With New SQLiteConnection
                ' connect to SQLite db
                .OpenDB PathName, SQLiteReadWrite

                ' connection is good?
                If .hDB <> NULL_PTR Then
                    Set DBConnection = .object
                End If
            End With
        Else ' if db not exists then create it and set up the new database with hard coded schema
            If MsgBox(PathName & " does not exist. Create new?", vbExclamation + vbOKCancel) <> vbCancel Then
                Call createDBFromScratch(PathName)
            End If
        End If
        
        With DBConnection
            .SetProgressHandler Me ' Registers the progress handler callback
        End With
        
        CommandInsert.Enabled = True
        List1.Enabled = True
        List1.Clear
        Call Requery
    Else
        MsgBox "Already connected.", vbExclamation
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
    
    List1.Clear
    Call insertRecords

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

    Call closeDatabase
    hiddenForm.CommandInsert.Enabled = False
    hiddenForm.List1.Clear
    hiddenForm.List1.Enabled = False

    On Error GoTo 0
    Exit Sub

CommandClose_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CommandClose_Click of Form hiddenForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : List1_KeyDown
' Author    : beededea
' Date      : 02/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim keyToDelete As String: keyToDelete = vbNullString
    
    On Error GoTo List1_KeyDown_Error

    If List1.ListCount > 0 Then
        If KeyCode = vbKeyDelete Then
            keyToDelete = Left$(List1.Text, InStr(List1.Text, "_") - 1)
        
            If MsgBox("Delete record number " & keyToDelete & " " & List1.Text & "?", vbQuestion + vbYesNo) <> vbNo Then
                Call deleteSpecificKey(keyToDelete)
                List1.RemoveItem List1.ListIndex
            End If
        End If
    End If
    
    On Error GoTo 0
    Exit Sub

List1_KeyDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure List1_KeyDown of Form hiddenForm"

End Sub




