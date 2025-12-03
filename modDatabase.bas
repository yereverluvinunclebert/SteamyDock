Attribute VB_Name = "modDatabase"
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

Public DBConnection As SQLiteConnection  ' requires the SQLLite project reference VBSQLLite12.DLL

' database schema (simplified)
'       iconRecordNumber As Integer
'       iconFilename As String
'       iconFileName2 As String
'       iconTitle As String
'       iconCommand As String
'       iconArguments As String
'       iconWorkingDirectory As String
'       iconShowCmd As String
'       iconOpenRunning As String
'       iconIsSeparator As String
'       iconUseContext As String
'       iconDockletFile As String
'       iconUseDialog As String
'       iconUseDialogAfter As String
'       iconQuickLaunch As String
'       iconAutoHideDock As String
'       iconSecondApp As String
'       iconRunElevated As String
'       iconRunSecondAppBeforehand As String
'       iconAppToTerminate As String
'       iconDisabled As String

'---------------------------------------------------------------------------------------
' Procedure : SaveToiconData
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Inserts or updates a single key/value pair in the iconDataTable.
'             On CONFLICT(key), the row is updated instead of inserting a duplicate.
'             The triggers on the table ensure update_counter is bumped appropriately.
'---------------------------------------------------------------------------------------
'
Public Sub SaveToiconData(ByVal p_Key As String, p_Data As Variant)
    On Error GoTo SaveToiconData_Error
    
    With DBConnection
        .Execute "INSERT INTO iconDataTable (key, data) VALUES ('" & p_Key & "','" & p_Data & "') ON CONFLICT (key) DO UPDATE SET data=excluded.data"
    End With
   
    On Error GoTo 0
    Exit Sub

SaveToiconData_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveToiconData of Form hiddenForm"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : GetFromiconData
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Retrieves the data BLOB for a given key.
'             Raises error 5 if the key is not found.
'---------------------------------------------------------------------------------------
'
Public Function GetFromiconData(ByVal p_Key As String) As Variant
    On Error GoTo GetFromiconData_Error

    Dim DataSet As SQLiteDataSet
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable  WHERE key= " & p_Key)
    
    ' No matching row: raise a generic "Invalid procedure call or argument" (5)
    ' with a more descriptive message.
    If DataSet.RecordCount = 0 Then
       Err.Raise 5, , "Data not found for key " & p_Key
    End If
    
    ' Return the first (and only) column: data
    GetFromiconData = DataSet.Columns(0).Value

    On Error GoTo 0
    Exit Function

GetFromiconData_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFromiconData of Form hiddenForm"
End Function



'---------------------------------------------------------------------------------------
' Procedure : MaxUpdateCounter
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Returns the maximum update_counter value.
' Using Currency here gives enough range for SQLite INTEGER values.
' Returns 0 if table is empty or MAX() is NULL.
'---------------------------------------------------------------------------------------
'
Public Function MaxUpdateCounter() As Currency
    On Error GoTo MaxUpdateCounter_Error
    
    Dim DataSet As SQLiteDataSet
    Set DataSet = DBConnection.OpenDataSet("SELECT MAX(update_counter) FROM iconDataTable")

    If DataSet.RecordCount > 0 Then
       ' If the result is NULL, this will default to 0 when assigned to Currency.
       MaxUpdateCounter = DataSet.Columns(0)
    End If

    On Error GoTo 0
    Exit Function

MaxUpdateCounter_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MaxUpdateCounter of Form hiddenForm"
End Function






'---------------------------------------------------------------------------------------
' Procedure : GetDataSinceUpdateCounter
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Returns all rows with update_counter greater than the specified value.
'
' Result:
'   - A collection of:
'       Keys: iconDataTable key (String)
'       Items: iconDataTable data (Variant / BLOB)
'---------------------------------------------------------------------------------------
'
Public Function GetDataSinceUpdateCounter(ByVal p_UpdateCounter As Currency)
    On Error GoTo GetDataSinceUpdateCounter_Error
    
    Dim DataSet As SQLiteDataSet
    Set DataSet = DBConnection.OpenDataSet("SELECT key, data FROM iconDataTable WHERE update_counter>" & p_UpdateCounter)
        
    'dictionary for the database access
    Set GetDataSinceUpdateCounter = CreateObject("Scripting.Dictionary")
    GetDataSinceUpdateCounter.CompareMode = 1 'case-insenitive Key-Comparisons
    
    ' Select only rows whose update_counter is greater than the given value
    With DataSet
       Do Until .EOF
          GetDataSinceUpdateCounter.Add .Columns("data").Value, .Columns("key").Value
          
          .MoveNext
       Loop
    End With

    On Error GoTo 0
    Exit Function

GetDataSinceUpdateCounter_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDataSinceUpdateCounter of Form hiddenForm"
End Function




' END of JBPro's suggested subs and functions




'---------------------------------------------------------------------------------------
' Procedure : closeDatabase
' Author    : Krool
' Date      : 01/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub closeDatabase()
    On Error GoTo closeDatabase_Error

    If DBConnection Is Nothing Then
        MsgBox "Not connected.", vbExclamation
    Else
        DBConnection.SetProgressHandler Nothing ' Unregisters the progress handler callback
        DBConnection.CloseDB
        Set DBConnection = Nothing
        hiddenForm.CommandInsert.Enabled = False
        hiddenForm.List1.Clear
        hiddenForm.List1.Enabled = False
    End If

    On Error GoTo 0
    Exit Sub

closeDatabase_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure closeDatabase of Form hiddenForm"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : insertRecords
' Author    : Krool
' Date      : 01/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub insertRecords()

    Dim Text As String: Text = vbNullString
    
    On Error GoTo insertRecords_Error

    Text = VBA.InputBox("iconCommand")
    If StrPtr(Text) = NULL_PTR Then Exit Sub

    With DBConnection
        .Execute "INSERT INTO iconDataTable (iconCommand) VALUES ('" & Text & "')"
    End With
    Call Requery

    On Error GoTo 0
    Exit Sub

insertRecords_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertRecords of Form hiddenForm"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : Requery
' Author    : beededea
' Date      : 02/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Requery()

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo Requery_Error

    Set DataSet = DBConnection.OpenDataSet("SELECT key, iconCommand FROM iconDataTable")
    DataSet.MoveFirst
    Do Until DataSet.EOF
        hiddenForm.List1.AddItem DataSet!key & "_" & DataSet!iconCommand
        DataSet.MoveNext
    Loop

    On Error GoTo 0
    Exit Sub

Requery_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Requery of Module modDatabase"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : deleteSpecificKey
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : need to test Krool's original
'---------------------------------------------------------------------------------------
'
Public Sub deleteSpecificKey(ByVal keyToDelete As String)
            
    Dim Command As SQLiteCommand
    
    On Error GoTo deleteSpecificKey_Error
    
    If keyToDelete = vbNullString Then Exit Sub

    Set Command = DBConnection.CreateCommand("DELETE FROM iconDataTable WHERE key = @oid")
    Command.SetParameterValue Command![@oid], keyToDelete
    Command.Execute

    On Error GoTo 0
    Exit Sub

deleteSpecificKey_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deleteSpecificKey of Module modDatabase"
        
End Sub




'---------------------------------------------------------------------------------------
' Procedure : createDBFromScratch
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : used to recreate the databse from scratch if required
'             In any case, this code is NOT required as an empty databse is never going to be shipped with the program.
'             This is just retained retained here for later investigation and for education (mine).
'---------------------------------------------------------------------------------------
'
Public Sub createDBFromScratch(ByVal pathToFile As String)
    
    On Error GoTo createDBFromScratch_Error

    With New SQLiteConnection

         .OpenDB pathToFile, SQLiteReadWriteCreate

         ' Create main iconDataTable table:
         '  - key: logical identifier (case-insensitive primary key)
         '  - data: payload stored as text in general
         
         ' note some lines have been concatentated as VB6 does not allow more than a certain number of line continuations
         .Execute _
             "CREATE TABLE iconDataTable (" & _
             " key TEXT COLLATE NOCASE," & _
             " iconRecordNumber INTEGER DEFAULT 0, iconFilename TEXT," & _
             " iconFileName2 TEXT," & _
             " iconTitle TEXT," & _
             " iconCommand TEXT," & _
             " iconArguments TEXT," & _
             " iconWorkingDirectory TEXT," & _
             " iconShowCmd TEXT," & _
             " iconOpenRunning TEXT," & _
             " iconIsSeparator TEXT," & _
             " iconUseContext TEXT," & _
             " iconDockletFile TEXT," & _
             " iconUseDialog TEXT," & _
             " iconUseDialogAfter TEXT," & _
             " iconQuickLaunch TEXT," & _
             " iconAutoHideDock TEXT," & _
             " iconSecondApp TEXT," & _
             " iconRunElevated TEXT," & _
             " iconRunSecondAppBeforehand TEXT," & _
             " iconAppToTerminate TEXT," & _
             " iconDisabled TEXT," & _
             " PRIMARY KEY(key))"
             
         ' Create updateTable table:
         '  - update_counter: monotonically increasing integer, used for change tracking
        .Execute _
             "CREATE TABLE updateTable (" & _
             " key TEXT COLLATE NOCASE," & _
             " update_counter INTEGER NOT NULL DEFAULT 0," & _
             " PRIMARY KEY(key))"

'
'         ' Trigger to bump update_counter on UPDATE of data:
'         '  - AFTER UPDATE OF data: only fires when the data column changes
'         '  - Sets update_counter to current max(update_counter)+1 across the table
'         '  - WHERE key = NEW.key ensures only the updated row is changed
         .Execute _
               "CREATE TRIGGER iconData_update_counter " & _
               "AFTER UPDATE OF data ON iconDataTable " & _
               "FOR EACH ROW " & _
               "BEGIN " & _
               "  UPDATE iconDataTable " & _
               "  SET update_counter = (SELECT COALESCE(MAX(update_counter), 0) + 1 FROM updateTable) " & _
               "  WHERE key = NEW.key; " & _
               "END;"
'
'         ' Trigger to bump update_counter on INSERT:
'         '  - AFTER INSERT: runs after the row is inserted
'         '  - Sets update_counter for just the new row (NEW.key)
'         '  - Uses same global max(update_counter)+1 logic
         .Execute _
               "CREATE TRIGGER iconData_insert_counter " & _
               "AFTER INSERT ON iconDataTable " & _
               "BEGIN " & _
               "  UPDATE updateTable " & _
               "  SET update_counter = (SELECT COALESCE(MAX(update_counter), 0) + 1 FROM updateTable) " & _
               "  WHERE key = NEW.key; " & _
               "END;"
               
        .OpenDB pathToFile, SQLiteReadWriteCreate

        If .hDB <> NULL_PTR Then
            Set DBConnection = .object
        End If

    End With

    On Error GoTo 0
    Exit Sub

createDBFromScratch_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createDBFromScratch of Module modDatabase"

End Sub
