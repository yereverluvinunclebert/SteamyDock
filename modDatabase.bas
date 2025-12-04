Attribute VB_Name = "modDatabase"
Option Explicit

#If (VBA7 = 0) Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
#If Win64 Then
    'Private Const NULL_PTR As LongPtr = 0 ' this may glow red but is NOT an error, suitable for 64bit TwinBasic
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
Public Sub insertRecords(Optional ByVal thisKeyValue As Integer)

    Dim Text As String: Text = vbNullString
    Dim useloop As Integer
    Dim Command As SQLiteCommand
    
    On Error GoTo insertRecords_Error

    Text = VBA.InputBox("fIconCommand")
    If StrPtr(Text) = NULL_PTR Then Exit Sub
    
    thisKeyValue = 2
    
    ' now load the user specified icons to the dictionary
    For useloop = iconArrayLowerBound To iconArrayUpperBound
    
        ' extract filenames from the random access data file
        readIconSettingsIni useloop, False
        thisKeyValue = useloop
        With DBConnection
        
        If useloop >= 87 Then
            useloop = useloop
        End If
        
            .Execute "INSERT INTO iconDataTable (Key, fIconRecordNumber) VALUES ('" & thisKeyValue & "','" & thisKeyValue & "')"
            .Execute "INSERT INTO iconDataTable (Key, fIconFilename) VALUES ('" & thisKeyValue & "','" & sFilename & "') ON CONFLICT (Key) DO UPDATE SET fIconFilename=excluded.fIconFilename"
            .Execute "INSERT INTO iconDataTable (Key, fIconFileName2) VALUES ('" & thisKeyValue & "','" & sFileName2 & "') ON CONFLICT (Key) DO UPDATE SET fIconFileName2=excluded.fIconFileName2"
            '.Execute "INSERT INTO iconDataTable (Key, fIconTitle) VALUES ('" & thisKeyValue & "','" & sTitle & "') ON CONFLICT (Key) DO UPDATE SET fIconTitle=excluded.fIconTitle"
            
            Set Command = DBConnection.CreateCommand("INSERT INTO iconDataTable (Key, fIconTitle) VALUES (@oid,@opo) ON CONFLICT (Key) DO UPDATE SET fIconTitle=excluded.fIconTitle")
            Command.SetParameterValue Command![@oid], thisKeyValue
            Command.SetParameterValue Command![@opo], sTitle
            Command.Execute
            
            '.Execute "INSERT INTO iconDataTable (Key, fIconCommand) VALUES ('" & thisKeyValue & "','" & sCommand & "') ON CONFLICT (Key) DO UPDATE SET fIconCommand=excluded.fIconCommand"
            
            Set Command = DBConnection.CreateCommand("INSERT INTO iconDataTable (Key, fIconCommand) VALUES (@oid,@opo) ON CONFLICT (Key) DO UPDATE SET fIconCommand=excluded.fIconCommand")
            Command.SetParameterValue Command![@oid], thisKeyValue
            Command.SetParameterValue Command![@opo], sCommand
            Command.Execute
           
                  
            .Execute "INSERT INTO iconDataTable (Key, fIconArguments) VALUES ('" & thisKeyValue & "','" & sArguments & "') ON CONFLICT (Key) DO UPDATE SET fIconArguments=excluded.fIconArguments"
            .Execute "INSERT INTO iconDataTable (Key, fIconWorkingDirectory) VALUES ('" & thisKeyValue & "','" & sWorkingDirectory & "') ON CONFLICT (Key) DO UPDATE SET fIconWorkingDirectory=excluded.fIconWorkingDirectory"
            .Execute "INSERT INTO iconDataTable (Key, fIconShowCmd) VALUES ('" & thisKeyValue & "','" & sShowCmd & "') ON CONFLICT (Key) DO UPDATE SET fIconShowCmd=excluded.fIconShowCmd"
            .Execute "INSERT INTO iconDataTable (Key, fIconOpenRunning) VALUES ('" & thisKeyValue & "','" & sOpenRunning & "') ON CONFLICT (Key) DO UPDATE SET fIconOpenRunning=excluded.fIconOpenRunning"
            .Execute "INSERT INTO iconDataTable (Key, fIconIsSeparator) VALUES ('" & thisKeyValue & "','" & sIsSeparator & "') ON CONFLICT (Key) DO UPDATE SET fIconIsSeparator=excluded.fIconIsSeparator"
            .Execute "INSERT INTO iconDataTable (Key, fIconUseContext) VALUES ('" & thisKeyValue & "','" & sUseContext & "') ON CONFLICT (Key) DO UPDATE SET fIconUseContext=excluded.fIconUseContext"
            .Execute "INSERT INTO iconDataTable (Key, fIconDockletFile) VALUES ('" & thisKeyValue & "','" & sDockletFile & "') ON CONFLICT (Key) DO UPDATE SET fIconDockletFile=excluded.fIconDockletFile"
            .Execute "INSERT INTO iconDataTable (Key, fIconUseDialog) VALUES ('" & thisKeyValue & "','" & sUseDialog & "') ON CONFLICT (Key) DO UPDATE SET fIconUseDialog=excluded.fIconUseDialog"
            .Execute "INSERT INTO iconDataTable (Key, fIconUseDialogAfter) VALUES ('" & thisKeyValue & "','" & sUseDialogAfter & "') ON CONFLICT (Key) DO UPDATE SET fIconUseDialogAfter=excluded.fIconUseDialogAfter"
            .Execute "INSERT INTO iconDataTable (Key, fIconQuickLaunch) VALUES ('" & thisKeyValue & "','" & sQuickLaunch & "') ON CONFLICT (Key) DO UPDATE SET fIconQuickLaunch=excluded.fIconQuickLaunch"
            .Execute "INSERT INTO iconDataTable (Key, fIconAutoHideDock) VALUES ('" & thisKeyValue & "','" & sAutoHideDock & "') ON CONFLICT (Key) DO UPDATE SET fIconAutoHideDock=excluded.fIconAutoHideDock"
            .Execute "INSERT INTO iconDataTable (Key, fIconSecondApp) VALUES ('" & thisKeyValue & "','" & sSecondApp & "') ON CONFLICT (Key) DO UPDATE SET fIconSecondApp=excluded.fIconSecondApp"
            .Execute "INSERT INTO iconDataTable (Key, fIconRunElevated) VALUES ('" & thisKeyValue & "','" & sRunElevated & "') ON CONFLICT (Key) DO UPDATE SET fIconRunElevated=excluded.fIconRunElevated"
            .Execute "INSERT INTO iconDataTable (Key, fIconRunSecondAppBeforehand) VALUES ('" & thisKeyValue & "','" & sRunSecondAppBeforehand & "') ON CONFLICT (Key) DO UPDATE SET fIconRunSecondAppBeforehand=excluded.fIconRunSecondAppBeforehand"
            .Execute "INSERT INTO iconDataTable (Key, fIconAppToTerminate) VALUES ('" & thisKeyValue & "','" & sAppToTerminate & "') ON CONFLICT (Key) DO UPDATE SET fIconAppToTerminate=excluded.fIconAppToTerminate"
            .Execute "INSERT INTO iconDataTable (Key, fIconDisabled) VALUES ('" & thisKeyValue & "','" & sDisabled & "') ON CONFLICT (Key) DO UPDATE SET fIconDisabled=excluded.fIconDisabled"
        End With
        

    
    Next useloop
    Call Requery

    On Error GoTo 0
    Exit Sub

insertRecords_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertRecords of Form hiddenForm"

End Sub

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
    
' database schema (simplified)
'      fIconRecordNumber As Integer
'      fIconFilename As String
'      fIconFileName2 As String
'      fIconTitle As String
'      fIconCommand As String
'      fIconArguments As String
'      fIconWorkingDirectory As String
'      fIconShowCmd As String
'      fIconOpenRunning As String
'      fIconIsSeparator As String
'      fIconUseContext As String
'      fIconDockletFile As String
'      fIconUseDialog As String
'      fIconUseDialogAfter As String
'      fIconQuickLaunch As String
'      fIconAutoHideDock As String
'      fIconSecondApp As String
'      fIconRunElevated As String
'      fIconRunSecondAppBeforehand As String
'      fIconAppToTerminate As String
'      fIconDisabled As String
    
    With DBConnection
        .Execute "INSERT INTO iconDataTable (key, data) VALUES ('" & p_Key & "','" & p_Data & "') ON CONFLICT (key) DO UPDATE SET data=excluded.data"
    End With
   
    On Error GoTo 0
    Exit Sub

SaveToiconData_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveToiconData of Form hiddenForm"
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

    Set DataSet = DBConnection.OpenDataSet("SELECT key, fIconCommand FROM iconDataTable")
    DataSet.MoveFirst
    Do Until DataSet.EOF
        hiddenForm.List1.AddItem DataSet!key & "_" & DataSet!fIconCommand
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
' Purpose   : Delete a database record
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
' Purpose   : used to recreate the database from scratch if required
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
             " key INTEGER UNIQUE," & _
             " fIconRecordNumber INTEGER DEFAULT 0, fIconFilename TEXT," & _
             " fIconFileName2 TEXT," & _
             " fIconTitle TEXT," & _
             " fIconCommand TEXT," & _
             " fIconArguments TEXT," & _
             " fIconWorkingDirectory TEXT," & _
             " fIconShowCmd TEXT," & _
             " fIconOpenRunning TEXT," & _
             " fIconIsSeparator TEXT," & _
             " fIconUseContext TEXT," & _
             " fIconDockletFile TEXT," & _
             " fIconUseDialog TEXT," & _
             " fIconUseDialogAfter TEXT," & _
             " fIconQuickLaunch TEXT," & _
             " fIconAutoHideDock TEXT," & _
             " fIconSecondApp TEXT," & _
             " fIconRunElevated TEXT," & _
             " fIconRunSecondAppBeforehand TEXT," & _
             " fIconAppToTerminate TEXT," & _
             " fIconDisabled TEXT," & _
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
