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





'---------------------------------------------------------------------------------------
' Procedure : getAllFieldsFromSingleRecord
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Retrieves the data BLOB for a given key.
'             Raises error 5 if the key is not found.
'---------------------------------------------------------------------------------------
'
Public Function getAllFieldsFromSingleRecord(ByVal p_Key As String) As Variant

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo getAllFieldsFromSingleRecord_Error

    ' select one record matching the supplied key pulling all fields/columns into a dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable WHERE key= " & p_Key)
    
    ' No matching row: raise a generic "Invalid procedure call or argument" (5)
    ' with a more descriptive message.
    If DataSet.RecordCount = 0 Then
       Err.Raise 5, , "Data not found for this key " & p_Key
    End If
    
    ' Return the fifth column: fIconTitle
    getAllFieldsFromSingleRecord = DataSet.Columns(5).Value

    On Error GoTo 0
    Exit Function

getAllFieldsFromSingleRecord_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getAllFieldsFromSingleRecord of Form hiddenForm"
End Function



'---------------------------------------------------------------------------------------
' Procedure : getSingleFieldFromSingleRecord
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : select one record matching the supplied key pulling just one fields/column into a dataset
'---------------------------------------------------------------------------------------
'
Public Function getSingleFieldFromSingleRecord(ByVal fieldName As String, ByVal p_Key As String) As Variant

    Dim DataSet As SQLiteDataSet
    Dim returnedValue As Variant
    
    On Error GoTo getSingleFieldFromSingleRecord_Error

    ' select one record matching the supplied key pulling just one fields/column into a dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT " & fieldName & " FROM iconDataTable WHERE key= " & p_Key)
    
    ' No matching row: raise a generic "Invalid procedure call or argument" (5)
    ' with a more descriptive message.
    If DataSet.RecordCount = 0 Then
       Err.Raise 5, , "Data not found for this key " & p_Key
    End If
    
    ' assign the value in the required field from the dataset to the function return value
    
    ' if this seems a bit wordy, it is, I cannot replace the DataSet!fieldName as a variable
        
    If fieldName = "fIconRecordNumber" Then returnedValue = DataSet!fIconRecordNumber
    If fieldName = "fIconFilename" Then returnedValue = DataSet!fIconFilename
    If fieldName = "fIconFileName2" Then returnedValue = DataSet!fIconFileName2
    If fieldName = "fIconTitle" Then returnedValue = DataSet!fIconTitle
    If fieldName = "fIconCommand" Then returnedValue = DataSet!fIconCommand
    If fieldName = "fIconArguments" Then returnedValue = DataSet!fIconArguments
    If fieldName = "fIconWorkingDirectory" Then returnedValue = DataSet!fIconWorkingDirectory
    If fieldName = "fIconShowCmd" Then returnedValue = DataSet!fIconShowCmd
    If fieldName = "fIconOpenRunning" Then returnedValue = DataSet!fIconOpenRunning
    If fieldName = "fIconIsSeparator" Then returnedValue = DataSet!fIconIsSeparator
    If fieldName = "fIconUseContext" Then returnedValue = DataSet!fIconUseContext
    If fieldName = "fIconDockletFile" Then returnedValue = DataSet!fIconDockletFile
    If fieldName = "fIconUseDialog" Then returnedValue = DataSet!fIconUseDialog
    If fieldName = "fIconUseDialogAfter" Then returnedValue = DataSet!fIconUseDialogAfter
    If fieldName = "fIconQuickLaunch" Then returnedValue = DataSet!fIconQuickLaunch
    If fieldName = "fIconAutoHideDock" Then returnedValue = DataSet!fIconAutoHideDock
    If fieldName = "fIconSecondApp" Then returnedValue = DataSet!fIconSecondApp
    If fieldName = "fIconRunElevated" Then returnedValue = DataSet!fIconRunElevated
    If fieldName = "fIconRunSecondAppBeforehand" Then returnedValue = DataSet!fIconRunSecondAppBeforehand
    If fieldName = "fIconAppToTerminate" Then returnedValue = DataSet!fIconAppToTerminate
    If fieldName = "fIconDisabled" Then returnedValue = DataSet!fIconDisabled
    
    getSingleFieldFromSingleRecord = returnedValue

    On Error GoTo 0
    Exit Function

getSingleFieldFromSingleRecord_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getSingleFieldFromSingleRecord of Module modDatabase"

End Function



'---------------------------------------------------------------------------------------
' Procedure : getAllFieldsFromAllRecords
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : select all records pulling the key and all fields into the dataset
'---------------------------------------------------------------------------------------
'
Public Sub getAllFieldsFromAllRecords()

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo getAllFieldsFromAllRecords_Error

    ' select all records pulling the key and all fields into the dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable")
    
    ' move to the first record in a Recordset and makes it current
    DataSet.MoveFirst
    
    ' list all records in the dataset to the listbox but only show one field from the dataset
    Do Until DataSet.EOF
    
        ' need to insert these into a collection or read the resulting values into the global var cache
            
        hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
        DataSet.MoveNext
    Loop

    On Error GoTo 0
    Exit Sub

getAllFieldsFromAllRecords_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getAllFieldsFromAllRecords of Module modDatabase"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : getSingleFieldFromMultipleRecords
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : select ALL records pulling the key and the chosen field only
'---------------------------------------------------------------------------------------
'
Public Sub getSingleFieldFromMultipleRecords(ByVal fieldName As String)

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo getSingleFieldFromMultipleRecords_Error

    ' select ALL records pulling the key and the chosen field only
    Set DataSet = DBConnection.OpenDataSet("SELECT key, " & fieldName & " FROM iconDataTable")
    
    ' move to the first record in a Recordset and makes it current
    DataSet.MoveFirst
    
    ' list all records in the dataset to the listbox
    Do Until DataSet.EOF
    
        ' need to insert these into a collection or read the resulting values into the global var cache
        
        If fieldName = "fIconRecordNumber" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRecordNumber
        If fieldName = "fIconFilename" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconFilename
        If fieldName = "fIconFileName2" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconFileName2
        If fieldName = "fIconTitle" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
        If fieldName = "fIconCommand" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconCommand
        If fieldName = "fIconArguments" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconArguments
        If fieldName = "fIconWorkingDirectory" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconWorkingDirectory
        If fieldName = "fIconShowCmd" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconShowCmd
        If fieldName = "fIconOpenRunning" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconOpenRunning
        If fieldName = "fIconIsSeparator" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconIsSeparator
        If fieldName = "fIconUseContext" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseContext
        If fieldName = "fIconDockletFile" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconDockletFile
        If fieldName = "fIconUseDialog" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseDialog
        If fieldName = "fIconUseDialogAfter" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseDialogAfter
        If fieldName = "fIconQuickLaunch" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconQuickLaunch
        If fieldName = "fIconAutoHideDock" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconAutoHideDock
        If fieldName = "fIconSecondApp" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconSecondApp
        If fieldName = "fIconRunElevated" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRunElevated
        If fieldName = "fIconRunSecondAppBeforehand" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRunSecondAppBeforehand
        If fieldName = "fIconAppToTerminate" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconAppToTerminate
        If fieldName = "fIconDisabled" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconDisabled
        DataSet.MoveNext
    Loop

    On Error GoTo 0
    Exit Sub

getSingleFieldFromMultipleRecords_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getSingleFieldFromMultipleRecords of Module modDatabase"

End Sub


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
        ' do nothing if nothing
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
' Procedure : insertRecordsFromRandomFile
' Author    : beededea
' Date      : 01/12/2025
' Purpose   : Inserts or updates a single key/value pair in the iconDataTable.
'             On CONFLICT(key), the row is updated instead of inserting a duplicate.
'             The triggers on the table ensure update_counter is bumped appropriately.
'---------------------------------------------------------------------------------------
'
Public Sub insertRecordsFromRandomFile()

    Dim Text As String: Text = vbNullString
    Dim useloop As Integer: useloop = 0
    Dim thisKeyValue As Integer: thisKeyValue = 0
    Dim Command As SQLiteCommand
    
    On Error GoTo insertRecordsFromRandomFile_Error
    
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

    ' now load the user specified icons to the dictionary
    For useloop = iconArrayLowerBound To iconArrayUpperBound
    
        ' extract filenames from the random access data file
        readIconSettingsIni useloop, False
        
        thisKeyValue = useloop
        With DBConnection
            hiddenForm.lblRecordNum.Caption = " Record Number being written now: " & useloop
            hiddenForm.lblRecordNum.Refresh
        
            ' insert values that do not need to be sanitised
            .Execute "INSERT INTO iconDataTable (Key, fIconRecordNumber) VALUES ('" & thisKeyValue & "','" & thisKeyValue & "')"

            ' insert values into fields that can possibly contain dodgy characters as they are user-typed
            Call insertFieldToSingleRecord(thisKeyValue, "fIconFilename", sFilename)
            Call insertFieldToSingleRecord(thisKeyValue, "fIconFilename2", sFileName2)
            Call insertFieldToSingleRecord(thisKeyValue, "fIconTitle", sTitle)
            Call insertFieldToSingleRecord(thisKeyValue, "fIconCommand", sCommand)
                  
            ' insert more values that do not need to be sanitised
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
            
            Call insertFieldToSingleRecord(thisKeyValue, "fIconSecondApp", sSecondApp)
          
            .Execute "INSERT INTO iconDataTable (Key, fIconRunElevated) VALUES ('" & thisKeyValue & "','" & sRunElevated & "') ON CONFLICT (Key) DO UPDATE SET fIconRunElevated=excluded.fIconRunElevated"
            .Execute "INSERT INTO iconDataTable (Key, fIconRunSecondAppBeforehand) VALUES ('" & thisKeyValue & "','" & sRunSecondAppBeforehand & "') ON CONFLICT (Key) DO UPDATE SET fIconRunSecondAppBeforehand=excluded.fIconRunSecondAppBeforehand"
            
            Call insertFieldToSingleRecord(thisKeyValue, "fIconAppToTerminate", sAppToTerminate)
            
            .Execute "INSERT INTO iconDataTable (Key, fIconDisabled) VALUES ('" & thisKeyValue & "','" & sDisabled & "') ON CONFLICT (Key) DO UPDATE SET fIconDisabled=excluded.fIconDisabled"
        End With
        
    Next useloop
    
    ' at the end we prove that this has been achieved
    Call getSingleFieldFromMultipleRecords("fIconTitle")

    On Error GoTo 0
    Exit Sub

insertRecordsFromRandomFile_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertRecordsFromRandomFile of Form hiddenForm"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : insertAllFieldsIntoRandomFile
' Author    : beededea
' Date      : 05/12/2025
' Purpose   : keep the random access data file in sunch with the SQLite database,
'             writing all the data from the iconSettings.db to the iconSettings.dat
'---------------------------------------------------------------------------------------
'
Public Sub insertAllFieldsIntoRandomFile()

    Dim DataSet As SQLiteDataSet
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo insertAllFieldsIntoRandomFile_Error

    ' select all records pulling the key and all fields into the dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable")
    
    ' move to the first record in a Recordset and makes it current
    DataSet.MoveFirst
    
    ' list all records in the dataset to the listbox but only show one field from the dataset
    Do Until DataSet.EOF
        
        hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
        DataSet.MoveNext
        
        useloop = useloop + 1
        Call writeIconSettingsIni(useloop, False)
    Loop
    On Error GoTo 0
    Exit Sub

insertAllFieldsIntoRandomFile_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertAllFieldsIntoRandomFile of Module modDatabase"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : insertFieldToSingleRecord
' Author    : beededea
' Date      : 04/12/2025
' Purpose   : user-entered text or file/folder names can contain characters that an SQL statement can baulk at.
'             Instead the text is entered as a parameter
'---------------------------------------------------------------------------------------
'
Public Sub insertFieldToSingleRecord(ByVal thisKeyValue As Integer, ByVal fieldName As String, ByVal iconVariable As String)

    Dim thisSQL As String: thisSQL = vbNullString
    Dim Command As SQLiteCommand
    
    On Error GoTo insertFieldToSingleRecord_Error

        thisSQL = "INSERT INTO iconDataTable (Key, " & fieldName & ") VALUES (@oid,@opo) ON CONFLICT (Key) DO UPDATE SET " & fieldName & "=excluded." & fieldName

        Set Command = DBConnection.CreateCommand(thisSQL)
        Command.SetParameterValue Command![@oid], thisKeyValue
        Command.SetParameterValue Command![@opo], iconVariable '
        Command.Execute

    On Error GoTo 0
    Exit Sub

insertFieldToSingleRecord_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertFieldToSingleRecord of Module modDatabase"

End Sub






'---------------------------------------------------------------------------------------
' Procedure : deleteSpecificKey
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : Delete a single database record
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
