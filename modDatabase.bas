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

Public DBConnection As SQLiteConnection  ' requires the SQLite project reference VBSQLite12.DLL

'---------------------------------------------------------------------------------------
' Procedure : connectDatabase
' Author    : beededea
' Date      : 07/12/2025
' Purpose   : test connection to DB exists, if not then connect or create.
'---------------------------------------------------------------------------------------
'
Public Function connectDatabase() As String

    Dim PathName As String: PathName = vbNullString

    On Error GoTo connectDatabase_Error
    
    connectDatabase = "No Database"

    If DBConnection Is Nothing Then
            
        PathName = App.Path
        If Not Right$(PathName, 1) = "\" Then PathName = PathName & "\"
        PathName = "C:\Users\beededea\AppData\Roaming\steamyDock\iconSettings.db"
        PathName = gblsIconDataBase
        
        ' check database file exists on the system
        If fFExists(PathName) = True Then
            With New SQLiteConnection
                ' connect to SQLite db
                .OpenDB PathName, SQLiteReadWrite

                ' connection is good?
                If .hDB <> NULL_PTR Then
                    Set DBConnection = .Object
                End If
            End With
            connectDatabase = "Database Connected."

        Else ' if db not exists then create it and set up the new database with hard coded schema
            If MsgBox(PathName & " does not exist. Create new?", vbExclamation + vbOKCancel) <> vbCancel Then
                Call createUnpopulatedDBFromSchema(PathName)
                connectDatabase = "New Empty Database Created with Good Schema & Connected."
            Else
                Exit Function
            End If
        End If
    End If

    On Error GoTo 0
    Exit Function

connectDatabase_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure connectDatabase of Form dock"

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
'Public Function MaxUpdateCounter() As Currency
'    On Error GoTo MaxUpdateCounter_Error
'
'    Dim DataSet As SQLiteDataSet
'    Set DataSet = DBConnection.OpenDataSet("SELECT MAX(update_counter) FROM iconDataTable")
'
'    If DataSet.RecordCount > 0 Then
'       ' If the result is NULL, this will default to 0 when assigned to Currency.
'       MaxUpdateCounter = DataSet.Columns(0)
'    End If
'
'    On Error GoTo 0
'    Exit Function
'
'MaxUpdateCounter_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MaxUpdateCounter of module ModDatase"
'End Function






'---------------------------------------------------------------------------------------
' Procedure : GetDataSinceUpdateCounter
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Returns all rows with update_counter greater than the specified value.
'
' Result:
'   - A collection of:
'       Keys: iconDataTable key (String)
'       Items: iconDataTable data ()
'---------------------------------------------------------------------------------------
'
'Public Function GetDataSinceUpdateCounter(ByVal p_UpdateCounter As Currency)
'    On Error GoTo GetDataSinceUpdateCounter_Error
'
'    Dim DataSet As SQLiteDataSet
'    Set DataSet = DBConnection.OpenDataSet("SELECT key, data FROM iconDataTable WHERE update_counter>" & p_UpdateCounter)
'
'    'dictionary for the database access
'    Set GetDataSinceUpdateCounter = CreateObject("Scripting.Dictionary")
'    GetDataSinceUpdateCounter.CompareMode = 1 'case-insenitive Key-Comparisons
'
'    ' Select only rows whose update_counter is greater than the given value
'    With DataSet
'       Do Until .EOF
'          GetDataSinceUpdateCounter.Add .Columns("data").Value, .Columns("key").Value
'
'          .MoveNext
'       Loop
'    End With
'
'    On Error GoTo 0
'    Exit Function
'
'GetDataSinceUpdateCounter_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDataSinceUpdateCounter of module ModDatase"
'End Function



'---------------------------------------------------------------------------------------
' Procedure : closeDatabase
' Author    : Krool
' Date      : 01/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CloseDatabase()
    On Error GoTo closeDatabase_Error

    If DBConnection Is Nothing Then
        ' do nothing if nothing
    Else
        'DBConnection.SetProgressHandler Nothing ' Unregisters the progress handler callback
        'DBConnection.CloseDatabase
        DBConnection.CloseDB
        Set DBConnection = Nothing
    End If

    On Error GoTo 0
    Exit Sub

closeDatabase_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure closeDatabase of module ModDatase"

End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : putIconSettingsIntoDatabase
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save icon values to database rather than to the random access data file
'---------------------------------------------------------------------------------------
'
Public Function putIconSettingsIntoDatabase(ByVal thisKeyValue As Integer) As Integer

    On Error GoTo putIconSettingsIntoDatabase_Error
    
    Dim DataSet As SQLiteDataSet
    
    With DBConnection
    
        ' We don't have an UPSERT with this SQLite DLL so we have to test first whether the record exists or not.
        
        ' select one record matching the supplied key pulling all fields/columns into a dataset
        Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable WHERE key= " & thisKeyValue)
        
        ' Matching row found
        If DataSet.RecordCount > 0 Then

            ' UPDATE values into fields as a parameter as they are user-typed that could possibly contain dodgy characters
            
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconFilename", sFilename)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconFilename2", sFileName2)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconTitle", sTitle)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconCommand", sCommand)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconArguments", sArguments)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconWorkingDirectory", sWorkingDirectory)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconShowCmd", sShowCmd)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconOpenRunning", sOpenRunning)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconIsSeparator", sIsSeparator)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconUseContext", sUseContext)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconDockletFile", sDockletFile)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconUseDialog", sUseDialog)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconUseDialogAfter", sUseDialogAfter)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconQuickLaunch", sQuickLaunch)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconAutoHideDock", sAutoHideDock)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconSecondApp", sSecondApp)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconRunSecondAppBeforehand", sRunSecondAppBeforehand)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconAppToTerminate", sAppToTerminate)
            Call UPDATEFieldInSingleRecord(thisKeyValue, "fIconDisabled", sDisabled)
        
            ' no error count
            putIconSettingsIntoDatabase = 0
            
        Else
            
            ' if no matching record found then we insert a new record with all the icon values
            ' INSERT new values into fields as a parameter as they are user-typed that could possibly contain dodgy characters
            
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconFilename", sFilename)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconFilename2", sFileName2)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconTitle", sTitle)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconCommand", sCommand)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconArguments", sArguments)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconWorkingDirectory", sWorkingDirectory)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconShowCmd", sShowCmd)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconOpenRunning", sOpenRunning)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconIsSeparator", sIsSeparator)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseContext", sUseContext)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconDockletFile", sDockletFile)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseDialog", sUseDialog)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseDialogAfter", sUseDialogAfter)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconQuickLaunch", sQuickLaunch)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconAutoHideDock", sAutoHideDock)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconSecondApp", sSecondApp)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunSecondAppBeforehand", sRunSecondAppBeforehand)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconAppToTerminate", sAppToTerminate)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconDisabled", sDisabled)
            
            ' no error count
            putIconSettingsIntoDatabase = 0
       
        End If
        
    End With
                
   On Error GoTo 0
   Exit Function

putIconSettingsIntoDatabase_Error:

    ' error count of 1 passed back to calling routine
    putIconSettingsIntoDatabase = 1

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure putIconSettingsIntoDatabase of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : querySingleRecordFromDatabase
' Author    : beededea
' Date      : 24/11/2025
' Purpose   : Retrieves all the data fields for a given key.
'             Raises error 5 if the key is not found.
'---------------------------------------------------------------------------------------
'
Public Function querySingleRecordFromDatabase(ByVal thisKeyValue As String) As Boolean

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo querySingleRecordFromDatabase_Error

    ' select one record matching the supplied key pulling all fields/columns into a dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable WHERE key= " & thisKeyValue)
    
    ' No matching row
    If DataSet.RecordCount = 0 Then
        querySingleRecordFromDatabase = False
    Else
        querySingleRecordFromDatabase = True
    End If
    
    On Error GoTo 0
    Exit Function

querySingleRecordFromDatabase_Error:
    
    ' error count of 1 passed back to calling routine
    querySingleRecordFromDatabase = 1

     MsgBox "Error Data not found for this key " & thisKeyValue & " or other error - " & Err.Number & " (" & Err.Description & ") in procedure querySingleRecordFromDatabase of module ModDatase"
End Function



'---------------------------------------------------------------------------------------
' Procedure : getIconSettingsFromDatabase
' Author    : beededea
' Date      : 24/11/2025
' Purpose   : Retrieves all the data fields for a given key.
'             Raises error 5 if the key is not found.
'---------------------------------------------------------------------------------------
'
Public Function getIconSettingsFromDatabase(ByVal thisKeyValue As String) As Integer

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo getIconSettingsFromDatabase_Error

    ' select one record matching the supplied key pulling all fields/columns into a dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable WHERE key= " & thisKeyValue)
    
    ' No matching row
    If DataSet.RecordCount = 0 Then
       GoTo getIconSettingsFromDatabase_Error
    End If
    
    sFilename = DataSet!fIconFilename
    sFileName2 = DataSet!fIconFileName2
    sTitle = DataSet!fIconTitle
    sCommand = DataSet!fIconCommand
    sArguments = DataSet!fIconArguments
    sWorkingDirectory = DataSet!fIconWorkingDirectory
    sShowCmd = DataSet!fIconShowCmd
    sOpenRunning = DataSet!fIconOpenRunning
    sIsSeparator = DataSet!fIconIsSeparator
    sUseContext = DataSet!fIconUseContext
    sDockletFile = DataSet!fIconDockletFile
    sUseDialog = DataSet!fIconUseDialog
    sUseDialogAfter = DataSet!fIconUseDialogAfter
    sQuickLaunch = DataSet!fIconQuickLaunch
    sAutoHideDock = DataSet!fIconAutoHideDock
    sSecondApp = DataSet!fIconSecondApp
    sRunElevated = DataSet!fIconRunElevated
    sRunSecondAppBeforehand = DataSet!fIconRunSecondAppBeforehand
    sAppToTerminate = DataSet!fIconAppToTerminate
    sDisabled = DataSet!fIconDisabled
    
    ' no error count
    getIconSettingsFromDatabase = 0
    
    On Error GoTo 0
    Exit Function

getIconSettingsFromDatabase_Error:
    
    ' error count of 1 passed back to calling routine
    getIconSettingsFromDatabase = 1

     MsgBox "Error Data not found for this key " & thisKeyValue & " or other error - " & Err.Number & " (" & Err.Description & ") in procedure getIconSettingsFromDatabase of module ModDatase"
End Function





'---------------------------------------------------------------------------------------
' Procedure : INSERTFieldToSingleRecord
' Author    : beededea
' Date      : 04/12/2025
' Purpose   : user-entered text or file/folder names can contain characters that an SQL statement can baulk at.
'             Instead the text is entered as a parameter
'---------------------------------------------------------------------------------------
'
Public Sub INSERTFieldToSingleRecord(ByVal thisKeyValue As Integer, ByVal fieldName As String, ByVal iconVariable As String)

    Dim thisSQL As String: thisSQL = vbNullString
    Dim Command As SQLiteCommand
    
    On Error GoTo INSERTFieldToSingleRecord_Error

    ' I'm unsure whether the ON CONFLICT is working at all, instead we also have an UPDATE version below
    thisSQL = "INSERT INTO iconDataTable (Key, " & fieldName & ") VALUES (@oid,@opo) ON CONFLICT (Key) DO UPDATE SET " & fieldName & "=excluded." & fieldName

    Set Command = DBConnection.CreateCommand(thisSQL)
    Command.SetParameterValue Command![@oid], thisKeyValue
    Command.SetParameterValue Command![@opo], iconVariable
    Command.Execute

    On Error GoTo 0
    Exit Sub

INSERTFieldToSingleRecord_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure INSERTFieldToSingleRecord of Module modDatabase"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : UPDATEFieldInSingleRecord
' Author    : beededea
' Date      : 04/12/2025
' Purpose   : user-entered text or file/folder names can contain characters that an SQL statement can baulk at.
'             Instead the text is entered as a parameter
'---------------------------------------------------------------------------------------
'
Public Sub UPDATEFieldInSingleRecord(ByVal thisKeyValue As Integer, ByVal fieldName As String, ByVal iconVariable As String)

    Dim thisSQL As String: thisSQL = vbNullString
    Dim Command As SQLiteCommand
    
    On Error GoTo UPDATEFieldInSingleRecord_Error

   ' ON CONFLICT (Key) DO UPDATE SET " & fieldName & "=excluded." & fieldName ' does not work with an UPDATE

    thisSQL = "UPDATE iconDataTable SET (" & fieldName & ") = (@opo) WHERE key = " & thisKeyValue
    Set Command = DBConnection.CreateCommand(thisSQL)
    Command.SetParameterValue Command![@opo], iconVariable '
    Command.Execute

    On Error GoTo 0
    Exit Sub

UPDATEFieldInSingleRecord_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UPDATEFieldInSingleRecord of Module modDatabase"

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
' Procedure : getRecordCount
' Author    : beededea
' Date      : 11/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getRecordCount() As Integer

    Dim DataSet As SQLiteDataSet
    
    On Error GoTo getRecordCount_Error

    ' select one record matching the supplied key pulling all fields/columns into a dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable")
    getRecordCount = DataSet.RecordCount

    On Error GoTo 0
    Exit Function

getRecordCount_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getRecordCount of Module modDatabase"
End Function



'---------------------------------------------------------------------------------------
' From this point on these are test and administration only routines
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : insertAllFieldsIntoRandomDataFile
' Author    : beededea
' Date      : 05/12/2025
' Purpose   : keep the random access data file in synch. with the SQLite database,
'             writing all the data from the iconSettings.db to the iconSettings.dat
'---------------------------------------------------------------------------------------
'
Public Sub insertAllFieldsIntoRandomDataFile()

    Dim DataSet As SQLiteDataSet
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo insertAllFieldsIntoRandomDataFile_Error

    ' select all records pulling the key and all fields into the dataset
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM iconDataTable")
    
    ' move to the first record in a Recordset and makes it current
    DataSet.MoveFirst
    
    ' list all records in the dataset to the listbox but only show one field from the dataset
    Do Until DataSet.EOF
        
        'hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
        DataSet.MoveNext
        
        useloop = useloop + 1
        Call putIconSettings(useloop)
    Loop
    On Error GoTo 0
    Exit Sub

insertAllFieldsIntoRandomDataFile_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertAllFieldsIntoRandomDataFile of Module modDatabase"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : createUnpopulatedDBFromSchema
' Author    : beededea
' Date      : 02/12/2025
' Purpose   : used to recreate the database from scratch if required
'             In any case, this code is NOT required as an empty databse is never going to be shipped with the program.
'             This is just retained retained here for later investigation and for education (mine).
'---------------------------------------------------------------------------------------
'
Public Sub createUnpopulatedDBFromSchema(ByVal pathToFile As String)
    
    On Error GoTo createUnpopulatedDBFromSchema_Error

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
            Set DBConnection = .Object
        End If

    End With

    On Error GoTo 0
    Exit Sub

createUnpopulatedDBFromSchema_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createUnpopulatedDBFromSchema of Module modDatabase"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : insertRecordsFromRandomDataFileIntoDatabase
' Author    : beededea
' Date      : 01/12/2025
' Purpose   : Reading from the random access data file
'             Inserts or updates multiple records into the iconDataTable.
'             On CONFLICT(key), the row is updated instead of inserting a duplicate.
'             The triggers on the table ensure update_counter is bumped appropriately.
'---------------------------------------------------------------------------------------
'
Public Sub insertRecordsFromRandomDataFileIntoDatabase()

    Dim useloop As Integer: useloop = 0
    Dim thisKeyValue As Integer: thisKeyValue = 0
    
    On Error GoTo insertRecordsFromRandomDataFileIntoDatabase_Error
    
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

    ' loop through all the records in data file
    For useloop = iconArrayLowerBound To iconArrayUpperBound
    
        ' extract filenames from the random access data file
        Call getIconSettings(useloop)
        
        thisKeyValue = useloop
        With DBConnection
            'If hiddenForm.IsLoaded = True Then Call writeHiddenFormLabel(" Record Number being written now: ", useloop)
            
            ' this is slow but will probably improve with a BEGIN TRANSACTION" then execute with "END TRANSACTION" but not worth the development time as this will seldom ever be used.
            
        
            ' insert a value that does not need to be sanitised
            .Execute "INSERT INTO iconDataTable (Key, fIconRecordNumber) VALUES ('" & thisKeyValue & "','" & thisKeyValue & "')"
            
            'retained for example
            '.Execute "INSERT INTO iconDataTable (Key, fIconDisabled) VALUES ('" & thisKeyValue & "','" & sDisabled & "') ON CONFLICT (Key) DO UPDATE SET fIconDisabled=excluded.fIconDisabled"
            
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconFilename", sFilename)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconFilename2", sFileName2)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconTitle", sTitle)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconCommand", sCommand)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconArguments", sArguments)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconWorkingDirectory", sWorkingDirectory)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconShowCmd", sShowCmd)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconOpenRunning", sOpenRunning)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconIsSeparator", sIsSeparator)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseContext", sUseContext)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconDockletFile", sDockletFile)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseDialog", sUseDialog)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconUseDialogAfter", sUseDialogAfter)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconQuickLaunch", sQuickLaunch)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconAutoHideDock", sAutoHideDock)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconSecondApp", sSecondApp)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunElevated", sRunElevated)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconRunSecondAppBeforehand", sRunSecondAppBeforehand)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconAppToTerminate", sAppToTerminate)
            Call INSERTFieldToSingleRecord(thisKeyValue, "fIconDisabled", sDisabled)
            
        End With
        
    Next useloop
    
    ' at the end we prove that this has been achieved
    Call getSingleFieldFromMultipleRecords("fIconTitle")

    On Error GoTo 0
    Exit Sub

insertRecordsFromRandomDataFileIntoDatabase_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertRecordsFromRandomDataFileIntoDatabase of module ModDatase"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : getAllFieldsFromSingleRecord
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Retrieves the data fields for a given key.
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getAllFieldsFromSingleRecord of module ModDatase"
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
            
        'hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
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
        
'        If fieldName = "fIconRecordNumber" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRecordNumber
'        If fieldName = "fIconFilename" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconFilename
'        If fieldName = "fIconFileName2" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconFileName2
'        If fieldName = "fIconTitle" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconTitle
'        If fieldName = "fIconCommand" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconCommand
'        If fieldName = "fIconArguments" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconArguments
'        If fieldName = "fIconWorkingDirectory" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconWorkingDirectory
'        If fieldName = "fIconShowCmd" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconShowCmd
'        If fieldName = "fIconOpenRunning" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconOpenRunning
'        If fieldName = "fIconIsSeparator" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconIsSeparator
'        If fieldName = "fIconUseContext" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseContext
'        If fieldName = "fIconDockletFile" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconDockletFile
'        If fieldName = "fIconUseDialog" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseDialog
'        If fieldName = "fIconUseDialogAfter" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconUseDialogAfter
'        If fieldName = "fIconQuickLaunch" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconQuickLaunch
'        If fieldName = "fIconAutoHideDock" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconAutoHideDock
'        If fieldName = "fIconSecondApp" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconSecondApp
'        If fieldName = "fIconRunElevated" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRunElevated
'        If fieldName = "fIconRunSecondAppBeforehand" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconRunSecondAppBeforehand
'        If fieldName = "fIconAppToTerminate" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconAppToTerminate
'        If fieldName = "fIconDisabled" Then hiddenForm.List1.AddItem DataSet!key & " " & DataSet!fIconDisabled
        DataSet.MoveNext
    Loop

    On Error GoTo 0
    Exit Sub

getSingleFieldFromMultipleRecords_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getSingleFieldFromMultipleRecords of Module modDatabase"

End Sub



''---------------------------------------------------------------------------------------
'' Procedure : writeHiddenFormLabel
'' Author    : beededea
'' Date      : 08/12/2025
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub writeHiddenFormLabel(textForLabel As String, Optional ByVal count As Integer)
'    Dim textToDisplay As String: textToDisplay = vbNullString
'
'    On Error GoTo writeHiddenFormLabel_Error
'
'    textToDisplay = textForLabel
'    If count > 0 Then textToDisplay = textToDisplay & count
'    hiddenForm.lblRecordNum.Caption = textToDisplay
'    hiddenForm.lblRecordNum.Refresh
'
'    On Error GoTo 0
'    Exit Sub
'
'writeHiddenFormLabel_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeHiddenFormLabel of Module modDatabase"
'End Sub


