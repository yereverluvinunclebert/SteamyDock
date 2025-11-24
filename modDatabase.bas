Attribute VB_Name = "modDatabase"
'Option Explicit




'---------------------------------------------------------------------------------------
' Procedure : SaveToiconData
' Author    : jbPro
' Date      : 24/11/2025
' Purpose   : Inserts or updates a single key/value pair in the iconData.
' On CONFLICT(key), the row is updated instead of inserting a duplicate.
' The triggers on the table ensure update_counter is bumped appropriately.
'---------------------------------------------------------------------------------------
'
Public Sub SaveToiconData(ByVal p_Key As String, p_Data As Variant)
    On Error GoTo SaveToiconData_Error
    
    With DBConnection
        .Execute "INSERT INTO iconData (key, data) VALUES ('" & p_Key & "','" & p_Data & "') ON CONFLICT (key) DO UPDATE SET data=excluded.data"
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
    Set DataSet = DBConnection.OpenDataSet("SELECT * FROM " & iconData & "  WHERE key= " & p_Key)
    
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
    Set DataSet = DBConnection.OpenDataSet("SELECT MAX(update_counter) FROM iconData")

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
'       Keys: iconData key (String)
'       Items: iconData data (Variant / BLOB)
'---------------------------------------------------------------------------------------
'
Public Function GetDataSinceUpdateCounter(ByVal p_UpdateCounter As Currency)
    On Error GoTo GetDataSinceUpdateCounter_Error
    
    Dim DataSet As SQLiteDataSet
    Set DataSet = DBConnection.OpenDataSet("SELECT key, data FROM iconData WHERE update_counter>" & p_UpdateCounter)
    
    ' replace with scripting collection

    'Set GetDataSinceUpdateCounter = New_c.Collection(False)
    
    ' Select only rows whose update_counter is greater than the given value
    With DataSet
       Do Until .EOF
          'GetDataSinceUpdateCounter.Add .Fields("data").Value, .Fields("key").Value
          
          .MoveNext
       Loop
    End With

    On Error GoTo 0
    Exit Function

GetDataSinceUpdateCounter_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDataSinceUpdateCounter of Form hiddenForm"
End Function




' END of JBPro's suggested subs and functions


