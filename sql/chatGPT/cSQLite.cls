VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_db As Long

Public Sub OpenDatabase(ByVal FileName As String)
    If sqlite3_open(ToUTF8(FileName), m_db) <> SQLITE_OK Then
        Err.Raise vbObjectError, , "Cannot open database"
    End If
End Sub

Public Sub Execute(ByVal SQL As String)
    Dim errPtr As Long
    If sqlite3_exec(m_db, ToUTF8(SQL), 0, 0, errPtr) <> SQLITE_OK Then
        Err.Raise vbObjectError, , FromUTF8(sqlite3_errmsg(m_db))
    End If
End Sub

Public Function QueryScalar(ByVal SQL As String) As String
    Dim stmt As Long
    If sqlite3_prepare_v2(m_db, ToUTF8(SQL), -1, stmt, 0) <> SQLITE_OK Then
        Err.Raise vbObjectError, , "Prepare failed"
    End If

    If sqlite3_step(stmt) = SQLITE_ROW Then
        QueryScalar = FromUTF8(sqlite3_column_text(stmt, 0))
    End If

    sqlite3_finalize stmt
End Function

Public Sub CloseDatabase()
    If m_db <> 0 Then
        sqlite3_close m_db
        m_db = 0
    End If
End Sub

Public Function OpenDataSet(ByVal SQL As String) As cSQLiteDataSet
    Dim stmt As Long
    Dim tail As Long
    Dim colCount As Long
    Dim i As Long
    Dim dataset As New cSQLiteDataSet
    
    ' Prepare statement
    If sqlite3_prepare_v2(m_db, ToUTF8(SQL), -1, stmt, tail) <> SQLITE_OK Then
        Err.Raise vbObjectError, , "Failed to prepare SQL statement"
    End If
    
    ' Get column count and names
    colCount = sqlite3_column_count(stmt)
    dataset.Init colCount
    
    Dim colNames() As String
    ReDim colNames(0 To colCount - 1)
    For i = 0 To colCount - 1
        colNames(i) = FromUTF8(sqlite3_column_name(stmt, i))
    Next i
    dataset.SetColumnNames colNames
    
    ' Loop over rows
    Do While sqlite3_step(stmt) = SQLITE_ROW
        Dim rowData() As String
        ReDim rowData(0 To colCount - 1)
        For i = 0 To colCount - 1
            rowData(i) = FromUTF8(sqlite3_column_text(stmt, i))
        Next i
        dataset.AddRow rowData
    Loop
    
    ' Finalize statement
    sqlite3_finalize stmt
    
    Set OpenDataSet = dataset
End Function

Private Sub Class_Terminate()
    CloseDatabase
End Sub

