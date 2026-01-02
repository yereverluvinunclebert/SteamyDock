VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cSQLiteDataSet
' Author    : beededea
' Date      : 02/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

' Store data in memory
Private m_Columns() As String
Private m_Rows() As Variant
Private m_RecordCount As Long
Private m_ColCount As Long
Private m_CurRow As Long

' Initialize the dataset
Friend Sub Init(ByVal colCount As Long)
    m_ColCount = colCount
    ReDim m_Columns(0 To colCount - 1)
    ReDim m_Rows(0 To 0)
    m_RecordCount = 0
    m_CurRow = -1
End Sub

' Set column names
Friend Sub SetColumnNames(Names() As String)
    Dim i As Long
    For i = 0 To UBound(Names)
        m_Columns(i) = Names(i)
    Next i
End Sub

' Add a row of data
Friend Sub AddRow(rowData() As String)
    If m_RecordCount > 0 Then ReDim Preserve m_Rows(0 To m_RecordCount)
    m_Rows(m_RecordCount) = rowData
    m_RecordCount = m_RecordCount + 1
End Sub

' Move to first row
Public Sub MoveFirst()
    If m_RecordCount > 0 Then
        m_CurRow = 0
    Else
        m_CurRow = -1
    End If
End Sub

' Move to next row
Public Sub MoveNext()
    If m_CurRow >= 0 And m_CurRow < m_RecordCount - 1 Then
        m_CurRow = m_CurRow + 1
    Else
        m_CurRow = -1
    End If
End Sub

' EOF check
Public Function EOF() As Boolean
    EOF = (m_CurRow = -1)
End Function

' Get value by column name
Public Property Get Field(ByVal ColName As String) As String
    Dim i As Long
    For i = 0 To m_ColCount - 1
        If StrComp(m_Columns(i), ColName, vbTextCompare) = 0 Then
            Field = m_Rows(m_CurRow)(i)
            Exit Property
        End If
    Next i
    Field = ""
End Property

' Get value by column index
Public Property Get FieldByIndex(ByVal Index As Long) As String
    If Index >= 0 And Index < m_ColCount Then
        FieldByIndex = m_Rows(m_CurRow)(Index)
    Else
        FieldByIndex = ""
    End If
End Property

' Get column count
Public Property Get ColumnCount() As Long
    ColumnCount = m_ColCount
End Property

' Get row count
Public Property Get RecordCount() As Long
    RecordCount = m_RecordCount
End Property

' Get column names
Public Property Get ColumnName(ByVal Index As Long) As String
    If Index >= 0 And Index < m_ColCount Then
        ColumnName = m_Columns(Index)
    Else
        ColumnName = ""
    End If
End Property

