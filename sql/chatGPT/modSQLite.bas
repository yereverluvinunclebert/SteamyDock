Attribute VB_Name = "modSQLite"
'---------------------------------------------------------------------------------------
' Module    : modSQLite
' Author    : chatGPT
' Date      : 02/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

' === SQLite return codes ===
Public Const SQLITE_OK As Long = 0
Public Const SQLITE_ROW As Long = 100
Public Const SQLITE_DONE As Long = 101

' === SQLite handle types ===
Public Type sqlite3
    dummy As Long
End Type

Public Type sqlite3_stmt
    dummy As Long
End Type

' === Core API ===
Public Declare Function sqlite3_open Lib "sqlite3win32.dll" _
    (ByVal FileName As String, ByRef db As Long) As Long

Public Declare Function sqlite3_close Lib "sqlite3win32.dll" _
    (ByVal db As Long) As Long

Public Declare Function sqlite3_errmsg Lib "sqlite3win32.dll" _
    (ByVal db As Long) As Long

Public Declare Function sqlite3_exec Lib "sqlite3win32.dll" _
    (ByVal db As Long, ByVal SQL As String, _
     ByVal callback As Long, ByVal arg As Long, _
     ByRef errmsg As Long) As Long

' === Prepared statements ===
Public Declare Function sqlite3_prepare_v2 Lib "sqlite3win32.dll" _
    (ByVal db As Long, ByVal SQL As String, ByVal nBytes As Long, _
     ByRef stmt As Long, ByRef tail As Long) As Long

Public Declare Function sqlite3_step Lib "sqlite3win32.dll" _
    (ByVal stmt As Long) As Long

Public Declare Function sqlite3_finalize Lib "sqlite3win32.dll" _
    (ByVal stmt As Long) As Long

Public Declare Function sqlite3_column_text Lib "sqlite3win32.dll" _
    (ByVal stmt As Long, ByVal col As Long) As Long

Public Declare Function sqlite3_column_int Lib "sqlite3win32.dll" _
    (ByVal stmt As Long, ByVal col As Long) As Long
    
Private Declare Function WideCharToMultiByte Lib "kernel32" _
    (ByVal CodePage As Long, ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
     ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32" _
    (ByVal CodePage As Long, ByVal dwFlags As Long, _
     ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
     ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

Public Declare Function sqlite3_column_count Lib "sqlite3.dll" _
    (ByVal stmt As Long) As Long

Public Declare Function sqlite3_column_name Lib "sqlite3.dll" _
    (ByVal stmt As Long, ByVal col As Long) As Long


Public Function ToUTF8(ByVal s As String) As String
    Dim cb As Long
    cb = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), -1, 0, 0, 0, 0)
    ToUTF8 = String$(cb - 1, 0)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), -1, _
        StrPtr(ToUTF8), cb, 0, 0
End Function

Public Function FromUTF8(ByVal pStr As Long) As String
    Dim cb As Long
    cb = MultiByteToWideChar(CP_UTF8, 0, pStr, -1, 0, 0)
    FromUTF8 = String$(cb - 1, 0)
    MultiByteToWideChar CP_UTF8, 0, pStr, -1, _
        StrPtr(FromUTF8), cb
End Function
    


