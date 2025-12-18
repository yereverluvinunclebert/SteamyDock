VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMainIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cwMainIcon
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : A currently unused class that will form the basis of a class to encapsulate the properties of the main animated icons
'---------------------------------------------------------------------------------------

Option Explicit

' local property holders for this class

Private mIconHOffset As Long
Private mIconVOffset As Long
Private mIconHeight As Long
Private mIconWidth As Long
Private mIconIndex As Integer
Private mIconOpacity As Integer
Private mIconImage As String
Private mFilename As String
Private mFileName2 As String
Private mTitle As String
Private mCommand As String
Private mArguments As String
Private mWorkingDirectory As String
Private mShowCmd As String
Private mOpenRunning As Boolean
Private mIsSeparator As Boolean
Private mUseContext As Boolean
Private mDockletFile As Boolean
Private mUseDialog As Boolean
Private mUseDialogAfter As Boolean
Private mQuickLaunch As Boolean
Private mAutoHideDock As Boolean
Private mSecondApp As String
Private mRunElevated As Boolean
Private mRunSecondAppBeforehand As String
Private mAppToTerminate As String
Private mDisabled As Boolean
Private mRefresh As Boolean



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 13/12/2025
' Purpose   : Only the main animated icons require this.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    'mainIconxo = 0: mainIconyo = 0

    On Error GoTo Class_Initialize_Error

    

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwMainIcon"
End Sub
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IconIndex() As Integer
   On Error GoTo IconIndexGet_Error

   IconIndex = mIconIndex

   On Error GoTo 0
   Exit Property

IconIndexGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IconIndex(ByVal newValue As Integer)
   On Error GoTo IconIndexLet_Error

    If mIconIndex <> newValue Then mIconIndex = newValue Else Exit Property
    
    If mIconIndex = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IconIndexLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'
'---------------------------------------------------------------------------------------
' Property  : iconHOffset
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get iconHOffset() As Long
   On Error GoTo iconHOffsetGet_Error

   iconHOffset = mIconHOffset

   On Error GoTo 0
   Exit Property

iconHOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconHOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : iconHOffset
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let iconHOffset(ByVal newValue As Long)
   On Error GoTo iconHOffsetLet_Error

    If mIconHOffset <> newValue Then mIconHOffset = newValue Else Exit Property
    
    If mIconHOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

iconHOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconHOffset of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : iconVOffset
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get iconVOffset() As Long
   On Error GoTo iconVOffsetGet_Error

   iconVOffset = mIconVOffset

   On Error GoTo 0
   Exit Property

iconVOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconVOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : iconVOffset
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let iconVOffset(ByVal newValue As Long)
   On Error GoTo iconVOffsetLet_Error

    If mIconVOffset <> newValue Then mIconVOffset = newValue Else Exit Property
    
    If mIconVOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

iconVOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconVOffset of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Filename
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FileName() As String
   On Error GoTo FilenameGet_Error

   FileName = mFilename

   On Error GoTo 0
   Exit Property

FilenameGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Filename of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Filename
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FileName(ByVal newValue As String)
   On Error GoTo FilenameLet_Error

    If mFilename <> newValue Then mFilename = newValue Else Exit Property
    
    If mFilename = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

FilenameLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Filename of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : FileName2
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FileName2() As String
   On Error GoTo FileName2Get_Error

   FileName2 = mFileName2

   On Error GoTo 0
   Exit Property

FileName2Get_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FileName2 of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : FileName2
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FileName2(ByVal newValue As String)
   On Error GoTo FileName2Let_Error

    If mFileName2 <> newValue Then mFileName2 = newValue Else Exit Property
    
    If mFileName2 = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

FileName2Let_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FileName2 of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Title
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Title() As String
   On Error GoTo TitleGet_Error

   Title = mTitle

   On Error GoTo 0
   Exit Property

TitleGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Title of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Title
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Title(ByVal newValue As String)
   On Error GoTo TitleLet_Error

    If mTitle <> newValue Then mTitle = newValue Else Exit Property
    
    If mTitle = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

TitleLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Title of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Command
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Command() As String
   On Error GoTo CommandGet_Error

   Command = mCommand

   On Error GoTo 0
   Exit Property

CommandGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Command of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Command
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Command(ByVal newValue As String)
   On Error GoTo CommandLet_Error

    If mCommand <> newValue Then mCommand = newValue Else Exit Property
    
    If mCommand = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

CommandLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Command of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : Arguments
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Arguments() As String
   On Error GoTo ArgumentsGet_Error

   Arguments = mArguments

   On Error GoTo 0
   Exit Property

ArgumentsGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Arguments of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Arguments
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Arguments(ByVal newValue As String)
   On Error GoTo ArgumentsLet_Error

    If mArguments <> newValue Then mArguments = newValue Else Exit Property
    
    If mArguments = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

ArgumentsLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Arguments of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : WorkingDirectory
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get WorkingDirectory() As String
   On Error GoTo WorkingDirectoryGet_Error

   WorkingDirectory = mWorkingDirectory

   On Error GoTo 0
   Exit Property

WorkingDirectoryGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WorkingDirectory of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : WorkingDirectory
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let WorkingDirectory(ByVal newValue As String)
   On Error GoTo WorkingDirectoryLet_Error

    If mWorkingDirectory <> newValue Then mWorkingDirectory = newValue Else Exit Property
    
    If mWorkingDirectory = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

WorkingDirectoryLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WorkingDirectory of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : ShowCmd
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ShowCmd() As Boolean
   On Error GoTo ShowCmdGet_Error

   ShowCmd = mShowCmd

   On Error GoTo 0
   Exit Property

ShowCmdGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ShowCmd  of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : ShowCmd
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ShowCmd(ByVal newValue As Boolean)
   On Error GoTo ShowCmdLet_Error

    If mShowCmd <> newValue Then mShowCmd = newValue Else Exit Property
    
    If mShowCmd = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

ShowCmdLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ShowCmd  of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : OpenRunning
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get OpenRunning() As Boolean
   On Error GoTo OpenRunningGet_Error

   OpenRunning = mOpenRunning

   On Error GoTo 0
   Exit Property

OpenRunningGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property OpenRunning of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : OpenRunning
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let OpenRunning(ByVal newValue As Boolean)
   On Error GoTo OpenRunningLet_Error

    If mOpenRunning <> newValue Then mOpenRunning = newValue Else Exit Property
    
    If mOpenRunning = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

OpenRunningLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property OpenRunning of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : IsSeparator
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IsSeparator() As Boolean
   On Error GoTo IsSeparatorGet_Error

   IsSeparator = mIsSeparator

   On Error GoTo 0
   Exit Property

IsSeparatorGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IsSeparator of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : IsSeparator
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IsSeparator(ByVal newValue As Boolean)
   On Error GoTo IsSeparatorLet_Error

    If mIsSeparator <> newValue Then mIsSeparator = newValue Else Exit Property
    
    If mIsSeparator = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IsSeparatorLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IsSeparator of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : UseContext
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseContext() As Boolean
   On Error GoTo UseContextGet_Error

   UseContext = mUseContext

   On Error GoTo 0
   Exit Property

UseContextGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseContext of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : UseContext
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseContext(ByVal newValue As Boolean)
   On Error GoTo UseContextLet_Error

    If mUseContext <> newValue Then mUseContext = newValue Else Exit Property
    
    If mUseContext = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

UseContextLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseContext of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : DockletFile
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get DockletFile() As Boolean
   On Error GoTo DockletFileGet_Error

   DockletFile = mDockletFile

   On Error GoTo 0
   Exit Property

DockletFileGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property DockletFile of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : DockletFile
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let DockletFile(ByVal newValue As Boolean)
   On Error GoTo DockletFileLet_Error

    If mDockletFile <> newValue Then mDockletFile = newValue Else Exit Property
    
    If mDockletFile = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

DockletFileLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property DockletFile of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : UseDialog
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseDialog() As Boolean
   On Error GoTo UseDialogGet_Error

   UseDialog = mUseDialog

   On Error GoTo 0
   Exit Property

UseDialogGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseDialog of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : UseDialog
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseDialog(ByVal newValue As Boolean)
   On Error GoTo UseDialogLet_Error

    If mUseDialog <> newValue Then mUseDialog = newValue Else Exit Property
    
    If mUseDialog = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

UseDialogLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseDialog of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : UseDialogAfter
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseDialogAfter() As Boolean
   On Error GoTo UseDialogAfterGet_Error

   UseDialogAfter = mUseDialogAfter

   On Error GoTo 0
   Exit Property

UseDialogAfterGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseDialogAfter of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : UseDialogAfter
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseDialogAfter(ByVal newValue As Boolean)
   On Error GoTo UseDialogAfterLet_Error

    If mUseDialogAfter <> newValue Then mUseDialogAfter = newValue Else Exit Property
    
    If mUseDialogAfter = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

UseDialogAfterLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UseDialogAfter of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : QuickLaunch
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get QuickLaunch() As Boolean
   On Error GoTo QuickLaunchGet_Error

   QuickLaunch = mQuickLaunch

   On Error GoTo 0
   Exit Property

QuickLaunchGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property QuickLaunch of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : QuickLaunch
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let QuickLaunch(ByVal newValue As Boolean)
   On Error GoTo QuickLaunchLet_Error

    If mQuickLaunch <> newValue Then mQuickLaunch = newValue Else Exit Property
    
    If mQuickLaunch = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

QuickLaunchLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property QuickLaunch of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : AutoHideDock
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AutoHideDock() As Boolean
   On Error GoTo AutoHideDockGet_Error

   AutoHideDock = mAutoHideDock

   On Error GoTo 0
   Exit Property

AutoHideDockGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property AutoHideDock of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : AutoHideDock
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let AutoHideDock(ByVal newValue As Boolean)
   On Error GoTo AutoHideDockLet_Error

    If mAutoHideDock <> newValue Then mAutoHideDock = newValue Else Exit Property
    
    If mAutoHideDock = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

AutoHideDockLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property AutoHideDock of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : SecondApp
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SecondApp() As String
   On Error GoTo SecondAppGet_Error

   SecondApp = mSecondApp

   On Error GoTo 0
   Exit Property

SecondAppGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SecondApp of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : SecondApp
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SecondApp(ByVal newValue As String)
   On Error GoTo SecondAppLet_Error

    If mSecondApp <> newValue Then mSecondApp = newValue Else Exit Property
    
    If mSecondApp = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

SecondAppLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SecondApp of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : RunElevated
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RunElevated() As Boolean
   On Error GoTo RunElevatedGet_Error

   RunElevated = mRunElevated

   On Error GoTo 0
   Exit Property

RunElevatedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RunElevated of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : RunElevated
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RunElevated(ByVal newValue As Boolean)
   On Error GoTo RunElevatedLet_Error

    If mRunElevated <> newValue Then mRunElevated = newValue Else Exit Property
    
    If mRunElevated = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

RunElevatedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RunElevated of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : RunSecondAppBeforehand
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RunSecondAppBeforehand() As String
   On Error GoTo RunSecondAppBeforehandGet_Error

   RunSecondAppBeforehand = mRunSecondAppBeforehand

   On Error GoTo 0
   Exit Property

RunSecondAppBeforehandGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RunSecondAppBeforehand of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : RunSecondAppBeforehand
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RunSecondAppBeforehand(ByVal newValue As String)
   On Error GoTo RunSecondAppBeforehandLet_Error

    If mRunSecondAppBeforehand <> newValue Then mRunSecondAppBeforehand = newValue Else Exit Property
    
    If mRunSecondAppBeforehand = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

RunSecondAppBeforehandLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property RunSecondAppBeforehand of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : AppToTerminate
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AppToTerminate() As String
   On Error GoTo AppToTerminateGet_Error

   AppToTerminate = mAppToTerminate

   On Error GoTo 0
   Exit Property

AppToTerminateGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property AppToTerminate  of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : AppToTerminate
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let AppToTerminate(ByVal newValue As String)
   On Error GoTo AppToTerminateLet_Error

    If mAppToTerminate <> newValue Then mAppToTerminate = newValue Else Exit Property
    
    If mAppToTerminate = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

AppToTerminateLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property AppToTerminate of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Disabled
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Disabled() As Boolean
   On Error GoTo DisabledGet_Error

   Disabled = mDisabled

   On Error GoTo 0
   Exit Property

DisabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Disabled of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Disabled
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Disabled(ByVal newValue As Boolean)
   On Error GoTo DisabledLet_Error

    If mDisabled <> newValue Then mDisabled = newValue Else Exit Property
    
    If mDisabled = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

DisabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Disabled of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : IconHeight
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IconHeight() As Long
   On Error GoTo IconHeightGet_Error

   IconHeight = mIconHeight

   On Error GoTo 0
   Exit Property

IconHeightGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconHeight of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : IconHeight
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IconHeight(ByVal newValue As Long)
   On Error GoTo IconHeightLet_Error

    If mIconHeight <> newValue Then mIconHeight = newValue Else Exit Property
    
    If mIconHeight = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IconHeightLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconHeight of Class Module cwIcon"
End Property


'---------------------------------------------------------------------------------------
' Property  : IconWidth
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IconWidth() As Long
   On Error GoTo IconWidthGet_Error

   IconWidth = mIconWidth

   On Error GoTo 0
   Exit Property

IconWidthGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconWidth of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : IconWidth
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IconWidth(ByVal newValue As Long)
   On Error GoTo IconWidthLet_Error

    If mIconWidth <> newValue Then mIconWidth = newValue Else Exit Property
    
    If mIconWidth = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IconWidthLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconWidth of Class Module cwIcon"
End Property



'
'---------------------------------------------------------------------------------------
' Property  : IconOpacity
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IconOpacity() As Long
   On Error GoTo IconOpacityGet_Error

   IconOpacity = mIconOpacity

   On Error GoTo 0
   Exit Property

IconOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconOpacity of Class Module cwIcon"
End Property

'---------------------------------------------------------------------------------------
' Property  : IconOpacity
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IconOpacity(ByVal newValue As Long)
   On Error GoTo IconOpacityLet_Error

    If mIconOpacity <> newValue Then mIconOpacity = newValue Else Exit Property
    
    If mIconOpacity = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IconOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconOpacity of Class Module cwIcon"
End Property




'---------------------------------------------------------------------------------------
' Property  : IconImage
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IconImage() As String
   On Error GoTo IconImageGet_Error

   IconImage = mIconImage

   On Error GoTo 0
   Exit Property

IconImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconImage of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : IconImage
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IconImage(ByVal newValue As String)
   On Error GoTo IconImageLet_Error

    If mIconImage <> newValue Then mIconImage = newValue Else Exit Property
    
    If mIconImage = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

IconImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IconImage of Class Module cwIcon"
End Property


'---------------------------------------------------------------------------------------
' Property  : Refresh
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Refresh() As String
   On Error GoTo RefreshGet_Error

   Refresh = mRefresh

   On Error GoTo 0
   Exit Property

RefreshGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Refresh of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : Refresh
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Refresh(ByVal newValue As String)
   On Error GoTo RefreshLet_Error

    If mRefresh <> newValue Then mRefresh = newValue Else Exit Property
    
    Paint

   On Error GoTo 0
   Exit Property

RefreshLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Refresh of Class Module cwIcon"
End Property




'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : beededea
' Date      : 20/11/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Paint()
' ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)

    Dim someOpacity As Double: someOpacity = 0
   
    On Error GoTo W_Paint_Error
    
    'someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    'Call updateScreenUsingGDIPBitmap

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 


