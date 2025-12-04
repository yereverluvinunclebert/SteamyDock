VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwMainIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cwMainIcon
' Author    : beededea
' Date      : 20/11/2025
' Purpose   : A currently unused class that will form the basis of a class to encapsulate the properties of an icon
'---------------------------------------------------------------------------------------

Option Explicit

Private mHOffset As Long
Private mVOffset As Long
Private mIconWidth As Long
Private mIconIndex As Integer
Private mIconOpacity As Integer
Private mIconImage As String

Private mFilename As String
Private mFileName2 As String ' ' conFileName2)
Private mTitle As String ' ' conTitle)
Private mCommand As String ' ' conCommand)
Private mArguments As String ' ' conArguments)
Private mWorkingDirectory As String ' ' conWorkingDirectory)
Private mShowCmd As String  'ShowCmd
Private mOpenRunning As String  'OpenRunning
Private mIsSeparator As String  'IsSeparator
Private mUseContext As String  'UseContext
Private mDockletFile As String ' ' conDockletFile)
Private mUseDialog As String  'UseDialog
Private mUseDialogAfter As String  'UseDialogAfter
Private mQuickLaunch As String  'QuickLaunch
Private mAutoHideDock As String  'AutoHideDock
Private mSecondApp As String ' ' conSecondApp)
Private mRunElevated As String  'RunElevated
Private mRunSecondAppBeforehand As String  'RunSecondAppBeforehand
Private mAppToTerminate As String ' ' conAppToTerminate)
Private mDisabled As String  'Disabled


Private Sub Class_Initialize()
' Only the main icons
' require this.


    'mainIconxo = 0: mainIconyo = 0

End Sub
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get iconIndex() As Integer
   On Error GoTo iconIndexGet_Error

   iconIndex = mIconIndex

   On Error GoTo 0
   Exit Property

iconIndexGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : iconIndex
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let iconIndex(ByVal newValue As Integer)
   On Error GoTo iconIndexLet_Error

    If mIconIndex <> newValue Then mIconIndex = newValue Else Exit Property
    
    If mIconIndex = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

iconIndexLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property iconIndex of Class Module cwIcon"
End Property
'
'---------------------------------------------------------------------------------------
' Property  : hOffset
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get hOffset() As Long
   On Error GoTo hOffsetGet_Error

   hOffset = mHOffset

   On Error GoTo 0
   Exit Property

hOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property hOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : hOffset
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let hOffset(ByVal newValue As Long)
   On Error GoTo hOffsetLet_Error

    If mHOffset <> newValue Then mHOffset = newValue Else Exit Property
    
    If mHOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

hOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property hOffset of Class Module cwIcon"
End Property


'
'---------------------------------------------------------------------------------------
' Property  : vOffset
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get vOffset() As Long
   On Error GoTo vOffsetGet_Error

   vOffset = mVOffset

   On Error GoTo 0
   Exit Property

vOffsetGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property vOffset of Class Module cwIcon"
End Property
'---------------------------------------------------------------------------------------
' Property  : vOffset
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let vOffset(ByVal newValue As Long)
   On Error GoTo vOffsetLet_Error

    If mVOffset <> newValue Then mVOffset = newValue Else Exit Property
    
    If mVOffset = True Then
    
    Else
    
    End If

   On Error GoTo 0
   Exit Property

vOffsetLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property vOffset of Class Module cwIcon"
End Property

'
'---------------------------------------------------------------------------------------
' Property  : Filename
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FileName() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FileName(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FileName2() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FileName2(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Title() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Title(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Command() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Command(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Arguments() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Arguments(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get WorkingDirectory() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let WorkingDirectory(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ShowCmd() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ShowCmd(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get OpenRunning() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let OpenRunning(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IsSeparator() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IsSeparator(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseContext() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseContext(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get DockletFile() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let DockletFile(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseDialog() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseDialog(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UseDialogAfter() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UseDialogAfter(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get QuickLaunch() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let QuickLaunch(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AutoHideDock() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let AutoHideDock(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SecondApp() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SecondApp(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RunElevated() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RunElevated(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get RunSecondAppBeforehand() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let RunSecondAppBeforehand(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get AppToTerminate() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let AppToTerminate(ByVal newValue As Long)
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
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Disabled() As Long
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
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Disabled(ByVal newValue As Long)
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

