Attribute VB_Name = "mdlSystray"
' .01 DAEB mdlSysTray 20/02/2021 Added new mdlSysTray module containing the code required to analyse the icons in the systray
 
' CREDIT:
'
' Dragokas on the VBforum

Option Explicit

Private Type TBBUTTON_32
    iBitmap         As Long
    idCommand       As Long
    fsState         As Byte
    fsStyle         As Byte
    bReserved(1)    As Byte
    dwData          As Long
    iString         As Long
End Type

Private Type TBBUTTON_64
    iBitmap         As Long
    idCommand       As Long
    fsState         As Byte
    fsStyle         As Byte
    bReserved(5)    As Byte
    dwData          As Currency
    iString         As Currency
End Type

Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll.dll" (ByVal ProcessHandle As Long, ByVal BaseAddress As Currency, ByVal Buffer As Long, ByVal Size As Currency, ByVal NumberOfBytesRead As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByVal lpSystemInfo As Long)
Public Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Public Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long

Public Const MAX_PATH As Long = 260&

Public Const TB_GETBUTTON As Long = 1047&
Public Const TB_BUTTONCOUNT As Long = 1048&

Public Const PROCESS_VM_OPERATION As Long = &H8&
Public Const PROCESS_VM_READ As Long = 16&
Public Const PROCESS_QUERY_INFORMATION As Long = 1024&
Public Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Public Const MEM_COMMIT As Long = &H1000&
Public Const PAGE_READWRITE As Long = 4&
Public Const MEM_RELEASE As Long = &H8000&

Public Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9&


Private Type RTL_OSVERSIONINFOEXW
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(127) As Integer
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Declare Function RtlGetVersion Lib "ntdll.dll" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long

Public Const ERROR_PARTIAL_COPY            As Long = 299&
Public Const ERROR_ACCESS_DENIED           As Long = 5&


Public Function FindWindow_NotifyTray() As Long
    Dim hWnd As Long
    hWnd = FindWindow(StrPtr("Shell_TrayWnd"), 0&)
    hWnd = FindWindowEx(hWnd, 0, StrPtr("TrayNotifyWnd"), 0)
    hWnd = FindWindowEx(hWnd, 0, StrPtr("SysPager"), 0)
    hWnd = FindWindowEx(hWnd, 0, StrPtr("ToolbarWindow32"), 0)
    FindWindow_NotifyTray = hWnd
End Function

Public Function FindWindow_NotifyOverflow() As Long
    Dim hWnd As Long
    hWnd = FindWindow(StrPtr("NotifyIconOverflowWindow"), 0&)
    hWnd = FindWindowEx(hWnd, 0, StrPtr("ToolbarWindow32"), 0)
    FindWindow_NotifyOverflow = hWnd
End Function

Public Function GetIconCount(hWnd As Long) As Long
    GetIconCount = SendMessage(hWnd, TB_BUTTONCOUNT, 0, ByVal 0)
End Function

Public Function GetIconHandles(hTray As Long, Count As Long, hIcon() As Long) As Boolean

    Dim pid         As Long
    Dim tb_32       As TBBUTTON_32
    Dim tb_64       As TBBUTTON_64
    Dim Extra(1)    As Long
    Dim hProc       As Long
    Dim pMem        As Long
    Dim index       As Long
    Dim OS_64       As Boolean
    Dim si          As SYSTEM_INFO

    GetNativeSystemInfo VarPtr(si)
    OS_64 = (si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64)

    ReDim hIcon(Count - 1)

    GetWindowThreadProcessId hTray, pid

    If pid Then

        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ, False, pid)

        If hProc Then

            pMem = VirtualAllocEx(hProc, 0&, IIf(OS_64, LenB(tb_64), LenB(tb_32)), MEM_COMMIT, PAGE_READWRITE)

            If pMem Then

                For index = 0 To Count - 1

                    If SendMessage(hTray, TB_GETBUTTON, index, ByVal pMem) Then

                        If OS_64 Then

                            If ReadProcessMemory64(hProc, IntToInt64(pMem), VarPtr(tb_64), LenB(tb_64)) Then

                                If tb_64.dwData <> 0 Then

                                    If ReadProcessMemory64(hProc, tb_64.dwData, VarPtr(Extra(0)), 8&) Then

                                        hIcon(index) = Extra(0)
                                        GetIconHandles = True
                                    End If
                                End If
                            End If
                        Else
                            If ReadProcessMemory(hProc, pMem, ByVal VarPtr(tb_32), LenB(tb_32), ByVal 0&) Then

                                If tb_32.dwData <> 0 Then

                                    If ReadProcessMemory(hProc, tb_32.dwData, ByVal VarPtr(Extra(0)), 8&, ByVal 0&) Then

                                        hIcon(index) = Extra(0)
                                        GetIconHandles = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                VirtualFreeEx hProc, pMem, 0, MEM_RELEASE
            End If
            CloseHandle hProc
        End If
    End If
End Function

Public Function ReadProcessMemory64(hProcess As Long, lpBaseAddress As Currency, lpBuffer As Long, nSize As Long) As Boolean
    ReadProcessMemory64 = NT_SUCCESS(NtWow64ReadVirtualMemory64(hProcess, lpBaseAddress, lpBuffer, IntToInt64(nSize), 0&))
End Function

Public Function NT_SUCCESS(NT_Code As Long) As Boolean
    NT_SUCCESS = (NT_Code >= 0)
End Function

Public Function IntToInt64(numInt As Long) As Currency
    IntToInt64 = CCur(numInt / 10000&)
End Function

Public Function GetPidByWindow(hWnd As Long) As Long
     GetWindowThreadProcessId hWnd, GetPidByWindow
End Function




Public Function GetFilePathByPid(pid As Long) As String

    Dim hProc       As Long
    Dim ProcPath    As String
    Dim cnt         As Long
    Dim osi         As RTL_OSVERSIONINFOEXW
    Dim bIsWinVistaAndNewer As Boolean

    osi.dwOSVersionInfoSize = Len(osi)
    Call RtlGetVersion(osi)
    bIsWinVistaAndNewer = (osi.dwMajorVersion >= 6)

    hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, pid)
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(bIsWinVistaAndNewer, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, pid)
        End If
    End If

    If hProc Then
        If bIsWinVistaAndNewer Then
            cnt = MAX_PATH + 1
            ProcPath = String$(cnt, 0&)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If

        If 0 <> Err.LastDllError Or Not bIsWinVistaAndNewer Then
            ProcPath = String$(MAX_PATH, 0&)
            cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
        End If

        If ERROR_PARTIAL_COPY = Err.LastDllError Or cnt = 0 Then
            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
        End If
        CloseHandle hProc
    End If

    If cnt <> 0 Then GetFilePathByPid = Left$(ProcPath, cnt)
End Function


'Public Function GetFilePathByPid(pid As Long) As String
'
'    Dim hProc As Long
'    Dim ProcPath As String
'    Dim cnt As Long
'
'    hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION Or PROCESS_VM_READ, 0&, pid)
'
'    If hProc Then
'        cnt = MAX_PATH + 1
'        ProcPath = String$(cnt, 0&)
'        Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
'
'        If 0 <> Err.LastDllError Then
'            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
'        End If
'        CloseHandle hProc
'    End If
'
'    If cnt <> 0 Then GetFilePathByPid = Left$(ProcPath, cnt)
'End Function


