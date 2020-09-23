Attribute VB_Name = "modShellCommand"
Option Explicit

Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const MOVEFILE_COPY_ALLOWED = &H2
Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4
        
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Private Const DEBUG_PROCESS = &H1
Private Const DEBUG_ONLY_THIS_PROCESS = &H2

Private Const CREATE_SUSPENDED = &H4

Private Const DETACHED_PROCESS = &H8

Private Const CREATE_NEW_CONSOLE = &H10

Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Private Const CREATE_NEW_PROCESS_GROUP = &H200

Private Const CREATE_NO_WINDOW = &H8000000

Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&
Private Const WAIT_TIMEOUT = &H102&

Private Const SW_SHOW = 5

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function CreateProcessBynum Lib "kernel32" Alias "CreateProcessA" _
        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long

Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long
        
Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
        
Private Declare Function WaitForInputIdle Lib "user32" _
        (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
        
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
        (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) _
        As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" _
        (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) _
        As Long
        
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias _
        "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
        
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
        (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
         ByVal bFailIfExists As Long) As Long

Public Function CreateProcess(strApp As String, strWorkDir As String, Optional strCommandLine As String) As Boolean
   
On Error Resume Next
    
    Dim res&
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    sinfo.cb = Len(sinfo)
    sinfo.lpReserved = vbNullString
    sinfo.lpDesktop = vbNullString
    sinfo.lpTitle = vbNullString
    sinfo.dwFlags = 0
    
    If Len(Trim(strCommandLine)) = 0 Then
      strCommandLine = vbNullString
    End If
    
    res = CreateProcessBynum(strWorkDir & strApp, strCommandLine, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, sinfo, pinfo)
    If res Then
        WaitForTerm2 pinfo
        CreateProcess = True
    Else
        CreateProcess = False
    End If
    
End Function

Public Sub WaitForTerm1(pid&)

On Error Resume Next
    
    Dim phnd&
    phnd = OpenProcess(SYNCHRONIZE, 0, pid)
    If phnd <> 0 Then
        Call WaitForSingleObject(phnd, INFINITE)
        Call CloseHandle(phnd)
    End If
End Sub

Public Sub WaitForTerm2(pinfo As PROCESS_INFORMATION)
On Error Resume Next
    Dim res&
    ' Let the process initialize
    Call WaitForInputIdle(pinfo.hProcess, INFINITE)
    ' We don't need the thread handle
    Call CloseHandle(pinfo.hThread)
    ' Disable the button to prevent reentrancy

    Do
        res = WaitForSingleObject(pinfo.hProcess, 0)
        If res <> WAIT_TIMEOUT Then
            ' No timeout, app is terminated
            Exit Do
        End If
        DoEvents
    Loop While True
    
    ' Kill the last handle of the process
    Call CloseHandle(pinfo.hProcess)
End Sub

Public Function CreateDir(ByVal strNewDir As String) As Boolean

On Error Resume Next
    
   Dim sinfo As SECURITY_ATTRIBUTES
   Dim intRtn As Long
    
   strNewDir = Replace(strNewDir, "\\", "\", 1, -1, vbTextCompare)
   
   sinfo.lpSecurityDescriptor = 0
   sinfo.bInheritHandle = 0
   sinfo.nLength = Len(sinfo)
   
   CreateDir = False
   
   intRtn = CreateDirectory(strNewDir, sinfo)
   CreateDir = intRtn

End Function
