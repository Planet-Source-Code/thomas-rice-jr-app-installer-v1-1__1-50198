VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegisterServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************************************************************************************
'* Class       : CRegisterServer
'* Description : Class for register or unregsiter a component to the system registry.
'* Notes       : Use this class to register or unregister a component (exe, dll, ocx file)
'*               whitout the use of regsvr32.exe.
'*****************************************************************************************

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17980 + vbObjectError
Public Enum EErrRegisterServer
    eErrRegisterServer_InvalidFileName = E_ERR_BASE + 1
    eErrRegisterServer_ComponentFailure
End Enum
Private Const S_ERR_InvalidFileName = "Invalid file name"
Private Const S_ERR_ComponentFailure = "CRegisterServer component failure"

' Public class enums
Public Enum ERegisterServerReturn
    eRegisterServerReturn_CannotBeLoaded = 1
    eRegisterServerReturn_NotAValidComponent = 2
    eRegisterServerReturn_RegFailed = 3
    eRegisterServerReturn_RegSuccess = 4
    eRegisterServerReturn_UnregSuccess = 5
    eRegisterServerReturn_UnregFailed = 6
End Enum

' Private class API function declarations
Private Declare Function LoadLibraryRegister Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

' Private class constants
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

' Private variables to hold property values
Private m_sFileName As String


'*****************************************************************************************
'* Property    : FileName
'* Notes       : Property used to store/retrieve the file name (*.dll, *.ocx, etc) of
'*               the COM server that will be / was registered/unregistered
'*****************************************************************************************
Public Property Get FileName() As String
    On Error GoTo hComponentFailure
    
    FileName = m_sFileName

    Exit Property

hComponentFailure:
    Err.Raise eErrRegisterServer_ComponentFailure, App.EXEName & ".CRegisterServer", S_ERR_ComponentFailure
End Property

Public Property Let FileName(Value As String)
    On Error GoTo hComponentFailure
    
    If Len(Value) = 0 Then
        On Error GoTo 0
        Err.Raise eErrRegisterServer_InvalidFileName, App.EXEName & ".CRegisterServer", S_ERR_InvalidFileName
    End If
    
    m_sFileName = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrRegisterServer_ComponentFailure, App.EXEName & ".CRegisterServer", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : Register
'* Notes       : Registers the current component (ActiveX DLL/EXE/OCX) to the registry.
'*               Returns the result of the operation.
'*****************************************************************************************
Public Function Register() As ERegisterServerReturn
    On Error GoTo hComponentFailure
    
    Dim dwExitCode   As Long
    Dim fSuccess     As Long
    Dim hThread      As Long
    Dim lLib         As Long
    Dim lProcAddress As Long
    Dim lpThreadID   As Long
    
    If m_sFileName = "" Then Exit Function
    
    lLib = LoadLibraryRegister(m_sFileName)
    If lLib = 0 Then
        Register = eRegisterServerReturn_CannotBeLoaded
        Exit Function
    End If

    lProcAddress = GetProcAddressRegister(lLib, "DllRegisterServer")

    If lProcAddress = 0 Then
        
        Register = eRegisterServerReturn_NotAValidComponent
        If lLib Then Call FreeLibraryRegister(lLib)
        
        Exit Function
    
    Else
        
        hThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lpThreadID)
        
        If hThread Then
            
            fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
            
            If Not fSuccess Then
                Call GetExitCodeThread(hThread, dwExitCode)
                Call ExitThread(dwExitCode)
                
                Register = eRegisterServerReturn_RegFailed
                
                If lLib Then Call FreeLibraryRegister(lLib)
                
                Exit Function
            Else
                Register = eRegisterServerReturn_RegSuccess
            End If
            
            Call CloseHandle(hThread)
            
            If lLib Then Call FreeLibraryRegister(lLib)
        
        End If
    
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegisterServer_ComponentFailure, App.EXEName & ".CRegisterServer", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : Unregister
'* Notes       : Unregisters the current component (ActiveX DLL/EXE/OCX) from the registry.
'*               Returns the result of the operation.
'*****************************************************************************************
Public Function Unregister() As ERegisterServerReturn
    On Error GoTo hComponentFailure
    
    Dim dwExitCode   As Long
    Dim fSuccess     As Long
    Dim hThread      As Long
    Dim lLib         As Long
    Dim lProcAddress As Long
    Dim lpThreadID   As Long
    
    If m_sFileName = "" Then Exit Function
    
    lLib = LoadLibraryRegister(m_sFileName)
    If lLib = 0 Then
        Unregister = eRegisterServerReturn_CannotBeLoaded
        Exit Function
    End If

    lProcAddress = GetProcAddressRegister(lLib, "DllUnregisterServer")

    If lProcAddress = 0 Then
        
        Unregister = eRegisterServerReturn_NotAValidComponent
        If lLib Then Call FreeLibraryRegister(lLib)
        
        Exit Function
    
    Else
        
        hThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lpThreadID)
        
        If hThread Then
            
            fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
            
            If Not fSuccess Then
                Call GetExitCodeThread(hThread, dwExitCode)
                Call ExitThread(dwExitCode)
                
                Unregister = eRegisterServerReturn_UnregFailed
                
                If lLib Then Call FreeLibraryRegister(lLib)
                
                Exit Function
            Else
                Unregister = eRegisterServerReturn_UnregSuccess
            End If
            
            Call CloseHandle(hThread)
            
            If lLib Then Call FreeLibraryRegister(lLib)
        
        End If
    
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrRegisterServer_ComponentFailure, App.EXEName & ".CRegisterServer", S_ERR_ComponentFailure
End Function
