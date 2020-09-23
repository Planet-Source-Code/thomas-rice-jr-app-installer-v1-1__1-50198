Attribute VB_Name = "modRegistry"
Option Explicit

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Declare Function RegSetValueEx Lib "advapi32.dll" _
   Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpData As Any, ByVal cbData As Long) As Long
   
Private Declare Function RegQueryValueExString Lib "advapi32.dll" _
   Alias "RegQueryValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, _
   lpcbData As Long) As Long
   
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" _
   Alias "RegCreateKeyExA" _
   (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, _
   ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As Any, _
   hKeyHandle As Long, lpdwDisposition As Long) As Long
   
Public Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
   Alias "RegOpenKeyExA" _
   (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
   ByVal samDesired As Long, hKeyHandle As Long) As Long

Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwNumberOfProcessors As Long
    dwActiveProcessorMask As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type MEMORYSTATUS
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
End Type


Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003

Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000

Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
   KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or _
   KEY_CREATE_SUB_KEY Or _
   KEY_ENUMERATE_SUB_KEYS Or _
   KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Method fnGetRegistryKey
'
' This function is designed to retrieve a registry key from a particular
' section of the registry. Instead of making the caller worry about the
' various constants that specify each of the hives, this function has
' optional Boolean arguments that can be set in order to select a particular
' hive.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modification History
' Date      Description
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function fnGetRegistryKey(sKey As String, sEntry As String, _
   Optional bHKeyClassesRoot As Boolean = False, _
   Optional bHKeyCurrentConfig As Boolean = False, _
   Optional bHKeyCurrentUser As Boolean = False, _
   Optional bHKeyDynamicData As Boolean = False, _
   Optional bHKeyLocalMachine As Boolean = True, _
   Optional bHKeyPerformanceData As Boolean = False, _
   Optional bHKeyUsers As Boolean = False, _
   Optional bDirectory As Boolean = False) As String
 
   Const BUFFER_LENGTH = 255
 
   Dim sKeyName As String
   Dim sReturnBuffer As String
   Dim lBufLen As Long
   Dim lReturn As Long
   Dim hKeyHandle As Long
   Dim lKeyType As Long
   '
   ' Set up return buffer
   '
   sReturnBuffer = Space(BUFFER_LENGTH)
   lBufLen = BUFFER_LENGTH
   
   lKeyType = fnDetermineKeyType(bHKeyClassesRoot, _
      bHKeyCurrentConfig, _
      bHKeyCurrentUser, _
      bHKeyDynamicData, _
      bHKeyLocalMachine, _
      bHKeyPerformanceData, _
      bHKeyUsers)
 
   lReturn = RegOpenKeyEx(lKeyType, sKey, _
      0, KEY_ALL_ACCESS, hKeyHandle)
   If lReturn = ERROR_SUCCESS Then
      lReturn = RegQueryValueExString(hKeyHandle, sEntry, _
         0, 0, sReturnBuffer, lBufLen)
      If lReturn = ERROR_SUCCESS Then
         '
         ' Have to remove the null terminator at end of string
         '
         sReturnBuffer = Trim$(left$(sReturnBuffer, lBufLen - 1))
         '
         ' Add a backslash if one isn't already on a
         ' directory entry.
         '
         If bDirectory Then
            If Right$(sReturnBuffer, 1) <> "\" Then
               sReturnBuffer = sReturnBuffer & "\"
            End If
         End If
         fnGetRegistryKey = sReturnBuffer
 
      Else
         fnGetRegistryKey = ""
      End If
   Else
      fnGetRegistryKey = ""
   End If
   '
   ' Close the key
   '
   RegCloseKey hKeyHandle
 
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Public Method subSaveRegistryKey
'
' This function is designed to save a registry key to a particular
' section of the registry. Instead of making the caller worry about the
' various constants that specify each of the hives, this function has
' optional Boolean arguments that can be set in order to select a particular
' hive.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub subSaveRegistryKey(sKey As String, _
   sEntry As String, sValue As String, _
   Optional bHKeyClassesRoot As Boolean = False, _
   Optional bHKeyCurrentConfig As Boolean = False, _
   Optional bHKeyCurrentUser As Boolean = False, _
   Optional bHKeyDynamicData As Boolean = False, _
   Optional bHKeyLocalMachine As Boolean = True, _
   Optional bHKeyPerformanceData As Boolean = False, _
   Optional bHKeyUsers As Boolean = False, _
   Optional bDirectory As Boolean = False)

 
   Dim lReturn As Long
   Dim hKeyHandle As Long
   Dim lKeyType As Long
   
   lKeyType = fnDetermineKeyType(bHKeyClassesRoot, _
      bHKeyCurrentConfig, _
      bHKeyCurrentUser, _
      bHKeyDynamicData, _
      bHKeyLocalMachine, _
      bHKeyPerformanceData, _
      bHKeyUsers)
  
   '
   ' RegCreateKeyEx will open the named key if it exists, and
   ' create a new one if it doesn't.
   '
   lReturn = RegCreateKeyEx(lKeyType, sKey, 0&, _
      vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
      0&, hKeyHandle, lReturn)
 
   '
   ' RegSetValueEx saves the value to the string within the
   ' key that was just opened.
   '
   lReturn = RegSetValueEx(hKeyHandle, sEntry, _
      0&, REG_SZ, ByVal sValue, Len(sValue))
 
 RegCloseKey hKeyHandle
 
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Private Method fnDetermineKeyType
'
' This function takes each of the seven booleans used for the other
' two functions in this file and determines which section of the
' registry is being used.  It then returns the appropriate constant
' to the caller.  This function is basically designed to simplify the
' code in each of the functions in this file.

Public Function fnDetermineKeyType(bHKeyClassesRoot As Boolean, _
   bHKeyCurrentConfig As Boolean, _
   bHKeyCurrentUser As Boolean, _
   bHKeyDynamicData As Boolean, _
   bHKeyLocalMachine As Boolean, _
   bHKeyPerformanceData As Boolean, _
   bHKeyUsers As Boolean) As Long

   Dim lResult As Long
   
   If bHKeyClassesRoot Then
      lResult = HKEY_CLASSES_ROOT
   ElseIf bHKeyCurrentConfig Then
      lResult = HKEY_CURRENT_CONFIG
   ElseIf bHKeyCurrentUser Then
      lResult = HKEY_CURRENT_USER
   ElseIf bHKeyDynamicData Then
      lResult = HKEY_DYN_DATA
   ElseIf bHKeyLocalMachine Then
      lResult = HKEY_LOCAL_MACHINE
   ElseIf bHKeyPerformanceData Then
      lResult = HKEY_PERFORMANCE_DATA
   ElseIf bHKeyUsers Then
      lResult = HKEY_USERS
   End If

   fnDetermineKeyType = lResult
End Function

Public Function MemoryAvailable() As Long
Dim memsts As MEMORYSTATUS
GlobalMemoryStatus memsts
MemoryAvailable = memsts.dwAvailPhys
End Function

Public Function MemoryTotal() As Long
Dim memsts As MEMORYSTATUS
GlobalMemoryStatus memsts
MemoryTotal = memsts.dwTotalPhys
End Function
