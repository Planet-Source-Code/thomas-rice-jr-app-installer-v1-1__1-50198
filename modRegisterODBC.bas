Attribute VB_Name = "modRegisterODBC"
'**************************************
' Name: Create/Check Access' DSN in ODBC
'
' Description:Code You can use for check
'     and (if not exist) create DSN for Access
'     DB in ODBC.
' By: Tair Abdurman
'
'**************************************
Option Explicit

'Constants
Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4

'API Declares
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Public Function isSZKeyExist(ByVal szKeyPath As String, _
    ByVal szKeyName As String, _
    ByRef szKeyValue As String) As Boolean
    Dim bRes As Boolean
    Dim lRes As Long
    Dim hkey As Long
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    0&, _
    KEY_QUERY_VALUE, _
    hkey)


    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    lRes = RegQueryValueEx(hkey, _
    szKeyName, _
    0&, _
    REG_SZ, _
    ByVal szKeyValue, _
    Len(szKeyValue))
    RegCloseKey (hkey)


    If lRes <> ERROR_SUCCESS Then
        isSZKeyExist = False
        Exit Function
    End If
    isSZKeyExist = True
End Function


Public Function checkAccessDriver(ByRef szDriverName As String) As Boolean

On Error Resume Next
    
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    
    bRes = False
    
    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"
    szKeyName = "Driver"
    szKeyValue = String(255, Chr(32))
    


    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        szDriverName = szKeyValue
        bRes = True
    Else
        bRes = False
    End If
    
    checkAccessDriver = bRes
End Function


Public Function checkWantedAccessDSN(ByVal szWantedDSN As String) As Boolean
    
On Error Resume Next
    
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    szKeyName = szWantedDSN
    szKeyValue = String(255, Chr(32))
    


    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then
        bRes = True
    Else
        bRes = False
    End If
    
    checkWantedAccessDSN = bRes
    
End Function


Public Function createAccessDSN(ByVal szDriverName As String, _
    ByVal szWantedDSN As String, ByVal szMDB As String) As Boolean
    
On Error Resume Next
    
    Dim hkey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\" & _
    szWantedDSN, _
    hkey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    
    lRes = RegSetValueExString(hkey, "UID", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    
    szKeyValue = szMDB 'gstrAppPath & "\DB\ssmdb.mdb"
    lRes = RegSetValueExString(hkey, "DBQ", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = szWantedDSN
    lRes = RegSetValueExString(hkey, "Description", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = szDriverName
    lRes = RegSetValueExString(hkey, "Driver", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    szKeyValue = "MS Access;"
    lRes = RegSetValueExString(hkey, "FIL", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lKeyValue = 25
    lRes = RegSetValueExLong(hkey, "DriverId", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 0
    lRes = RegSetValueExLong(hkey, "SafeTransactions", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lRes = RegCloseKey(hkey)
    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
    
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    szKeyPath, _
    hkey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    lRes = RegSetValueExString(hkey, "ImplicitCommitSync", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    szKeyValue = "Yes"
    lRes = RegSetValueExString(hkey, "UserCommitSync", 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    lKeyValue = 2048
    lRes = RegSetValueExLong(hkey, "MaxBufferSize", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 5
    lRes = RegSetValueExLong(hkey, "PageTimeout", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 3
    lRes = RegSetValueExLong(hkey, "Threads", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lRes = RegCloseKey(hkey)
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hkey)
    


    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If
    
    szKeyValue = "Microsoft Access Driver (*.mdb)"
    lRes = RegSetValueExString(hkey, szWantedDSN, 0&, REG_SZ, _
    szKeyValue, Len(szKeyValue))
    
    lRes = RegCloseKey(hkey)
    createAccessDSN = True
End Function
