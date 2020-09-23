Attribute VB_Name = "modSQLDSN"

Option Explicit

Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long


Public Function DoesKeyExist(RegKeyPath As String, _
    RegKeyName As String, _
    ByRef RegKeyValue As String) As Boolean
    Dim DoesIt As Boolean
    Dim Result As Long
    Dim hKey As Long
    Result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, RegKeyPath, 0&, KEY_QUERY_VALUE, hKey)


    If Result <> ERROR_SUCCESS Then
        DoesKeyExist = False
        Exit Function
    End If
    Result = RegQueryValueEx(hKey, RegKeyName, 0&, REG_SZ, ByVal RegKeyValue, Len(RegKeyValue))
    RegCloseKey (hKey)


    If Result <> ERROR_SUCCESS Then
        DoesKeyExist = False
        Exit Function
    End If
    DoesKeyExist = True
End Function


Public Function checkSQLDriver(ByRef DriverODBC As String) As Boolean
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim DoesIt As Boolean
    
    
    DoesIt = False
    
    RegKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\SQL Server"
    RegKeyName = "Driver"
    RegKeyValue = String(255, Chr(32))
    


    If DoesKeyExist(RegKeyPath, RegKeyName, RegKeyValue) Then
        DriverODBC = RegKeyValue
        DoesIt = True
    Else
        DoesIt = False
    End If
    
    checkSQLDriver = DoesIt
End Function


Public Function SQLDSNWanted(NameDSN As String) As Boolean
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim DoesIt As Boolean
    
    RegKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    RegKeyName = NameDSN
    RegKeyValue = String(255, Chr(32))
    


    If DoesKeyExist(RegKeyPath, RegKeyName, RegKeyValue) Then
        DoesIt = True
    Else
        DoesIt = False
    End If
    
    SQLDSNWanted = DoesIt
    
End Function


Public Function MakeSQLDSN(DriverODBC As String, _
    NameDSN As String) As Boolean
    
    Dim hKey As Long
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim lKeyValue As Long
    Dim Result As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
    NameDSN, hKey)
    

    If Result <> ERROR_SUCCESS Then
        MakeSQLDSN = False
        Exit Function
    End If
    ' For User ID
    'Result = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
    'szEmpty, Len(szEmpty))
    
    'Camel for Coyote
' Edit the next line to reflect your Server Name
    RegKeyValue = "Coite"
    Result = RegSetValueExString(hKey, "Server", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    RegKeyValue = DriverODBC
    Result = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
' Edit the Next Line to Revise Description
    RegKeyValue = "Coyote Server Conn"
    Result = RegSetValueExString(hKey, "Description", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
' Edit the next line to Revise Database Name
    RegKeyValue = "Northwind"
    Result = RegSetValueExString(hKey, "Database", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
' Working with this for now this is the user logged on
    RegKeyValue = "Coyote"
    Result = RegSetValueExString(hKey, "LastUser", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
'    lKeyValue = 25
'    Result = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, _
'    lKeyValue, 4)
    
    RegKeyValue = "Yes"
    Result = RegSetValueExString(hKey, "Trusted_Connection", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
' To require password validation comment the above three lines
' and un-comment the following three lines
'    lKeyValue = 0
'    Result = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, _
'    lKeyValue, 4)
    
'    Result = RegCloseKey(hKey)
'    RegKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & NameDSN
    
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, _
    RegKeyPath, _
    hKey)
    


    If Result <> ERROR_SUCCESS Then
        MakeSQLDSN = False
        Exit Function
    End If
    Result = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    RegKeyValue = "Yes"
    Result = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    lKeyValue = 2048
    Result = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 5
    Result = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 3
    Result = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    Result = RegCloseKey(hKey)
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)
    


    If Result <> ERROR_SUCCESS Then
        MakeSQLDSN = False
        Exit Function
    End If
    
    RegKeyValue = "SQL Server"
    Result = RegSetValueExString(hKey, NameDSN, 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    Result = RegCloseKey(hKey)
    MakeSQLDSN = True
End Function

