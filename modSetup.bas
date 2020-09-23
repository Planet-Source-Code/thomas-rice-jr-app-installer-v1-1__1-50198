Attribute VB_Name = "modSetup"
Option Explicit

Public arrArgArray(10)     As String
Public blnCommandtest      As Boolean
Public mblnSQLList         As Boolean

Public Sub CommandArgs()

On Error Resume Next
    
    ' Dim blnCommandtest As Boolean
    Dim intCmdCnt As Integer
    Dim intMaxArgs As Integer
    
    intMaxArgs = 1
    
    blnCommandtest = GetCommandLine(intCmdCnt, intMaxArgs)
    
End Sub

'Function SetupODBC() As Boolean
'
'    Dim strDriverName As String
'    Dim strWantedDSN As String
'    Dim strMDB As String
'    Dim intRtn As Integer
'
'    strDriverName = String(255, Chr(32))
'    strWantedDSN = "TransmitDetail"
'    strMDB = gstrAppPath & "\earnest.mdb"
'    SetupODBC = False
'
'    'are access drivers installed?
'    If Not checkAccessDriver(strDriverName) Then
'        intRtn = MsgBox("You must Install Access ODBC Drivers before use this program.", vbOK + vbCritical)
'    End If
'
'    'does our dsn exist?
'    If Not (checkWantedAccessDSN(strWantedDSN)) Then
'        If strDriverName = "" Then
'            intRtn = MsgBox("Can't find access ODBC driver.", vbOK + vbCritical)
'        Else
'            If Not createAccessDSN(strDriverName, strWantedDSN, strMDB) Then
'                intRtn = MsgBox("Can't create database ODBC.", vbOK + vbCritical)
'            End If
'        End If
'    End If
'
'    SetupODBC = True
'
'End Function

Public Function RunProgram(ByVal strProgram As String, ByVal strMessage As String, Optional strCommandLine As String = "", Optional strLocation As String = "") As Boolean
    
On Error GoTo Error_Handler

   Dim blnResult       As Boolean
   Dim strApp          As String
   Dim strDir          As String
   Dim intRtn          As Integer
   
   RunProgram = False
   
   strDir = fnGetDirectory(gstrInstallDir & "\")
   
   If Len(Trim(strLocation)) <> 0 Then
      strDir = strLocation
   End If
   
   'Microsoft Installer Package?
   If InStr(1, LCase(strProgram), ".msi") > 0 Then
      strProgram = strLocation & strProgram
      If InStr(1, strProgram, " ", vbTextCompare) > 0 Then
         strProgram = Chr(34) & strProgram & Chr(34)
      End If
      intRtn = SetCurrentDirectory(gstrInstallDir)
      strCommandLine = " /i " & strProgram
      strProgram = "msiexec.exe "
      strDir = fnGetDirectory(gstrSystemPath & "\")
   End If
   
'   If UCase(Trim(strProgram)) = "MARKWAIT.EXE" Then
'      strDir = gstrAppPath & "\Library\"
'      strCommandLine = ""
'   End If
   
   blnResult = CreateProcess(strProgram, strDir, strCommandLine)
   If Not blnResult And Len(Trim(strMessage)) > 0 Then
      intRtn = MsgBox(strMessage & vbCrLf & "Please Exit this Setup, Close ALL Running programs and Try Again.", vbCritical + vbApplicationModal + vbOKOnly)
      Exit Function
   Else
      RunProgram = True
   End If
    
Error_Handler:

End Function

'Public Function GetServerNames() As String()
'
'On Error GoTo GetServerNames_Error
'
'   Dim strNames()    As String
'   Dim intLoop       As Long
'   Dim objSQL        As New SQLDMO.SQLServer
'   Dim objNameList   As SQLDMO.NameList
'   
'   Set objNameList = objSQL.Application.ListAvailableSQLServers
'
'   For intLoop = 1 To objNameList.Count
'     ReDim Preserve strNames(intLoop - 1)
'     strNames(intLoop - 1) = objNameList(intLoop)
'   Next
'
'   GetServerNames = strNames
'   mblnSQLList = True
'   
'GetServerNames_Exit:
'
'   Set objSQL = Nothing
'   Set objNameList = Nothing
'   
'   Exit Function
'   
'GetServerNames_Error:
'
'   mblnSQLList = False
'   GoTo GetServerNames_Exit
'   
'End Function

Public Function SetRegEntries() As Boolean
   
On Error Resume Next

   Dim strVersion          As String
   Dim strKeys             As String
   Dim strValue            As String
   Dim intLoop             As Integer
   Dim strTemp             As String
   Dim strRecord()         As String
   Dim blnRtn              As Boolean
   Dim strKey              As String
   Dim strKeyRoot          As String
   Dim strSubKey           As String
   Dim strKeyValue         As String
   Dim KeyType             As hkey
   Dim KeyDataType         As DataType
   
   strVersion = gstrAppVersion
   If Len(Trim(strVersion)) = 0 Then
      strVersion = App.Major & "." & App.Minor & "." & App.Revision
   End If
   
   'First Check for the Keys to Remove
   strKeys = IniGetString("DELETEKEYS", "Keys", "")
    
   If Val(strKeys) > 0 Then
       
      For intLoop = 1 To Val(strKeys)
         strTemp = "KEY" & Trim(Str(intLoop))
         strValue = IniGetString("DELETEKEYS", strTemp, "")
         
         If Len(strValue) Then
         
            strRecord = Split(strValue, ",", -1, vbTextCompare)
            strKey = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
            If Mid(strKey, 1, 2) = "$(" Then
               KeyType = fnGetKeyType(strKey)
               strKey = Mid(strKey, 9)
            Else
               KeyType = HKEY_LOCAL_MACHINE
            End If
            strSubKey = Replace(strRecord(1), "@", "", 1, -1, vbTextCompare)
            
            blnRtn = DeleteKey(KeyType, strKey, strSubKey)

         End If
           
      Next
   
   End If
   
   
   'Now Add These Keys
   strKeys = IniGetString("ADDKEYS", "Keys", "")
    
   If Val(strKeys) > 0 Then
       
      For intLoop = 1 To Val(strKeys)
         strTemp = "KEY" & Trim(Str(intLoop))
         strValue = IniGetString("ADDKEYS", strTemp, "")
         
         If Len(strValue) Then
         
            strRecord = Split(strValue, ",", -1, vbTextCompare)
            strKey = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
            If Mid(strKey, 1, 2) = "$(" Then
               KeyType = fnGetKeyType(strKey)
               strKey = Mid(strKey, 9)
            Else
               KeyType = HKEY_LOCAL_MACHINE
            End If
            strKeyRoot = Mid(strKey, 1, InStrRev(strKey, "\", -1, vbTextCompare) - 1)
            strKey = Mid(strKey, InStrRev(strKey, "\", -1, vbTextCompare) + 1)
            strSubKey = Replace(strRecord(1), "@", "", 1, -1, vbTextCompare)
            KeyDataType = fnGetDataType(strRecord(2))
            strKeyValue = strRecord(3)
            If InStr(1, strKeyValue, "$(", vbTextCompare) = 1 Then
               strKeyValue = fnGetDirectory(strKeyValue)
            End If
            
            blnRtn = SetKeyValue(KeyType, strKeyRoot, strKey, strSubKey, strKeyValue, KeyDataType)

         End If
           
      Next
   
   End If
   
   
End Function

Public Function SetKeyValue(ByVal KeyType As hkey, ByVal strRoot As String, ByVal strKey As String, ByVal strValue As String, ByVal strData As String, Optional ByVal DType As DataType = REG_BINARY) As Boolean

On Error Resume Next
   
   Dim blnRtn          As Boolean
   Dim objReg          As New Registry
   
   objReg.hkey = KeyType
   objReg.KeyRoot = strRoot    '"Software\Rice WebDesigns\eLinkBB"
   objReg.Subkey = strKey
   
   If Not objReg.KeyExists Then
      objReg.CreateKey
   End If
   
   SetKeyValue = objReg.SetRegistryValue(strValue, strData, DType)

   Set objReg = Nothing
   
End Function

Public Function DeleteKey(ByVal KeyType As hkey, ByVal strRoot As String, ByVal strKeyName As String) As Boolean

On Error Resume Next
   
   Dim blnRtn          As Boolean
   Dim objReg          As New Registry
   
   objReg.hkey = KeyType
   objReg.KeyRoot = strRoot    '"Software\Rice WebDesigns\eLinkBB"
   'objReg.Subkey = strKey
   
   DeleteKey = objReg.DeleteKey(strKeyName)

   Set objReg = Nothing
   
End Function

Public Function GetKeyValue(ByVal strRoot As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal blnDecrypt As Boolean = False) As Variant

On Error Resume Next
   
   Dim blnRtn          As Boolean
   Dim objReg          As New Registry
'   Dim objEncrypt      As New clsRC4

   objReg.hkey = HKEY_LOCAL_MACHINE
   objReg.KeyRoot = strRoot '"Software\Rice WebDesigns\eLinkBB"
   objReg.Subkey = strKey '"eLinkShip\UPS"

'   If blnDecrypt Then
'      GetKeyValue = objEncrypt.DecryptString(objReg.GetRegistryValue(strValue, "EMPTY"))
'   Else
      GetKeyValue = objReg.GetRegistryValue(strValue, "EMPTY")
'   End If

'   Set objEncrypt = Nothing
   Set objReg = Nothing

End Function

Public Function Unzip(ByVal intCount As Integer) As Boolean

On Error GoTo UnZip_Error

   Dim strMsgTmp        As String
   Dim intStatus        As Integer
   Dim strValue         As String
   Dim intRtn           As Integer
   Dim strRecord()      As String
   Dim strZipFile       As String
   Dim strZipDir        As String
   Dim strDestDir       As String
   
   Unzip = True
      
   strValue = "FILE" & Trim(Str(intCount))
   strValue = IniGetString("ZIPFILES", strValue, "")
   
   If Len(Trim(strValue)) Then
      strRecord = Split(strValue, ",", -1, vbTextCompare)
      
      'intPos = InStr(1, strValue, "@")
      'intPos1 = InStr(intPos, strValue, ",")
      strZipFile = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
      
      '(strValue, intPos + 1, intPos1 - 2)
      
      strZipDir = strRecord(1)
      strZipDir = fnGetDirectory(strZipDir)
      
      strDestDir = strRecord(2)
      strDestDir = fnGetDirectory(strDestDir)
      
      intRtn = SetCurrentDirectory(strZipDir)

   End If
   
'   If gblnDebugging And intCount = 5 Then
'      intRtn = MsgBox("strValue: " & strValue & vbCrLf & "strZipFile: " & strZipFile & vbCrLf & "Dir: " & strDirectory, vbOKOnly + vbApplicationModal)
'   End If

   '-- Init Global Message Variables
   uZipInfo = ""
   uZipNumber = 0   ' Holds The Number Of Zip Files
   
   '-- Select UNZIP32.DLL Options - Change As Required!
   uPromptOverWrite = 0  ' 1 = Prompt To Overwrite
   uOverWriteFiles = 1   ' 1 = Always Overwrite Files
   uDisplayComment = 0   ' 1 = Display comment ONLY!!!
   
   '-- Change The Next Line To Do The Actual Unzip!
   uExtractList = 0       ' 1 = List Contents Of Zip 0 = Extract
   uHonorDirectories = 1  ' 1 = Honour Zip Directories
   
   '-- Select Filenames If Required
   '-- Or Just Select All Files
   uZipNames.uzFiles(0) = vbNullString
   uNumberFiles = 0
   
   '-- Select Filenames To Exclude From Processing
   ' Note UNIX convention!
   '   vbxnames.s(0) = "VBSYX/VBSYX.MID"
   '   vbxnames.s(1) = "VBSYX/VBSYX.SYX"
   '   numx = 2
   
   '-- Or Just Select All Files
   uExcludeNames.uzFiles(0) = vbNullString
   uNumberXFiles = 0
   
   '-- Change The Next 2 Lines As Required!
   '-- These Should Point To Your Directory
   uZipFileName = strZipFile  'ZipFName.Text
   uExtractDir = strDestDir 'ExtractRoot.Text
   If uExtractDir <> "" Then
     uExtractList = 0 ' unzip if dir specified
   End If
   
   '-- Let's Go And Unzip Them!
   Call VBUnZip32
   
   '-- Tell The User What Happened
   'If Len(uZipMessage) > 0 Then
   '    strMsgTmp = uZipMessage
   'End If
   
   '-- Display Zip File Information.
   'If Len(uZipInfo) > 0 Then
   '    strMsgTmp = strMsgTmp & vbNewLine & "uZipInfo is:" & vbNewLine & uZipInfo
   'End If
   
   '-- Display The Number Of Extracted Files!
   'If uZipNumber > 0 Then
   '    strMsgTmp = strMsgTmp & vbNewLine & "Number Of Files: " & Str(uZipNumber)
   'End If
   
   'MsgOut.Text = MsgOut.Text & strMsgTmp & vbNewLine
    
Unzip_Exit:

   Exit Function
   
UnZip_Error:

   Unzip = False
   GoTo Unzip_Exit
End Function

Public Function RegisterDLL(ByVal strFileName As String) As Boolean

On Error GoTo RegisterDLL_Error

    ' Registers the current component
    Dim oRegSvr As CRegisterServer
    Dim lRet    As ERegisterServerReturn
    Dim intRtn As Integer
    
    ' Basic object initialization
    Set oRegSvr = New CRegisterServer
    
    strFileName = Replace(strFileName, "\\", "\", 1, -1, vbTextCompare)
    
    With oRegSvr
        .FileName = strFileName
        
        lRet = .Register
        
'        Select Case lRet
'            Case eRegisterServerReturn_CannotBeLoaded
'
'                intRtn = MsgBox("The file cannot be loaded.", vbCritical)
'
'            Case eRegisterServerReturn_NotAValidComponent
'
'                intRtn = MsgBox("The file is not a valid component.", vbCritical)
'
'            Case eRegisterServerReturn_RegFailed
'
'                intRtn = MsgBox("Registration failed", vbCritical)
'
'            'Case eRegisterServerReturn_RegSuccess
'            '    intRtn = MsgBox("The file was successfully registered.", vbInformation)
'        End Select

    End With
    
RegisterDLL_Exit:

    Set oRegSvr = Nothing
   
   Exit Function
   
RegisterDLL_Error:

   RegisterDLL = False
   GoTo RegisterDLL_Exit
   
End Function

Public Function UnRegisterDll(ByVal strFileName As String) As Boolean

On Error GoTo UnRegisterDLL_Error

    ' Registers the current component
    Dim oRegSvr As CRegisterServer
    Dim lRet    As ERegisterServerReturn
    Dim intRtn As Integer
    
    ' Basic object initialization
    Set oRegSvr = New CRegisterServer
    
    strFileName = Replace(strFileName, "\\", "\", 1, -1, vbTextCompare)
    
    With oRegSvr
        .FileName = strFileName
        
        lRet = .Unregister
        
'        Select Case lRet
'            Case eRegisterServerReturn_CannotBeLoaded
'
'                intRtn = MsgBox("The file cannot be loaded.", vbCritical)
'
'            Case eRegisterServerReturn_NotAValidComponent
'
'                intRtn = MsgBox("The file is not a valid component.", vbCritical)
'
'            Case eRegisterServerReturn_RegFailed
'
'                intRtn = MsgBox("Registration failed", vbCritical)
'
'            'Case eRegisterServerReturn_RegSuccess
'            '    intRtn = MsgBox("The file was successfully registered.", vbInformation)
'        End Select
    End With
    
UnRegisterDLL_Exit:

    Set oRegSvr = Nothing
    
    Exit Function

UnRegisterDLL_Error:
   
   UnRegisterDll = False
   GoTo UnRegisterDLL_Exit
   
End Function

Public Sub SetForgroundWindow(ByVal hndSelected As Long)
    
    Dim lngWinHnd As Long
    
    lngWinHnd = SetActiveWindow(hndSelected)
    lngWinHnd = SetForegroundWindow(hndSelected)

End Sub

Public Function fnGetKeyType(ByVal strKey As String) As hkey

On Error Resume Next

   If InStr(1, UCase(strKey), "$(HKCR)", vbTextCompare) Then
      fnGetKeyType = HKEY_CLASSES_ROOT
   ElseIf InStr(1, UCase(strKey), "$(HKLM)", vbTextCompare) Then
      fnGetKeyType = HKEY_LOCAL_MACHINE
   ElseIf InStr(1, UCase(strKey), "$(HKCU)", vbTextCompare) Then
      fnGetKeyType = HKEY_CURRENT_USER
   ElseIf InStr(1, UCase(strKey), "$(HKUS)", vbTextCompare) Then
      fnGetKeyType = HKEY_USERS
   ElseIf InStr(1, UCase(strKey), "$(HKPD)", vbTextCompare) Then
      fnGetKeyType = HKEY_PERFORMANCE_DATA
   ElseIf InStr(1, UCase(strKey), "$(HKCC)", vbTextCompare) Then
      fnGetKeyType = HKEY_CURRENT_CONFIG
   ElseIf InStr(1, UCase(strKey), "$(HKDD)", vbTextCompare) Then
      fnGetKeyType = HKEY_DYN_DATA
   End If
   
End Function

Public Function fnGetDataType(ByVal strKey As String) As hkey

On Error Resume Next

   If InStr(1, UCase(strKey), "REG_SZ", vbTextCompare) Then
      fnGetDataType = REG_SZ
   ElseIf InStr(1, UCase(strKey), "REG_EXPAND_SZ", vbTextCompare) Then
      fnGetDataType = REG_EXPAND_SZ
   ElseIf InStr(1, UCase(strKey), "REG_BINARY", vbTextCompare) Then
      fnGetDataType = REG_BINARY
   ElseIf InStr(1, UCase(strKey), "REG_DWORD", vbTextCompare) Then
      fnGetDataType = REG_DWORD
   ElseIf InStr(1, UCase(strKey), "REG_MULTI_SZ", vbTextCompare) Then
      fnGetDataType = REG_MULTI_SZ
   End If
   
End Function

