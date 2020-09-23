Attribute VB_Name = "modExecInitialize"
Option Explicit
Option Compare Text

Public hndMasterWindow     As Long
Public gstrINIFile         As String
Public gstrSystemPath      As String
Public gstrProgramFiles    As String
Public gstrUserPath        As String
Public gstrCommonPrograms  As String
Public gstrSystemDrive     As String
Public gstrAppPath         As String
Public gblnSQLInstalled    As Boolean
Public gblnIISInstalled    As Boolean
Public gstrWWWRoot         As String
Public gstrZipFile         As String
Public gstrServerName      As String
Public gstrSQLServer       As String
Public gstrSQLUser         As String
Public gstrSQLPassword     As String
Public gblnError           As Boolean
Public gblnCancel          As Boolean
Public gblnDebugging       As Boolean
Public gstrInstallDir      As String
Public gstrStartUpDir      As String
Public gblnReboot          As Boolean
Public gstrPackageName     As String
Public gstrSetupDesc       As String
Public gstrWebURL          As String
Public gstrCompanyName     As String
Public gstrDesktopPath     As String
Public gstrAppVersion      As String

Public Sub SetINIFile()

On Error Resume Next
    
   Dim strKey       As String
   Dim strPath      As String
   Dim strSQLKey    As String
   Dim intRtn       As Integer
   Dim blnRtn       As Boolean
   Dim dblAvlSpace  As Double
   Dim idl          As Long
   Dim pidl         As ITEMIDLIST
   
   Dim objReg       As New Registry
   Dim objSystem    As New clsSystemInfo
   Dim objVersion   As New clsOSVersion
   
   strKey = RemoveFromString(App.Title, " ")
   
   strPath = Space(256)
   intRtn = GetCurrentDirectory(256, strPath)
   gstrStartUpDir = TrimNull(strPath)
   
   gstrINIFile = gstrStartUpDir & "\setup.ini"
   gstrINIFile = Replace(gstrINIFile, "\\", "\", 1, -1, vbTextCompare)
   
   gblnDebugging = CBool(FnNumberOnly(IniGetString("GENERAL", "DEBUG", "")))
   
   If gblnDebugging Then
     intRtn = MsgBox(gstrStartUpDir, vbApplicationModal + vbCritical + vbOKCancel)
     If intRtn = vbCancel Then
        Set objReg = Nothing
        Set objVersion = Nothing
        Set objSystem = Nothing
        End
     End If
   End If
   
   dblAvlSpace = objSystem.FreeMegaBytesOnDisk
   If dblAvlSpace < 125 Then
      intRtn = MsgBox("There is not enough room on the System Drive to Install eLinkBB." & vbCrLf & _
                      "This installation requires 125mb of free disk space. Please cleanup " & vbCrLf & _
                      "your system drive before running this installation again.", vbApplicationModal + vbCritical + vbOKOnly)
                      
      Set objReg = Nothing
      Set objVersion = Nothing
      Set objSystem = Nothing
      End
                    
   End If
   
   gstrServerName = objSystem.ComputerName
   
   gstrSystemPath = objSystem.SystemDir
   blnRtn = objVersion.IsWin2K
   If Not blnRtn And Not gblnDebugging Then
      intRtn = MsgBox("This software requires Windows 2000 or above ONLY. " & vbCrLf, vbApplicationModal + vbCritical + vbOKOnly)
      Set objReg = Nothing
      Set objVersion = Nothing
      Set objSystem = Nothing
      End
   End If
   
   If Len(gstrSystemPath) = "\" Then
      gstrSystemPath = Mid(gstrSystemPath, 1, Len(gstrSystemPath) - 1)
   End If
   
   intRtn = InStr(1, gstrSystemPath, "\")
   gstrSystemDrive = Mid(gstrSystemPath, 1, intRtn)
   
'    intRtn = SetCurrentDirectory(gstrInstallDir)
'    If intRtn = 0 Then
'       intRtn = MsgBox("Could not locate the Install Directory. Please run 'SETUP.BAT' from the Install CD again.", vbApplicationModal + vbCritical + vbOKOnly)
'       Set objReg = Nothing
'       Set objVersion = Nothing
'       Set objSystem = Nothing
'       End
'    End If
      
    
   SHGetSpecialFolderLocation 0&, CSIDL_PROGRAM_FILES, pidl
   idl = pidl.mkid.cb
   gstrProgramFiles = Space(512)
   SHGetPathFromIDList ByVal idl&, ByVal gstrProgramFiles
   gstrProgramFiles = TrimNull(gstrProgramFiles) & "\"
   
   SHGetSpecialFolderLocation 0&, CSIDL_COMMON_PROGRAMS, pidl
   idl = pidl.mkid.cb
   gstrUserPath = Space(512)
   SHGetPathFromIDList ByVal idl&, ByVal gstrUserPath
   gstrUserPath = TrimNull(gstrUserPath) & "\"
   
   
   SHGetSpecialFolderLocation 0&, CSIDL_COMMON_DESKTOPDIRECTORY, pidl
   idl = pidl.mkid.cb
   gstrDesktopPath = Space(512)
   SHGetPathFromIDList ByVal idl&, ByVal gstrDesktopPath
   gstrDesktopPath = TrimNull(gstrDesktopPath) & "\"
   
   SHGetSpecialFolderLocation 0&, CSIDL_PROGRAM_FILES_COMMON, pidl
   idl = pidl.mkid.cb
   gstrCommonPrograms = Space(512)
   SHGetPathFromIDList ByVal idl&, ByVal gstrCommonPrograms
   gstrCommonPrograms = TrimNull(gstrCommonPrograms) & "\"
   
   'Sometimes sProgDir Is Empty When The User Has Program Files Folder In Different Drives
   If gstrProgramFiles = "\" Then
     gstrProgramFiles = gstrSystemDrive & "Program Files\"
   End If
    
'    'Get System Path
'    gstrSystemPath = GetKeyValue("SOFTWARE\Microsoft\Windows NT\", "CurrentVersion", "SystemRoot")
'    If Len(gstrSystemPath) = 0 Then
'        gstrSystemPath = GetKeyValue("", "SOFTWARE\Microsoft\Windows\CurrentVersion", "SystemRoot")
'        gstrSystemPath = gstrSystemPath & "\System\"
'    Else
'        gstrSystemPath = gstrSystemPath & "\System32\"
'    End If
    
   'Check for MS SQL Server
   strKey = GetKeyValue("SOFTWARE\Microsoft\Microsoft SQL Server\80\", "Registration", "CD_KEY")
   If Len(Trim(strKey)) > 0 Then
      gblnSQLInstalled = True
   End If
    
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MMC\SnapIns\{A841B6C2-7577-11D0-BB1F-00A0C922E79C}
    'Check for IIS Server
   strKey = GetKeyValue("SOFTWARE\Microsoft\", "INetStp", "MajorVersion")
   If Len(Trim(strKey)) > 0 Then
      strKey = Trim(strKey)
      If Val(Mid(strKey, 1, 1)) >= 5 Then
        gblnIISInstalled = True
        gstrWWWRoot = GetKeyValue("SOFTWARE\Microsoft\", "INetStp", "PathWWWRoot")
      End If
   Else
      gstrWWWRoot = gstrSystemDrive & "\InetPub\wwwroot\"
   End If
    
'    gstrINIFile = gstrInstallDir & "\Install.Ini"
    
   gstrInstallDir = IniGetString("STARTUP", "InstallDir", "")
   gstrInstallDir = fnGetDirectory(gstrInstallDir)
    
   strKey = IniGetString("STARTUP", "RESTART", "")
   gblnReboot = (UCase(Trim(strKey)) = "YES")

   gstrPackageName = IniGetString("STARTUP", "PackageName", "")
   gstrSetupDesc = IniGetString("STARTUP", "SetupDesc", "")
   gstrSetupDesc = Replace(gstrSetupDesc, "$(PackageName)", Trim(gstrPackageName), 1, -1, vbTextCompare)
   gstrWebURL = IniGetString("STARTUP", "WebURL", "")
   gstrAppPath = IniGetString("STARTUP", "DefaultDir", "")
   gstrAppPath = fnGetDirectory(gstrAppPath)
   gstrCompanyName = IniGetString("STARTUP", "CompanyName", "")
   gstrAppVersion = IniGetString("STARTUP", "AppVersion", "")
   If Len(Trim(gstrAppVersion)) = 0 Then
      gstrAppVersion = App.Major & "." & App.Minor & "." & App.Revision
   End If
   
   
'   If Not gblnIISInstalled And Not gblnDebugging Then
'     intRtn = MsgBox("You MUST have MS Internet Information Services installed before running this installation." & vbCrLf & _
'                     "Please try again after you have Installed IIS ! ", vbApplicationModal + vbCritical + vbOKOnly)
'      Set objReg = Nothing
'      Set objVersion = Nothing
'      Set objSystem = Nothing
'     End
'   End If
    
SetINIFile_Exit:

   Set objReg = Nothing
   Set objVersion = Nothing
   Set objSystem = Nothing
   
End Sub

Public Sub Main()

On Error Resume Next
   
    If App.PrevInstance Then
        ActivatePrevInstance
        End
    End If
    
    SetINIFile
    Load frmSetup
    
End Sub

Private Sub ActivatePrevInstance()
    
On Error Resume Next
    
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim Result As Long
    
    'Save the title of the application.
    OldTitle = App.Title
    'Rename the title of this application so
    '     FindWindow
    'will not find this application instance
    '     .
    App.Title = "unwanted instance"
    'Attempt to get window handle using VB4
    '     class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    'Check for no success.


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB5
        '     class name.
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB6
        '     class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    End If
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    Result = OpenIcon(PrevHndl)
    'Activate the application.
    Result = SetForegroundWindow(PrevHndl)
    'End the application.
    End
End Sub

