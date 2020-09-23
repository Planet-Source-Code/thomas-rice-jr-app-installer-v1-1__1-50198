VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Installing $(PackageName)"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetup1.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5790
      Picture         =   "frmSetup1.frx":1272
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   4680
      Width           =   255
   End
   Begin VB.Frame fraInstall 
      Caption         =   "Installation Progress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   7
      Top             =   2610
      Width           =   5940
      Begin VB.TextBox txtSetup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1335
         TabIndex        =   10
         Top             =   345
         Width           =   4365
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Setup"
      Height          =   375
      Left            =   3780
      TabIndex        =   6
      Top             =   4230
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4890
      TabIndex        =   5
      Top             =   4230
      Width           =   900
   End
   Begin VB.PictureBox picBGHead 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6090
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6090
      Begin VB.PictureBox picBGHeadIcon 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         Picture         =   "frmSetup1.frx":17B4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblBGHeadInfo1 
         BackStyle       =   0  'Transparent
         Caption         =   "$(PackageName) - Software Installalation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1065
         TabIndex        =   2
         Top             =   330
         Width           =   3975
      End
   End
   Begin VB.Label lblInst 
      BackStyle       =   0  'Transparent
      Caption         =   "Please refer to the documentation included on this CD for complele installation instructions."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3630
      Width           =   5535
   End
   Begin VB.Label lblThanks 
      BackStyle       =   0  'Transparent
      Caption         =   "Thank you for selecting $(PackageName) software. "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   1140
      Width           =   5685
   End
   Begin VB.Label lblBGHeadInfo2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please register your copy of $(PackageName)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   9
      Top             =   4260
      Width           =   3630
   End
   Begin VB.Label lblHomepage 
      Alignment       =   2  'Center
      Caption         =   "http://www.ricewebdesigns.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   75
      MouseIcon       =   "frmSetup1.frx":1FF6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4575
      Width           =   2460
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSetup1.frx":2148
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   5700
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6090
      Y1              =   975
      Y2              =   975
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
    End
End Sub

Private Sub cmdInfo_Click()
On Error Resume Next
    'Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub cmdStart_Click()
On Error Resume Next
   'Load frmBrowse
   txtSetup.Text = "Requesting Installation Directory"
   DoEvents
   frmBrowse.Show vbModal
   hndMasterWindow = frmSetup.hwnd
   Call Setup
   Unload Me
   End
    
End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
    Set IE = Nothing
End Sub

Private Sub lblHomepage_Click()

On Error Resume Next

   If Len(Trim(gstrWebURL)) > 0 Then
      'goto Website using Internet Explorer object and make it visible
      Set IE = New InternetExplorer
      IE.Visible = True
      IE.Navigate gstrWebURL '"http://www.ricewebdesigns.com"
   End If

End Sub

Private Sub Form_Load()
    
On Error Resume Next

    Dim intRtn       As Long
    Dim strTemp      As String
    
    strTemp = frmSetup.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    frmSetup.Caption = strTemp
    
    strTemp = lblThanks.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    lblThanks.Caption = strTemp
    
    strTemp = lblBGHeadInfo1.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    lblBGHeadInfo1.Caption = strTemp
    
    strTemp = lblBGHeadInfo2.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    lblBGHeadInfo2.Caption = strTemp
    
    lblHomepage.Caption = gstrWebURL
    lblInfo.Caption = gstrSetupDesc
    
    Me.Show
    hndMasterWindow = frmSetup.hwnd
    SetForgroundWindow (hndMasterWindow)
    intRtn = centerForm(Me)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

   Set IE = Nothing

End Sub

Private Sub Setup()

On Error GoTo Setup_Error

   Dim blnRtn           As Boolean
   Dim intRtn           As Integer
   Dim intWait          As Integer
   Dim strMessage       As String
   Dim strPath          As Variant
   Dim intLimit         As Integer
   Dim intLoop          As Integer
   Dim strDir           As String
   Dim strProgPath      As String
   Dim strCommandLine   As String
   
   Dim objExit          As New clsExitWindows
   
   Screen.MousePointer = vbHourglass
   cmdStart.Enabled = False
   cmdCancel.Enabled = False
    
   txtSetup.Text = "Creating Installation Directory"
   cmdInfo.SetFocus
   DoEvents

'GoTo RegistryTest

   'Now Create Install Directory
   blnRtn = CreateDirTree(gstrInstallDir)
   blnRtn = SetCurrentDirectory(gstrInstallDir)
   If Not blnRtn Then
      intRtn = MsgBox("Could not locate the Install Directory. Please re-start the Install CD again.", vbApplicationModal + vbCritical + vbOKOnly)
      GoTo Setup_Exit
   End If
   intRtn = FileCopy(gstrStartUpDir & "\install.zip", gstrInstallDir & "\install.zip", 0)
   intRtn = FileCopy(gstrStartUpDir & "\setup.exe", gstrInstallDir & "\SETUP.exe", 0)
   intRtn = FileCopy(gstrStartUpDir & "\setup.ini", gstrInstallDir & "\setup.ini", 0)
   intRtn = FileCopy(gstrStartUpDir & "\msvbvm60.dll", gstrInstallDir & "\msvbvm60.dll", 0)
   intRtn = FileCopy(gstrStartUpDir & "\unzip32.dll", gstrInstallDir & "\unzip32.dll", 0)
   
   blnRtn = SetCurrentDirectory(gstrInstallDir)
   
   'Now Create Program Files Path
   blnRtn = CreateDirTree(gstrAppPath)
   blnRtn = SetCurrentDirectory(gstrAppPath)
   If Not blnRtn Then
      intRtn = MsgBox("Could Not Create Application Install Directories." & vbCrLf & "Please close ALL Windows programs and restart this Setup.", vbCritical + vbOKOnly + vbApplicationModal)
      GoTo Setup_Error
   End If
  
   'Setup New Directories
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Setting Up Installation Directories"
   cmdInfo.SetFocus
   DoEvents
   blnRtn = SetupDirs
   
   blnRtn = SetCurrentDirectory(gstrInstallDir)
   
   blnRtn = ProcessZipFiles
   If Not blnRtn Then
      intRtn = MsgBox("Could Not Process Install Archive Files." & vbCrLf & "Please close ALL Windows programs and restart this Setup.", vbCritical + vbOKOnly + vbApplicationModal)
      GoTo Setup_Error
   End If
   
   'Move Required files to Windows System Directory
   txtSetup.Text = "Installing Windows Shared Components"
   cmdInfo.SetFocus
   DoEvents
   blnRtn = CopyFiles("SystemFiles")
   If Not blnRtn Then
      intRtn = MsgBox("Could Not Copy Shared Components." & vbCrLf & "Please close ALL Windows programs and restart this Setup.", vbCritical + vbOKOnly + vbApplicationModal)
      GoTo Setup_Error
   End If

   'Copy and Extra Files
   txtSetup.Text = "Installing Application Specific Files."
   cmdInfo.SetFocus
   DoEvents
   blnRtn = CopyFiles("AppFiles")
   If Not blnRtn Then
      intRtn = MsgBox("Could Not Copy Application Files." & vbCrLf & "Please close ALL Windows programs and restart this Setup.", vbCritical + vbOKOnly + vbApplicationModal)
      GoTo Setup_Error
   End If
   intRtn = SetCurrentDirectory(gstrInstallDir)

   'Register eLink DLL's
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Registering Server Components"
   cmdInfo.SetFocus
   DoEvents
   Call RegisterFiles
       
   intRtn = SetCurrentDirectory(gstrInstallDir)

   'Now Install System Add-On Software
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Installing Add-On Software"
   cmdInfo.SetFocus
   DoEvents
'   blnRtn = RunPrograms
'   If Not blnRtn Then
'      GoTo Setup_Error
'   End If
   
   'Now Setup the Program Group and Links
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Creating Start Menu Items"
   cmdInfo.SetFocus
   DoEvents
   blnRtn = SetupLinks
   If Not blnRtn Then
      GoTo Setup_Error
   End If
   
   'Setup Install Registry Entries
   blnRtn = SetRegEntries

Register:
   'last but not Least, Register the Product
'   DoEvents
'   SetForgroundWindow (hndMasterWindow)
'   txtSetup.Text = "Running Product Registration Software"
'   cmdInfo.SetFocus
'   DoEvents
'   strMessage = "Unable to Register eLinkCart. Please contact your Reseller"
'   intRtn = SetCurrentDirectory(strProgPath)
'   blnRtn = RunProgram("eLinkRegister.exe", strMessage, "", strProgPath & "\")
'   If Not blnRtn Then
'      GoTo Setup_Error
'   End If
   
   'Remove Installation Files
   Screen.MousePointer = vbNormal
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Cleaning Up Installation Files"
   Screen.MousePointer = vbHourglass
   cmdInfo.SetFocus
   DoEvents
   Call DeleteDirectory(gstrInstallDir)

   'Completed the Install Process
   Screen.MousePointer = vbNormal
   txtSetup.Text = "Setup Complete"
   cmdInfo.SetFocus
   DoEvents
   Me.Enabled = True
   
   If gblnReboot Then
      frmReStart.Show vbModal
   End If
                 
   Wait (3)
   
Setup_Exit:

   Set objExit = Nothing
   
   If gblnReboot Then
      objExit.ExitWindows WE_REBOOT
   End If
   
   Exit Sub

Setup_Error:

   gblnReboot = False
   GoTo Setup_Exit

End Sub

Private Sub RegisterFiles()

On Error Resume Next

    Dim strFiles     As String
    Dim strValue     As String
    Dim intLoop      As Integer
    Dim strTemp      As String
    Dim blnRtn       As Boolean
    Dim blnReg       As Boolean
    Dim strRecord()  As String
    Dim strProgram   As String
    Dim strRegDir    As String
    
    strFiles = IniGetString("LibraryFiles", "Files", "")
    
    If Val(strFiles) > 0 Then
        
        For intLoop = 1 To Val(strFiles)
            strTemp = "FILE" & Trim(Str(intLoop))
            strValue = IniGetString("LibraryFiles", strTemp, "")
            
            If Len(strValue) Then
             
               strRecord = Split(strValue, ",", -1, vbTextCompare)
               strProgram = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
               strRegDir = strRecord(1)
               strRegDir = fnGetDirectory(strRegDir)
               blnReg = UCase(Trim(strRecord(2))) = "$(DLLSELFREGISTER)"
               
               If blnReg Then
                  blnRtn = UnRegisterDll(strRegDir & strProgram)
                  blnRtn = RegisterDLL(strRegDir & strProgram)
               End If
               
             End If
        Next
        
    End If
            
End Sub

Private Function SetupDirs() As Boolean

On Error GoTo SetupDirs_Error

   Dim intStatus        As Integer
   Dim strDirs          As String
   Dim strValue         As String
   Dim intLoop          As Integer
   Dim strTemp          As String
   
   SetupDirs = True
   
   strDirs = IniGetString("DIRECTORIES", "LIST", "")
    
   If Val(strDirs) > 0 Then
   
      For intLoop = 1 To Val(strDirs)
         strTemp = "DIR" & Trim(Str(intLoop))
         strValue = IniGetString("DIRECTORIES", strTemp, "")
         
         If Len(Trim(strValue)) Then
            intStatus = CreateDir(gstrAppPath & "\" & strValue)
            intStatus = SetCurrentDirectory(gstrAppPath & "\" & strValue)
            If intStatus = False Then
               intStatus = MsgBox("Unable to Create Directiry. Please try Again.", vbCritical + vbOKCancel)
               If intStatus = vbCancel Then
                  Exit Function
               End If
            End If
         End If
      Next
   End If

SetupDirs_Exit:

   Exit Function
   
SetupDirs_Error:

   SetupDirs = False
   GoTo SetupDirs_Exit

End Function

Private Function SetupLinks() As Boolean

On Error GoTo SetupLinks_Error

   Dim strProgPath         As String
   Dim strLnkFile          As String      ' Link file name
   Dim strExeFile          As String      ' Link - Exe file name
   Dim strWorkDir          As String      '      - Working directory
   Dim strExeArgs          As String      '      - Command line arguments
   Dim strIconFile         As String      '      - Icon File name
   Dim lngIconIdx          As Long        '      - Icon Index
   Dim lngShowCmd          As Long        '      - Program start state...
   Dim strLinkDir          As String
   Dim blnRtn              As Integer
   Dim blnDesktop          As Boolean
   Dim intStatus           As Integer
   Dim strDirs             As String
   Dim strValue            As String
   Dim intLoop             As Integer
   Dim strTemp             As String
   Dim strDescription      As String
   Dim strRecord()         As String
    
   Dim objsLnk             As New cShellLink         ' ShellLink class variable
    
   SetupLinks = False
    
'   strProgPath = Trim(strProgPath)
'   If Mid(strProgPath, Len(strProgPath)) <> "\" Then
'      strProgPath = strProgPath & "\"
'   End If
   
   'Create Program Group
'   strLinkDir = gstrUserPath & "eLinkCart"
'   blnRtn = CreateDirTree(strLinkDir)
'   If Not blnRtn Then
'      GoTo SetupLinks_Exit
'   End If
   
   strDirs = IniGetString("PROGRAMLINKS", "FILES", "")
    
   If Val(strDirs) > 0 Then
   
      For intLoop = 1 To Val(strDirs)
         strTemp = "FILE" & Trim(Str(intLoop))
         strValue = IniGetString("PROGRAMLINKS", strTemp, "")
         
         If Len(Trim(strValue)) Then
         
            strRecord = Split(strValue, ",", -1, vbTextCompare)
            strExeFile = Trim(strRecord(0))
            strProgPath = fnGetDirectory(strRecord(1))
            strLinkDir = fnGetDirectory(strRecord(2))
            strLnkFile = Trim(strLinkDir) & Trim(Mid(strExeFile, 1, InStr(1, strExeFile, ".") - 1)) & ".lnk"
            strExeFile = strProgPath & strExeFile
            strWorkDir = strProgPath
            strExeArgs = ""
            strDescription = strRecord(3)
            strIconFile = strExeFile
            lngIconIdx = 0
            lngShowCmd = 5
            
            blnDesktop = UCase(Trim(strRecord(4))) = "$(DESKTOP)"
            
            'Make sure the Link Directoy Exists
            blnRtn = CreateDirTree(strLinkDir)

            ' Create a ShellLink (ShortCut)
            intStatus = objsLnk.CreateShellLink(strLnkFile, strExeFile, strWorkDir, strExeArgs, strIconFile, lngIconIdx, lngShowCmd, strDescription)
            
            If blnDesktop Then
               strLinkDir = fnGetDirectory(gstrDesktopPath & "\")
               strExeFile = Trim(strRecord(0))
               strLnkFile = Trim(strLinkDir) & Trim(Mid(strExeFile, 1, InStr(1, strExeFile, ".") - 1)) & ".lnk"
               strExeFile = strProgPath & strExeFile
               intStatus = objsLnk.CreateShellLink(strLnkFile, strExeFile, strWorkDir, strExeArgs, strIconFile, lngIconIdx, lngShowCmd, strDescription)
            End If
         
         End If
      Next
   End If
    
   SetupLinks = True
   
SetupLinks_Exit:

   Set objsLnk = Nothing
   
   Exit Function
   
SetupLinks_Error:

   SetupLinks = False
   GoTo SetupLinks_Exit
   
End Function


Private Function CopyFiles(ByVal strType As String) As Boolean

On Error GoTo CopyFiles_Error

   Dim strFiles         As String
   Dim strValue         As String
   Dim intLoop          As Integer
   Dim strTemp          As String
   Dim intStatus        As Integer
   Dim blnReg           As Boolean
   Dim blnRtn           As Boolean
   Dim strCopyDir       As String
   Dim strFromDir       As String
   Dim strRecord()      As String
   Dim strProgram       As String
   
   CopyFiles = True
   strFiles = IniGetString(strType, "Files", "")
    
   If Val(strFiles) > 0 Then
       
      For intLoop = 1 To Val(strFiles)
         strTemp = "FILE" & Trim(Str(intLoop))
         strValue = IniGetString(strType, strTemp, "")
         
         If Len(strValue) Then
            
            strRecord = Split(strValue, ",", -1, vbTextCompare)
            strProgram = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
            strFromDir = strRecord(1)
            strFromDir = fnGetDirectory(strFromDir)
            strCopyDir = strRecord(2)
            strCopyDir = fnGetDirectory(strCopyDir)
            
            CreateDirTree (strCopyDir)
            
            blnReg = UCase(Trim(strRecord(3))) = "$(DLLSELFREGISTER)"
            
            intStatus = FileCopy(strFromDir & strProgram, strCopyDir & strProgram, False)
            
            If blnReg Then
               blnRtn = UnRegisterDll(strCopyDir & strValue)
               blnRtn = RegisterDLL(strCopyDir & strValue)
            End If
            
         End If
      Next
   End If

CopyFiles_Exit:

   Exit Function
   
CopyFiles_Error:

   CopyFiles = False
   On Error GoTo CopyFiles_Exit
   
End Function

Private Function ProcessZipFiles() As Boolean

On Error GoTo ProcessZipFiles_Error

   Dim strFiles         As String
   Dim strValue         As String
   Dim strValue1        As String
   Dim intLoop          As Integer
   Dim strTemp          As String
   Dim blnRtn           As Boolean
   
   ProcessZipFiles = True
   
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Extracting Installation Files"
   cmdInfo.SetFocus
   DoEvents
   
   strFiles = IniGetString("ZipFiles", "Files", "")
    
   If Val(strFiles) > 0 Then
       
      For intLoop = 1 To Val(strFiles)
         strTemp = "FILE" & Trim(Str(intLoop))
         strValue = IniGetString("ZipFiles", strTemp, "")
         
         If Len(strValue) Then
   
            blnRtn = Unzip(intLoop)
            If Not blnRtn Then
               GoTo ProcessZipFiles_Error
            End If
            Wait 2
      
            cmdInfo.SetFocus
            DoEvents
         
         End If
           
      Next
   
   End If

ProcessZipFiles_Exit:

Exit Function

ProcessZipFiles_Error:

   ProcessZipFiles = False
   GoTo ProcessZipFiles_Exit
   
End Function

Private Function RunPrograms() As Boolean

On Error GoTo RunPrograms_Error

   Dim strFiles         As String
   Dim strValue         As String
   Dim intLoop          As Integer
   Dim strTemp          As String
   Dim strRecord()      As String
   Dim strProgram       As String
   Dim strRunDir        As String
   Dim strMessage       As String
   Dim blnRtn           As Boolean
   
   RunPrograms = True
   
   SetForgroundWindow (hndMasterWindow)
   txtSetup.Text = "Running Installation Programs"
   cmdInfo.SetFocus
   DoEvents
   
   strFiles = IniGetString("RunProgs", "Files", "")
    
   If Val(strFiles) > 0 Then
       
      For intLoop = 1 To Val(strFiles)
         strTemp = "FILE" & Trim(Str(intLoop))
         strValue = IniGetString("RunProgs", strTemp, "")
         
         If Len(strValue) Then
         
            strRecord = Split(strValue, ",", -1, vbTextCompare)
            strProgram = Replace(strRecord(0), "@", "", 1, -1, vbTextCompare)
            strRunDir = strRecord(1)
            strRunDir = fnGetDirectory(strRunDir)
            strMessage = strRecord(2)
         
            SetForgroundWindow (hndMasterWindow)
            txtSetup.Text = strMessage
            cmdInfo.SetFocus
            DoEvents
            blnRtn = RunProgram(strProgram, "Unable to Run " & strProgram & "." & vbCrLf & " Please close ALL Windows programs and Re-Start this Intallation.", "", strRunDir)
            If Not blnRtn Then
               GoTo RunPrograms_Error
            End If
   
            cmdInfo.SetFocus
            DoEvents
            
         End If
           
      Next
   
   End If

RunPrograms_Exit:

Exit Function

RunPrograms_Error:

   RunPrograms = False
   GoTo RunPrograms_Exit
   
End Function

