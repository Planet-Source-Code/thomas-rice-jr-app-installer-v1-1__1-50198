Attribute VB_Name = "modGeneral"
Option Explicit

Private Declare Function GetClientRect& Lib "user32" _
                        (ByVal hwnd&, Rct As RECT)
Private Declare Function GetParent& Lib "user32" _
                        (ByVal hwnd&)

Public Declare Function SetActiveWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function GetVersionEx Lib "Kernel32" _
    Alias "GetVersionExA" (lpVersionInformation As _
    OSVERSIONINFOEX) As Long
Declare Function GetTickCount& Lib "Kernel32" ()

Public Declare Function IsWindowVisible Lib "user32" _
    (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Stuff to check for Previous Instance already running
Public Const GW_HWNDPREV = 3

Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long


Public Enum DataType
 REG_SZ = &H1
 REG_EXPAND_SZ = &H2
 REG_BINARY = &H3
 REG_DWORD = &H4
 REG_MULTI_SZ = &H7
End Enum

Private Type RECT
    Lft As Long
    top As Long
    Rgt As Long
    Bot As Long
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2

Function centerForm(Frm As Form) As Boolean
    
On Error Resume Next
    
    ' constants
    Const MB_ICONEXCLAMATION = 48
    Const FIFTHEEN = 15
    Const ONE_HALF = 0.5
    
    ' variables
    Dim x As Integer
    Dim y As Integer
    Dim ParWid As Integer
    Dim ParHgt As Integer
    Dim ChildFrm As Boolean
    Dim msg As String
    Dim Rct As RECT
    
    
    ' in-line error handling
    On Error Resume Next
    
    If Frm.MDIChild Then
        If Err = False Then
            On Error GoTo ET
            ChildFrm = True
            GetClientRect GetParent(Frm.hwnd), Rct
            ParWid = (Rct.Rgt - Rct.Lft) * Screen.TwipsPerPixelY
            ParHgt = (Rct.Bot - Rct.top) * Screen.TwipsPerPixelX
            x = (ParWid - Frm.Width) * ONE_HALF
            y = (ParHgt - Frm.Height) * ONE_HALF
        End If
    End If
    
    If Not ChildFrm Then
        On Error GoTo ET
        x = (Screen.Width - Frm.Width) * ONE_HALF
        y = (Screen.Height - Frm.Height) * ONE_HALF
    End If
    
    ' center the form and return True to indicate success
    Frm.Move x, y
    Err = False
    centerForm = True
    Exit Function
    
ET:
    'msg = "Run-time error " & Err & " occured."
    'MsgBox msg, MB_ICONEXCLAMATION, "Centerform"
    Err = False
    
End Function

Public Function WindowsRunTime() As Long
On Error Resume Next
    WindowsRunTime = GetTickCount()
End Function

Public Function OSVersion() As String
    
On Error Resume Next
    
    Dim udtOSVersion As OSVERSIONINFOEX
    Dim lMajorVersion As Long
    Dim lMinorVersion As Long
    Dim lPlatformID As Long
    Dim sAns As String
    udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
    GetVersionEx udtOSVersion
    lMajorVersion = udtOSVersion.dwMajorVersion
    lMinorVersion = udtOSVersion.dwMinorVersion
    lPlatformID = udtOSVersion.dwPlatformId


    Select Case lMajorVersion
        Case 5
        sAns = "Windows 2000"
        Case 4


        If lPlatformID = VER_PLATFORM_WIN32_NT Then
            sAns = "Windows NT 4.0"
        Else
            sAns = IIf(lMinorVersion = 0, _
            "Windows 95", "Windows 98")
        End If
        Case 3


        If lPlatformID = VER_PLATFORM_WIN32_NT Then
            sAns = "Windows NT 3.x"
            
            'below should only happen if person has
            '     Win32s
            'installed
        Else
            sAns = "Windows 3.x"
        End If
        Case Else
        sAns = "Unknown Windows Version"
    End Select
OSVersion = sAns
End Function

Public Function WinVerRuntime() As String
    
On Error Resume Next
    
    Dim tMilliseconds As Double, days As Integer, hours As Integer
    Dim minutes As Integer, seconds As Integer, milliseconds As Integer
    Dim mLeft1 As Double, mLeft2 As Double, mLeft3 As Double
    tMilliseconds = WindowsRunTime
    Const MilliPerDay = 86400000
    Const MilliPerHour = 3600000
    Const MilliPerMinute = 60000
    days = Int(tMilliseconds / MilliPerDay)
    mLeft1 = tMilliseconds Mod MilliPerDay
    hours = Int(mLeft1 / MilliPerHour)
    mLeft2 = mLeft1 Mod MilliPerHour
    minutes = Int(mLeft2 / MilliPerMinute)
    mLeft3 = mLeft2 Mod MilliPerMinute
    seconds = Int(mLeft3 / 1000)
    milliseconds = mLeft3 Mod 1000
    WinVerRuntime = OSVersion & " has been running for " & _
    days & " day(s) " & hours & " hour(s) " & minutes & _
    " minutes " & seconds & " seconds."
End Function

Public Function Wait(ByVal lngNumSeconds As Long)

On Error Resume Next

    Dim dteStart        As Date
    Dim dteRightNow     As Date
    Dim lngHourDiff     As Long
    Dim lngMinuteDiff   As Long
    Dim lngSecondDiff   As Long
    Dim lngTotalMinDiff As Long
    Dim lngTotalSecDiff As Long
    
    dteStart = Now

    While True
      
      dteRightNow = Now
      lngHourDiff = Hour(dteRightNow) - Hour(dteStart)
      lngMinuteDiff = Minute(dteRightNow) - Minute(dteStart)
      lngSecondDiff = Second(dteRightNow) - Second(dteStart) + 1


      If lngSecondDiff = 60 Then
          lngMinuteDiff = lngMinuteDiff + 1 ' Add 1 to minute.
          lngSecondDiff = 0 ' Zero seconds.
      End If


      If lngMinuteDiff = 60 Then
          lngHourDiff = lngHourDiff + 1 ' Add 1 to hour.
          lngMinuteDiff = 0 ' Zero minutes.
      End If
      
      lngTotalMinDiff = (lngHourDiff * 60) + lngMinuteDiff ' Get totals.
      lngTotalSecDiff = (lngTotalMinDiff * 60) + lngSecondDiff

      If lngTotalSecDiff >= lngNumSeconds Then
          Exit Function
      End If

      DoEvents
          'Debug.Print dteRightNow
   Wend
   
End Function

Public Function CreateDirTree(ByVal strProgPath As String) As Boolean

On Error GoTo CreateDirTree_Exit

On Error Resume Next
   
   Dim intLimit         As Integer
   Dim strDir           As String
   Dim strPath()        As String
   Dim intLoop          As Integer
   Dim intRtn           As Integer
   Dim lngtest          As Long
   
   CreateDirTree = False
   
   strProgPath = Replace(strProgPath, "\\", "\", 1, -1, vbTextCompare)
   strPath = Split(strProgPath, "\")
   intLimit = UBound(strPath)
   
   strDir = strPath(0)
   For intLoop = 1 To intLimit
      strDir = strDir & "\" & strPath(intLoop)
      intRtn = CreateDir(strDir)
   Next

   lngtest = SetCurrentDirectory(strProgPath)
   If lngtest Then
      CreateDirTree = True
   End If
   
CreateDirTree_Exit:

End Function

Public Sub DeleteDirectory(ByRef strDirectory As String)

On Error GoTo DeleteDirectory_Error

   Dim blnExists     As Boolean
   Dim objFSO        As New FileSystemObject
   
   blnExists = objFSO.FolderExists(strDirectory)
   
   If blnExists Then
   
      objFSO.DeleteFolder strDirectory, True
   
   End If


DeleteDirectory_Exit:

Set objFSO = Nothing

Exit Sub

DeleteDirectory_Error:

   GoTo DeleteDirectory_Exit
   
End Sub

Public Function FnNumberOnly(strVar As Variant, Optional strStrip As String)

On Error Resume Next
    
    Dim intLen As Integer
    Dim intLoop As Integer
    Dim strNew As String
    Dim blnNeg As Boolean
    
    intLen = Len(strVar)
    intLoop = 1
    strNew = ""
    
    If Len(strStrip) > 0 Then
        strVar = FnStrip(strVar, strStrip)
    End If
    
    If Len(strVar) > 0 Then
        blnNeg = CBool((Mid(strVar, Len(strVar)) = "-") + (Mid(strVar, 1, 1) = "-"))
    End If
    
    If intLen > 0 Then
        
        For intLoop = 1 To intLen
            If (Asc(Mid(strVar, intLoop, 1)) >= 48 And Asc(Mid(strVar, intLoop, 1)) <= 57) _
                Or Asc(Mid(strVar, intLoop, 1)) = 46 Then
                    If Asc(Mid(strVar, intLoop, 1)) = 46 And InStr(1, strNew, ".") = 0 Then
                        strNew = strNew & Mid(strVar, intLoop, 1)
                    ElseIf (Asc(Mid(strVar, intLoop, 1)) >= 48 And Asc(Mid(strVar, intLoop, 1)) <= 57) Then
                        strNew = strNew & Mid(strVar, intLoop, 1)
                    End If
                    
            End If
        Next
    End If
    
    If Len(strNew) = 0 Then
        strNew = "0"
    Else
        If blnNeg Then
            strNew = "-" & strNew
        End If
    End If
    
    FnNumberOnly = strNew
               
End Function

Public Function FnStrip(strVar As Variant, Optional strStrip As String = ",")

On Error Resume Next
    
    Dim intPos As Integer
    
    If IsNull(strVar) Or Len(strVar) = 0 Then
        FnStrip = ""
        Exit Function
    End If
        
Next_Test:
    intPos = InStr(1, strVar, strStrip)
    If intPos = 0 Then
        FnStrip = strVar
        Exit Function
    End If
    
    strVar = Mid(strVar, 1, intPos - 1) & Mid(strVar, intPos + 1)
    GoTo Next_Test

End Function

Public Function TrimNull(ByVal item As String) As String
On Error Resume Next
    Dim nPos As Long
    nPos = InStr(item, Chr$(0))
    If nPos Then item = left$(item, nPos - 1)
    TrimNull = item
End Function

Public Function FileCopy(ByVal strFromFile As String, ByVal strToFile As String, Optional blnOverwrite As Boolean = False) As Long

On Error Resume Next

   strFromFile = Replace(strFromFile, "\\", "\", 1, -1, vbTextCompare)
   strToFile = Replace(strToFile, "\\", "\", 1, -1, vbTextCompare)
   
   FileCopy = CopyFile(strFromFile, strToFile, Abs(blnOverwrite))
   
End Function

Public Function fnGetDirectory(ByVal strDir As String) As String

On Error Resume Next
   
   Dim strTemp    As String
   Dim intPos     As Integer
   
   strTemp = strDir
   
   strTemp = Replace(strTemp, "$(InstallDir)", gstrInstallDir & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(ProgramFiles)", gstrProgramFiles & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(AppPath)", gstrAppPath & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(CommonFiles)", gstrCommonPrograms & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(WinSysPath)", gstrSystemPath & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(PathWWWRoot)", gstrWWWRoot & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(UserDir)", gstrUserPath & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(SystemDrive)", gstrSystemDrive & "\", 1, -1, vbTextCompare)
   strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName & "\", 1, -1, vbTextCompare)
   strTemp = Trim(strTemp) & "\"
   
Correct_Slash:
   intPos = InStr(1, strTemp, "\\", vbTextCompare)
   If intPos > 0 Then
      strTemp = Replace(strTemp, "\\", "\", 1, -1, vbTextCompare)
      GoTo Correct_Slash
   End If
   
   fnGetDirectory = Trim(strTemp)

End Function
