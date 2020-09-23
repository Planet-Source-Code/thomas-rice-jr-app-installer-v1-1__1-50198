VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'
'   Class Module    : OSVersion
'   Description     : Operating System Version Information Class
'   Author          : Eswar Santhosh
'   Last Updated    : 08th February, 2001
'
'   Copyright Info  :
'
'   This Class module is provided AS-IS. This Class module can be used as a part of a compiled
'   executable whether freeware or not. This Class module may not be posted to any web site
'   or BBS or any redistributable media like CD-ROM without the consent of the author.
'
'   Web Site : http://www.quickvb.com/
'
'   e-mail   : esanthosh@quickvb.com
'
'
'   Revision History :
'
'---------------------------------------------------------------------------------------------

Option Explicit


'
'  API Declarations - Constants
'
Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

'
' The following declarations are valid only for Windows 2000 and later
'

' Suite Types
Private Const VER_SUITE_SMALLBUSINESS As Long = &H1
Private Const VER_SUITE_ENTERPRISE As Long = &H2
Private Const VER_SUITE_BACKOFFICE As Long = &H4
Private Const VER_SUITE_COMMUNICATIONS As Long = &H8
Private Const VER_SUITE_TERMINAL As Long = &H10
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
Private Const VER_SUITE_EMBEDDEDNT As Long = &H40
Private Const VER_SUITE_DATACENTER As Long = &H80
Private Const VER_SUITE_SINGLEUSERTS As Long = &H100

' Product Types
Private Const VER_NT_WORKSTATION As Long = &H1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Private Const VER_NT_SERVER As Long = &H3

' VerifyVersionInfo Masks
Private Const VER_MINORVERSION As Long = &H1
Private Const VER_MAJORVERSION As Long = &H2
Private Const VER_BUILDNUMBER As Long = &H4
Private Const VER_PLATFORMID As Long = &H8
Private Const VER_SERVICEPACKMINOR As Long = &H10
Private Const VER_SERVICEPACKMAJOR As Long = &H20
Private Const VER_SUITENAME As Long = &H40
Private Const VER_PRODUCT_TYPE As Long = &H80

' VerifyVersionInfo comparison Constants
Private Const VER_EQUAL = 1
Private Const VER_GREATER = 2
Private Const VER_GREATER_EQUAL = 3
Private Const VER_LESS = 4
Private Const VER_LESS_EQUAL = 5
Private Const VER_AND = 6
Private Const VER_OR = 7

Private Const VER_CONDITION_MASK = 7
Private Const VER_NUM_BITS_PER_CONDITION_MASK = 3


'
' API Declarations - Types
'
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type

'
' This structure is for Windows 2000 and later
'
Private Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128        ' Maintenance string for PSS usage
        wSPMajor As Integer                 ' Service Pack Major Version
        wSPMinor As Integer                 ' Service Pack Minor Version
        wSuiteMask As Integer               ' Suite Identifier
        bProductType As Byte                ' Server / Workstation / Domain Controller ?
        bReserved As Byte                   ' Reserved
End Type

'
' API Declarations - Functions
'
Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetOSVersionEx Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Declare Function VerifyVersionInfo Lib "kernel32" _
Alias "VerifyVersionInfoA" (lpVersionInformation As OSVERSIONINFOEX, _
ByVal dwTypeMask As Long, ByVal dwlConditionMask As Currency) _
As Long

Private Declare Function VerSetConditionMask Lib "kernel32" ( _
  ByVal ConditionMask As Currency, _
  ByVal TypeMask As Long, _
  ByVal Condition As Byte) As Currency

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" _
(dest As Any, ByVal numBytes As Long)

'
' Enums
'
Public Enum OSType
    Win32s
    Win95
    Win95OSR2
    Win98
    Win98R2
    WinME
    WinNT3
    WinNT31
    WinNT35
    WinNT4
    Win2000
    Win2003
    WinXP
    WinDontKnow
End Enum

' The Following Enums are for Windows 2000 or later.
Public Enum ProductType
    OSTypeUnSupported = 0
    OSTypeWorkStation = VER_NT_WORKSTATION
    OSTypeDomainController = VER_NT_DOMAIN_CONTROLLER
    OSTypeServer = VER_NT_SERVER
End Enum

Public Enum SuiteType
    SuiteUnSupported = 0
    SuiteBackOffice = VER_SUITE_BACKOFFICE
    SuiteDataCenterServer = VER_SUITE_DATACENTER
    SuiteAdvancedServer = VER_SUITE_ENTERPRISE
    SuiteSmallBusinessServer = VER_SUITE_SMALLBUSINESS
    SuiteSmallBusinessRestricted = VER_SUITE_SMALLBUSINESS_RESTRICTED    ' Restricted Client Licnese
    SuiteTerminalServices = VER_SUITE_TERMINAL
    
    ' Don't know what these are. They are defined in WinNT.H. But, No description
    ' exists in the Help file
    [_SuiteCommunications] = VER_SUITE_COMMUNICATIONS
    [_SuiteEmbeddedNT] = VER_SUITE_EMBEDDEDNT
    [_SuiteSingleUser] = VER_SUITE_SINGLEUSERTS
End Enum

Public Enum CheckOS
    Win9x = 0
    WinNT = 1
End Enum

Public Enum OSVerComparison
    Less = VER_LESS
    LessOrEqual = VER_LESS_EQUAL
    Equal = VER_EQUAL
    GreaterOrEqual = VER_GREATER_EQUAL
    Greater = VER_GREATER
End Enum

'
' Class Variables
'
Dim m_OS As OSType                  ' OS Enum
Dim m_OSString() As String          ' The Name of the OS

Dim m_Major As Long                 ' Major Version
Dim m_Minor As Long                 ' Minor Version
Dim m_BuildNumber As Long           ' Build Number of the OS
Dim m_SPMajor As Long               ' Service Pack Minor Version. Win2000 or later
Dim m_SPMinor As Long               ' Service Pack Minor Version. Win2000 or later
Dim m_PSSInfo As String             ' PSS Info. Contains Service Pack Info in NT

' The Following vars are for Windows 2000 or later
Dim m_IsVerifyAvailable As Boolean  ' Does the OS support VerifyVersionInfo ?
Dim m_SuiteType As SuiteType
Dim m_ProductType As ProductType

Private Sub Class_Initialize()
'
' Calls GetVersionInfo to Initialize the Object
'
Dim OSInfo As OSVERSIONINFO, OSInfoEx As OSVERSIONINFOEX, APIRetVal As Long
Dim i As Long

'
' Before anything, initialize the OSString Array with elements corresponding to
' the Enum values of OSType
'
ReDim m_OSString(0 To OSType.WinDontKnow)
m_OSString(Win32s) = "Win32s"
m_OSString(Win95) = "Windows '95"
m_OSString(Win95OSR2) = "Windows '95 OSR 2"
m_OSString(Win98) = "Windows '98"
m_OSString(Win98R2) = "Windows '98 Release 2"
m_OSString(WinME) = "Windows ME"
m_OSString(WinNT3) = "Windows NT 3.0"
m_OSString(WinNT31) = "Windows NT 3.1"
m_OSString(WinNT35) = "Windows NT 3.5"
m_OSString(WinNT4) = "Windows NT 4.0"
m_OSString(Win2000) = "Windows 2000"
m_OSString(Win2003) = "Windows 2003"
m_OSString(WinXP) = "Windows XP"
m_OSString(WinDontKnow) = "[Unknown]"

' Initialize the Structure
OSInfo.dwOSVersionInfoSize = Len(OSInfo)
OSInfoEx.dwOSVersionInfoSize = Len(OSInfoEx)

APIRetVal = GetOSVersion(OSInfo)

'
' Check if the OS is Windows 2000 or later. If so, it provides additional information
' through OSVERSIONINFOEX and Supports VerifyVersionInfo API.
'
If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
   
    '
    ' Use Windows 2000 OSINFOEX Structure
    '
    GetOSVersionEx OSInfoEx
    
    '
    ' Build the Results
    '
    m_OS = Win2000
    
    m_Major = OSInfoEx.dwMajorVersion
    m_Minor = OSInfoEx.dwMinorVersion
    m_BuildNumber = OSInfoEx.dwBuildNumber
    m_SPMajor = OSInfoEx.wSPMajor
    m_SPMinor = OSInfoEx.wSPMinor
    m_PSSInfo = OSInfoEx.szCSDVersion
    
    ' Get PSSInfo String
    If Len(m_PSSInfo) > 0 Then
        If InStr(m_PSSInfo, Chr$(0)) > 0 Then
            m_PSSInfo = left$(m_PSSInfo, InStr(m_PSSInfo, Chr$(0)) - 1)
        End If
    End If
    
    m_IsVerifyAvailable = True
    
    m_ProductType = OSInfoEx.bProductType
    m_SuiteType = OSInfoEx.wSuiteMask
Else
'
' May be one of those Win9x or NT 3 to 4.51
' The Common Items are Major and Minor Verions and the PSS Strings
' As far as the Build Number goes, the NT/2000 systems return the Build Number in
' dwBuildNumber of the OSINFO structure. For Win9x, the HighWord of dwBuildNumber
' member contains the Major and Minor Versions and the BuildNumber is stored in the Lower
' order byte of the member
'

' Fill in the variables
m_Major = OSInfo.dwMajorVersion
m_Minor = OSInfo.dwMinorVersion
m_BuildNumber = OSInfo.dwBuildNumber

m_PSSInfo = OSInfo.szCSDVersion
    
' Features not available other than Windows 2000 or later
m_IsVerifyAvailable = False
m_ProductType = OSTypeUnSupported
m_SuiteType = SuiteUnSupported

' Get PSSInfo String
If Len(m_PSSInfo) > 0 Then
    If InStr(m_PSSInfo, Chr$(0)) > 0 Then
        m_PSSInfo = left$(m_PSSInfo, InStr(m_PSSInfo, Chr$(0)) - 1)
    End If
End If

'
' Now, Check for Various versions
'
Select Case OSInfo.dwPlatformId

Case VER_PLATFORM_WIN32_WINDOWS
    ' Check for Various Win9x Versions
        
    m_BuildNumber = LoWord(m_BuildNumber)
    
    If m_Major = 4 And m_Minor = 0 Then
        If m_BuildNumber = 950 Then
            m_OS = Win95
        Else        ' 1111 for OSR 2. For OSR 2.5 = ???
            m_OS = Win95OSR2
        End If
        
    ElseIf m_Major = 4 And m_Minor = 10 Then
        If m_BuildNumber = 1998 Then
            m_OS = Win98
        Else        ' Build Number is 2222.
            m_OS = Win98R2
        End If
        
    ElseIf m_Major >= 4 And m_Minor > 10 Then
        m_OS = WinME      ' Version = 4.90 ?
        
        ' Perhaps, Win ME may support VerifyVersionInfo
        ' m_IsVerifyAvailable=True
    End If


Case VER_PLATFORM_WIN32_NT
    ' Windows NT
    If m_Major = 3 And m_Minor = 0 Then
        m_OS = WinNT3
    ElseIf m_Major = 3 And m_Minor = 1 Then
        m_OS = WinNT31
    ElseIf m_Major = 4 Then
        m_OS = WinNT4
    End If
    
    m_BuildNumber = m_BuildNumber And &HFFFF&
    
    '
    ' Get the SP Version from the PSS String (Not Tested)
    '
    For i = 1 To 6
        If InStr(m_PSSInfo, "Service Pack " & Trim$(Str$(i))) > 0 Then
            m_SPMajor = i
            Exit For
        End If
    Next

Case Else
    ' The case that will never occur on Win32. It's Win 3.x and on win16 our app will not run
    m_OS = Win32s
    
End Select
End If
End Sub

Public Function OSEnum() As OSType
Attribute OSEnum.VB_Description = "The Operating Systems Enum"
    OSEnum = m_OS
End Function

Public Function OSName() As String
Attribute OSName.VB_Description = "Descriptive Name of the OS"
Attribute OSName.VB_UserMemId = 0
    OSName = m_OSString(m_OS)
End Function

Public Function OSMajorVersion() As Long
Attribute OSMajorVersion.VB_Description = "OS Major Version"
    OSMajorVersion = m_Major
End Function

Public Function OSMinorVersion() As Long
Attribute OSMinorVersion.VB_Description = "OS Minor Version"
    OSMinorVersion = m_Minor
End Function

Public Function OSSPMajorVersion() As Long
Attribute OSSPMajorVersion.VB_Description = "Installed Service Pack Major Version (only for Win 2000 and above)"
    OSSPMajorVersion = m_SPMajor
End Function

Public Function OSSPMinorVersion() As Long
Attribute OSSPMinorVersion.VB_Description = "Installed Service Pack Minor Version (only for Win 2000 and above)"
    OSSPMinorVersion = m_SPMinor
End Function

Public Function OSBuildNumber() As Long
Attribute OSBuildNumber.VB_Description = "OS Build Number"
    OSBuildNumber = m_BuildNumber
End Function

Public Function PSSInfo() As String
Attribute PSSInfo.VB_Description = "Product Support Services String"
    PSSInfo = m_PSSInfo
End Function

Public Property Get IsNT() As Boolean
Attribute IsNT.VB_Description = "Returns True if the OS is Windows NT/2000"
    IsNT = (m_OS >= WinNT3)
End Property

Public Property Get IsWin98() As Boolean
Attribute IsWin98.VB_Description = "Returns True if the OS is Windows 98"
    IsWin98 = (m_OS = Win98 Or m_OS = Win98R2)
End Property
Public Property Get IsWin95() As Boolean
    IsWin95 = (m_OS = Win95 Or m_OS = Win95OSR2)
End Property

Public Property Get IsWin2K() As Boolean
Attribute IsWin2K.VB_Description = "Returns True if the OS is Windows 2000 or later"
    IsWin2K = (m_OS = Win2000)
End Property

'Public Property Get IsWinXP() As Boolean
'    IsWin2K = (m_OS = Win2000)
'End Property

Public Property Get SupportsVerify() As Boolean
Attribute SupportsVerify.VB_Description = "Returns True if the OS Supports the VerifyVersionInfo API"
    SupportsVerify = m_IsVerifyAvailable
End Property

Public Property Get OSProductType() As ProductType
    OSProductType = m_ProductType
End Property

Public Property Get OSSuiteType() As SuiteType
    OSSuiteType = m_SuiteType
End Property

Public Function CheckVersion(ByVal CheckType As OSVerComparison, _
ByVal MajorVersion As Long, ByVal MinorVersion As Long, _
Optional BuildNumber As Long = -1, Optional SPMajor As Long = -1, _
Optional SPMinor As Long = -1) As Boolean
Attribute CheckVersion.VB_Description = "Checks if the Version no. of the OS againsta user specified Version conditions"
'
'Checks for a particular version of the OS
'
Dim OSCheckStruct As OSVERSIONINFOEX, OSConditionMask As Currency
Dim CheckMask As Long, APIRetVal As Long

If Not m_IsVerifyAvailable Then
    '
    ' Manual Checking
    ' Note : VB doesn't short circuit expressions, so we take pain
    ' to write a few more lines of code
    '
    Select Case CheckType
    Case Greater
        CheckVersion = (m_Major > MajorVersion)
        
        ' If Major Version is greater, then we need not proceed further.
        ' If not, we check if the versions are equal and then proceed to check Minor Version
                
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_Major = MajorVersion)
        If Not CheckVersion Then Exit Function
        
        If MinorVersion < 0 Then Exit Function
        CheckVersion = (m_Minor > MinorVersion)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_Minor = MinorVersion)
        If Not CheckVersion Then Exit Function
        
        If BuildNumber < 0 Then Exit Function
        CheckVersion = (m_BuildNumber > BuildNumber)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_BuildNumber = BuildNumber)
        If Not CheckVersion Then Exit Function
        
        ' SP Version Can be checked only for Windows NT
        If Not IsNT() Then Exit Function
        
        If SPMajor < 0 Then Exit Function
        CheckVersion = (m_SPMajor > SPMajor)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_SPMajor = SPMajor)

    Case GreaterOrEqual
        CheckVersion = (m_Major >= MajorVersion)
        If Not CheckVersion Then Exit Function
        
        If MinorVersion < 0 Then Exit Function
        CheckVersion = (m_Minor >= MinorVersion)
        If Not CheckVersion Then Exit Function
        
        If BuildNumber < 0 Then Exit Function
        CheckVersion = (m_BuildNumber >= BuildNumber)
        If Not CheckVersion Then Exit Function
        
        If Not IsNT() Then Exit Function
        
        If SPMajor < 0 Then Exit Function
        CheckVersion = (m_SPMajor >= SPMajor)
        
    Case Equal
        CheckVersion = (m_Major = MajorVersion)
        If Not CheckVersion Then Exit Function
        
        If MinorVersion < 0 Then Exit Function
        CheckVersion = (m_Minor = MinorVersion)
        If Not CheckVersion Then Exit Function
        
        If BuildNumber < 0 Then Exit Function
        CheckVersion = (m_BuildNumber = BuildNumber)
        If Not CheckVersion Then Exit Function
        
        ' SP Version Can be checked only for Windows NT
        If Not IsNT() Then Exit Function
        
        If SPMajor < 0 Then Exit Function
        CheckVersion = (m_SPMajor = SPMajor)
                
    Case LessOrEqual
        CheckVersion = (m_Major < MajorVersion)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_Major = MajorVersion)
        If Not CheckVersion Then Exit Function
        
        If MinorVersion < 0 Then Exit Function
        CheckVersion = (m_Minor < MinorVersion)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_Minor = MinorVersion)
        If Not CheckVersion Then Exit Function
        
        If BuildNumber < 0 Then Exit Function
        CheckVersion = (m_BuildNumber < BuildNumber)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_BuildNumber = BuildNumber)
        If Not CheckVersion Then Exit Function
        
        ' SP versions are available only under NT
        If Not IsNT() Then Exit Function
        
        If SPMajor < 0 Then Exit Function
        CheckVersion = (m_SPMajor < SPMajor)
        If CheckVersion Then Exit Function
        
        CheckVersion = (m_SPMajor = SPMajor)
        
    Case Less
        CheckVersion = (m_Major < MajorVersion)
        If CheckVersion Then Exit Function
        
        ' Check if Major versions are equal. If so, Proceed to check Minor Version
        CheckVersion = (m_Major = MajorVersion)
        
        If MinorVersion < 0 Then Exit Function
        CheckVersion = (m_Minor < MinorVersion)
        If CheckVersion Then Exit Function
        
        ' Check if Minor Versions are equal. If so, check Build Number
        CheckVersion = (m_Minor = MinorVersion)
        If Not CheckVersion Then Exit Function
        
        If BuildNumber < 0 Then Exit Function
        CheckVersion = (m_BuildNumber < BuildNumber)
        If CheckVersion Then Exit Function
        
        ' Check if Build Numbers are equal. If so, check Service Pack Version if It is NT
        CheckVersion = (m_BuildNumber = BuildNumber)
        If Not CheckVersion Then Exit Function
                
                
        If Not IsNT() Then Exit Function
        If SPMajor < 0 Then Exit Function
        CheckVersion = (m_SPMajor < SPMajor)
    End Select
Else
    ' OS Supports the VerifyVersionInfo API
    OSCheckStruct.dwOSVersionInfoSize = Len(OSCheckStruct)
    
    
    If MajorVersion >= 0 Then
        OSCheckStruct.dwMajorVersion = MajorVersion
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_MAJORVERSION, CheckType)
        CheckMask = VER_MAJORVERSION
    End If
    
    If MinorVersion >= 0 Then
        OSCheckStruct.dwMinorVersion = MinorVersion
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_MINORVERSION, CheckType)
        CheckMask = CheckMask Or VER_MINORVERSION
    End If
    
    If BuildNumber >= 0 Then
        OSCheckStruct.dwBuildNumber = BuildNumber
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_BUILDNUMBER, CheckType)
        CheckMask = CheckMask Or VER_BUILDNUMBER
    End If
    
    If SPMajor >= 0 Then
        OSCheckStruct.wSPMajor = SPMajor
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_SERVICEPACKMAJOR, CheckType)
        CheckMask = CheckMask Or VER_SERVICEPACKMAJOR
    End If
    
    If SPMinor >= 0 Then
        OSCheckStruct.wSPMinor = SPMinor
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_SERVICEPACKMINOR, CheckType)
        CheckMask = CheckMask Or VER_SERVICEPACKMINOR
    End If
'
'    If chkProductType > 0 Then
'        OSCheckStruct.bProductType = CByte(chkProductType)
'        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_PRODUCT_TYPE, CheckType)
'        CheckMask = CheckMask Or VER_PRODUCT_TYPE
'    End If

    APIRetVal = VerifyVersionInfo(OSCheckStruct, CheckMask, OSConditionMask)
    
    CheckVersion = (APIRetVal <> 0)
End If
End Function

Public Function CheckSuites(chkSuitesInstalled As SuiteType, _
Optional AllSuitesInstalled As Boolean = False) As Boolean

'
' AllSuitesInstalled = True if all the suites indicated in chkSuitesInstalled (by Or-ing)
' must be installed on the System. The default is to check if one of the Suites is
' installed on the System.
'
Dim OSCheckStruct As OSVERSIONINFOEX, OSConditionMask As Currency
Dim CheckMask As Long, APIRetVal As Long

If Not m_IsVerifyAvailable Then
    CheckSuites = False
Else
    ' Suites installation can be verified
    OSCheckStruct.dwOSVersionInfoSize = Len(OSCheckStruct)
        
    If chkSuitesInstalled > 0 Then
        OSCheckStruct.wSuiteMask = chkSuitesInstalled
        OSConditionMask = VerSetConditionMask(OSConditionMask, VER_SUITENAME, _
        IIf(AllSuitesInstalled, VER_AND, VER_OR))
        CheckMask = CheckMask Or VER_SUITENAME
    End If
    
    APIRetVal = VerifyVersionInfo(OSCheckStruct, CheckMask, OSConditionMask)
    
    CheckSuites = (APIRetVal <> 0)
End If
End Function

Private Function LoWord(ByVal lData As Long)
'
' Returns the LoWORD of the Long
'
If (lData And &HFFFF&) > &H7FFF& Then
    lData = lData - &H10000
Else
    lData = lData And &HFFFF&
End If

LoWord = lData
End Function

