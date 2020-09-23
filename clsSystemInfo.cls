VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DRIVE_TYPE
    DRIVE_DOESNT_EXIST = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum

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

Private Const BITSPIXEL = 12
Private Const PLANES = 14

Private Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private pUdtMemStatus As MEMORYSTATUS

Private plMajorVersion  As Long
Private plMinorVersion As Long
Private plPlatformID As Long

Private psComputerName As String
Private plLastDllError As Long
'This Function Written By Sudheer
Public Function IsDosFileCompatible(sFileName As String) As Boolean
Dim sFol() As String, i As Integer, J%
IsDosFileCompatible = True

If Len(sFileName) > 255 Then IsDosFileCompatible = False: Exit Function
sFol = Split(sFileName, "\", , vbTextCompare)

For i = 1 To UBound(sFol)
   If InStr(1, sFol(i), " ", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "[", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "]", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   
   'All Other Conventions Like FileName Should Not Contain \/:*?<>|" are to be checked before as these are not valid neither in FAT32 OR NTFS
   If InStr(1, sFol(i), "\", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "/", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "|", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "?", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "*", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), "<", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), ">", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), ":", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   If InStr(1, sFol(i), """", vbTextCompare) > 1 Then IsDosFileCompatible = False: Exit Function
   
   For J = 0 To 31 'These Are Not Allowed In Dos
    If InStr(1, sFol(i), Chr$(J)) > 1 Then IsDosFileCompatible = False: Exit Function
   Next
    
   For J = 127 To 255 'These Are Not Allowed In Dos
    If InStr(1, sFol(i), Chr$(J)) > 1 Then IsDosFileCompatible = False: Exit Function
   Next
      
Next
If Len(sFol(UBound(sFol()))) > 12 Then IsDosFileCompatible = False: Exit Function 'we check len To Be 12 As filenamelen=8,fileext len=3,"." len is 1 so 8+3+1=12

End Function
'This Function Written By Sudheer
Public Function GetShortPath(sFileName As String) As String
    Dim lRes As Long, sPath As String
    sPath = String$(255, 0)   'Create a buffer
    lRes = GetShortPathName(sFileName, sPath, 254)
    GetShortPath = Left$(sPath, lRes) 'remove all unnecessary vbnull`s
End Function

Private Function FreeBytesOnDisk(Drive As String) As Long
    On Error Resume Next
    plLastDllError = 0
    
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double
    
    
    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    lFreeBytes = lBytesPerCluster * lFreeClusters
    FreeBytesOnDisk = lFreeBytes
    plLastDllError = Err.LastDllError
End Function

Public Function WindowsDir() As String
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        WindowsDir = Temp & "\"
    Else
        WindowsDir = Temp
    End If
End Function

Public Function SystemDir() As String
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetSystemDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        SystemDir = Temp & "\"
    Else
        SystemDir = Temp
    End If
End Function

Private Function TotalBytesOnDisk(Drive As String) As Double
    On Error Resume Next
    plLastDllError = 0
    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lTotalBytes As Double

    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    'dblAns = (Bytes / 1024) / 1024
    TotalBytesOnDisk = lBytesPerCluster * lTotalClusters
    If TotalBytesOnDisk = 0 Then
        TotalBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lTotalClusters
    End If
    plLastDllError = Err.LastDllError
End Function

Public Function FreeMegaBytesOnDisk(Optional Drive As String = "C:\") As Double
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long

    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lFreeBytes As Double

    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"

    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector

    FreeMegaBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lFreeClusters
    FreeMegaBytesOnDisk = Format(FreeMegaBytesOnDisk, "###,###,##0.00")
End Function

Public Function TotalMegaBytesOnDisk(Optional Drive As String = "C:\") As Double
    'some time in the future disk may be to large to calculate
    'like this so resume next on any errors
    On Error Resume Next

    Dim lAns As Long
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    Dim lBytesPerCluster As Long
    Dim lTotalBytes As Double
    
    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    
    lAns = GetDiskFreeSpace(Drive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
    lBytesPerCluster = lSectorsPerCluster * lBytesPerSector
    
    TotalMegaBytesOnDisk = ((lBytesPerCluster / 1024) / 1024) * lTotalClusters
    TotalMegaBytesOnDisk = Format(TotalMegaBytesOnDisk, "###,###,##0.00")
End Function

Public Property Get HasSoundCard() As Boolean
    plLastDllError = 0
    HasSoundCard = (waveOutGetNumDevs > 0)
    plLastDllError = Err.LastDllError
End Property

Public Property Get ComputerName() As String
    Dim sBuffer As String
    Dim lAns As Long

    plLastDllError = 0
    sBuffer = Space$(255)
    lAns = GetComputerName(sBuffer, 255)
    If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        plLastDllError = Err.LastDllError
    End If
End Property

Public Property Get CurrentUser() As String
    Dim l As Long
    Dim sUser As String

    plLastDllError = 0
    sUser = Space(255)
    l = GetUserName(sUser, 255)
    'strip null terminator
    If l <> 0 Then
        CurrentUser = Left(sUser, InStr(sUser, Chr(0)) - 1)
    Else
        plLastDllError = Err.LastDllError
    End If
End Property

Public Function VolumeName(Optional Drive As String = "C:\")
    Dim sBuffer As String

    plLastDllError = 0
    sBuffer = Space$(255)
    'fix bad parameter values
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    If GetVolumeInformation(Drive, sBuffer, Len(sBuffer), 0, 0, 0, Space$(255), 255) = 0 Then
        plLastDllError = Err.LastDllError
    Else
        VolumeName = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    End If
End Function


Public Function DriveType(Drive As String) As DRIVE_TYPE
    'fix bad parameter values
    plLastDllError = 0
    If Len(Drive) = 1 Then Drive = Drive & ":\"
    If Len(Drive) = 2 And Right$(Drive, 1) = ":" Then Drive = Drive & "\"
    DriveType = GetDriveType(Drive)
    plLastDllError = Err.LastDllError
End Function

Public Function AvailablePhysicalMemory() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    AvailablePhysicalMemory = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function TotalPhysicalMemory() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    TotalPhysicalMemory = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function PercentMemoryFree() As Double
    PercentMemoryFree = Format(AvailableMemory / TotalMemory * 100, "0#")
End Function

Public Function AvailablePageFile() As Double
    'Return Value in Megabytes
    Dim dblAns As Double

    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPageFile
    AvailablePageFile = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function PageFileSize() As Double
    'Return Value in Megabytes
    Dim dblAns As Double

    plLastDllError = 0
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    PageFileSize = BytesToMegabytes(dblAns)
    plLastDllError = Err.LastDllError
End Function

Public Function AvailableMemory() As Double
    'Return Value in Megabytes
     AvailableMemory = AvailablePhysicalMemory + AvailablePageFile
End Function

Public Function TotalMemory() As Double
    'Return Value in Megabytes
    TotalMemory = PageFileSize + TotalPhysicalMemory
End Function

Private Function BytesToMegabytes(Bytes As Double) As Double
  Dim dblAns As Double
  dblAns = (Bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
End Function

Public Property Get MaxScreenColors(lwndHandle As Long) As Double
    'Returns the maximum number of colors supported
    'by the system - e.g.,  256, 16,777,216
    Dim lngBits As Long
    Dim lngPlanes As Long
    
    Dim dblAns As Double

    plLastDllError = 0
    'bits per pixel
    lngBits = GetDeviceCaps(lwndHandle, BITSPIXEL)
    'number of color planes
    lngPlanes = GetDeviceCaps(lwndHandle, PLANES)
    'maximum colors available
    MaxScreenColors = (2 ^ (lngBits * lngPlanes))
    plLastDllError = Err.LastDllError
End Property

Public Property Get ScreenPixelWidth() As Integer
    plLastDllError = 0
    ScreenPixelWidth = Screen.Width \ Screen.TwipsPerPixelX
End Property

Public Property Get ScreenPixelHeight() As Integer
    plLastDllError = 0
    ScreenPixelHeight = Screen.Height \ Screen.TwipsPerPixelY
End Property

Public Property Get ScreenResolution() As String
    plLastDllError = 0
    ScreenResolution = ScreenPixelWidth & " X " & ScreenPixelHeight
End Property

Public Property Get SystemErrorCode() As Long
    SystemErrorCode = plLastDllError
End Property
