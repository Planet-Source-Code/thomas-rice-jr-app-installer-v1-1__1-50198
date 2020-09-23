Attribute VB_Name = "modBrowse"
Option Explicit

Public Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateDirectoryEx Lib "kernel32.dll" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDListA Lib "shell32.dll" (pidl As Any, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As ITEMIDLIST) As Long

Public Declare Function SHGetPathFromIDList Lib _
      "Shell32" (ByVal pidList As Long, ByVal lpBuffer _
      As String) As Long


Public Declare Function lstrcat Lib "kernel32" _
      Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
      lpString2 As String) As Long

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Enum BROWSE_FLAGS
  BIF_BROWSEFORCOMPUTER = &H1000
  BIF_BROWSEFORPRINTER = &H2000
  BIF_BROWSEINCLUDEFILES = &H4000
  BIF_DONTGOBELOWDOMAIN = &H2
  BIF_EDITBOX = &H10
  BIF_RETURNFSANCESTORS = &H8
  BIF_RETURNONLYFSDIRS = &H1
  BIF_STATUSTEXT = &H4
  BIF_USENEWUI = &H40
  BIF_VALIDATE = &H20
End Enum

Public Const MAX_PATH = 260
Public Const MAX_NAME = 40

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum efbrCSIDLConstants
    CSIDL_DESKTOP = &H0                   '(desktop)
    CSIDL_INTERNET = &H1                  'Internet Explorer (icon on desktop)
    CSIDL_PROGRAMS = &H2                  'Start Menu\Programs
    CSIDL_CONTROLS = &H3                  'My Computer\Control Panel
    CSIDL_PRINTERS = &H4                  'My Computer\Printers
    CSIDL_PERSONAL = &H5                  'My Documents
    CSIDL_FAVORITES = &H6                 '(user name)\Favorites
    CSIDL_STARTUP = &H7                   'Start Menu\Programs\Startup
    CSIDL_RECENT = &H8                    '(user name)\Recent
    CSIDL_SENDTO = &H9                    '(user name)\SendTo
    CSIDL_BITBUCKET = &HA                 '(desktop)\Recycle Bin
    CSIDL_STARTMENU = &HB                 '(user name)\Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10         '(user name)\Desktop
    CSIDL_DRIVES = &H11                   'My Computer
    CSIDL_NETWORK = &H12                  'Network Neighborhood
    CSIDL_NETHOOD = &H13                  '(user name)\nethood
    CSIDL_FONTS = &H14                    'windows\fonts
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16         'All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17          'All Users\Programs
    CSIDL_COMMON_STARTUP = &H18           'All Users\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19  'All Users\Desktop
    CSIDL_APPDATA = &H1A                  '(user name)\Application Data
    CSIDL_PRINTHOOD = &H1B                '(user name)\PrintHood
    CSIDL_LOCAL_APPDATA = &H1C            '(user name)\Local Settings\Applicaiton Data (non roaming)
    CSIDL_ALTSTARTUP = &H1D               'non localized startup
    CSIDL_COMMON_ALTSTARTUP = &H1E        'non localized common startup
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23           'All Users\Application Data
    CSIDL_WINDOWS = &H24                  'GetWindowsDirectory()
    CSIDL_SYSTEM = &H25                   'GetSystemDirectory()
    CSIDL_PROGRAM_FILES = &H26            'C:\Program Files
    CSIDL_MYPICTURES = &H27               'C:\Program Files\My Pictures
    CSIDL_PROFILE = &H28                  'USERPROFILE
    CSIDL_PROGRAM_FILES_COMMON = &H2B     'C:\Program Files\Common
    CSIDL_COMMON_TEMPLATES = &H2D         'All Users\Templates
    CSIDL_COMMON_DOCUMENTS = &H2E         'All Users\Documents
    CSIDL_COMMON_ADMINTOOLS = &H2F        'All Users\Start Menu\Programs\Administrative Tools
    CSIDL_ADMINTOOLS = &H30               '(user name)\Start Menu\Programs\Administrative Tools

    CSIDL_FLAG_CREATE = &H8000            'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
    CSIDL_FLAG_DONT_VERIFY = &H4000       'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
    CSIDL_FLAG_MASK = &HFF00              'mask for all possible flag values
End Enum


