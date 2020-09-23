Attribute VB_Name = "modIni"
Option Explicit
Option Compare Text

Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
    As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize _
    As Long, ByVal lpFileName As String) As Long

Function IniGetString(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String) As String
On Error Resume Next
    
    IniGetString = IniPrivateGetString(strSection, strKey, strDefault, gstrINIFile)

End Function

Function IniPrivateGetString(ByVal strSection As String, ByVal strKey As String, _
    ByVal strDefault As String, ByVal strINIFile As String) As String
On Error Resume Next
    
    Const INI_BUF_SIZE = 255
    Dim intcch As Integer
    Dim strBuf As String
    
    strBuf = String$(INI_BUF_SIZE + 1, 0)
   
    intcch = GetPrivateProfileString(strSection, strKey, strDefault, strBuf, _
        INI_BUF_SIZE, strINIFile)
    IniPrivateGetString = C2ABStr(strBuf)

End Function

Public Function C2ABStr(ByVal strCString As String) As String
On Error Resume Next
Dim intPos As Integer
    
    intPos = InStr(strCString, Chr$(0))
    If intPos Then
        C2ABStr = left$(strCString, intPos - 1)
    Else
        C2ABStr = strCString
    End If
End Function

Public Function RemoveFromString(ByVal strStr As String, ByVal strChar As String) As String
On Error Resume Next
Dim intStrLen As Integer
Dim intPos As Integer
    
    intStrLen = Len(strStr)
    For intPos = 1 To intStrLen
        If Mid(strStr, intPos, 1) = strChar Then
            strStr = left(strStr, intPos - 1) & Right(strStr, intStrLen - intPos)
            intStrLen = intStrLen - 1
            intPos = intPos - 1
        End If
    Next intPos
    RemoveFromString = strStr
    
End Function

