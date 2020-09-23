Attribute VB_Name = "modCommandLine"
Option Explicit

Function GetCommandLine(ByVal intNumArgs As Integer, Optional ByVal intMaxArgs As Integer = 0) As Boolean

On Error Resume Next
   
   Dim strChr       As String
   Dim blnInArg     As Boolean
   ' Dim intNumArgs   As Integer
   Dim intCmdLnLen  As Integer
   Dim strCmdLine   As String
   Dim i            As Integer
   
   'See if MaxArgs was provided.
   If intMaxArgs = 0 Then
      intMaxArgs = 10
   End If
   
   'Make array of the correct size.
   'ReDim Preserve arrArgArray(intMaxArgs)
   
   intNumArgs = 0
   blnInArg = False
   
   'Get command line arguments.
   strCmdLine = Command()
   intCmdLnLen = Len(strCmdLine)
   
   For i = 1 To intCmdLnLen
      strChr = Mid(strCmdLine, i, 1)
 
      If (strChr <> " " And strChr <> vbTab) Then
          If Not blnInArg Then
             If intNumArgs = intMaxArgs Then
                Exit For
             End If
            intNumArgs = intNumArgs + 1
            blnInArg = True
         End If

         arrArgArray(intNumArgs) = arrArgArray(intNumArgs) & strChr
      Else
         blnInArg = False
      End If
   Next i
 
   'ReDim Preserve arrArgArray(intNumArgs)
   
   GetCommandLine = True
   
End Function
