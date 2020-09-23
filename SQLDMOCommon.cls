VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDMOCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' SQL Server Database  Class

Dim m_SQLServer As SQLDMO.SQLServer
Private m_DataTypes() As String

Type FindTableResult
  Database As String
  Table As String
End Type

Type FindFieldResult
  Database As String
  Table As String
  Field As String
  DataType As String
End Type

Public Enum SQLLoginSecure
  SQL_LoginSecure = False
  NT_LoginSecure = True
End Enum

' PROPERTIES =============

Public Property Get DataTypes() As String()
  DataTypes = m_DataTypes
End Property


Public Property Get SQLServer() As SQLDMO.SQLServer
  Set SQLServer = m_SQLServer
End Property

' PRIVATES================

Private Sub Class_Initialize()
  Set m_SQLServer = New SQLDMO.SQLServer
  ReDim m_DataTypes(22)
  m_DataTypes(0) = "binary"
  m_DataTypes(1) = "bit"
  m_DataTypes(2) = "char"
  m_DataTypes(3) = "datetime"
  m_DataTypes(4) = "decimal"
  m_DataTypes(5) = "float"
  m_DataTypes(6) = "image"
  m_DataTypes(7) = "int"
  m_DataTypes(8) = "money"
  m_DataTypes(9) = "nchar"
  m_DataTypes(10) = "ntext"
  m_DataTypes(11) = "numeric"
  m_DataTypes(12) = "nvarchar"
  m_DataTypes(13) = "real"
  m_DataTypes(14) = "smalldatetime"
  m_DataTypes(15) = "smallint"
  m_DataTypes(16) = "smallmoney"
  m_DataTypes(17) = "text"
  m_DataTypes(18) = "timestamp"
  m_DataTypes(19) = "tinyint"
  m_DataTypes(20) = "uniqueidentifier"
  m_DataTypes(21) = "varbinary"
  m_DataTypes(22) = "varchar"
End Sub

Private Sub Class_Terminate()
  Set m_SQLServer = Nothing
End Sub

' METHODS================

Public Function GetServerNames() As String()
' This function finds servers declared by "Client Network Utility" or by an OBDC declaration.
' Note: an error in call procedure may occurs if no server found: GetServerNames is empty
On Error GoTo EH
Dim Names() As String
Dim i As Long
Dim myNameList As SQLDMO.NameList
Set myNameList = m_SQLServer.Application.ListAvailableSQLServers

For i = 1 To myNameList.Count
  ReDim Preserve Names(i - 1)
  Names(i - 1) = myNameList(i)
Next

GetServerNames = Names

EH:

End Function

Function Connect(ServerName As String, LoginSecure As SQLLoginSecure, UserName As String, Password As String) As Integer
'RETURN:
'0 - No Error
'1 - Undefined Error , Raise Error
'2 - Server not found or Network Error
'3 - Login failed on defined User and Password
'-----------------------------------------
  On Error GoTo ErrorTrap
  Connect = 0
  m_SQLServer.LoginSecure = LoginSecure
  m_SQLServer.Connect ServerName, UserName, Password
  
Exit Function
ErrorTrap:
  Connect = 1
  Select Case Err.Number
    Case &H80040000
      Connect = 2 ' Server not found or Network Error
    Case &H80044818
      Connect = 3 ' Login failed on defined User and Password  (Password was wrong)
    Case &H80045021 ' Already connected
      m_SQLServer.DisConnect
      Resume
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function

Function Disconect()
  On Error Resume Next
  m_SQLServer.DisConnect
  Disconect = 0
End Function

Public Function CreateDatabase(DatabaseName As String, Directory As String, _
                         Optional DatabaseFileName As String) As Integer
'INPUT:
'DatabaseName : Name of the Database to create
'Directory : Location of new Database
'DatabaseFileName : Filename only
'DatabaseLogFileName : FileName Only
'------------------------------------
'RETURN:
'0 - no error
'1 - Undefined Error,  Raise Error
'2 - cannot create database because it's existing.
'3 - Phisical data file is existing
'4 - File can not created
'--------------------------------------
  Dim strDBFile As String
  Dim strDBFileName As String
  Dim strDBLogFile As String
  Dim strDBDirectory As String
   
  Dim oDatabase As New SQLDMO.Database
  Dim oDBFileData As New SQLDMO.DBFile
  Dim oLogFile As New SQLDMO.LogFile
  
  On Error GoTo ErrorTrap
  CreateDatabase = 0
  strDBDirectory = Directory
  If Right(strDBDirectory, 1) <> "\" Then strDBDirectory = strDBDirectory & "\"
  
'Getting Filenames of Database :
  
  If DatabaseFileName = "" Then
    strDBFile = DatabaseName & "_Data.mdf"
    strDBLogFile = DatabaseName & "_Log.ldf"
  Else
    'Make sure strDBFileName is not includes "\" (Filename Only):
    strDBFileName = Trim(DatabaseFileName)
    strDBFileName = Split(strDBFileName, "\")(UBound(Split(strDBFileName, "\")))
    strDBFile = strDBFileName
    If Right(strDBFile, 4) <> ".mdf" Then
      strDBFile = strDBFile & "_Data.mdf"
      strDBLogFile = left(strDBFile, Len(strDBFile) - Len("_Data.mdf")) & "_Log.ldf"
    Else
      strDBLogFile = left(strDBFile, Len(strDBFile) - Len(".mdf")) & "_Log.ldf"
    End If
  End If
                 
    ' Define the PRIMARY data file.
    oDBFileData.name = DatabaseName
    oDBFileData.PhysicalName = strDBDirectory & strDBFile
    oDBFileData.PrimaryFile = True
    
    'Set up the parameters for the transaction log:
    ' Define the database transaction log.
    oLogFile.name = DatabaseName & "_Log"
    oLogFile.PhysicalName = strDBDirectory & strDBLogFile
    oDatabase.TransactionLog.LogFiles.Add oLogFile

    'Set the database growth options:
    'Specify file growth in chunks of fixed size for all data files.
    oDBFileData.FileGrowthType = SQLDMOGrowth_MB
    oDBFileData.FileGrowth = 10
    oDatabase.FileGroups("PRIMARY").DBFiles.Add oDBFileData
    oDatabase.name = DatabaseName

    'Create the database as defined.
    m_SQLServer.Databases.Add oDatabase
    
Exit Function
ErrorTrap:
  CreateDatabase = 1
  Select Case Err.Number
    Case &H80040709
      CreateDatabase = 2 ' Database name is existing
    Case &H80041432
      CreateDatabase = 3 ' Phisical data file is existing
    Case &H800413F1
      CreateDatabase = 4 ' File can not created
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function

Public Function DeleteDatabase(DatabaseName As String) As Integer
'RETURN:
'0 - no error
'1 - Undefined Error,  Raise Error
'2 - Database not found
'3 - Database in use
'--------------------------------------
On Error GoTo ErrorTrap
DeleteDatabase = 0
m_SQLServer.Databases.Remove (DatabaseName)
    
Exit Function
ErrorTrap:
  DeleteDatabase = 1
  Select Case Err.Number
    Case &H80045510
      DeleteDatabase = 2 ' Database not found
    Case &H80040E76
      DeleteDatabase = 3 ' Database in use
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function

Public Function DatabaseExists(DatabaseName As String) As Boolean
'RETURN:
'False - No Existing
'True - Existing
'--------------------------------------
Dim db As SQLDMO.Database

On Error GoTo ErrorTrap
DatabaseExists = False
For Each db In m_SQLServer.Databases
  If UCase(db.name) = UCase(DatabaseName) Then DatabaseExists = True: Exit For
Next

Exit Function
ErrorTrap:
  Select Case Err.Number
    'Case ...
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function


Public Function TableExists(DatabaseName As String, TableName As String) As Boolean
'RETURN:
'False - No Existing
'True - Existing
'--------------------------------------
Dim tbl As SQLDMO.Table

On Error GoTo ErrorTrap
TableExists = False
For Each tbl In m_SQLServer.Databases(DatabaseName).Tables
  If UCase(tbl.name) = UCase(TableName) Then TableExists = True: Exit For
Next

Exit Function
ErrorTrap:
  Select Case Err.Number
    'Case ...
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function


Public Sub ExecuteImmediate(DatabaseName As String, Command As String, Optional ExecuteType As SQLDMO.SQLDMO_EXEC_TYPE = SQLDMOExec_Default, Optional Length As Long, Optional ByRef ErrorMessage As String)
On Error GoTo ErrorTrap
m_SQLServer.Databases(DatabaseName).ExecuteImmediate Command, SQLDMOExec_ContinueOnError
ErrorMessage = ""
Exit Sub
ErrorTrap:
  Select Case Err.Number
    Case &H80045703
      If Err.Description = "[SQL-DMO]" Then ErrorMessage = "": Exit Sub ' No Error
      ErrorMessage = "Error: " & Err.Description
    Case Else
      ErrorMessage = "Error: " & Err.Description
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Sub

    
Public Function ExecuteWithResults(DatabaseName As String, Command As String, Optional Length As Long) As SQLDMO.QueryResults
Dim QResults As SQLDMO.QueryResults
On Error GoTo ErrorTrap
Set QResults = m_SQLServer.Databases(DatabaseName).ExecuteWithResults(Command, Length)
Set ExecuteWithResults = QResults

Exit Function
ErrorTrap:
  Select Case Err.Number
    'Case ...
    Case Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
  End Select
End Function

    
Public Function ExecuteWithResultsAndMessages(DatabaseName As String, Command As String, Message As String, Optional Length As Long) As SQLDMO.QueryResults
Dim QResults As SQLDMO.QueryResults
On Error Resume Next  '
Set QResults = m_SQLServer.Databases(DatabaseName).ExecuteWithResultsAndMessages(Command, Length, Message)
Set ExecuteWithResultsAndMessages = QResults

'Exit Function
'ErrorTrap:
'  Select Case Err.Number
'    'Case ...
'    Case Else
'      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'  End Select
End Function

Public Function FindDatabase( _
  DatabaseName As String, _
  Optional SystemDatabase As Boolean = False, _
  Optional Exact As Boolean = False, _
  Optional Refresh As Boolean) _
  As String()
'If Database="" then find database in all the server

Dim Databases() As String ' Found Tables array
Dim Count As Long
Dim Database As SQLDMO.Database
Dim i As Long
Count = -1
If Refresh Then m_SQLServer.Databases.Refresh
For i = 1 To m_SQLServer.Databases.Count
  Set Database = m_SQLServer.Databases(i)
  If IIf(Exact, StrComp(Database.name, DatabaseName, vbTextCompare) = 0, (InStr(1, Database.name, DatabaseName, vbTextCompare) <> 0)) And _
    (Database.SystemObject = SystemDatabase) Then
    Count = Count + 1
    ReDim Preserve Databases(Count)
    Databases(Count) = Database.name
  End If
Next
FindDatabase = Databases
End Function

Public Function FindTable( _
  TableName As String, _
  Optional DatabaseName As String = "", _
  Optional SystemDatabase As Boolean = False, _
  Optional SystemTable As Boolean = False, _
  Optional ExactDatabase As Boolean = False, _
  Optional ExactTable As Boolean = False, _
  Optional Refresh As Boolean) _
  As FindTableResult()
'If Table="" then find Table in all databases
'If Database="" then find database in the server

Dim Result() As FindTableResult 'Found Tables array
Dim Databases() As String ' Found Database array
Dim Count As Long
Dim DatabaseCount As Long
Dim Database As SQLDMO.Database
Dim Tables As SQLDMO.Tables
Dim Table As SQLDMO.Table
Dim i As Long, j As Long

Databases = FindDatabase(DatabaseName, SystemDatabase, ExactDatabase, Refresh)
On Error Resume Next  ' prevents error when count of array =0
DatabaseCount = UBound(Databases) - LBound(Databases) + 1
On Error GoTo 0
If DatabaseCount = 0 Then GoTo EXIT_

Count = -1
For i = 0 To DatabaseCount - 1
  Set Tables = m_SQLServer.Databases(Databases(i)).Tables
  If Refresh Then Tables.Refresh
  For j = 1 To Tables.Count
    Set Table = Tables(j)
    If IIf(ExactTable, StrComp(Table.name, TableName, vbTextCompare) = 0, (InStr(1, Table.name, TableName, vbTextCompare) <> 0)) And _
      (Table.SystemObject = SystemTable) Then
      Count = Count + 1
      ReDim Preserve Result(Count)
      Result(Count).Database = Databases(i)
      Result(Count).Table = Table.name
    End If
  Next j
Next i

EXIT_:
  FindTable = Result
End Function


Public Function FindField( _
  FieldName As String, _
  Optional DataType As String = "", _
  Optional TableName As String = "", _
  Optional DatabaseName As String = "", _
  Optional SystemDatabase As Boolean = False, _
  Optional SystemTable As Boolean = False, _
  Optional ExactDatabase As Boolean = False, _
  Optional ExactTable As Boolean = False, _
  Optional ExactField As Boolean = False, _
  Optional Refresh As Boolean) _
  As FindFieldResult()
'If Field="" then find Field in all tables
'If Table="" then find Table in all databases
'If Database="" then find database in the server

Dim Result() As FindFieldResult 'Found Tables array
Dim FoundTable() As FindTableResult 'Found Tables array
Dim Count As Long
Dim TableCount As Long
'Dim Fields As SQLDMO.Columns
Dim Field As SQLDMO.Column
Dim i As Long, j As Long

FoundTable = FindTable(TableName, DatabaseName, SystemDatabase, SystemTable, ExactDatabase, ExactTable, Refresh)
On Error Resume Next  ' prevents error when count of array =0
TableCount = UBound(FoundTable) - LBound(FoundTable) + 1
On Error GoTo 0
If TableCount = 0 Then GoTo EXIT_

Count = -1
For i = 0 To TableCount - 1
'  Set Fields = m_SQLServer.Databases(FoundTable(i).Database).Tables(FoundTable(i).Table).Columns
  If Refresh Then m_SQLServer.Databases(FoundTable(i).Database).Tables(FoundTable(i).Table).Columns.Refresh 'Fields.Refresh
  For j = 1 To m_SQLServer.Databases(FoundTable(i).Database).Tables(FoundTable(i).Table).Columns.Count ' Fields.Count
    Set Field = m_SQLServer.Databases(FoundTable(i).Database).Tables(FoundTable(i).Table).Columns(j) ' Fields(j)
    If IIf(ExactField, _
            StrComp(Field.name, FieldName, vbTextCompare) = 0, _
            (InStr(1, Field.name, FieldName, vbTextCompare) <> 0)) _
            And _
            IIf(DataType = "", True, StrComp(Field.DataType, DataType, vbTextCompare) = 0) _
            Then
      Count = Count + 1
      ReDim Preserve Result(Count)
      Result(Count).Database = FoundTable(i).Database
      Result(Count).Table = FoundTable(i).Table
      Result(Count).Field = Field.name
      Result(Count).DataType = Field.DataType
    End If
  Next j
Next i

EXIT_:
  FindField = Result
End Function


