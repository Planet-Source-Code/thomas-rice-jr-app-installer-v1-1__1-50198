VERSION 5.00
Begin VB.Form frmSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eLinkCart -  Select SQL Install Server"
   ClientHeight    =   4560
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
   Icon            =   "frmSQL.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSQL 
      Caption         =   "Install Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   75
      TabIndex        =   7
      Top             =   1875
      Width           =   5940
      Begin VB.TextBox txtServer 
         Height          =   315
         Left            =   1365
         TabIndex        =   15
         Text            =   "Server"
         Top             =   255
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1035
         Width           =   3045
      End
      Begin VB.ComboBox cboServer 
         Height          =   315
         Left            =   1365
         TabIndex        =   11
         Text            =   "SQL Server"
         Top             =   255
         Width           =   3060
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1365
         TabIndex        =   8
         Text            =   "sa"
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   225
         Left            =   405
         TabIndex        =   13
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name:"
         Height          =   225
         Left            =   405
         TabIndex        =   12
         Top             =   690
         Width           =   930
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         Caption         =   "SQL Server:"
         Height          =   225
         Left            =   405
         TabIndex        =   9
         Top             =   315
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Next"
      Height          =   375
      Left            =   3825
      TabIndex        =   6
      Top             =   3675
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4935
      TabIndex        =   5
      Top             =   3675
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
         Picture         =   "frmSQL.frx":0ABA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblBGHeadInfo1 
         BackStyle       =   0  'Transparent
         Caption         =   "eLinkCart - Select SQL Server"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   330
         Width           =   3540
      End
   End
   Begin VB.Label lblBGHeadInfo2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please register your copy of eLinkCart"
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
      Left            =   45
      TabIndex        =   10
      Top             =   3795
      Width           =   3030
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
      Left            =   30
      MouseIcon       =   "frmSQL.frx":1574
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4110
      Width           =   2460
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSQL.frx":16C6
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6090
      Y1              =   975
      Y2              =   975
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is the Registration Module for eLinkBB
' Copyright (c) 2002 by Rice WebDesigns, Inc. All right Reserved
'

'Variables must be declared
Option Explicit

Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1

Private Sub cmdCancel_Click()

    gblnCancel = True
    Me.Hide
    Unload Me
    
End Sub

Private Sub cmdContinue_Click()

On Error GoTo Cmd_Error

   Dim blnRtn           As Boolean
   Dim strConnection    As String
   Dim intRtn           As Integer
   
   cmdCancel.Enabled = False
   cmdContinue.Enabled = False
   
   'Server=RWDServer;Database=master;User=sa;Password=;
   strConnection = "Database=master;User=sa;Password=" & txtPassword.Text & ";"
   If cboServer.Enabled Then
      strConnection = "Server=" & cboServer.Text & ";" & strConnection
   Else
      strConnection = "Server=" & txtServer.Text & ";" & strConnection
   End If
   
   blnRtn = TestConnection(strConnection)
   If Not blnRtn Then
      intRtn = MsgBox("Server Connection Failed. Please check your entries and try again", vbApplicationModal + vbCritical + vbOKOnly)
      cmdCancel.Enabled = True
      cmdContinue.Enabled = True
      cmdCancel_Click
      Exit Sub
   End If
   
   If cboServer.Enabled Then
      gstrSQLServer = cboServer.Text
   Else
      gstrSQLServer = txtServer.Text
   End If
   
   gstrSQLUser = "sa"
   gstrSQLPassword = txtPassword.Text
   
   DoEvents
   gblnError = False
   
Cmd_Exit:

   Unload Me
   Exit Sub
   
Cmd_Error:
   
   'gblnError = True
   Call cmdCancel_Click
   Exit Sub
   
End Sub

Private Sub Form_Load()

On Error GoTo Load_error

   Dim strServers()        As String
   Dim intLoop             As Integer
   Dim intSelected         As Integer
   
   mblnSQLList = False
   
'   If gblnSQLInstalled Then
   txtServer.Visible = False
   strServers = GetServerNames
   DoEvents
   If Not mblnSQLList Then
      cboServer.Visible = False
      cboServer.Enabled = False
      txtServer.Visible = True
      txtServer.Text = gstrServerName
   Else
   
      cboServer.Clear
      For intLoop = 0 To UBound(strServers)
         If strServers(intLoop) = "(local)" Then
            cboServer.AddItem (gstrServerName)
            intSelected = intLoop
         Else
            cboServer.AddItem (strServers(intLoop))
         End If
         
      Next
      
      cboServer.ListIndex = intSelected
   End If
'   Else
'      cboServer.Visible = False
'      cboServer.Enabled = False
'      txtServer.Visible = True
'      txtServer.Text = gstrServerName
'   End If
   
   'Me.Show
   Screen.MousePointer = vbNormal
   
Load_Exit:
   
   Exit Sub
   
Load_error:
      
   cboServer.Visible = False
   cboServer.Enabled = False
   txtServer.Visible = True
   txtServer.Text = gstrServerName
   GoTo Load_Exit
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

   Set IE = Nothing
   DoEvents

End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Set IE = Nothing
End Sub

Private Sub lblHomepage_Click()

On Error Resume Next

    'goto RWD Website using Internet Explorer object and make it visible
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.Navigate "http://www.ricewebdesigns.com"

End Sub

Private Function TestConnection(strConnection As String, Optional intCn As Integer = 0, Optional strDatabaseType As String = "MSSQL") As Boolean

    Dim cnADO As clsAdoConnection
    Dim cn As New ADODB.Connection
    Dim lngRtn As Long
    Dim strTemp As String
    Dim strTag As String
    Dim strValue As String
    Dim strUser As String
    Dim strPassword As String
    Dim strServer As String
    Dim strDataBase As String
    Dim intRtn As Integer
    Dim strNewConnection As String
    Dim intPos As Integer

   TestConnection = False
   
    'Parse the incoming connection string for req's
    While Len(strConnection) > 0
        intPos = InStr(1, strConnection, ";")
        If intPos = 0 Then
            GoTo Continue_Loop
        End If

        strTemp = Mid(strConnection, 1, intPos - 1)
        strConnection = Mid(strConnection, intPos + 1)

        intPos = InStr(1, strTemp, "=")
        If intPos = 0 Then
            GoTo Continue_Loop
        End If
        strTemp = strTemp & Space(1)

        strTag = Mid(strTemp, 1, intPos - 1)
        strValue = Mid(strTemp, intPos + 1)

        Select Case UCase(strTag)
            Case "SERVER"
                strServer = Trim(strValue)
            Case "DATABASE"
                strDataBase = Trim(strValue)
            Case "USER"
                strUser = Trim(strValue)
            Case "PASSWORD"
                strPassword = Trim(strValue)
            Case Else
                intRtn = MsgBox("Invalid Database Tag in DBConvert.ini. Please Correct and restart the Conversion", vbCritical + vbApplicationModal + vbOKOnly)
                GoTo testConnection_Error
        End Select
Continue_Loop:
    Wend

    If Len(strDataBase) = 0 Then
        intRtn = MsgBox("No Database listed in DBConvert.ini. Please Correct and restart the Conversion.", vbCritical + vbApplicationModal + vbOKOnly)
        GoTo testConnection_Error
    End If

    Set cnADO = New clsAdoConnection

    Select Case UCase(strDatabaseType)
        Case "ACCESS" 'Open an Access 97 database
            cnADO.ProviderConst = pdsajet40
            cnADO.DataSource = strDataBase
            cnADO.Password = strPassword

        Case "MSSQL" 'Open an SQL Server database

            cnADO.ProviderConst = pdsasqlserver
            cnADO.DataSource = strServer
            cnADO.InitialCatalog = strDataBase
            cnADO.UserID = strUser
            cnADO.Password = strPassword
            'Conn.UseNTSecurity = True

        Case "DBASE" 'Open an Dbase III database directory
            cnADO.ProviderConst = pdsadbase
            cnADO.DataSource = strDataBase

        Case "ORACLE"
            cnADO.ProviderConst = pdsaoracle
            cnADO.DataSource = strServer
            cnADO.InitialCatalog = strDataBase
            cnADO.UserID = strUser
            cnADO.Password = strPassword
            'Conn.UseNTSecurity = True
        Case Else
            intRtn = MsgBox("This Application only supports Access, MS SQL, DBase and Oracle. Please correct DBConvert.ini and restart the Conversion.", vbCritical + vbApplicationModal + vbOKOnly)
            GoTo testConnection_Error
    End Select

    strNewConnection = cnADO.ConnectionString

    Select Case intCn
        Case 0  ' eLinkBB
            cn.ConnectionString = strNewConnection
            cn.CursorLocation = adUseClient
            cn.CommandTimeout = 5
            cn.Open
            
        Case Else
            intRtn = MsgBox("No Database listed in .ini file. Please Correct and restart the Conversion.", vbCritical + vbApplicationModal + vbOKOnly)
    End Select

    TestConnection = True
    strConnection = strNewConnection

testConnection_Exit:

On Error Resume Next
    
    cn.Close
    Set cn = Nothing
    Set cnADO = Nothing
    
    Exit Function

testConnection_Error:

    TestConnection = False
    GoTo testConnection_Exit
    
End Function

