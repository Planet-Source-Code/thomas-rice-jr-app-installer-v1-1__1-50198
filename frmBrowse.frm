VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "$(PackageName) -  Select Install Folder"
   ClientHeight    =   3795
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
   Icon            =   "frmBrowse.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInstall 
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
      Height          =   705
      Left            =   120
      TabIndex        =   7
      Top             =   1950
      Width           =   5940
      Begin VB.CommandButton cmdBrowse 
         Height          =   255
         Left            =   5475
         Picture         =   "frmBrowse.frx":0582
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Left            =   1035
         TabIndex        =   8
         Text            =   "C:\InetPub\wwwroot\eLinkCart"
         Top             =   240
         Width           =   4410
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Folder:"
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Next"
      Height          =   375
      Left            =   3990
      TabIndex        =   6
      Top             =   2865
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   2865
      Width           =   840
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
         Picture         =   "frmBrowse.frx":0AC4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblBGHeadInfo1 
         BackStyle       =   0  'Transparent
         Caption         =   "$(PackageName) - Install Folder"
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
         Width           =   3180
      End
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
      Left            =   60
      TabIndex        =   10
      Top             =   3120
      Width           =   3780
   End
   Begin VB.Label lblHomepage 
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
      Left            =   45
      MouseIcon       =   "frmBrowse.frx":3266
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3420
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBrowse.frx":33B8
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
      Index           =   0
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
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Copyright (c) 2002 by Rice WebDesigns, Inc. All right Reserved
'

'Variables must be declared
Option Explicit

Private blnRepeat    As Boolean
Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1

Private Sub SetupDir()

On Error Resume Next

    Dim lpIDList        As Long ' Declare Varibles
    Dim sBuffer         As String
    Dim szTitle         As String
    Dim tBrowseInfo     As BROWSEINFO
    Dim intRtn          As Integer
    
    szTitle = "Please select the Installation folder for eLinkCart. " & vbCrLf
    szTitle = szTitle & "To create a new folder, press 'Make New Folder'"


    With tBrowseInfo
         .hwndOwner = Me.hwnd ' Owner Form
         .lpszTitle = szTitle & vbNullChar
         .ulFlags = BROWSE_FLAGS.BIF_RETURNONLYFSDIRS + BROWSE_FLAGS.BIF_DONTGOBELOWDOMAIN + BROWSE_FLAGS.BIF_USENEWUI
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
         sBuffer = Space(MAX_PATH)
         SHGetPathFromIDList lpIDList, sBuffer
         sBuffer = left(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
         txtDirectory.Text = sBuffer
         If Len(sBuffer) > 6 Then
            If LCase(Mid(sBuffer, Len(sBuffer) - 6)) = "wwwroot" Or _
               LCase(Mid(sBuffer, Len(sBuffer) - 6)) = "wwroot\" Or _
               LCase(Mid(sBuffer, Len(sBuffer) - 6)) = "wwroot/" Then
               intRtn = MsgBox("It is not recommended to install eLinkCart in the IIS Root Directory. " & vbCrLf & _
                             "Please select a different directory or select Continue to proceed.", vbApplicationModal + vbInformation + vbOKOnly)
               If intRtn = vbCancel Then
                  blnRepeat = True
                  Exit Sub
               End If
            End If
         End If
    End If

End Sub

Private Sub cmdBrowse_Click()

On Error Resume Next

Call_Setup:

   blnRepeat = False
   SetupDir
   If blnRepeat Then
      GoTo Call_Setup
   End If
   
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdContinue_Click()
    
On Error GoTo cmdContinue_Error

   Dim intStatus        As Boolean
   Dim intRtn           As Integer
   Dim strPath          As Variant
   Dim intLimit         As Integer
   Dim intLoop          As Integer
   Dim strDir           As String
   
   gstrAppPath = Trim(txtDirectory.Text)
    
   If Len(gstrAppPath) > 6 Then
      If LCase(Mid(gstrAppPath, Len(gstrAppPath) - 6)) = "wwwroot" Or _
         LCase(Mid(gstrAppPath, Len(gstrAppPath) - 6)) = "wwroot\" Or _
         LCase(Mid(gstrAppPath, Len(gstrAppPath) - 6)) = "wwroot/" Then
         intRtn = MsgBox("It is not recommended to install eLinkCart in the IIS Root Directory. " & vbCrLf & _
                       "Please select a different directory or select Continue to proceed.", vbApplicationModal + vbInformation + vbOKCancel)
         If intRtn = vbCancel Then
            Exit Sub
         End If
      End If
   End If
   
   gstrAppPath = Replace(gstrAppPath, "/", "\", 1, -1, vbTextCompare)
   strPath = Split(gstrAppPath, "\")
   intLimit = UBound(strPath)
   
   strDir = strPath(0) '& "\"
   For intLoop = 1 To intLimit
      strDir = strDir & "\" & strPath(intLoop)
      intStatus = CreateDir(strDir)
   Next
   
   intStatus = SetCurrentDirectory(gstrAppPath)
   If intStatus = False Then
       intRtn = MsgBox("Unable to Create Directiry. Please try Again.", vbCritical + vbOKCancel)
       If intRtn = vbCancel Then
           cmdCancel_Click
       Else
           Exit Sub
       End If
   Else
       intStatus = SetCurrentDirectory(gstrInstallDir)
   End If
      
   If Mid(gstrAppPath, 1, Len(gstrAppPath)) = "\" Then
       gstrAppPath = Mid(gstrAppPath, 1, Len(gstrAppPath) - 1)
   End If
   
   SetCurrentDirectory (gstrInstallDir)
   Unload Me
   'frmSetup.Setup
    
cmdContinue_Error:

End Sub

Private Sub Form_Load()

On Error Resume Next

   Dim strTemp    As String

    strTemp = frmBrowse.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    frmBrowse.Caption = strTemp
    
    strTemp = lblBGHeadInfo1.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    lblBGHeadInfo1.Caption = strTemp
    
    strTemp = lblBGHeadInfo2.Caption
    strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
    lblBGHeadInfo2.Caption = strTemp
    
    lblHomepage.Caption = gstrWebURL
    
    'Me.Show
    txtDirectory.Text = gstrAppPath  'gstrWWWRoot & "\eLinkCart"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

   Set IE = Nothing

End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
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

