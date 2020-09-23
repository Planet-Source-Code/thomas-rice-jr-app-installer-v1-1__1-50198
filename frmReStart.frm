VERSION 5.00
Begin VB.Form frmReStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "$(PackageName) -  Installation Complete"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
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
   Icon            =   "frmReStart.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRestart 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmReStart.frx":0582
      Top             =   1095
      Width           =   6105
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&ReStart"
      Height          =   375
      Left            =   4635
      TabIndex        =   4
      Top             =   2865
      Width           =   855
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
      ScaleWidth      =   6405
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6405
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
         Picture         =   "frmReStart.frx":06CC
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblBGHeadInfo1 
         BackStyle       =   0  'Transparent
         Caption         =   " Restart Required"
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
         Width           =   2415
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
      TabIndex        =   5
      Top             =   3120
      Width           =   4185
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
      MouseIcon       =   "frmReStart.frx":094E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3435
      Width           =   2895
   End
   Begin VB.Line linBorder 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6345
      Y1              =   975
      Y2              =   975
   End
End
Attribute VB_Name = "frmReStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Copyright (c) 2002 by Rice WebDesigns, Inc. All right Reserved
'

'Variables must be declared
Option Explicit

Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1

Private Sub cmdContinue_Click()

   Unload Me
   
End Sub

Private Sub Form_Load()

On Error Resume Next

   Dim strTemp    As String

   'Me.Show
   txtRestart.Text = vbCrLf & "Windows will now Restart to enable these Changes. For " & vbCrLf & _
   "additional setup instructions, please refer to the Installation Instructions on your " & vbCrLf & _
   "Installation CD. " & vbCrLf & vbCrLf & _
   "Thank you for Installing " & gstrPackageName & " Software from " & gstrCompanyName & "." & vbCrLf
   
   strTemp = frmReStart.Caption
   strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
   frmReStart.Caption = strTemp
   
   strTemp = lblBGHeadInfo2.Caption
   strTemp = Replace(strTemp, "$(PackageName)", gstrPackageName, 1, -1, vbTextCompare)
   lblBGHeadInfo2.Caption = strTemp
   
   lblHomepage.Caption = gstrWebURL
   


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

