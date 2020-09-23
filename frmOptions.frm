VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NNTP Options"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOptions 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   8
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox chkLogin 
         Caption         =   "NNTP Server Requires Login"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   1530
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   7
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   6
         Text            =   "25"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Text            =   "119"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtOptions 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "news.newsfeeds.com"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Organization"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Headers"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Out"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Email Address"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Display Name"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Password"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "Username"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "NNTP Port"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblStuff 
         Alignment       =   1  'Right Justify
         Caption         =   "NNTP Server"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLogin_Click()
    If chkLogin.Value = 1 Then
        txtOptions(3).Enabled = True
        txtOptions(4).Enabled = True
        txtOptions(3).BackColor = vbWindowBackground
        txtOptions(4).BackColor = vbWindowBackground
    Else
        txtOptions(3).Enabled = False
        txtOptions(4).Enabled = False
        txtOptions(3).BackColor = vbButtonFace
        txtOptions(4).BackColor = vbButtonFace
    End If
End Sub

Private Sub cmdOK_Click()
    NNTPServer = txtOptions(0)
    SaveSetting "NNTPWee", "Settings", "NNTPServer", NNTPServer
    NNTPPort = txtOptions(1)
    SaveSetting "NNTPWee", "Settings", "NNTPPort", NNTPPort
    TimeOut = txtOptions(2)
    SaveSetting "NNTPWee", "Settings", "NNTPPort", NNTPPort
    UserName = txtOptions(3)
    SaveSetting "NNTPWee", "Settings", "UserName", UserName
    Password = txtOptions(4)
    SaveSetting "NNTPWee", "Settings", "Password", Password
    SaveSetting "NNTPWee", "Settings", "Login", chkLogin.Value
    MaxHeaders = txtOptions(5)
    SaveSetting "NNTPWee", "Settings", "MaxHeaders", MaxHeaders
    DisplayName = txtOptions(6)
    SaveSetting "NNTPWee", "Settings", "DisplayName", DisplayName
    Email = txtOptions(7)
    SaveSetting "NNTPWee", "Settings", "Email", Email
    Organization = txtOptions(8)
    SaveSetting "NNTPWee", "Settings", "Organization", Organization
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbNormal
    txtOptions(0) = NNTPServer
    txtOptions(1) = NNTPPort
    txtOptions(2) = TimeOut
    txtOptions(3) = UserName
    txtOptions(4) = Password
    chkLogin.Value = Login
    txtOptions(5) = MaxHeaders
    txtOptions(6) = DisplayName
    txtOptions(7) = Email
    txtOptions(8) = Organization
End Sub

Private Sub txtOptions_GotFocus(Index As Integer)
    txtOptions(Index).SelStart = 0
    txtOptions(Index).SelLength = Len(txtOptions(Index))
End Sub
