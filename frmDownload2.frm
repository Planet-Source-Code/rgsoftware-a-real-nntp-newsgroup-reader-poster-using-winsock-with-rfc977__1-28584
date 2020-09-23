VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading Newsgroups from NNTP Server"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmDownload2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image imgNews 
      Height          =   480
      Left            =   600
      Picture         =   "frmDownload2.frx":222A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblDownloaded 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    frmNNTP.SendAndWait "CANCEL"
    DownloadingGroups = False
    Unload Me
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_Load()
    lblInfo.Caption = Replace(LoadResString(101), "{crlf}", vbCrLf)
End Sub
