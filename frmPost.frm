VERSION 5.00
Begin VB.Form frmPost 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmPost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtBody 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label lblStuff 
         Caption         =   "Message:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblStuff 
         Caption         =   "Subject:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Group As String

Private Sub cmdPost_Click()
    Dim Header As String
    MousePointer = vbHourglass
    Caption = "Posting message..."
    frmNNTP.SendAndWait "POST"
    'Build a header, body and post the message to the NNTP server.
    Header = "From: " & DisplayName & " <" & Email & ">" & vbCrLf & _
            "Organization: " & Organization & vbCrLf & _
            "Subject: " & txtSubject & vbCrLf & _
            "Newsgroups: " & Group & vbCrLf & vbCrLf
    frmNNTP.SendAndWait Header & txtBody & vbCrLf & vbCrLf & "." & vbCrLf
    If ResponseCode <> 240 Then
        MsgBox "Posting failed!", vbExclamation
    Else
        MsgBox "Message posted to " & Group & "!", vbInformation
    End If
    MousePointer = vbNormal
    Group = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub
