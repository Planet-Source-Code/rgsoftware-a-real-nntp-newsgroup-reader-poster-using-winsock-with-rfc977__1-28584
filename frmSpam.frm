VERSION 5.00
Begin VB.Form frmSpam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NNTPWee Spammer"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmSpam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ListBox lstView 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   6615
      End
      Begin VB.TextBox txtNotOk 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtOK 
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtBody 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3000
         Width           =   6615
      End
      Begin VB.Label lblStuff 
         Caption         =   "Subject:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label lblStuff 
         Caption         =   "Message:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label lblStuff 
         AutoSize        =   -1  'True
         Caption         =   "Do NOT post to newsgroups that contain these comma seperated keywords:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label lblStuff 
         AutoSize        =   -1  'True
         Caption         =   "Post to newsgroups that contain these comma seperated keywords:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5880
      Width           =   6375
   End
End
Attribute VB_Name = "frmSpam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Use this spam form with moderation!
'Please stay on topic or you may get your ISP account deactivated!

'Richard Gardner
'http://www.rgsoftware.com

Private CancelSpam As Boolean

Private OK() As String
Private NotOK() As String
Private Spam() As String
Private Preview As Boolean

Private Sub cmdCancel_Click()
    CancelSpam = True
    Unload Me
End Sub

Private Sub cmdPost_Click()
    Dim n As Long
    Dim Count As Long
    Dim Header As String
    If Not Preview Then ListGroups
    Count = getCount
    For n = 1 To Count
        If CancelSpam Then Exit For
        frmNNTP.SendAndWait "POST"
        'Build a header to send to the NNTP server
        Header = "From: " & DisplayName & " <" & Email & ">" & vbCrLf & _
                "Organization: " & Organization & vbCrLf & _
                "Subject: " & txtSubject & vbCrLf & _
                "Newsgroups: " & Spam(n) & vbCrLf & vbCrLf
        frmNNTP.SendAndWait Header & txtBody & vbCrLf & vbCrLf & "." & vbCrLf
        If ResponseCode <> 240 Then
            lblInfo.Caption = "WARNING: Posting failed for " & Spam(n) & "!"
        Else
            lstView.Selected(n - 1) = True
            lblInfo.Caption = "Message " & n & " posted to " & Spam(n)
        End If
    Next n
End Sub

Private Sub cmdPreview_Click()
    ListGroups
End Sub

Private Sub ListGroups()
    Dim Found As Integer
    Dim fNum As Long
    Dim n As Long, j As Long
    Dim Record As String
    Dim Temp As String
    Dim Ignore As Boolean
    ReDim OK(0) As String
    ReDim NotOK(0) As String
    ReDim Spam(0) As String
    Temp = txtOK
    Do
        If Trim$(Temp) = "" Then Exit Do
        Found = InStr(Temp, ",")
        ReDim Preserve OK(UBound(OK) + 1) As String
        If Found <> 0 Then
            OK(UBound(OK)) = Trim$(Mid$(Temp, 1, Found - 1))
            Temp = Mid$(Temp, Found + 1)
        Else
            OK(UBound(OK)) = Trim$(Temp)
            Exit Do
        End If
        DoEvents
    Loop
    Temp = txtNotOk
    Do
        If Trim$(Temp) = "" Then Exit Do
        Found = InStr(Temp, ",")
        ReDim Preserve NotOK(UBound(NotOK) + 1) As String
        If Found <> 0 Then
            NotOK(UBound(NotOK)) = Trim$(Mid$(Temp, 1, Found - 1))
            Temp = Mid$(Temp, Found + 1)
        Else
            NotOK(UBound(NotOK)) = Trim$(Temp)
            Exit Do
        End If
        DoEvents
    Loop
    fNum = FreeFile
    Open NewsGroupFile For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, Record
        For n = 1 To UBound(OK)
            Ignore = False
            If InStr(Record, OK(n)) <> 0 Then 'OK
                For j = 1 To UBound(NotOK)
                    If InStr(Record, NotOK(j)) <> 0 Then 'Not OK
                        Ignore = True
                        Exit For
                    End If
                Next j
                If Not Ignore Then
                    ReDim Preserve Spam(UBound(Spam) + 1) As String
                    Spam(UBound(Spam)) = Trim$(Record)
                End If
            End If
        Next n
        DoEvents
    Loop
    Close #fNum
    lblInfo.Caption = UBound(Spam) & " groups available (double click groups to remove)"
    lstView.Clear
    For n = 1 To UBound(Spam)
        lstView.AddItem (Spam(n))
    Next n
    Preview = True
End Sub

Private Sub Form_Load()
    txtSubject = GetSetting("NNTPWee", "Settings", "SpamSubject")
    txtBody = GetSetting("NNTPWee", "Settings", "SpamBody")
    txtOK = GetSetting("NNTPWee", "Settings", "OK")
    txtNotOk = GetSetting("NNTPWee", "Settings", "NotOK")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "NNTPWee", "Settings", "SpamSubject", txtSubject
    SaveSetting "NNTPWee", "Settings", "SpamBody", txtBody
    SaveSetting "NNTPWee", "Settings", "OK", txtOK
    SaveSetting "NNTPWee", "Settings", "NotOK", txtNotOk
End Sub

Private Sub lstView_DblClick()
    Dim Answer As VbMsgBoxResult
    Dim Record As String
    Dim n As Long
    Record = lstView.Text
    Answer = MsgBox("Remove " & Record & "?", vbQuestion + vbYesNo)
    If Answer = vbNo Then Exit Sub
    For n = 1 To UBound(Spam)
        If Spam(n) = Record Then
            Spam(n) = ""
            Exit For
        End If
    Next n
    lstView.Clear
    For n = 1 To UBound(Spam)
        If Spam(n) <> "" Then
            lstView.AddItem (Spam(n))
        End If
    Next n
    lblInfo.Caption = getCount & " groups available (double click groups to remove)"
End Sub

Private Function getCount()
    Dim n As Long
    Dim Count As Long
    For n = 1 To UBound(Spam)
        If Spam(n) <> "" Then Count = Count + 1
    Next n
    getCount = Count
End Function
