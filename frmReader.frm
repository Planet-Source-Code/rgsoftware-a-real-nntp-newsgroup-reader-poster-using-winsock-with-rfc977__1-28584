VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NNTPWee News Reader"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7965
   Icon            =   "frmReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   255
         Left            =   6720
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboGroups 
         Height          =   315
         ItemData        =   "frmReader.frx":222A
         Left            =   120
         List            =   "frmReader.frx":222C
         TabIndex        =   0
         Text            =   "alt.test*"
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtBody 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3120
         Width           =   7455
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin MSComctlLib.ListView lsvSubjects 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   327681
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuSpammer 
         Caption         =   "&Spammer"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2001 by RG Software Corporation http://www.rgsoftware.com

Option Explicit
Option Compare Text

Private Function EnableControls(Value As Boolean)
    cmdFind.Enabled = Value
    cmdGo.Enabled = Value
    cmdOptions.Enabled = Value
    cmdPost.Enabled = Value
    txtBody.Enabled = Value
    lsvSubjects.Enabled = Value
    If Value Then
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbHourglass
    End If
End Function

Public Function FindGroup(FileName As String, Search As String) As String
    Dim fNum As Long
    Dim n As Long
    Dim Record As String
    Dim Temp As String
    On Error GoTo ErrHndl
    cboGroups.Clear
    fNum = FreeFile
    Open FileName For Input Lock Read As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, Record
        If Trim$(Record) Like Search Then
            cboGroups.AddItem Trim$(Record)
        End If
        DoEvents
    Loop
    Close #fNum
    Exit Function
ErrHndl:
    Close #fNum
    FindGroup = "FAILURE"
End Function

Private Sub cmdFind_Click()
    EnableControls False
    lblInfo.Caption = "Searching..."
    FindGroup NewsGroupFile, cboGroups.Text
    lblInfo.Caption = "Search Complete"
    cboGroups.Text = "[Select a group]"
    EnableControls True
End Sub

Private Sub cmdGo_Click()
    Dim n As Integer
    EnableControls False
    Headers(1).ArticleID = ""
    frmNNTP.DownloadHeaders cboGroups.Text
    'Add headers to listview
    lsvSubjects.ListItems.Clear
    On Error Resume Next
    For n = 1 To MaxHeaders
        If Headers(n).ArticleID <> "" Then
            lsvSubjects.ListItems.Add
            lsvSubjects.ListItems.Item(n) = Headers(n).ArticleID
            lsvSubjects.ListItems.Item(n).SubItems(1) = Headers(n).From
            lsvSubjects.ListItems.Item(n).SubItems(2) = Headers(n).Subject
            lsvSubjects.ListItems.Item(n).SubItems(3) = Headers(n).PostDate
        End If
        DoEvents
    Next n
    If Headers(1).ArticleID = "" Then
        lblInfo.Caption = "Error getting headers. Invalid newsgroup?"
    End If
    EnableControls True
End Sub

Private Sub cmdOptions_Click()
    frmNNTP.updateOptions
End Sub

Private Sub cmdPost_Click()
    frmPost.Group = cboGroups.Text
    frmPost.Show vbModal
    Call cmdGo_Click
End Sub

Private Sub Form_Load()
    EnableControls False
    lsvSubjects.View = lvwReport
    lsvSubjects.ColumnHeaders.Add 1, , "ArticleID"
    lsvSubjects.ColumnHeaders.Add 2, , "From"
    lsvSubjects.ColumnHeaders.Add 3, , "Subject"
    lsvSubjects.ColumnHeaders.Add 4, , "Date"
    Show
    frmNNTP.Initialize
    lblInfo.Caption = "Loading newsgroups..."
    frmNNTP.Load NewsGroupFile
    If UBound(Newsgroups) = 0 Then
        frmNNTP.Reset NewsGroupFile
    End If
    lblInfo.Caption = UBound(Newsgroups) & " newsgroups loaded"
    EnableControls True
    fraMain.Caption = "Connected to " & NNTPServer & " on port " & NNTPPort
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmNNTP.SendAndWait "QUIT"
    frmNNTP.Dissconnect
    End 'Don't even think about complaining! ;-)
End Sub

Private Sub lsvSubjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim Body As String
    lblInfo.Caption = "Downloading article " & Item.Text
    frmNNTP.SendAndWait "ARTICLE " & Item.Text, True
    Body = SCResponse
    txtBody = Body
    For n = 1 To MaxHeaders
        If Headers(n).ArticleID = Item.Text Then Exit For
    Next n
    lblInfo.Caption = Headers(n).Subject
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuOptions_Click()
    frmNNTP.updateOptions
End Sub

Private Sub mnuSpammer_Click()
    frmSpam.Show
End Sub
