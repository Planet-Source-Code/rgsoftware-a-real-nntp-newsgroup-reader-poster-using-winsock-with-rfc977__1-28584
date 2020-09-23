VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NNTP News Reader"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGroups 
      Height          =   315
      ItemData        =   "frmMain.frx":222A
      Left            =   120
      List            =   "frmMain.frx":222C
      TabIndex        =   5
      Text            =   "alt.test*"
      Top             =   480
      Width           =   6015
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtBody 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   7455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin MSComctlLib.ListView lsvSubjects 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   840
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
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NNTP client by Richard Gardner http://www.rgsoftware.com
'This is a *very* basic news reader. There are more bells & whistles to be added:
'http://www.networksorcery.com/enp/protocol/nntp.htm
'Such as threads, better implementation of commands, etc.

Option Compare Text

Private Type Header
    From As String
    Subject As String
    PostDate As String
    ArticleID As String
End Type

Private Response As String 'Response from the server
Private SCResponse As String
Private ResponseCode As Integer
Private SendComplete As Boolean

Private Newsgroups() As String 'Holds all available newsgroups
Public DownloadingGroups As Boolean

Private Const MAX_HEADERS As Integer = 25 'How many headers should be downloaded
Private Const GROUP_FILE As String = "NewsGroups.txt"

Private Headers(1 To MAX_HEADERS) As Header

Public TimeOut As Long

'Encapsulate response
Public Function getResponse() As String
    getResponse = Response
End Function

'Encapsulate response code
Public Function getResponseCode() As String
    getResponseCode = ResponseCode
End Function

Public Function Load(FileName As String) As String
    'Loads newsgroups from file.
    Dim fNum As Long
    Dim Record As String
    On Error GoTo ErrHndl
    ReDim Newsgroups(0) As String
    fNum = FreeFile
    Open FileName For Input Lock Read As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, Record
        If Trim$(Record) <> "" Then
            ReDim Preserve Newsgroups(UBound(Newsgroups) + 1) As String
            Newsgroups(UBound(Newsgroups)) = Record
        End If
        DoEvents
    Loop
    Close #fNum
    Load = "LOADED"
    Exit Function
ErrHndl:
    Close #fNum
    Load = "FAILURE"
End Function

Public Function Reset(FileName As String) As String
    'Downloads newsgroups and saves a newsgroups file.
    Dim fNum As Long
    Dim n As Long
    On Error GoTo ErrHndl
    ReDim Newsgroups(0) As String
    fNum = FreeFile
    Open FileName For Output Lock Write As #fNum
    frmDownload.Show
    DownloadingGroups = True
    SendAndWait "LIST " & vbCrLf
    If ResponseCode <> "215" Then GoTo ErrHndl
    Do While DownloadingGroups
        DoEvents
    Loop
    Unload frmDownload
    For n = 1 To UBound(Newsgroups)
        Print #fNum, Newsgroups(n)
        DoEvents
    Next n
    Close #fNum
    Reset = "GROUPS LISTED"
    Exit Function
ErrHndl:
    Close #fNum
    Reset = "FAILURE"
End Function

Public Sub Dissconnect()
    Winsock1.Close
    Caption = "Connection Closed: " & Time
End Sub

Public Function Connect(Server As String, Port As Long)
    On Error GoTo ErrHndl
    Winsock1.RemoteHost = Server
    Winsock1.RemotePort = Port
    Winsock1.Connect
    Connect = WaitForResponse
    If ResponseCode <> 200 Then Connect = "FAILED"
    Exit Function
ErrHndl:
    MsgBox Err.Description, vbExclamation
End Function

Public Function WaitForResponse(Optional WaitForSendComplete As Boolean) As String
    Dim StartTime As Long
    Dim Found As Integer
    Dim Code As String
    On Error Resume Next
    'Wait for a response
    If TimeOut = 0 Then TimeOut = 30
    'Clear the last message
    Response = ""
    SCResponse = ""
    ResponseCode = 0
    StartTime = Timer
    SendComplete = False
    If WaitForSendComplete Then
        Do While Response = "" And (Timer - StartTime) < TimeOut And Not SendComplete
            DoEvents
        Loop
    Else
        Do While Response = "" And (Timer - StartTime) < TimeOut
            DoEvents
        Loop
    End If
    If IgnoreFirst Then Stop
    If Response <> "" Then
        Found = InStr(Response, " ")
        If Found <> 0 Then
            Code = Mid$(Response, 1, Found - 1)
        End If
    Else
        Response = "TIME OUT"
    End If
    ResponseCode = CInt(Code)
    WaitForResponse = Response
    Exit Function
ErrHndl:
    WaitForResponse = "ERROR" & vbCrLf & "CONNECTION ERROR"
    Winsock1.Close
End Function

Public Function SendAndWait(Message As String, _
            Optional WaitForSendComplete As Boolean) As String
    On Error GoTo ErrHndl
    EnableControls False
    ResponseCode = 0
    Response = ""
    'Send the data
    Winsock1.SendData Message & vbCrLf
    'Wait for a response
    WaitForResponse WaitForSendComplete
    DoEvents
    EnableControls True
    Exit Function
ErrHndl:
    SendAndWait = "ERROR" & vbCrLf & "CONNECTION RESET"
    Winsock1.Close
    EnableControls True
End Function

Private Function EnableControls(Value As Boolean)
    cmdFind.Enabled = Value
    cmdGo.Enabled = Value
    txtBody.Enabled = Value
    lsvSubjects.Enabled = Value
    If Value Then
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbHourglass
    End If
End Function

Public Function Login(UserName As String, Password As String) As String
    'Login with AUTHINFO command
    SendAndWait "AUTHINFO user " & UserName
    If ResponseCode <> 381 Then
        Login = "FAILED"
    Else
        SendAndWait "AUTHINFO pass " & Password
        If ResponseCode <> 281 Then
            Login = "FAILED"
        Else
            Login = "USER " & UserName & " LOGGED IN"
        End If
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
    FindGroup GROUP_FILE, cboGroups.Text
    lblInfo.Caption = "Search Complete"
    cboGroups.Text = "[Select a group]"
    EnableControls True
End Sub

Private Sub cmdGo_Click()
    Dim n As Integer
    Headers(1).ArticleID = ""
    DownloadHeaders
    'Add headers to listview
    lsvSubjects.ListItems.Clear
    For n = 1 To MAX_HEADERS
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
End Sub

Public Function DownloadHeaders()
    Dim n As Integer
    Dim Found As Integer
    Dim Temp As String
    Dim r As String
    Dim Article As String
    Dim LastID As Long
    Dim ArticleNum As Long

    lblInfo.Caption = "Downloading headers..."
    r = SendAndWait("GROUP " & cboGroups.Text)
    
    'Looking for last message id
    'Response looks like: 211 66950 1001065 1069029 alt.test selected
    Article = Response
    For n = 1 To 3
        Found = InStr(Article, " ")
        If Found = 0 Then Exit Function
        Article = Mid$(Article, Found + 1)
    Next n
    Found = InStr(Article, " ")
    If Found = 0 Then Exit Function
    Article = Trim$(Mid$(Article, 1, Found - 1))
    LastID = CLng(Article)
    n = 0
    For ArticleNum = LastID - 1 To (LastID - MAX_HEADERS) Step -1
        n = n + 1
        Headers(n).ArticleID = CStr(ArticleNum)
        r = SendAndWait("HEAD " & CStr(ArticleNum))
        'Parse the header
        Article = Response
        'From
        Found = InStr(Article, "From:")
        If Found <> 0 Then
            Temp = Mid$(Article, Found + 6)
            Found = InStr(Temp, vbCrLf)
            If Found <> 0 Then Headers(n).From = Mid$(Temp, 1, Found - 1)
            Headers(n).From = Replace(Headers(n).From, """", "")
            Found = InStr(Headers(n).From, "<")
            If Found <> 0 Then
                Headers(n).From = Trim$(Mid$(Headers(n).From, 1, Found - 1))
            End If
        End If
        'Subject
        Found = InStr(Article, "Subject:")
        If Found <> 0 Then
            Temp = Trim$(Mid$(Article, Found + 8))
            Found = InStr(Temp, vbCrLf)
            If Found <> 0 Then Headers(n).Subject = Mid$(Temp, 1, Found - 1)
        End If
        'Post Date
        Found = InStr(Article, "Date:")
        If Found <> 0 Then
            Temp = Trim$(Mid$(Article, Found + 5))
            Found = InStr(Temp, vbCrLf)
            If Found <> 0 Then Headers(n).PostDate = Mid$(Temp, 1, Found - 1)
        End If
    Next ArticleNum
    lblInfo.Caption = ""
End Function

Private Sub Form_Load()
    EnableControls False
    lsvSubjects.View = lvwReport
    lsvSubjects.ColumnHeaders.Add 1, , "ArticleID"
    lsvSubjects.ColumnHeaders.Add 2, , "From"
    lsvSubjects.ColumnHeaders.Add 3, , "Subject"
    lsvSubjects.ColumnHeaders.Add 4, , "Date"
    Show

    TimeOut = 60 'Seconds
    'Connect "news.newsfeeds.com", 119
    Connect "news", 119
    'Login "RGSoftware", "endsub"

    lblInfo.Caption = "Loading newsgroups..."

    Load GROUP_FILE
    If UBound(Newsgroups) = 0 Then
        Reset GROUP_FILE
    End If

    lblInfo.Caption = UBound(Newsgroups) & " newsgroups loaded"
    EnableControls True
End Sub

Public Sub WaitForGroups()
    DownloadingGroups = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendAndWait "QUIT"
    Dissconnect
End Sub

Private Sub lsvSubjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim Body As String
    lblInfo.Caption = "Downloading article " & Item.Text
    SendAndWait "ARTICLE " & Item.Text, True
    Body = SCResponse
    txtBody = Body
    For n = 1 To MAX_HEADERS
        If Headers(n).ArticleID = Item.Text Then Exit For
    Next n
    lblInfo.Caption = Headers(n).Subject
End Sub

Public Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim Found As Integer
    Dim Display As Integer
    Dim Data As String
    Dim Groups As String
    Dim Group As String
    Dim Current As Long

    Display = 100
    'Get data from NNTP server
    Winsock1.GetData Data, vbString

    Response = Data
    SCResponse = SCResponse & Data

    SendComplete = False
    If InStr(Trim$(Data), vbCrLf & "." & vbCrLf) Then
        SendComplete = True 'That's all folks!
    End If

    If DownloadingGroups Then 'If we are downloading groups, do some stuff
        'These groups come in large batches so we will have to parse them.
        If Mid$(Data, 1, 3) <> "215" Then 'If this is a command response then ignore
            If SendComplete Then DownloadingGroups = False
            Do
                Found = InStr(Data, " ")
                If Found = 0 Then
                    Record = Trim$(Data)
                Else
                    Record = Trim$(Mid$(Data, 1, Found - 1))
                End If
                If Record <> "" And Len(Record) > 2 And InStr(Record, ".") <> 0 Then
                    ReDim Preserve Newsgroups(UBound(Newsgroups) + 1) As String
                    Newsgroups(UBound(Newsgroups)) = Record
                    Found = InStr(Data, vbCrLf)
                    If Found = 0 Then Exit Do
                    Data = Mid$(Data, Found + 2)
                    Current = Current + 1
                    If Current >= Display Then
                        Current = 0
                        frmDownload.lblDownloaded = "Downloading newsgroups: " & _
                                UBound(Newsgroups) & " received..."
                    End If
                Else
                    'Remove this junk group record
                    Found = InStr(Data, vbCrLf)
                    If Found = 0 Then Exit Do
                    Data = Mid$(Data, Found + 2)
                End If
                DoEvents
            Loop
        End If
    Else
        'Print Data
    End If

End Sub

Public Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Public Sub Winsock1_Connect()
    Caption = "Connected: " & Time
End Sub

Public Sub Winsock1_SendComplete()
    Caption = "Send Complete: " & Time
End Sub
