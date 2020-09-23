VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmNNTP 
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
End
Attribute VB_Name = "frmNNTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2001 by RG Software Corporation http://www.rgsoftware.com


Option Explicit
Option Compare Text

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
    ResponseCode = 0
    Response = ""
    'Send the data
    Winsock1.SendData Message & vbCrLf
    'Wait for a response
    WaitForResponse WaitForSendComplete
    DoEvents
    Exit Function
ErrHndl:
    SendAndWait = "ERROR" & vbCrLf & "CONNECTION RESET"
    Winsock1.Close
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

Public Sub updateOptions()
    frmOptions.Show vbModal
    frmNNTP.Dissconnect
    frmNNTP.Winsock1.LocalPort = 0
    frmNNTP.Connect NNTPServer, NNTPPort 'Connect to newsgroup server
    If modNNTP.Login = "1" Then 'If user must login to server
        Login UserName, Password
    End If
End Sub

Public Function Initialize()
    If NNTPServer = "" Then
        NNTPServer = GetSetting("NNTPWee", "Settings", "NNTPServer", "")
    End If
    If NNTPPort = 0 Then
        NNTPPort = CLng(GetSetting("NNTPWee", "Settings", "NNTPPort", "119"))
    End If
    If NNTPServer = "" Then updateOptions
    If TimeOut = 0 Then
        TimeOut = CLng(GetSetting("NNTPWee", "Settings", "TimeOut", "60"))
    End If
    If MaxHeaders = 0 Then
        MaxHeaders = CLng(GetSetting("NNTPWee", "Settings", "MaxHeaders", "25"))
    End If
    DisplayName = GetSetting("NNTPWee", "Settings", "DisplayName", "")
    Email = GetSetting("NNTPWee", "Settings", "Email", "")
    Organization = GetSetting("NNTPWee", "Settings", "Organization", "")
    NewsGroupFile = GetSetting("NNTPWee", "Settings", "NewsGroupFile", "NewsGroups.txt")
    UserName = GetSetting("NNTPWee", "Settings", "UserName")
    Password = GetSetting("NNTPWee", "Settings", "Password")
    modNNTP.Login = GetSetting("NNTPWee", "Settings", "Login", "1")
    ReDim Headers(1 To MaxHeaders) As Header
    frmNNTP.Connect NNTPServer, NNTPPort 'Connect to newsgroup server
    If modNNTP.Login = "1" Then 'If user must login to server
        Login UserName, Password
    End If
End Function

Public Function DownloadHeaders(Group As String)
    Dim n As Integer
    Dim Found As Integer
    Dim Temp As String
    Dim r As String
    Dim Article As String
    Dim LastID As Long
    Dim ArticleNum As Long

    r = SendAndWait("GROUP " & Group)

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
    If Not IsNumeric(Article) Then Exit Function
    LastID = CLng(Article)
    n = 0
    For ArticleNum = LastID - 1 To (LastID - MaxHeaders) Step -1
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
End Function


Public Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim Found As Integer
    Dim Display As Integer
    Dim Data As String
    Dim Groups As String
    Dim Group As String
    Dim Record As String
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
    WinsockMessage = "Connected: " & Time
End Sub

Public Sub Winsock1_SendComplete()
    WinsockMessage = "Send Complete: " & Time
End Sub

