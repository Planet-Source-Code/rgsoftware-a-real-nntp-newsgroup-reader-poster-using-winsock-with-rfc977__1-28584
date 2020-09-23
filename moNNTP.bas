Attribute VB_Name = "modNNTP"
'NNTP client by Richard Gardner
'This is a *very* basic nntp client. There are more bells & whistles to be added:
'http://www.networksorcery.com/enp/protocol/nntp.htm
'Such as threads, better implementation of commands, etc.

'This copyright must remain:
'Copyright 2001 by RG Software Corporation
'http://www.rgsoftware.com

Public Type Header
    From As String
    Subject As String
    PostDate As String
    ArticleID As String
End Type

Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Response As String 'Response from the server
Public SCResponse As String
Public ResponseCode As Integer
Public SendComplete As Boolean
Public WinsockMessage As String

Public NNTPServer As String
Public NNTPPort As Long

Public Newsgroups() As String  'Holds all available newsgroups
Public DownloadingGroups As Boolean

Public UserName As String
Public Password As String
Public DisplayName As String
Public Email As String
Public Organization As String
Public Login As Integer

Public MaxHeaders As Integer  'How many headers should be downloaded
Public NewsGroupFile As String

Public Headers() As Header

Public TimeOut As Long
