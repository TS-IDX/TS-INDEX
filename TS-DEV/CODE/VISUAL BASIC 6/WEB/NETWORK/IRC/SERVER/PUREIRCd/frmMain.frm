VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IRC Server"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   400
      Left            =   1980
      Top             =   3000
   End
   Begin VB.Timer tmrFloodProt 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1250
      Left            =   4860
      Top             =   3000
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   1500
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7000
   End
   Begin VB.Timer tmrKill 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   2460
      Top             =   3000
   End
   Begin VB.Timer tmrNS 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   4000
      Left            =   3420
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2940
      Top             =   3000
   End
   Begin VB.Timer tmrKlined 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10000
      Left            =   3900
      Top             =   3000
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   4380
      Top             =   3000
   End
   Begin VB.CommandButton cmdUserInfo 
      Caption         =   "User Info..."
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdBroadcast 
      Caption         =   "Broadcast"
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Kill"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   2580
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3330
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "7:39 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/13/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Connection Count"
            TextSave        =   "Connection Count"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4366
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   3000
      Width           =   5175
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuMainStartServer 
         Caption         =   "(Re)Start Server"
      End
      Begin VB.Menu mnuMainCloseServer 
         Caption         =   "Close Server"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSession 
      Caption         =   "Session"
      Begin VB.Menu mnuSessionUsers 
         Caption         =   "Users..."
      End
      Begin VB.Menu mnuSessionChannels 
         Caption         =   "Channels..."
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayStartServer 
         Caption         =   "(Re)Start Server"
      End
      Begin VB.Menu mnuTrayCloseServer 
         Caption         =   "Close Server"
      End
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show/Hide"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1

Private Sub cmdBroadcast_Click()
Dim Comment As String, i As Long
Comment = InputBox("Enter a Message: ")
For i = 1 To lvwUsers.ListItems.Count
    SendNotice lvwUsers.ListItems(i).Text, "Global - Notice *** " & Comment, "Global"
Next i
End Sub

Private Sub cmdKill_Click()
Dim NickName As String, Comment As String, User As clsUser
NickName = lvwUsers.SelectedItem.Text
Comment = InputBox("Please enter a reason:")
Set User = NickToObject(NickName)
If Not User Is Nothing Then
    User.Killed = True
    SendWsock User.Index, ":Server!~admin@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
    SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
'K-Line (Ban) User from Network for 10 seconds
    Dim Kline As Long
    Kline = GetRand
    Load tmrKlined(Kline)
    tmrKlined(Kline).Tag = wsock(User.Index).RemoteHostIP
    tmrKlined(Kline).Enabled = True
    Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
    SendQuit User.Index, "Killed by Server Admin (" & Comment & ")", True
    Dim NN As Long
    NN = GetRand
    Load tmrKill(NN)
    tmrKill(NN).Tag = User.Index
    tmrKill(NN).Enabled = True
End If
End Sub

Private Sub cmdSendMsg_Click()
SendNotice lvwUsers.SelectedItem.Text, InputBox("Enter a message"), ServerName
End Sub

Private Sub cmdUserInfo_Click()
Dim frmUI As Form
Set frmUI = New frmUserInfo
frmUI.Show , Me
End Sub

Private Sub Form_Load()
Dim i As Long, FS As New FileSystemObject
Call Log("IRCd started.")
Rehash
For i = LBound(Users) To UBound(Users)
    Set Users(i) = Nothing
    Set Channels(i) = Nothing
Next i
Started = Now
If FS.FolderExists(App.Path & "\users") = False Then FS.CreateFolder App.Path & "\users"
If FS.FolderExists(App.Path & "\channels") = False Then FS.CreateFolder App.Path & "\channels"
If FS.FolderExists(App.Path & "\memos") = False Then FS.CreateFolder App.Path & "\memos"
LoadChans
LoadMemos
lvwUsers.ColumnHeaders.Add , , "User"
lvwUsers.ColumnHeaders.Add , , "IP"
lvwUsers.ColumnHeaders.Add , , "BS/BR"
Dim GFS As clsUser
Set GFS = GetFreeSlot
GFS.Nick = "ChanServ"
GFS.ID = "ChanServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "admin@manshadow.org"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "NickServ"
GFS.ID = "NickServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "admin@manshadow.org"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "MemoServ"
GFS.ID = "MemoServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "admin@manshadow.org"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "OperServ"
GFS.ID = "OperServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "admin@manshadow.org"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
End Sub

Private Sub mnuMainCloseServer_Click()
wsock(0).Close
End Sub

Private Sub mnuMainExit_Click()
Unload Me
End Sub

Private Sub mnuMainStartServer_Click()
If Not wsock(0).State = 2 Then
    wsock(0).Close
    wsock(0).Listen
End If
End Sub

Private Sub mnuSessionChannels_Click()
frmChannels.Show , Me
End Sub

Private Sub mnuSessionUsers_Click()
frmUsers.Show , Me
End Sub

Private Sub mnuTrayCloseServer_Click()
mnuMainCloseServer_Click
End Sub

Private Sub mnuTrayExit_Click()
Unload Me
End Sub

Private Sub mnuTrayShow_Click()
frmMain.Visible = Not frmMain.Visible
End Sub

Private Sub mnuTrayStartServer_Click()
mnuMainStartServer_Click
End Sub

Private Sub Timer1_Timer()
If Not Me.Visible Then Exit Sub
Label1 = ServerTraffic & "  /  " & Uptime & "  /  " & UserCount
StatusBar1.Panels(3).Text = "Connection Count: " & (UserCount - 4)
End Sub

Private Sub tmrFloodProt_Timer(Index As Integer)
On Error Resume Next
If Users(Index).HasRegistered = False Then
    Users(Index).HasRegistered = True
    Users(Index).MsgsSent = 0
    Exit Sub
End If
If Users(Index).MsgsSent > 2000 Then
    SendQuit CLng(Index), "killed by sysadmin (excess flooding)", True
    SendWsock Users(Index).Index, ":Server!IRCd@" & ServerName & " KILL " & Users(Index).Nick & " :Excess flooding", True
    SendWsock Users(Index).Index, "ERROR :Closing Link: " & Users(Index).Nick & "[" & frmMain.wsock(Index).RemoteHostIP & ".] " & ServerName & " (excess flooding)", True
    Dim NN As Long
    NN = GetRand
    Load tmrKill(NN)
    tmrKill(NN).Tag = Index
    tmrKill(NN).Enabled = True
End If
Users(Index).MsgsSent = 0
End Sub

Private Sub tmrKill_Timer(Index As Integer)
If Index = 0 Then
    wsock_Close tmrKill(Index).Tag
    wsock(0).Listen
    tmrKill(0).Enabled = False
    tmrKill(0).Interval = 200
Else
    wsock_Close tmrKill(Index).Tag
    Unload tmrKill(Index)
End If
End Sub

Private Sub tmrKlined_Timer(Index As Integer)
Klines.Remove tmrKlined(Index).Tag
Unload tmrKlined(Index)
End Sub

Private Sub tmrNS_Timer(Index As Integer)
On Error Resume Next
If tmrNS(Index).Interval = 60 And ((Not Users(tmrNS(Index).Tag).Identified = False) And IsRegistered(Users(tmrNS(Index).Tag).Nick)) Then
    SendNotice Users(Index).Nick, "This nickname does not belong to you.", "NickServ"
    ChangeNick CLng(Index), "Guest" & OverAllMax + 1
    Unload tmrNS(Index)
End If
If Users(tmrNS(Index).Tag).NR Then
    SendNotice Users(tmrNS(Index).Tag).Nick, "You have 60 seconds to identify or change nicknames.", "NickServ"
    SendNotice Users(tmrNS(Index).Tag).Nick, "To authenticate your identity: /msg NickServ identify [password]", "NickServ"
    tmrNS(Index).Interval = 60
End If
Unload tmrNS(Index)
End Sub

Private Sub tmrSend_Timer(Index As Integer)
If Not wsock(Index).Tag = "" Then
    wsock(Index).SendData wsock(Index).Tag
    wsock(Index).Tag = ""
End If
End Sub

Private Sub tmrTimeOut_Timer(Index As Integer)
SendPing CLng(Index)
End Sub

Public Sub wsock_Close(Index As Integer)
On Error Resume Next
If Not Users(Index).SentQuit Then SendQuit CLng(Index), "Client exited."
Dim i As Long, CurChan As clsChannel
For i = 1 To Users(Index).Onchannels.Count
        Set CurChan = ChanToObject(Users(Index).Onchannels(i))
        CurChan.All.Remove Users(Index).Nick
        If CurChan.IsNorm(Users(Index).Nick) Then
            CurChan.NormUsers.Remove Users(Index).Nick
        ElseIf CurChan.IsVoice(Users(Index).Nick) Then
            CurChan.Voices.Remove Users(Index).Nick
        ElseIf CurChan.IsOp(Users(Index).Nick) Then
            CurChan.Ops.Remove Users(Index).Nick
        End If
Next i
Set Users(Index) = Nothing
wsock(Index).Close
Unload wsock(Index)
Unload tmrTimeOut(Index)
Unload tmrFloodProt(Index)
Unload tmrSend(Index)
UserCount = UserCount - 1
lvwUsers.ListItems.Remove GetListItem(CLng(Index)).Index
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim FS As clsUser, WelcomeStr As String
Set FS = GetFreeSlot
If FS Is Nothing Then
    wsock(0).Close
    wsock(0).Listen
    Exit Sub
End If
MaxUser = MaxUser + 1
Unload wsock(FS.Index)
Unload tmrTimeOut(FS.Index)
Load wsock(FS.Index)
Load tmrTimeOut(FS.Index)
Load tmrFloodProt(FS.Index)
Load tmrSend(FS.Index)
tmrTimeOut(FS.Index).Enabled = True
tmrSend(FS.Index).Enabled = True
wsock(FS.Index).Accept requestID
'Welcome User
WelcomeStr = ":" & ServerName & " NOTICE AUTH :*** Welcome to " & ServerName & "." & vbCrLf & _
                           ":" & ServerName & " NOTICE AUTH :*** Looking up your Hostname...." & vbCrLf
SendWsock FS.Index, WelcomeStr
Dim strDNS As String
strDNS = modDNS.AddressToName(wsock(FS.Index).RemoteHostIP)
If Not Mid(strDNS, 1, InStr(1, strDNS, vbNullChar) - 1) = "" Then strDNS = Mid(strDNS, 1, InStr(1, strDNS, vbNullChar) - 1)
FS.DNS = IIf(strDNS = "", wsock(FS.Index).RemoteHostIP, strDNS)
FS.NewUser = True
FS.SignOn = UnixTime
FS.Idle = UnixTime
SendWsock FS.Index, "PING " & GetRandom
lvwUsers.ListItems.Add , , "dummy"
lvwUsers.ListItems(lvwUsers.ListItems.Count).Tag = FS.Index
GetListItem(FS.Index).SubItems(1) = wsock(FS.Index).RemoteHostIP
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo parseerr
Dim strMsg As String, strCmd() As String, LB As Long, UB As Long, i As Long
wsock(Index).GetData strMsg, 8
Users(Index).BS = Users(Index).BS + Len(strMsg)
Users(Index).MsgsSent = Users(Index).MsgsSent + Len(strMsg)
GetListItem(CLng(Index)).SubItems(2) = Users(Index).BS & "/" & Users(Index).BR
ServerTraffic = ServerTraffic + Len(strMsg)
strCmd = Split(strMsg, vbLf)
LB = LBound(strCmd)
UB = UBound(strCmd)
Log "[Client]<" & Now & " (from " & Users(Index).Nick & ")> " & strMsg
For i = LB To UB
    If Users(Index).Killed Then Exit Sub
'*****************************
'|      Client Commands     ||
'*****************************
'NICK
    If Left$(strCmd(i), 5) = "NICK " Then
        Dim NewNick As String
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 400
        If InStr(1, strCmd(i), ":") <> 0 Then
            NewNick = Replace(strCmd(i), "NICK :", "")
        Else
            NewNick = Replace(strCmd(i), "NICK ", "")
        End If
        If Not (ChangeNick(CLng(Index), NewNick)) Then
            Dim NNick As String
            NNick = "Guest" & MaxUser
            SendWsock Index, ":" & ServerName & " 433 * " & NewNick & " :Nickname is already in use or may not be used, so you have been assigned a radnom one." & vbCrLf
            SendWsock Index, (":" & IIf((Users(Index).Nick = ""), NewNick, Users(Index).Nick) & " NICK " & NNick)
            Users(Index).Nick = NNick
            GetListItem(CLng(Index)).Text = NNick
            If IsKlined(wsock(Index).RemoteHostIP) Then
                SendWsock Index, ":You are banned from this server"
                wsock_Close (Index)
                Exit Sub
            End If
        ElseIf NewNick = "NickServ" And Users(Index).Nick = "" Then
            SendWsock Index, GetWelcome(CLng(Index))
            SendWsock Index, ReadMotd(Users(Index).Nick)
            Users(Index).NewUser = False
        Else
            Users(Index).Identified = False
            Users(Index).NR = False
            Users(Index).ClearOwnerShip
            If IsRegistered(Users(Index).Nick) Then
                Dim NNR As Long
                NNR = GetRand
                Load tmrNS(NNR)
                tmrNS(NNR).Tag = Users(Index).Index
                tmrNS(NNR).Enabled = True
                Users(Index).NR = True
            End If
        End If
        If IsKlined(wsock(Index).RemoteHostIP) Then
            SendWsock Index, ":Server!" & ServerName & "@" & ServerName & " KILL " & Users(Index).Nick & " :You are banned from this server"
            SendWsock Index, "ERROR :You are banned from this server."
            Dim NN As Long
            NN = GetRand
            Load tmrKill(NN)
            tmrKill(NN).Tag = Index
            tmrKill(NN).Enabled = True
        End If
    ElseIf Left$(strCmd(i), 5) = "USERHOST " Then
        Dim User As clsUser
        Set User = NickToObject(Replace(strCmd(i), "USERHOST ", ""))
        If User Is Nothing Then
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :"
        Else
            SendWsock Index, ":" & ServerName & " 302 " & Users(Index).Nick & " :" & Replace(strCmd(i), "USERHOST ", "") & "=+" & User.ID
        End If
'USER
    ElseIf Left$(strCmd(i), 5) = "USER " Then
        Dim Ident As String, Email As String, Name As String, NewIdent As String * 10
        Ident = Replace(strCmd(i), "USER ", "")
        Ident = Mid(Ident, 1, InStr(1, Ident, " ") - 1)
        Email = Replace(strCmd(i), "USER " & Ident, "")
        Email = Mid(Email, 3)
        Email = Mid(Email, 1, InStr(1, Email, " "))
        Email = Replace(Email, Chr(34), "")
        Email = Mid(Email, 1, Len(Email) - 1)
        Email = Ident & "@" & Email
        Name = Mid(strCmd(i), InStr(1, strCmd(i), ":") + 1)
        Users(Index).Ident = Mid(Ident, 1, 10)
        Ident = Mid(Ident, 1, 10) & "@" & wsock(Index).RemoteHostIP
        Users(Index).Email = Email
        Users(Index).ID = Ident
        Users(Index).Name = Name
'QUIT
    ElseIf Left$(strCmd(i), 5) = "QUIT " Then
        Dim Quit As String
        Quit = Mid(strCmd(i), InStr(1, strCmd(i), " :") + 2)
        SendQuit CLng(Index), Quit
        wsock_Close (Index)
'JOIN
    ElseIf Left$(strCmd(i), 5) = "JOIN " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim Chan As String, CK As String
        If CountSpaces(strCmd(i)) = 3 Then
            Chan = Replace(strCmd(i), "JOIN ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            CK = Replace(strCmd(i), "JOIN " & Chan & " ", "")
        Else
            Chan = Replace(strCmd(i), "JOIN ", "")
        End If
        Dim Chans() As String
        Chan = Replace(Chan, " ", "")
        Chans = Split(Chan, ",")
        Dim x As Long
        For x = 0 To UBound(Chans)
           Chan = Chans(x)
            If Not Users(Index).IsOnChan(Chan) Then
                If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 100
                If Not ChanExists(Chan) Then
                    Dim NewChannel As clsChannel
                    Set NewChannel = GetFreeChan
                    NewChannel.Name = Chan
                    NewChannel.Modes.Add "t", "t"
                    NewChannel.Modes.Add "n", "n"
                    NewChannel.Topic = DefTopic
                    NewChannel.Ops.Add Users(Index).Nick, Users(Index).Nick
                    NewChannel.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & Replace(NewChannel.GetOps & " " & NewChannel.GetVoices & " " & NewChannel.GetNorms, "  ", " ")
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list."
                Else
                    On Local Error Resume Next
                    Dim JoinChan As clsChannel
                    Set JoinChan = ChanToObject(Chan)
                    If (Not JoinChan.Key = "") Then
                        If Not JoinChan.Key = CK And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                            SendWsock Index, ":" & ServerName & " 475 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                            Exit Sub
                        End If
                    End If
                    If (JoinChan.All.Count >= JoinChan.limit And JoinChan.limit <> 0) And Not Users(Index).IRCOp And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 471 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+l)"
                        Exit Sub
                    End If
                    If JoinChan.IsBanned(Users(Index)) And (Users(Index).IRCOp = False) And (JoinChan.IsException(Users(Index)) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 474 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+b)"
                        Exit Sub
                    End If
                    If JoinChan.IsMode("i") And (Users(Index).IRCOp = False) And (JoinChan.IsInvited2(Users(Index)) = False) And (JoinChan.IsInvited(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        SendWsock Index, ":" & ServerName & " 473 " & Users(Index).Nick & " " & Chan & " :Cannot join channel (+i)"
                        Exit Sub
                    End If
                    If Not Users(Index).IRCOp And (JoinChan.ULOp(Users(Index).Nick) = False) And (JoinChan.ULVoice(Users(Index).Nick) = False) And (Not Users(Index).IsOwner(JoinChan.Name)) Then
                        JoinChan.NormUsers.Add Users(Index).Nick, Users(Index).Nick
                    ElseIf JoinChan.ULVoice(Users(Index).Nick) Then
                        JoinChan.Voices.Add Users(Index).Nick, Users(Index).Nick
                    Else
                        JoinChan.Ops.Add Users(Index).Nick, Users(Index).Nick
                    End If
                    JoinChan.All.Add Users(Index).Nick, Users(Index).Nick
                    Users(Index).Onchannels.Add Chan, Chan
                    SendWsock Index, ":" & Users(Index).Nick & " JOIN " & Chan
                    SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & Trim(Replace(JoinChan.GetOps & " " & JoinChan.GetVoices & " " & JoinChan.GetNorms, "  ", " "))
                    SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list."
                    SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & JoinChan.Topic
                    SendWsock Index, ":" & ServerName & " 333 " & JoinChan.TopicSetBy & " " & Chan & " " & JoinChan.TopicSetBy & " " & JoinChan.TopicSetOn
                    NotifyJoin CLng(Index), Chan
                    If Users(Index).IRCOp Or JoinChan.ULOp(Users(Index).Nick) Or (Users(Index).IsOwner(JoinChan.Name)) Then
                        OpUser JoinChan, Users(Index).Nick, "ChanServ", True
                    ElseIf JoinChan.ULVoice(Users(Index).Nick) Then
                        VoiceUser JoinChan, Users(Index).Nick, "ChanServ", True
                    End If
                End If
            End If
        Next x
'PART
    ElseIf Left$(strCmd(i), 5) = "PART " Then
        Chan = Replace(strCmd(i), "PART ", "")
        SendPart CLng(Index), Chan
'MODE
    ElseIf Left$(strCmd(i), 5) = "MODE " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 500
        Dim Mode As String, ToUser() As String, Modes() As String, Op As String, ToUsers As String, Channel As clsChannel, y As Long
        Chan = Replace(strCmd(i), "MODE ", "")
        If InStr(1, Chan, " ") <> 0 Then
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        End If
        Set Channel = ChanToObject(Chan)
        Set User = NickToObject(Chan)
        If Not User Is Nothing Then
            Dim UM As String
            UM = Mid(Replace(strCmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 1, 1)
            Select Case UM
                Case "+"
                    AddUserMode User.Index, Mid(Replace(strCmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
                Case "-"
                    RemoveUsermode User.Index, Mid(Replace(strCmd(i), "MODE " & User.Nick & " ", "", , , vbTextCompare), 2)
            End Select
            Exit Sub
        End If
        Dim cmdline() As String, UserMode As Boolean
        cmdline = Split(strCmd(i), " ")
        For x = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(x)) Is Nothing Then UserMode = True
        Next x
        If InStr(1, strCmd(i), "*") <> 0 Then UserMode = True
        If UserMode Then
            Mode = Replace(strCmd(i), "MODE " & Chan & " ", "")
            Op = Mid(Mode, 1, 1)
            Mode = Mid(Mode, 2, InStr(1, Mode, " ") - 2)
            ToUsers = Mid(strCmd(i), InStr(1, strCmd(i), Op) + Len(Mode) + 2)
            ParseModeNicks ToUsers, ToUser()
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
                Exit Sub
            End If
            For x = 1 To Len(Mode)
                ReDim Preserve Modes(x)
                Modes(x) = Mid(Mode, x, 1)
            Next x
            ReDim Preserve Modes(UBound(ToUser))
            For y = LBound(Modes) To UBound(Modes)
                If Not ToUser(y) = "" Then
                    Select Case Modes(IIf((y = 0), y + 1, y))
                        Case "o"
                            Select Case Op
                                Case "+"
                                    OpUser Channel, ToUser(y), Users(Index).Nick
                                Case "-"
                                    DeOpUser Channel, ToUser(y), Users(Index).Nick
                            End Select
                        Case "v"
                            Select Case Op
                                Case "+"
                                    VoiceUser Channel, ToUser(y), Users(Index).Nick
                                Case "-"
                                    DeVoiceUser Channel, ToUser(y), Users(Index).Nick
                            End Select
                        Case "b"
                            Select Case Op
                                Case "+"
                                    BanUser Channel, ToUser(y), Users(Index).Nick
                                Case "-"
                                    UnBanUser Channel, ToUser(y), Users(Index).Nick
                            End Select
                        Case "e"
                            Select Case Op
                                Case "+"
                                    ExceptionUser Channel, ToUser(y), Users(Index).Nick
                                Case "-"
                                    UnExceptionUser Channel, ToUser(y), Users(Index).Nick
                            End Select
                        Case "I"
                            Select Case Op
                                Case "+"
                                    InviteUser Channel, ToUser(y), Users(Index).Nick
                                Case "-"
                                    UnInviteUser Channel, ToUser(y), Users(Index).Nick
                            End Select
                    End Select
                End If
            Next y
        Else
            If InStr(1, strCmd(i), " +b", vbBinaryCompare) <> 0 Then
                For x = 1 To Channel.Bans.Count
                    SendWsock Index, ":" & ServerName & " 367 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Bans(x)
                Next x
                SendWsock Index, ":" & ServerName & " 368 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Ban List"
            ElseIf InStr(1, strCmd(i), " +e", vbBinaryCompare) <> 0 Then
                For x = 1 To Channel.Exceptions.Count
                    SendWsock Index, ":" & ServerName & " 348 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Exceptions(x)
                Next x
                SendWsock Index, ":" & ServerName & " 349 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Exceptions List"
            ElseIf InStr(1, strCmd(i), " +I", vbBinaryCompare) <> 0 Then
                For x = 1 To Channel.Invites.Count
                    SendWsock Index, ":" & ServerName & " 346 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.Invites(x)
                Next x
                SendWsock Index, ":" & ServerName & " 347 " & Users(Index).Nick & " " & Channel.Name & " :End of Channel Invites List"
            ElseIf InStr(1, strCmd(i), " +w", vbBinaryCompare) <> 0 Then
                SendWsock Index, ":" & ServerName & " 472 " & Users(Index).Nick & " w :is unknown mode char to me"
            ElseIf InStr(1, strCmd(i), "+") <> 0 Then
                AddChanModes Mid(strCmd(i), InStr(1, strCmd(i), "+") + 1), Chan, Users(Index)
            ElseIf InStr(1, strCmd(i), "-") <> 0 Then
                RemoveChanModes Mid(strCmd(i), InStr(1, strCmd(i), "-") + 1), Chan, Users(Index)
            Else
                SendWsock Index, ":" & ServerName & " 324 " & Users(Index).Nick & " " & Channel.Name & " " & Channel.GetModes
            End If
        End If
'TOPIC
    ElseIf Left$(strCmd(i), 6) = "TOPIC " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        If InStr(1, strCmd(i), " :") <> 0 Then
            Dim NewTopic As String
            Chan = Replace(strCmd(i), "TOPIC ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then Exit Sub
            NewTopic = strCmd(i)
            NewTopic = Mid(NewTopic, InStr(1, NewTopic, ":") + 1)
            If Channel.IsOp(Users(Index).Nick) = False And (Not Users(Index).IsOwner(Channel.Name)) Then
                SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " :You're not channel operator"
                Exit Sub
            End If
            SetTopic Chan, NewTopic, Users(Index).Nick
        Else
            Chan = Replace(strCmd(i), "TOPIC ", "")
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then Exit Sub
            SendWsock Index, ":" & ServerName & " 332 " & Users(Index).Nick & " " & Chan & " :" & Channel.Topic
            SendWsock Index, ":" & ServerName & " 333 " & Channel.TopicSetBy & " " & Chan & " " & Users(Index).Nick & " " & Channel.TopicSetOn
        End If
'INVITE
    ElseIf Left$(strCmd(i), 7) = "INVITE " Then
        Dim Target As String
        Target = Replace(strCmd(i), "INVITE ", "")
        Target = Mid(Target, 1, InStr(1, Target, " ") - 1)
        Chan = Mid(strCmd(i), Len("INVITE " & Target & " ") + 1)
        Set Channel = ChanToObject(Chan)
        If Channel.IsMode("i") And Channel.IsOp(Users(Index).Nick) = False Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            Exit Sub
        End If
        On Local Error Resume Next
        Channel.Invited.Add Target, Target
        SendWsock NickToObject(Target).Index, ":" & Users(Index).Nick & " INVITE " & Target & " " & Chan
'KICK
    ElseIf Left$(strCmd(i), 5) = "KICK " Then
        Dim Source As String, Reason As String
        Chan = Mid(strCmd(i), 6)
        Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
        Set Channel = ChanToObject(Chan)
        Source = Users(Index).Nick
        If Channel.IsOp(Source) = False And (Not Users(Index).IsOwner(Channel.Name)) Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan & " :You're not channel operator"
            Exit Sub
        End If
        If InStr(1, strCmd(i), ":") <> 0 Then
            Reason = Mid(strCmd(i), InStr(1, strCmd(i), " :") + 2)
            Target = Replace(strCmd(i), "KICK", "")
            Target = Mid(Target, 2)
            Target = Mid(Target, 1, InStr(1, Target, ":") - 2)
            Target = Replace(Target, Chan & " ", "")
            If Target = "ChanServ" Then
                SendWsock Index, ":" & ServerName & " NOTICE " & Users(Index).Nick & " :*** ChanServ is not currently Online!"
                Exit Sub
            End If
            KickUser Source, Chan, Target, Reason, True
            Exit Sub
        End If
        Target = Mid(strCmd(i), InStrRev(strCmd(i), " ", InStrRev(strCmd(i), " ")) + 1)
        If Target = "ChanServ" Then
            SendWsock Index, ":" & ServerName & " NOTICE " & Users(Index).Nick & " :*** ChanServ is not currently Online!"
            Exit Sub
        End If
        KickUser Source, Chan, Target
'PONG
    ElseIf Left$(strCmd(i), 5) = "PONG " Then
        If Users(Index).NewUser Then
            SendWsock Index, GetWelcome(CLng(Index))
            SendWsock Index, ReadMotd(Users(Index).Nick)
            Users(Index).NewUser = False
            Users(Index).MsgsSent = 0
            tmrFloodProt(Index).Enabled = True
            If DefUserModes <> "" Then AddUserMode CLng(Index), DefUserModes
        End If
'PING
    ElseIf Left$(strCmd(i), 5) = "PING " Then
        SendWsock Index, "PONG " & Replace(strCmd(i), "PING ", ""), True
'PRIVMSG
    ElseIf Left$(strCmd(i), 8) = "PRIVMSG " Then
        cmdline = Split(Mid(strCmd(i), 1, InStr(1, strCmd(i), " :")), " ")
        For x = LBound(cmdline) To UBound(cmdline)
            If Not NickToObject(cmdline(x)) Is Nothing Then UserMode = True
        Next x
        Target = Replace(strCmd(i), "PRIVMSG ", "")
        Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
        Select Case LCase(Target)
            Case "chanserv"
                UserMode = True
            Case "nickserv"
               UserMode = True
            Case "memoserv"
               UserMode = True
            Case "operserv"
               UserMode = True
        End Select
        If (Not UserMode) Then
            Dim msgstr As String, msg As String
            Chan = Replace(strCmd(i), "PRIVMSG ", "")
            Chan = Mid(Chan, 1, InStr(1, Chan, " ") - 1)
            Set Channel = ChanToObject(Chan)
            If Channel Is Nothing Then
                SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                Exit Sub
            End If
            If Not Channel.IsOnChan(Users(Index).Nick) And (Not Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                Exit Sub
            End If
            If Channel.IsMode("m") Then
                If Channel.IsOp(Users(Index).Nick) Then
                ElseIf Channel.IsVoice(Users(Index).Nick) Or (Users(Index).IsOwner(Channel.Name)) And Not Users(Index).IRCOp Then
                Else
                    SendWsock Index, ":" & ServerName & " 404 " & Users(Index).Nick & " " & Chan & " :Cannot send to channel"
                    Exit Sub
                End If
            End If
            msg = strCmd(i)
            msg = Mid(msg, InStr(1, msg, ":") + 1)
            SendMsg Chan, msg, Users(Index).Nick
        Else
            Target = Replace(strCmd(i), "PRIVMSG ", "")
            Target = Strings.Left(Target, InStr(1, Target, ":") - 2)
            msg = strCmd(i)
            msg = Mid(msg, InStr(1, msg, ":") + 1)
            If UCase(Target) = UCase("CHANSERV") Then
                ParseCSCmd "CS " & msg, CLng(Index)
                Exit Sub
            End If
            If UCase(Target) = UCase("NICKSERV") Then
                ParseNSCmd "NS " & msg, CLng(Index)
                Exit Sub
            End If
            If UCase(Target) = UCase("MEMOSERV") Then
                ParseMSCmd "MS " & msg, CLng(Index)
                Exit Sub
            End If
            If UCase(Target) = UCase("OPERSERV") Then
                ParseOSCmd "OS " & msg, CLng(Index)
                Exit Sub
            End If
            Set User = NickToObject(Target)
            If User Is Nothing Then
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                Exit Sub
            End If
            SendMsg Target, msg, Users(Index).Nick, False
        End If
'NOTICE
    ElseIf Left$(strCmd(i), 7) = "NOTICE " Then
        Target = Replace(strCmd(i), "NOTICE ", "")
        Target = Replace(Target, ":*", " ")
        Target = Left(Target, InStr(1, Target, ":") - 2)
        msg = strCmd(i)
        msg = Mid(msg, InStr(1, msg, ":") + 1)
        If InStr(1, Target, "#") = 0 Then
            If NickToObject(Target) Is Nothing Then
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                Exit Sub
            End If
            SendNotice Target, msg, Users(Index).Nick
        Else
            Dim CurChan As clsChannel
            Set CurChan = ChanToObject(Target)
            If CurChan Is Nothing Then
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
                Exit Sub
            End If
            For x = 1 To ChanToObject(Target).All.Count
                If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 75
                SendNotice Target, msg, Users(Index).Nick, True, NickToObject(CurChan.All(x)).Index
            Next x
        End If
'MOTD
    ElseIf Left$(strCmd(i), 4) = "MOTD" Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1200
        SendWsock Index, ReadMotd(Users(Index).Nick)
'WHOIS
    ElseIf Left$(strCmd(i), 6) = "WHOIS " Then
        Dim WhoisStr As String, Nick As String
        Set User = NickToObject(Replace(strCmd(i), "WHOIS ", ""))
        If Not User Is Nothing Then
            SendWsock Index, User.GetWhois(Users(Index).Nick)
        Else
            SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
        End If
'AWAY
    ElseIf Left$(strCmd(i), 5) = "AWAY " Then
        If Not Users(Index).Away Then
            Users(Index).AwayMsg = Replace(strCmd(i), "AWAY :", "")
            Users(Index).Away = True
            SendWsock Index, ":" & ServerName & " 306 " & Users(Index).Nick & " :You have been marked as being away"
            Users(Index).Modes.Add "a", "a"
        Else
            Users(Index).Away = False
            Users(Index).AwayMsg = ""
            RemoveUsermode CLng(Index), "a", True
        End If
'WALLOPS
    ElseIf Left$(strCmd(i), 8) = "WALLOPS " Then
        If Users(Index).IRCOp Then
            WallOps Replace(strCmd(i), "WALLOPS ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'WALL
    ElseIf Left$(strCmd(i), 5) = "WALL " Then
        If Users(Index).IRCOp Then
            Wall Replace(strCmd(i), "WALL ", ""), Index
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'*****************************
'|      Client Queries             ||
'*****************************
'VERSION
    ElseIf Left$(strCmd(i), 7) = "VERSION" Then
        SendWsock Index, GetWelcome(CLng(Index))
'TIME
    ElseIf Left$(strCmd(i), 4) = "TIME" Then
        SendWsock Index, ":" & ServerName & " 391" & Users(Index).Nick & " " & ServerName & " :" & Now
'INFO
    ElseIf Left$(strCmd(i), 4) = "INFO" Then
        SendWsock Index, GetWelcome(CLng(Index))
'ISON
    ElseIf Left$(strCmd(i), 5) = "ISON " Then
        Dim strIsOn As String, LoggedIn() As String, IsOnArr() As String
        ReDim LoggedIn(1)
        strIsOn = Replace(strCmd(i), "ISON ", "")
        IsOnArr = Split(strIsOn, " ")
        For x = LBound(IsOnArr) To UBound(IsOnArr)
            If Not NickToObject(IsOnArr(x)) Is Nothing Then
                ReDim Preserve LoggedIn(UBound(LoggedIn) + 1)
                LoggedIn(UBound(LoggedIn)) = IsOnArr(x)
            End If
        Next x
        strIsOn = Join(LoggedIn, " ")
        SendWsock Index, (":" & ServerName & " 303 " & Users(Index).Nick & " :" & strIsOn)
'LUSERS
    ElseIf Left$(strCmd(i), 7) = "LUSERS " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 400
        SendWsock Index, ":" & ServerName & " 254 " & Users(Index).Nick & " :channels formed = " & ChanCount
        SendWsock Index, ":" & ServerName & " 255 " & Users(Index).Nick & " :I have " & UserCount & " clients"
        SendWsock Index, ":Server uptime: " & Uptime
        SendWsock Index, ":ServerTraffic (bytes): " & (ServerTraffic)
        SendWsock Index, ":" & ServerName & " NOTICE " & Users(Index).Nick & " :Overall connection count: " & MaxUser
'STATS
    ElseIf Left$(strCmd(i), 6) = "STATS " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        Dim StatsParam As String
        StatsParam = Replace(strCmd(i), "STATS ", "")
        Select Case StatsParam
            Case "u"
                SendWsock Index, ":" & ServerName & " 242 " & Users(Index).Nick & " :" & Uptime
                SendWsock Index, ":" & ServerName & " 250 " & Users(Index).Nick & " :Highest Connection Count: " & MaxUser
                SendWsock Index, ":" & ServerName & " 219 " & Users(Index).Nick & " u :End of /STATS report"
        End Select
'INFO
    ElseIf Left$(strCmd(i), 5) = "INFO " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 250
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :" & ServerName & " running vbIRCd " & App.Major & "." & App.Minor & "." & App.Revision
        SendWsock Index, ":" & ServerName & " 371 " & Users(Index).Nick & " :This server was created Sun Aug 11 2002 at 21:38:12 (GMT +0100) by ~admin Fisch (Server@gmx.net)"
        SendWsock Index, ":" & ServerName & " 374 " & Users(Index).Nick & " :End of INFO list"
'LINKS
    ElseIf Left$(strCmd(i), 6) = "LINKS " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 200
        SendWsock Index, ":" & ServerName & " 364 " & Users(Index).Nick & " " & ServerName & " " & ServerName & " :0 " & ServerDesc
        SendWsock Index, ":" & ServerName & " 365 " & Users(Index).Nick & " * :End of /LINKS list"
'NAMES
    ElseIf Left$(strCmd(i), 6) = "NAMES " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 150
        Chan = Replace(strCmd(i), "NAMES ", "")
        Set Channel = ChanToObject(Chan)
        SendWsock Index, ":" & ServerName & " 353 " & Users(Index).Nick & " = " & Chan & " :" & Channel.GetOps & " " & Channel.GetVoices & " " & Channel.GetNorms
        SendWsock Index, ":" & ServerName & " 366 " & Users(Index).Nick & " " & Chan & " :End of /NAMES list."
'LIST
    ElseIf Left$(strCmd(i), 5) = "LIST " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 1000
        SendWsock Index, ":" & ServerName & " 321 " & Users(Index).Nick & " Channel :Users  Name"
        SendWsock Index, GetChanList(Users(Index).Nick)
        SendWsock Index, ":" & ServerName & " 323 " & Users(Index).Nick & " :End of /LIST"
'ADMIN
    ElseIf Left$(strCmd(i), 6) = "ADMIN " Then
        If Not Users(Index).IRCOp Then Users(Index).MsgsSent = Users(Index).MsgsSent + 750
        SendWsock Index, ":" & ServerName & " 256 " & Users(Index).Nick & " :Administrative info about " & ServerName
        SendWsock Index, ":" & ServerName & " 257 " & Users(Index).Nick & " :" & ServerDesc
        SendWsock Index, ":" & ServerName & " 258 " & Users(Index).Nick & " :" & AdminName
        SendWsock Index, ":" & ServerName & " 259 " & Users(Index).Nick & " :" & AdminEmail
'*****************************
'|      Operator Queries        ||
'*****************************
'OPER
    ElseIf Left$(strCmd(i), 5) = "OPER " Then
        Dim PW As String, UserName As String
        UserName = Replace(strCmd(i), "OPER ", "")
        PW = Mid(UserName, InStr(1, UserName, " ") + 1)
        UserName = Mid(UserName, 1, InStr(1, UserName, " ") - 1)
        PW = Replace(PW, ":", "")
        If Not HasOline(Users(Index).Nick, Users(Index).GetMask) Then
            SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :No O-lines for your host"
            Exit Sub
        End If
        With Olines(GetOline(Users(Index).DNS))
            If Not Users(Index).Nick = .UserName Then
                SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :your nickname must match the nickname with which the O-Line has been created"
                Exit Sub
            End If
            If Not Users(Index).Identified Then
                SendWsock Index, ":" & ServerName & " 491 " & Users(Index).Nick & " :You have to be identified!"
                Exit Sub
            End If
            If Not PW = .Password Then
                SendWsock Index, ":" & ServerName & " 464 " & Users(Index).Nick & " :Password incorrect"
                Exit Sub
            End If
            SendWsock Index, ":" & ServerName & " 381 " & Users(Index).Nick & " :You are now an IRC operator"
            SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " +o"
            On Local Error Resume Next
            Users(Index).AddModes "o"
            Users(Index).IRCOp = True
            SendSvrMsg Users(Index).Nick & " is now Operator"
            'WallOps " is now Operator", 1
        End With
'RESTART
    ElseIf Left$(strCmd(i), 7) = "RESTART" Then
        If Users(Index).IRCOp Then
            Restart Users(Index).Nick
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'DIE
    ElseIf Left$(strCmd(i), 4) = "DIE " Then
        If Users(Index).IRCOp Then
            End
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'K-LINE
    ElseIf Left$(strCmd(i), 6) = "KLINE " Then
        If Not Users(Index).IRCOp Then
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
            Exit Sub
        End If
        If Not NickToObject(Replace(strCmd(i), "KLINE ", "")) Is Nothing Then
            Klines.Add wsock(NickToObject(Replace(strCmd(i), "KLINE ", "")).Index).RemoteHostIP
        Else
            Klines.Add Replace(strCmd(i), "KLINE ", "")
        End If
'KILL
    ElseIf Left$(strCmd(i), 5) = "KILL " Then
        If Users(Index).IRCOp Then
            Dim NickName As String, Comment As String
            NickName = Replace(strCmd(i), "KILL ", "")
            NickName = Mid(NickName, 1, InStr(1, NickName, " :") - 1)
            Comment = Replace(strCmd(i), "KILL " & NickName & " :", "")
            Set User = NickToObject(NickName)
            If Not User Is Nothing Then
                User.Killed = True
                SendWsock User.Index, ":Server!~admin@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
                SendWsock User.Index, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
                'K-Line (Ban) User from Network for 10 seconds
                Dim Kline As Long
                Kline = GetRand
                Load tmrKlined(Kline)
                tmrKlined(Kline).Tag = wsock(User.Index).RemoteHostIP
                tmrKlined(Kline).Enabled = True
                Klines.Add wsock(User.Index).RemoteHostIP, wsock(User.Index).RemoteHostIP
                SendQuit User.Index, "Killed by " & Users(Index).Nick & " (" & Comment & ")", True
                NN = GetRand
                Load tmrKill(NN)
                tmrKill(NN).Tag = User.Index
                tmrKill(NN).Enabled = True
                SendSvrMsg "Recieved Kill message for " & User.Nick & "!" & User.Ident & "@" & User.DNS & " Path: " & Users(Index).Nick
                SendNotice Users(Index).Nick, "User " & User.Nick & " has been successfully removed and K-Lined from the network", "" & ServerName & ""
            Else
                SendWsock Index, ":" & ServerName & " 401 " & Users(Index).Nick & " :No such nick/channel"
            End If
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'REHASH
    ElseIf Left$(strCmd(i), 7) = "REHASH " Then
        If Users(Index).IRCOp Then
            Rehash Users(Index).Nick
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'WRITEHASH
    ElseIf Left$(strCmd(i), 10) = "WRITEHASH " Then
        If Users(Index).IRCOp Then
            WriteHash
        Else
            SendWsock Index, ":" & ServerName & " 481 " & Users(Index).Nick & " :Permission Denied- You're not an IRC operator"
        End If
'*****************************
'|     Service Commands    ||
'*****************************
'NICKSERV
    ElseIf Left$(strCmd(i), 3) = "NS " Then
        ParseNSCmd strCmd(i), CLng(Index)
    ElseIf Left$(strCmd(i), 9) = "NICKSERV " Then
        ParseNSCmd Replace(strCmd(i), "NICKSERV", "NS"), CLng(Index)
'MEMOSERV
    ElseIf Left$(strCmd(i), 3) = "MS " Then
        ParseMSCmd strCmd(i), CLng(Index)
    ElseIf Left$(strCmd(i), 9) = "MEMOSERV " Then
        ParseMSCmd Replace(strCmd(i), "MEMOSERV", "MS"), CLng(Index)
'CHANSERV
    ElseIf Left$(strCmd(i), 3) = "CS " Then
        ParseCSCmd strCmd(i), CLng(Index)
    ElseIf Left$(strCmd(i), 9) = "CHANSERV " Then
        ParseCSCmd Replace(strCmd(i), "CHANSERV", "CS"), CLng(Index)
'OPERSERV
    ElseIf Left$(strCmd(i), 3) = "OS " Then
        ParseOSCmd strCmd(i), CLng(Index)
    ElseIf Left$(strCmd(i), 9) = "OPERSERV " Then
        ParseOSCmd Replace(strCmd(i), "OPERSERV", "OS"), CLng(Index)
    ElseIf strCmd(i) = "" Then
    Else
        If InStr(1, strCmd(i), " ") <> 0 Then strCmd(i) = Mid(strCmd(i), 1, InStr(1, strCmd(i), " ") - 1)
        SendWsock Index, ":" & ServerName & " 421 " & Users(Index).Nick & " :" & strCmd(i) & " Unknown command."
    End If
Next i
Exit Sub
parseerr:
If Not Users(Index) Is Nothing Then SendWsock Index, ":" & ServerName & " 421 " & Users(Index).Nick & " :Unknown command or Parsing error | " & Err.Description & " - " & Err.Number
End Sub

Private Sub wsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
SendQuit CLng(Index), "Connection error: " & Description, False
wsock_Close (Index)
End Sub
