Attribute VB_Name = "modIRCUserCommands"
Option Explicit
Dim NickName As String, level As Long

Public Function ChangeNick(Index As Long, NewNick As String) As Boolean
On Error Resume Next
If NickInUse(NewNick) Then
    ChangeNick = False
    Exit Function
End If
If UCase(Users(Index).Nick) = UCase(NewNick) Then
    ChangeNick = True
    Exit Function
End If
If Users(Index).Nick = "" Then
    SendWsock Index, ":" & NewNick & " NICK " & NewNick
    ChangeNick = True
Else
    Dim i As Long, x As Long, CurChan As clsChannel, CurNick As clsUser
    For i = 1 To Users(Index).Onchannels.Count
        Set CurChan = ChanToObject(Users(Index).Onchannels(i))
        If CurChan.IsNorm(Users(Index).Nick) Then
            CurChan.NormUsers.Remove Users(Index).Nick
            CurChan.NormUsers.Add NewNick, NewNick
        ElseIf CurChan.IsVoice(Users(Index).Nick) Then
            CurChan.Voices.Remove Users(Index).Nick
            CurChan.Voices.Add NewNick, NewNick
        ElseIf CurChan.IsOp(Users(Index).Nick) Then
            CurChan.Ops.Remove Users(Index).Nick
            CurChan.Ops.Add NewNick, NewNick
        End If
        CurChan.All.Remove Users(Index).Nick
        CurChan.All.Add NewNick, NewNick
        For x = 1 To (CurChan.All.Count)
            If Not NewNick = CurChan.All(x) Then SendWsock NickToObject(CurChan.All(x)).Index, ": " & Users(Index).Nick & " NICK " & NewNick
        Next x
    Next i
    SendWsock Index, ":" & Users(Index).Nick & " NICK " & NewNick
    ChangeNick = True
End If
GetListItem(Index).Text = NewNick
Users(Index).Nick = NewNick
ChangeNick = True
End Function

Public Sub SendMsg(Target As String, Message As String, User As String, Optional SendToChan As Boolean = True)
On Error Resume Next
Dim Index As Long
Index = NickToObject(User).Index
NickToObject(User).Idle = UnixTime
If SendToChan Then
    Dim i As Long, x As Long, CurChan As clsChannel
    Set CurChan = ChanToObject(Target)
    For x = 1 To (CurChan.All.Count)
        If Not User = CurChan.All(x) Then SendWsock NickToObject(CurChan.All(x)).Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PRIVMSG " & Target & " :" & Message
    Next x
Else
    SendWsock NickToObject(Target).Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PRIVMSG " & Target & " :" & Message & vbCrLf
End If
End Sub

Public Sub SendQuit(Index As Long, QuitMsg As String, Optional Kill As Boolean = False)
On Error Resume Next
Dim i As Long, x As Long, CurChan As clsChannel
If QuitMsg = "" Then QuitMsg = DefQuit
Users(Index).SentQuit = True
For i = 1 To Users(Index).Onchannels.Count
    Set CurChan = ChanToObject(Users(Index).Onchannels(i))
    For x = 1 To (CurChan.All.Count)
        If Not Users(Index).Nick = CurChan.All(x) Then SendWsock NickToObject(CurChan.All(x)).Index, ": " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " QUIT :" & QuitMsg
    Next x
    CurChan.All.Remove Users(Index).Nick
    If CurChan.IsNorm(Users(Index).Nick) Then
        CurChan.NormUsers.Remove Users(Index).Nick
    ElseIf CurChan.IsVoice(Users(Index).Nick) Then
        CurChan.Voices.Remove Users(Index).Nick
    ElseIf CurChan.IsOp(Users(Index).Nick) Then
        CurChan.Ops.Remove Users(Index).Nick
    End If
Next i
If Not Kill Then SendWsock Index, ": " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " QUIT :" & QuitMsg
End Sub

Public Sub SendPart(Index As Long, Chan As String)
On Error Resume Next
Dim i As Long, x As Long, CurChan As clsChannel, Found As Boolean
Users(Index).Idle = UnixTime
For i = 1 To Users(Index).Onchannels.Count
    If Chan = Users(Index).Onchannels(i) Then Found = True
Next i
If Found = False Then Exit Sub
Set CurChan = ChanToObject(Chan)
For x = 1 To (CurChan.All.Count)
    If Not Users(Index).Nick = CurChan.All(x) Then SendWsock NickToObject(CurChan.All(x)).Index, ": " & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " PART " & Chan
Next x
CurChan.All.Remove Users(Index).Nick
If CurChan.IsNorm(Users(Index).Nick) Then
    CurChan.NormUsers.Remove Users(Index).Nick
ElseIf CurChan.IsVoice(Users(Index).Nick) Then
    CurChan.Voices.Remove Users(Index).Nick
ElseIf CurChan.IsOp(Users(Index).Nick) Then
    CurChan.Ops.Remove Users(Index).Nick
End If
Users(Index).Onchannels.Remove Chan
SendWsock Index, ": " & Users(Index).Nick & " PART " & Chan
End Sub

Public Sub SendPing(Index As Long)
SendWsock Index, "" '"PING " & GetRandom
End Sub

Public Sub NotifyJoin(Index As Long, Chan As String)
Dim i As Long, Channel As clsChannel
Set Channel = ChanToObject(Chan)
For i = 1 To Channel.All.Count
    If Not Channel.All(i) = Users(Index).Nick Or Channel.All(i) = "" Then SendWsock NickToObject(Channel.All(i)).Index, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " JOIN " & Chan
Next i
End Sub

Public Sub SendNotice(Target As String, Message As String, User As String, Optional ToChannel As Boolean = False, Optional Index As Integer)
On Error Resume Next
If Not ToChannel Then
    SendWsock NickToObject(Target).Index, ":" & User & " NOTICE " & Target & " :" & Message
Else
    SendWsock Index, ":" & User & " NOTICE " & Target & " :" & Message
End If
    '[Server] :Server!Fox_JK_Re@dm-8554.dip.t-dialin.net NOTICE #fox :hey
End Sub

Public Sub KickUser(Source As String, Chan As String, Target As String, Optional Reason As String, Optional Reasoning As Boolean = False)
On Error Resume Next
Dim i As Long, Channel As clsChannel, KickMsg As String
Set Channel = ChanToObject(Chan)
If Reasoning = True Then
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target & " :" & Reason
Else
    KickMsg = ":" & Source & " KICK " & Chan & " " & Target
End If
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, KickMsg
Next i
Channel.All.Remove Target
If Channel.IsNorm(Target) Then
    Channel.NormUsers.Remove Target
ElseIf Channel.IsVoice(Target) Then
    Channel.Voices.Remove Target
ElseIf Channel.IsOp(Target) Then
    Channel.Ops.Remove Target
End If
NickToObject(Target).Onchannels.Remove Chan
End Sub

Public Sub SetTopic(Chan As String, NewTopic As String, User As String)
Dim i As Long, Channel As clsChannel
Set Channel = ChanToObject(Chan)
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, ":" & User & " TOPIC " & Chan & " :" & NewTopic
Next i
Channel.Topic = NewTopic
Channel.TopicSetOn = UnixTime
Channel.TopicSetBy = User
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Topic", NewTopic
        .WriteEntry "General", "TopicSetOn", Channel.TopicSetOn
        .WriteEntry "General", "TopicSetBy", User
    End With
End If
End Sub

Public Sub OpUser(Channel As clsChannel, Target As String, User As String, Optional OpAnyway As Boolean = False)
On Error Resume Next
Dim i As Long, Chan As String
Chan = Channel.Name
If Channel.IsOp(Target) And OpAnyway = False Then Exit Sub
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " +o " & Target)
Next i
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
Channel.Ops.Add Target, Target
End Sub

Public Sub DeOpUser(Channel As clsChannel, Target As String, User As String)
On Error Resume Next
Dim i As Long, Chan As String
If Not Channel.IsOp(Target) Then Exit Sub
Chan = Channel.Name
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " -o " & Target)
Next i
Channel.Ops.Remove Target
If Channel.IsVoice(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
End Sub
Public Sub VoiceUser(Channel As clsChannel, Target As String, User As String, Optional VoiceAnyway As Boolean = False)
On Error Resume Next
Dim i As Long, Chan As String
Chan = Channel.Name
If Channel.IsVoice(Target) And VoiceAnyway = False Then Exit Sub
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " +v " & Target)
Next i
Channel.Voices.Add Target, Target
If Channel.IsNorm(Target) Then Channel.NormUsers.Remove Target
End Sub
Public Sub DeVoiceUser(Channel As clsChannel, Target As String, User As String)
On Error Resume Next
Dim i As Long, Chan As String
If Not Channel.IsVoice(Target) Then Exit Sub
Chan = Channel.Name
For i = 1 To Channel.All.Count
     SendWsock NickToObject(Channel.All(i)).Index, (":" & User & " MODE " & Chan & " -v " & Target)
Next i
Channel.Voices.Remove Target
If Channel.IsOp(Target) Then
Else
    Channel.NormUsers.Add Target, Target
End If
End Sub

Public Sub BanUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
If Not Channel.IsBanned2(Replace(Target, "*!", "")) Then
    Channel.Bans.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +b " & Target
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Bans", "Count", Channel.Bans.Count
        For i = 1 To Channel.Bans.Count
            .WriteEntry "Ban " & CStr(i), "Mask", Channel.Bans(i)
        Next i
    End With
End If
End Sub

Public Sub UnBanUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
For i = 1 To Channel.All.Count
    SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -b " & Target
Next i
Channel.Bans.Remove Replace(Target, "*!", "")
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Bans", "Count", Channel.Bans.Count
        For i = 1 To Channel.Bans.Count
            .WriteEntry "Ban " & CStr(i), "Mask", Channel.Bans(i)
        Next i
    End With
End If
End Sub

Public Sub RemoveChanModes(NewModes As String, Chan As String, User As clsUser)
Dim Found As Boolean, x As Long, Channel As clsChannel, Modes As String
Set Channel = ChanToObject(Chan)
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For x = 1 To Len(NewModes)
        If Channel.IsMode(Mid(NewModes, x, 1)) Then
            Channel.Modes.Remove Mid(NewModes, x, 1)
            Modes = Modes & Mid(NewModes, x, 1)
        End If
    Next x
    If Modes = "" Then Exit Sub
    Dim i As Long
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & Modes
    Next i
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, limit As String
    limit = Replace(NewModes, "lk ", "")
    If Channel.Key = limit Then Channel.Key = ""
    Channel.limit = 0
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "k" Then
    If Channel.Key = Replace(NewModes, "k ", "") Then
        Channel.Key = ""
        For i = 1 To Channel.All.Count
            SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
        Next i
    End If
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.limit = 0
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " -" & NewModes
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Modes", Channel.GetModesForFile
        .WriteEntry "General", "Key", Channel.Key
        .WriteEntry "General", "Limit", Channel.limit
    End With
End If
End Sub

Public Sub AddChanModes(NewModes As String, Chan As String, User As clsUser)
Dim Found As Boolean, x As Long, Channel As clsChannel, Modes As String
On Error Resume Next
Set Channel = ChanToObject(Chan)
If Not (Mid(NewModes, 1, 1) = "k" Or Mid(NewModes, 1, 1) = "l") Then
    For x = 1 To Len(NewModes)
        If Not Channel.IsMode(Mid(NewModes, x, 1)) Then
            Channel.Modes.Add Mid(NewModes, x, 1), Mid(NewModes, x, 1)
            Modes = Modes & Mid(NewModes, x, 1)
        End If
    Next x
    If Modes = "" Then Exit Sub
    Dim i As Long
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & Modes
    Next i
ElseIf Mid(NewModes, 1, 2) = "lk" Then
    Dim Key As String, limit As String
    limit = Replace(NewModes, "lk ", "")
    Key = Mid(limit, 1, InStr(1, limit, " ") - 1)
    limit = Replace(limit & " ", Key, "")
    limit = Trim(limit)
    Channel.Key = limit
    Channel.limit = Key
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "k" Then
    Channel.Key = Replace(NewModes, "k ", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
ElseIf Mid(NewModes, 1, 1) = "l" Then
    Channel.limit = Replace(NewModes, "l ", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User.Nick & "!" & User.ID & " MODE " & Channel.Name & " +" & NewModes
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "General", "Modes", Channel.GetModesForFile
        .WriteEntry "General", "Key", Channel.Key
        .WriteEntry "General", "Limit", Channel.limit
    End With
End If
End Sub

Public Function GetChanList(User As String)
':immortal.se.eu.darkmyst.org 322 Guest1 #crystalgate 2 :[+tnr] Spammers paradise! - 4w00t!
Dim i As Long, Chan As clsChannel
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        GetChanList = GetChanList & ":" & ServerName & " 322 " & User & " " & Channels(i).Name & " " & Channels(i).All.Count & " :[+" & Channels(i).GetModes & "] " & Channels(i).Topic & vbCrLf
    End If
Next i
End Function

Public Sub InviteUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
If Not Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +I " & Target
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Invites", "Count", Channel.Invites.Count
        For i = 1 To Channel.Invites.Count
            .WriteEntry "Invite " & CStr(i), "Mask", Channel.Invites(i)
        Next i
    End With
End If
End Sub

Public Sub UnInviteUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
If Channel.IsInvited3(Replace(Target, "*!", "")) Then
    Channel.Invites.Remove Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -I " & Target
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Invites", "Count", Channel.Invites.Count
        For i = 1 To Channel.Invites.Count
            .WriteEntry "Invite " & CStr(i), "Mask", Channel.Invites(i)
        Next i
    End With
End If
End Sub

Public Sub ExceptionUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
If Not Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Add Replace(Target, "*!", ""), Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " +e " & Target
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Exceptions", "Count", Channel.Exceptions.Count
        For i = 1 To Channel.Exceptions.Count
            .WriteEntry "Exception " & CStr(i), "Mask", Channel.Exceptions(i)
        Next i
    End With
End If
End Sub

Public Sub UnExceptionUser(Channel As clsChannel, Target As String, User As String)
'[Client] MODE #fox2 +b *!*@80.143.156.2847
'[Server] :Server!Fox_JK_Re@80.143.156.2847 MODE #fox2 +b *!*@80.143.156.2847
Dim i As Long
If Channel.IsException2(Replace(Target, "*!", "")) Then
    Channel.Exceptions.Remove Replace(Target, "*!", "")
    For i = 1 To Channel.All.Count
        SendWsock NickToObject(Channel.All(i)).Index, ":" & User & "!" & NickToObject(User).ID & " MODE " & Channel.Name & " -e " & Target
    Next i
End If
If Channel.IsMode("r") Then
    Dim dab As New clsDatabase
    With dab
        .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
        .WriteEntry "Exceptions", "Count", Channel.Exceptions.Count
        For i = 1 To Channel.Exceptions.Count
            .WriteEntry "Exception " & CStr(i), "Mask", Channel.Exceptions(i)
        Next i
    End With
End If
End Sub

Public Sub AddUserMode(Index As Long, Modes As String, Optional Silent As Boolean = False)
Dim NewModes As String, i As Long
For i = 1 To Len(Modes)
    If InStr("OoKkMAYIiwsRPpHLQ", Mid$(Modes$, i&, 1)) <> 0 Then
         NewModes$ = NewModes$ & Users(Index&).AddModes(Mid$(Modes$, i&, 1))
    End If
Next i
If Silent Then Exit Sub
If Not NewModes = "" Then SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " +" & NewModes
End Sub

Public Sub RemoveUsermode(Index As Long, Modes As String, Optional Silent As Boolean = False)
Dim NewModes As String
Modes = LCase(Modes)
Dim i As Long
For i = 1 To Len(Modes)
    Select Case Mid(Modes, i, 1)
        Case "s"
            If Users(Index).IsMode("s") Then
                NewModes = NewModes & "s"
                Users(Index).Modes.Remove "s"
            End If
        Case "w"
            If Users(Index).IsMode("w") Then
                NewModes = NewModes & "w"
                Users(Index).Modes.Remove "w"
            End If
        Case "a"
            If Users(Index).IsMode("a") Then
                NewModes = NewModes & "a"
                Users(Index).Modes.Remove "a"
                Users(Index).Away = False
                Users(Index).AwayMsg = ""
                SendWsock Index, ":" & ServerName & " 305 " & Users(Index).Nick & " :You are no longer marked as being away"
            End If
        Case "o"
            If Users(Index).IsMode("o") Then
                NewModes = NewModes & "o"
                Users(Index).Modes.Remove "o"
                Users(Index).IRCOp = False
            End If
    End Select
Next i
If Silent Then Exit Sub
If Not NewModes = "" Then SendWsock Index, ":" & Users(Index).Nick & " MODE " & Users(Index).Nick & " -" & NewModes
End Sub
