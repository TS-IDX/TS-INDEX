Attribute VB_Name = "modChanServ"
Option Explicit
Private dab As New clsDatabase
Private FS As New FileSystemObject

Private Sub AddToList(Channel As clsChannel, Nick As String, level As String, Index As Long)
Channel.AddToUserList Nick, CLng(level)
SendNotice Users(Index).Nick, "Added user " & Nick & " with level " & level & " to Userlist", "ChanServ"
Dim i As Long, NickName As String, Level2 As Long
With dab
    .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
    .WriteEntry "UserLevels", "Count", CStr(Channel.UserLevels.Count)
    For i = 1 To Channel.UserLevels.Count
        Channel.GetUserListItem i, NickName, Level2
        .WriteEntry "User " & CStr(i), "Nickname", NickName
        .WriteEntry "User " & CStr(i), "Level", CStr(Level2)
    Next i
End With
End Sub

Private Sub RemoveFromList(Channel As clsChannel, Nick, Index As Long)
Channel.RemoveFromUserList CStr(Nick)
SendNotice Users(Index).Nick, "Removed user " & Nick & " from Userlist", "ChanServ"
Dim i As Long, NickName As String, Level2 As Long
With dab
    .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
    .WriteEntry "UserLevels", "Count", CStr(Channel.UserLevels.Count)
    For i = 1 To Channel.UserLevels.Count
        Channel.GetUserListItem i, NickName, Level2
        .WriteEntry "User " & CStr(i), "Nickname", NickName
        .WriteEntry "User " & CStr(i), "Level", CStr(Level2)
    Next i
End With
End Sub

Private Sub Register(Channel As clsChannel, Password As String, Info As String, Index As Long)
Dim NickName As String, level As Long, i As Long
If FS.FileExists(App.Path & "\channels\" & Channel.Name & ".dat") Then
    SendNotice Users(Index).Nick, "Channel is already registered", "ChanServ"
    Exit Sub
End If
AddChanModes "rtn", Channel.Name, NickToObject(Users(Index).Nick)
SetTopic Channel.Name, Info & " <- (registered by " & Users(Index).Nick & ")", "ChanServ"
SendNotice Users(Index).Nick, "Channel has been registered, your password is " & Password, "ChanServ"
Channel.AddToUserList Users(Index).Nick, 100
Channel.Password = Password
With dab
    .FileName = App.Path & "\channels\" & Channel.Name & ".dat"
    .WriteEntry "General", "Name", Channel.Name
    .WriteEntry "General", "Topic", Channel.Topic
    .WriteEntry "General", "TopicSetOn", Channel.TopicSetOn
    .WriteEntry "General", "TopicSetBy", Channel.TopicSetBy
    .WriteEntry "General", "Modes", Channel.GetModes
    .WriteEntry "General", "Password", Channel.Password
    .WriteEntry "UserLevels", "Count", Channel.UserLevels.Count
    For i = 1 To Channel.UserLevels.Count
        Channel.GetUserListItem i, NickName, level
        .WriteEntry "User " & CStr(i), "Nickname", NickName
        .WriteEntry "User " & CStr(i), "Level", CStr(level)
    Next i
    .WriteEntry "Bans", "Count", Channel.Bans.Count
    For i = 1 To Channel.Bans.Count
        Channel.GetUserListItem i, NickName, level
        .WriteEntry "Ban " & CStr(i), "Mask", Channel.Bans(i)
    Next i
    .WriteEntry "Invites", "Count", Channel.Invites.Count
    For i = 1 To Channel.Invites.Count
        .WriteEntry "Invite " & CStr(i), "Mask", Channel.Invites(i)
    Next i
    .WriteEntry "Exceptions", "Count", Channel.Exceptions.Count
    For i = 1 To Channel.Exceptions.Count
        .WriteEntry "Exception " & CStr(i), "Mask", Channel.Exceptions(i)
    Next i
End With
End Sub

Private Function Identify(Nick As clsUser, Channel As clsChannel, pass As String) As Boolean
If Channel.Password = pass Then Identify = True
End Function

Private Sub Drop(Channel As clsChannel, Password, Index As Long)
If Channel.Password = Password Then
    SendNotice Users(Index).Nick, "Channel has been dropped", "ChanServ"
    FS.DeleteFile App.Path & "\channels\" & Channel.Name & ".dat"
    RemoveChanModes "r", Channel.Name, NickToObject(Users(Index).Nick)
ElseIf Channel.IsMode("r") = False Then
    SendNotice Users(Index).Nick, "Channel is not registered", "ChanServ"
    Exit Sub
Else
    SendNotice Users(Index).Nick, "Invalid Password", "ChanServ"
End If
End Sub

Private Sub ListUL(Channel As clsChannel, Index As Long)
Dim User As clsUser, i As Long, NickName As String * 20, Level2 As Long
Set User = NickToObject(Users(Index).Nick)
SendNotice User.Nick, "Userlist Listing for Channel " & Channel.Name, "ChanServ"
For i = 1 To Channel.UserLevels.Count
    Channel.GetUserListItem i, NickName, Level2
    SendNotice User.Nick, NickName & "                " & Level2, "ChanServ"
Next i
End Sub

Public Sub ParseCSCmd(Cmd As String, Index As Long)
On Error GoTo parseerr
Dim msg As String, CMDStr As String
msg = Replace(Cmd, "CS ", "")
On Local Error Resume Next
CMDStr = (Mid(msg, 1, InStr(1, msg, " ") - 1))
msg = Replace(msg, CMDStr & " ", "")
If CMDStr = "" Then CMDStr = msg
Dim User As clsUser
Set User = Users(Index)
Select Case LCase(CMDStr)
    Case "addtolist"
        Dim Chan As clsChannel, PW As String, Desc As String
        Set Chan = ChanToObject((Mid(msg, 1, InStr(1, msg, " ") - 1)))
        If Not Chan.ULOp(User.Nick) Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan.Name & " :You're not channel operator"
            Exit Sub
        End If
        msg = Replace(msg, Chan.Name & " ", "")
        PW = (Mid(msg, 1, InStr(1, msg, " ") - 1))
        Desc = Replace(msg, PW & " ", "")
        AddToList Chan, PW, Desc, Index
    Case "removefromlist"
        Set Chan = ChanToObject((Mid(msg, 1, InStr(1, msg, " ") - 1)))
        If Not Chan.ULOp(User.Nick) Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan.Name & " :You're not channel operator"
            Exit Sub
        End If
        msg = Replace(msg, Chan.Name & " ", "")
        RemoveFromList Chan, msg, Index
    Case "register"
        Set Chan = ChanToObject((Mid(msg, 1, InStr(1, msg, " ") - 1)))
        If Chan Is Nothing Then
            SendWsock Index, ":" & ServerName & " 442 " & Users(Index).Nick & " " & Chan.Name & " :You're not on that channel"
            Exit Sub
        ElseIf Chan.IsOp(Users(Index).Nick) = False Then
            SendWsock Index, ":" & ServerName & " 482 " & Users(Index).Nick & " " & Chan.Name & " :You're not channel operator"
            Exit Sub
        End If
        msg = Replace(msg, Chan.Name & " ", "")
        PW = (Mid(msg, 1, InStr(1, msg, " ") - 1))
        Desc = Replace(msg, PW & " ", "")
        Register Chan, PW, Desc, Index
    Case "identify"
        Dim NewChanName As String
        NewChanName = (Mid(msg, 1, InStr(1, msg, " ") - 1))
        Set Chan = ChanToObject(NewChanName)
        If Not Chan Is Nothing Then
            msg = Replace(msg, Chan.Name & " ", "")
            If Identify(Users(Index), Chan, msg) Then
                SendNotice Users(Index).Nick, "Password accepted: you are now considered the channel owner", "ChanServ"
                Users(Index).OwnerOf.Add Chan.Name, Chan.Name
            Else
                SendNotice Users(Index).Nick, "Invalid Password", "ChanServ"
            End If
        Else
            SendNotice Users(Index).Nick, NewChanName & " is not a registered channel", "ChanServ"
        End If
    Case "drop"
        Set Chan = ChanToObject((Mid(msg, 1, InStr(1, msg, " ") - 1)))
        Drop Chan, Replace(msg, Chan.Name & " ", ""), Index
    Case "list"
        Set Chan = ChanToObject(msg)
        ListUL Chan, Index
    Case "clear"
        Set Chan = ChanToObject(msg)
        If Users(Index).IRCOp Or Users(Index).IsOwner(Chan.Name) Then
            Dim Item As Variant
            For Each Item In Chan.All
                KickUser "ChanServ", Chan.Name, CStr(Item), "Clear command used by " & Users(Index).Nick, True
            Next
        Else
            SendNotice Users(Index).Nick, "You do not have the privileges to use this command on this channel", Users(Index).Nick
        End If
    Case "help"
        If msg = "" Or msg = "help" Then
            SendNotice User.Nick, "ChanServ Commands", "ChanServ"
            SendNotice User.Nick, "ADDTOLIST [chan] [User] [level]", "ChanServ"
            SendNotice User.Nick, "REMOVEFROMLIST [chan] [User]", "ChanServ"
            SendNotice User.Nick, "LIST [chan]", "ChanServ"
            SendNotice User.Nick, "REGISTER [chan] [Password] [Description]", "ChanServ"
            SendNotice User.Nick, "DROP [chan] [password]", "ChanServ"
            SendNotice User.Nick, "IDENTIFY [chan] [password]", "ChanServ"
            SendNotice User.Nick, "CLEAR [chan]", "ChanServ"
        Else
            Select Case LCase(msg)
                Case "addtolist"
                    SendNotice User.Nick, "AddToList [chan] [User] [level]", "ChanServ"
                    SendNotice User.Nick, "Add a user to the Userlevel's list", "ChanServ"
                Case "removefromlist"
                    SendNotice User.Nick, "RemoveFromList [chan] [User]", "ChanServ"
                    SendNotice User.Nick, "Remove a user from the Userlevel's list", "ChanServ"
                Case "list"
                    SendNotice User.Nick, "List [chan]", "ChanServ"
                    SendNotice User.Nick, "Lists user levels of [chan]", "ChanServ"
                Case "register"
                    SendNotice User.Nick, "Register [chan] [Password] [Description]", "ChanServ"
                Case "drop"
                    SendNotice User.Nick, "Drop [chan] [password]", "ChanServ"
                Case "identify"
                    SendNotice User.Nick, "Identify [chan] [password]", "ChanServ"
                    SendNotice User.Nick, "Identify to gain Channel owner privileges", "ChanServ"
                Case "clear"
                    SendNotice User.Nick, "Clear [chan]", "ChanServ"
                    SendNotice User.Nick, "You need to have IRCop or Channel owner privileges for this command", "ChanServ"
            End Select
        End If
End Select
Exit Sub
parseerr:
SendNotice User.Nick, "Unknown Command or missing parameters", "ChanServ"
End Sub
