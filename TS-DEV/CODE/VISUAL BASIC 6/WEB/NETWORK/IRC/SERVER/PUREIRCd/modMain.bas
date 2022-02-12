Attribute VB_Name = "modMain"
Option Explicit
Option Base 1
Global ChanCount As Long
Global UserCount As Long
Global MaxUser As Long
Global Users() As clsUser
Global Channels() As clsChannel
Global Memos As New MemoCol
Global Olines(100) As Oline
Public ServerName As String
Public Started As Date
Public Klines As New Collection
Public ServerTraffic As Long
Public OverAllMax As Long
Public DefTopic As String
Public DefUserModes As String
Public DefQuit As String
Public MaxChannels As String
Public ServerDesc As String
Public AdminName As String
Public AdminEmail As String

Public Type Oline
    UserName As String
    Password As String
    Mask As String
    InUse As Boolean
End Type

Public Function UnixTime() As Long
UnixTime = DateDiff("s", DateValue("1/1/1970"), Now)
End Function

Public Function Uptime() As String
Uptime = SecsToMins2(DateDiff("s", Started, Now))
End Function

Public Function ChanToObject(ChanName As String) As clsChannel
On Error Resume Next
Dim i As Long
Set ChanToObject = Nothing
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        If UCase(ChanName) = UCase(Channels(i).Name) Then
            Set ChanToObject = Channels(i)
            Exit Function
        End If
    End If
Next i
End Function

Public Function NickToObject(NickName As String) As clsUser
On Error Resume Next
Dim i As Long
For i = 1 To UBound(Users)
    'DoEvents
    If Not Users(i) Is Nothing Then
        If UCase(NickName) = UCase(Users(i).Nick) Then
    '        MsgBox UCase(Users(i).Nick)
            Set NickToObject = Users(i)
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetFreeSlot() As clsUser
Dim i As Long
For i = 1 To UBound(Users)
    DoEvents
    If (Users(i) Is Nothing) Then
        Set Users(i) = New clsUser
        Users(i).Index = i
        Set GetFreeSlot = Users(i)
        UserCount = UserCount + 1
        Exit Function
    End If
Next i
Set GetFreeSlot = Nothing
End Function

Public Sub SendWsock(Index, Message, Optional SendImmediately As Boolean = False)
On Error Resume Next
Users(Index).BR = Users(Index).BR + Len(Message & vbCrLf)
GetListItem(CLng(Index)).SubItems(2) = Users(Index).BS & "/" & Users(Index).BR
ServerTraffic = ServerTraffic + Len(Message & vbCrLf)
Log "[Server]<" & Now & " (to " & Users(Index).Nick & ")> " & Message
If Not SendImmediately Then
    frmMain.wsock(Index).Tag = frmMain.wsock(Index).Tag & Message & vbCrLf
Else
    frmMain.wsock(Index).SendData Message & vbCrLf
End If
Debug.Print Message
End Sub

Public Function NickInUse(NickName As String) As Boolean
On Error Resume Next
Dim i As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        If UCase(Users(i).Nick) = UCase(NickName) Then
            NickInUse = True
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetRandom() As Long
Randomize
Dim MyValue As Long, i As Long, r As Long
For i = 1 To 8
    MyValue = Int((9 * Rnd) + 0)
    r = CLng(CStr(r) & CStr(MyValue))
Next i
GetRandom = r
End Function

Public Function ChanExists(ChannelName As String) As Boolean
If Not ChanToObject(ChannelName) Is Nothing Then ChanExists = True
End Function

Public Function GetFreeChan() As clsChannel
Dim i As Long
For i = 1 To UBound(Channels)
    DoEvents
    If (Channels(i) Is Nothing) Then
        Set Channels(i) = New clsChannel
        Channels(i).Index = i
        Set GetFreeChan = Channels(i)
        ChanCount = ChanCount + 1
        Exit Function
    End If
Next i
Set GetFreeChan = Nothing
End Function

Public Function ReadMotd(Nick As String) As String
On Error Resume Next
Dim FS As New FileSystemObject
If FS.FileExists(App.Path & "\motd.txt") Then
    With FS.OpenTextFile(App.Path & "\Motd.txt", ForReading)
        ReadMotd = ":" & ServerName & " 375 " & Nick & " :- " & ServerName & " Message of the day, " & vbCrLf
        ReadMotd = ReadMotd & ":" & ServerName & " 372 " & Nick & " :- " & Now & vbCrLf
        Do While (Not .AtEndOfStream)
            DoEvents
            ReadMotd = ReadMotd & ":" & ServerName & " 372 " & Nick & " :- " & .ReadLine & vbCrLf
        Loop
        ReadMotd = ReadMotd & ":" & ServerName & " 376 " & Nick & " :End of /MOTD command." & vbCrLf
    End With
End If
End Function

Public Function SecsToMins2(Seconds)
Dim lnge, mins, secs
    lnge = Seconds
    mins = lnge \ 60
    secs = lnge Mod 60
    If Len(secs) = 0 Then
        secs = secs & "00"
    ElseIf Len(secs) = 1 Then
        secs = "0" & secs
    End If
    SecsToMins2 = MinsToHrs(mins) & (", ") & (secs) & " Seconds"
End Function

Public Function MinsToHrs(Minutes)
Dim lnge, mins, secs
    lnge = Minutes
    mins = lnge \ 60
    secs = lnge Mod 60
    If Len(secs) = 0 Then
        secs = secs & "00"
    ElseIf Len(secs) = 1 Then
        secs = "0" & secs
    End If
    MinsToHrs = HrsToDays(mins) & (", ") & (secs) & " Minutes"
End Function

Public Function HrsToDays(Hrs)
Dim lnge, mins, secs
    lnge = Hrs
    mins = lnge \ 24
    secs = lnge Mod 24
    If Len(secs) = 0 Then
        secs = secs & "00"
    ElseIf Len(secs) = 1 Then
        secs = "0" & secs
    End If
    HrsToDays = (mins) & (" Days, ") & (secs) & " Hours"
End Function

Public Function CountSpaces(strCount As String) As Long
Dim i As Long
For i = 1 To Len(strCount)
    If (Mid(strCount, i, 1) = " ") Then CountSpaces = CountSpaces + 1
Next i
CountSpaces = CountSpaces + 1
End Function

Public Sub ParseModeNicks(Nicks As String, ByRef Nickarr() As String)
If InStr(1, Nicks, " ") <> 0 Then
    Nickarr = Split(Nicks, " ")
Else
    ReDim Nickarr(1)
    Nickarr(1) = Nicks
End If
End Sub

Public Sub Restart(Optional Nick As String = "Dill.Mine.nu")
Dim i As Long
SendSvrMsg "Recieved restart command from " & Nick
On Error Resume Next
frmMain.lvwUsers.ListItems.Clear
For i = LBound(Users) To UBound(Users)
    Set Users(i) = Nothing
    Set Channels(i) = Nothing
    Unload frmMain.wsock(i)
    Unload frmMain.tmrTimeOut(i)
Next i
For i = 1 To Memos.Count
    Memos.Remove i
Next i
Dim GFS As clsUser
Set GFS = GetFreeSlot
GFS.Nick = "ChanServ"
GFS.ID = "ChanServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "Server@gmx.de"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "NickServ"
GFS.ID = "NickServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "Server@gmx.de"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "MemoServ"
GFS.ID = "MemoServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "Server@gmx.de"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
Set GFS = GetFreeSlot
GFS.Nick = "OperServ"
GFS.ID = "OperServ@" & ServerName & ""
GFS.DNS = "" & ServerName & ""
GFS.Email = "Server@gmx.de"
GFS.IRCOp = True
GFS.Name = "Service"
GFS.SignOn = UnixTime
LoadChans
LoadMemos
Rehash
End Sub

Public Function GetWelcome(Index As Long) As String
Dim User As clsUser
Set User = Users(Index)
GetWelcome = ":" & ServerName & " 001 " & User.Nick & " :Welcome to the # man shadow IRC Network " & User.Nick & "!" & User.ID & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 002 " & User.Nick & " :Server is " & ServerName & ", running version " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 003 " & User.Nick & " :This server was created Wed Sep 11 2002 at 21:38:12 GMT" & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 004 " & User.Nick & " " & ServerName & " " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
GetWelcome = GetWelcome & ":" & ServerName & " 005 " & User.Nick & " NOQUIT TOKEN SAFELIST :are available on this server" & vbCrLf
End Function

Public Function GetRand() As Long
Randomize
Dim MyValue As Long, i As Long, r As Long
For i = 1 To 4
    MyValue = Int((9 * Rnd) + 0)
    r = CLng(CStr(r) & CStr(MyValue))
Next i
GetRand = r
End Function

Public Function IsKlined(IP As String) As Boolean
Dim i As Long
For i = 1 To Klines.Count
    If Klines(i) = IP Then
        IsKlined = True
        Exit Function
    End If
Next i
End Function

Public Sub LoadChans()
On Error Resume Next
Dim dab As New clsDatabase, FS As New FileSystemObject, ChanFile As File
Dim Chan As clsChannel
For Each ChanFile In FS.GetFolder(App.Path & "\channels").Files
    Set Chan = GetFreeChan
    dab.FileName = ChanFile.Path
    Chan.Name = dab.ReadEntry("General", "Name", "")
    Chan.Topic = dab.ReadEntry("General", "Topic", "")
    Chan.TopicSetOn = CLng(dab.ReadEntry("General", "TopicSetOn", CStr(UnixTime)))
    Chan.TopicSetBy = dab.ReadEntry("General", "TopicSetBy", "")
    Chan.AddModes dab.ReadEntry("General", "Modes", "r")
    Chan.Password = dab.ReadEntry("General", "Password", "")
    Chan.Key = dab.ReadEntry("General", "Key", "")
    Chan.limit = dab.ReadEntry("General", "Limit", "0")
    Dim i As Long
    For i = 1 To dab.ReadEntry("UserLevels", "Count", "0")
        Chan.AddToUserList dab.ReadEntry("User " & CStr(i), "Nickname", ""), CLng(dab.ReadEntry("User " & CStr(i), "Level", "0"))
    Next i
    For i = 1 To dab.ReadEntry("Bans", "Count", "0")
        Chan.Bans.Add dab.ReadEntry("Ban " & CStr(i), "Mask", ""), dab.ReadEntry("Ban " & CStr(i), "Mask", "")
    Next i
    For i = 1 To dab.ReadEntry("Exceptions", "Count", "0")
        Chan.Exceptions.Add dab.ReadEntry("Exception " & CStr(i), "Mask", ""), dab.ReadEntry("Exception " & CStr(i), "Mask", "")
    Next i
    For i = 1 To dab.ReadEntry("Invites", "Count", "0")
        Chan.Invites.Add dab.ReadEntry("Invite " & CStr(i), "Mask", ""), dab.ReadEntry("Invite " & CStr(i), "Mask", "")
    Next i
Next
End Sub

Public Function GetListItem(Index As Long) As ListItem
Dim i As Long
For i = 1 To frmMain.lvwUsers.ListItems.Count
    If frmMain.lvwUsers.ListItems(i).Tag = Index Then Set GetListItem = frmMain.lvwUsers.ListItems(i)
Next i
End Function

Public Sub Log(LogStr As String)
On Error Resume Next
Dim FS As New FileSystemObject, sDir As String, nOpenFile As Integer
Call Kill(App.Path & "\IRCd.log")
nOpenFile% = FreeFile
sDir$ = App.Path & "\IRCd.log"
Open sDir$ For Output As #nOpenFile%
Close #nOpenFile%
FS.OpenTextFile(App.Path & "\IRCd.log", ForAppending, True).WriteLine LogStr
FS.OpenTextFile(App.Path & "\IRCd.log", ForAppending, True).WriteBlankLines 1
End Sub

Public Function SizeString(strData As String, Size As Long) As String
If Size <= Len(strData) Then
    SizeString = strData
    Exit Function
End If
strData = strData & Space(Size - Len(strData))
SizeString = strData
End Function

Public Sub LoadMemos()
On Error Resume Next
Dim FS As New FileSystemObject, MemoFile As File
For Each MemoFile In FS.GetFolder(App.Path & "\memos").Files
    With MemoFile.OpenAsTextStream
        Memos.Add .ReadLine, .ReadLine, .ReadLine
        Memos(Memos.Count).Read = False
        Memos(Memos.Count).MemoID = MemoFile.Name
    End With
Next
End Sub

Public Sub Rehash(Optional Nick As String = "Dill.mine.nu")
Dim DB As New clsDatabase, i As Long, Kline As String
DB.FileName = App.Path & "\ircd.conf"
'Server Settings
ServerName = DB.ReadEntry("General Settings", "Servername", "dill.mine.nu")
ServerDesc = DB.ReadEntry("General Settings", "Description", "KillTheDill")
frmMain.wsock(0).Close
frmMain.wsock(0).LocalPort = DB.ReadEntry("General Settings", "Port", "6667")
frmMain.wsock(0).Listen
MaxUser = DB.ReadEntry("General Settings", "MaxUsers", "100")
MaxChannels = DB.ReadEntry("General Settings", "MaxChannels", "100")
ReDim Preserve Users(MaxUser)
ReDim Preserve Channels(MaxChannels)
'Admin
AdminName = DB.ReadEntry("Admin", "Name", "")
AdminEmail = DB.ReadEntry("Admin", "Email", "")
'Channel Defaults
DefTopic = DB.ReadEntry("Channel Defaults", "Topic", "Unregistered Channel")
'Standard Usermodes during log in
DefUserModes = DB.ReadEntry("Default User Settings", "UserModes", "i")
DefQuit = DB.ReadEntry("Default User Settings", "Default Quit Msg", "Fox Servers IRCd")
For i = 1 To DB.ReadEntry("K-lines", "Count", "0")
    Kline = DB.ReadEntry("K-lines", CStr(i), "")
    Klines.Add Kline, Kline
Next i
'Olines
Dim OLineCount As Long
OLineCount = DB.ReadEntry("O-Lines", "Count", "0")
For i = 1 To OLineCount
    'ReDim Preserve Olines(i)
    Olines(i).UserName = DB.ReadEntry("O-Line " & i, "UserName", "")
    Olines(i).Password = DB.ReadEntry("O-Line " & i, "Password", "")
    Olines(i).Mask = DB.ReadEntry("O-Line " & i, "Mask", "")
    Olines(i).InUse = True
Next i
SendSvrMsg "Server has rehashed on the request of " & Nick
End Sub

Public Sub WriteHash()
Dim DB As New clsDatabase, i As Long, Kline As String, x As Long
DB.FileName = App.Path & "\.conf"
'Server Settings
DB.WriteEntry "General Settings", "Servername", ServerName
DB.WriteEntry "General Settings", "Description", ServerDesc
DB.WriteEntry "General Settings", "Port", CStr(frmMain.wsock(0).LocalPort)
DB.WriteEntry "General Settings", "MaxUsers", "100"
DB.WriteEntry "General Settings", "MaxChannels", "100"
'Channel Defaults
DB.WriteEntry "Channel Defaults", "Topic", DefTopic
'Standard Usermodes during log in
DB.WriteEntry "Default User Settings", "UserModes", DefUserModes
DB.WriteEntry "Default User Settings", "Default Quit Msg", DefQuit
'Admin
DB.WriteEntry "Admin", "Name", AdminName
DB.WriteEntry "Admin", "Email", AdminEmail
'Klines
DB.WriteEntry "K-lines", "Count", Klines.Count
For i = 1 To Klines.Count
    DB.WriteEntry "K-lines", CStr(i), Klines(i)
Next i
'Olines
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        DB.WriteEntry "O-Line " & CStr(i), "UserName", Olines(i).UserName
        DB.WriteEntry "O-Line " & CStr(i), "Password", Olines(i).Password
        DB.WriteEntry "O-Line " & CStr(i), "Mask", Olines(i).Mask
        x = x + 1
    End If
Next i
DB.WriteEntry "O-Lines", "Count", CStr(x)
End Sub

Public Function GetPercent(Base As Long, Cur As Long) As Long
Dim x As Long, z As Long, p2 As Long, BaseVal As Long, PercVal As Long, Percent As Long, max
If Cur = 0 Then
    GetPercent = 0
    Exit Function
End If
x = Base
z = Cur
p2 = x / 100
BaseVal = x / p2
PercVal = z / p2
Percent = PercVal / BaseVal * 100
GetPercent = Percent
End Function

Public Sub SaveOlines()
Dim DB As New clsDatabase, i As Long, x As Long
DB.FileName = App.Path & "\.conf"
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        DB.WriteEntry "O-Line " & CStr(i), "UserName", Olines(i).UserName
        DB.WriteEntry "O-Line " & CStr(i), "Password", Olines(i).Password
        DB.WriteEntry "O-Line " & CStr(i), "Mask", Olines(i).Mask
        x = x + 1
    End If
Next i
DB.WriteEntry "O-Lines", "Count", CStr(x)
End Sub

Public Function GetFreeOLine() As Long
Dim i As Long
For i = 1 To UBound(Olines)
    If Not Olines(i).InUse Then
        GetFreeOLine = i
        Exit Function
    End If
Next i
End Function

Public Function HasOline(Nick As String, Mask As String) As Boolean
Dim i As Long, UIndex As Long
UIndex = NickToObject(Nick).Index
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        If Users(UIndex).DNS Like Olines(i).Mask Then
            HasOline = True
            Exit Function
        End If
    End If
Next i
End Function

Public Function GetOline(DNS As String) As Long
Dim i As Long
For i = 1 To UBound(Olines)
    If Olines(i).InUse Then
        If DNS Like Olines(i).Mask Then
            GetOline = i
            Exit Function
        End If
    End If
Next i
End Function

Public Sub SendSvrMsg(msg As String)
Dim i As Long
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        If Users(i).IsMode("s") Then SendNotice Users(i).Nick, "*** Notice -- " & msg, ServerName
    End If
Next i
End Sub

Public Sub Wall(msg As String, Index As Integer)
WallOps msg, Index
End Sub

Public Sub WallOps(msg As String, Index As Integer)
Dim x As Long
For x = 1 To UBound(Users)
    If Not Users(x) Is Nothing Then
        If Users(x).IsMode("o") Or Users(x).IsMode("w") Then
            SendWsock x, ":" & Users(Index).Nick & "!" & Users(Index).Ident & "@" & Users(Index).DNS & " WALLOPS " & msg
        End If
    End If
Next x
End Sub
