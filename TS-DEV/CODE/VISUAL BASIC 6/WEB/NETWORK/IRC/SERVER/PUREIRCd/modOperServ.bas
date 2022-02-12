Attribute VB_Name = "modOperServ"
Option Explicit
Private DB As New clsDatabase
Private FS As New FileSystemObject

Public Sub ParseOSCmd(Cmd As String, Index As Long)
On Error GoTo parseerr
Dim msg As String, CMDStr As String, lcmd As Integer, arg1 As String, arg2 As String, cmd2 As String
Dim User As clsUser
Set User = Users(Index)
msg = Replace(Cmd, "OS ", "")
If Not InStr(1, msg, " ") <> 0 Then
    CMDStr = msg
Else
    CMDStr = (Mid(msg, 1, InStr(1, msg, " ") - 1))
End If
msg = Replace(msg, CMDStr & " ", "")
Select Case LCase(CMDStr)
    Case "stats"
        lcmd = 1
    Case "addstaff"
        lcmd = 2
    Case "delstaff"
        lcmd = 3
    Case "kill"
        lcmd = 4
    Case "akill"
        lcmd = 5
    Case "clear"
        lcmd = 6
    Case "global"
        lcmd = 7
    Case "logonnews"
        lcmd = 8
    Case "help"
        If msg = "" Or msg = "help" Then
            SendNotice User.Nick, "OperServ Commands", "OperServ"
            SendNotice User.Nick, "STATS (stats)", "OperServ"
            SendNotice User.Nick, "ADDSTAFF (addstaff [Nick] )", "OperServ"
            SendNotice User.Nick, "DELSTAFF (delstaff [Nick] )", "OperServ"
            SendNotice User.Nick, "KILL (kill [nick] [reason] )", "OperServ"
            SendNotice User.Nick, "AKILL (akill [nick] )", "OperServ"
            SendNotice User.Nick, "CLEAR (clear [channel] )", "OperServ"
            SendNotice User.Nick, "GLOBAL (global [message] )", "OperServ"
        Else
            Select Case LCase(msg)
                Case "identify"
                    SendNotice User.Nick, "Identify (identify [password] )", "OperServ"
                Case "drop"
                    SendNotice User.Nick, "Drop (drop [password] )", "OperServ"
                Case "register"
                    SendNotice User.Nick, "Register (register [password] [email] )", "OperServ"
                Case "kill"
                    SendNotice User.Nick, "Kill (kill [nick] [password] )", "OperServ"
                Case "info"
                    SendNotice User.Nick, "Info (info [nick] )", "OperServ"
                Case "changeinfo"
                    SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "OperServ"
            End Select
        End If
    Case Else
        SendNotice User.Nick, "Command Unknown", "OperServ"
End Select
Exit Sub
parseerr:
Select Case lcmd
    Case 1
        SendNotice User.Nick, "Identify (identify [password] )", "OperServ"
    Case 2
        SendNotice User.Nick, "Drop (drop [password] )", "OperServ"
    Case 3
        SendNotice User.Nick, "Register (register [password] [email] )", "OperServ"
    Case 4
        SendNotice User.Nick, "Kill (kill [nick] [password] )", "OperServ"
    Case 5
        SendNotice User.Nick, "Info (info [nick] )", "OperServ"
    Case 6
        SendNotice User.Nick, "ChangeInfo (changeinfo [newpass] [newemail] )", "OperServ"
    Case Else
        SendNotice User.Nick, "Unknown Command or missing parameters", "OperServ"
End Select
End Sub
