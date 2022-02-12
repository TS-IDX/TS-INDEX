Attribute VB_Name = "modMemoServ"
Option Explicit
Private DB As New clsDatabase
Private FS As New FileSystemObject

Public Sub ParseMSCmd(Cmd As String, Index As Long)
On Error GoTo parseerr
Dim msg As String, CMDStr As String, lcmd As Integer, arg1 As String, arg2 As String, cmd2 As String, i As Long, memocount As Long
Dim User As clsUser
Set User = Users(Index)
msg = Replace(Cmd, "MS ", "")
On Local Error Resume Next
CMDStr = (Mid(msg, 1, InStr(1, msg, " ") - 1))
If CMDStr = "" Then CMDStr = msg
msg = Replace(msg, CMDStr & " ", "")
Select Case LCase(CMDStr)
    Case "read"
        Dim RE As Long
        RE = Replace(Cmd, "MS read ", "", , , vbTextCompare)
        For i = 1 To Memos.Count
            If Memos(i).Target = User.Nick Then memocount = memocount + 1
            If memocount = RE Then
                SendNotice User.Nick, "Memo " & RE & " - from " & Memos(i).Source, "MemoServ"
                SendNotice User.Nick, "Memo Text:", "MemoServ"
                SendNotice User.Nick, Memos(i).Message, "MemoServ"
                Memos(i).Read = True
                Exit Sub
            End If
        Next i
    Case "send"
        lcmd = 2
        arg1 = Replace(Cmd, "MS send ", "", , , vbTextCompare)
        arg2 = arg1
        arg1 = Mid(arg1, InStr(1, arg1, " ") + 1)
        If Len(arg1) > 512 Then
            SendNotice Users(Index).Nick, "Memo size exceeds maximum of 512 Characters", "MemoServ"
            Exit Sub
        End If
        arg2 = Replace(arg2, " " & arg1, "")
        If Not IsRegistered(arg2) Then
            SendNotice Users(Index).Nick, "Cannot deliver memo, " & arg2 & " is not a registered nickname", "MemoServ"
            Exit Sub
        End If
        SendMemo arg2, arg1, Users(Index).Nick
        SendNotice Users(Index).Nick, "Memo has been successfully stored and can be viewed by " & arg2, "MemoServ"
    Case "list"
        lcmd = 3
        If Memos.Count = 0 Then
            SendNotice User.Nick, "You dont have any Memos", "MemoServ"
            Exit Sub
        End If
        Dim MemoList As Long, TableSent As Boolean
        For MemoList = 1 To Memos.Count
            If Memos(MemoList).Target = User.Nick Then
                If Not TableSent Then
                    SendNotice User.Nick, SizeString("Index", 10) & "From Nick", "MemoServ"
                    TableSent = True
                End If
                SendNotice User.Nick, SizeString(CStr(MemoList), 10) & SizeString(Memos(MemoList).Source, 25), "MemoServ"
            End If
        Next MemoList
    Case "del"
        RE = Replace(Cmd, "MS del ", "", , , vbTextCompare)
        For i = 1 To Memos.Count
            If Memos(i).Target = User.Nick Then memocount = memocount + 1
            If memocount = RE Then
                FS.DeleteFile App.Path & "\memos\" & Memos(i).MemoID
                Memos.Remove i
                SendNotice User.Nick, "Memo " & RE & " has been deleted", "MemoServ"
                Exit Sub
            End If
        Next i
        SendNotice User.Nick, "Memo " & RE & " not found!", "MemoServ"
    Case "help"
        If msg = "" Or msg = "help" Then
            SendNotice User.Nick, "MemoServ Commands", "MemoServ"
            SendNotice User.Nick, "READ [Memo index]", "MemoServ"
            SendNotice User.Nick, "SEND [User] [Memo]", "MemoServ"
            SendNotice User.Nick, "LIST", "MemoServ"
            SendNotice User.Nick, "DELETE [Memo index]", "MemoServ"
        Else
            Select Case LCase(msg)
                Case "read"
                    SendNotice User.Nick, "Read [Memo index]", "MemoServ"
                Case "send"
                    SendNotice User.Nick, "Send [User] [Memo]", "MemoServ"
                Case "list"
                    SendNotice User.Nick, "List", "MemoServ"
                Case "delete"
                    SendNotice User.Nick, "Delete [Memo index]", "MemoServ"
            End Select
        End If
    Case Else
        SendNotice User.Nick, "Command Unknown", "MemoServ"
End Select
Exit Sub
parseerr:
Select Case lcmd
    Case 1
        SendNotice User.Nick, "Read (read [index] )", "MemoServ"
        SendNotice User.Nick, "Reads memo [index]. Get the [index] value from /ms list", "MemoServ"
    Case 2
        SendNotice User.Nick, "Send (send [Nick] [Message])", "MemoServ"
        SendNotice User.Nick, "Sends a memo to [nick]", "MemoServ"
    Case 3
        SendNotice User.Nick, "List (list)", "MemoServ"
    Case 4
        SendNotice User.Nick, "Delete (del [index] )", "MemoServ"
        SendNotice User.Nick, "Delete Memo [index]. Get the [index] value from /ms list", "MemoServ"
    Case Else
        SendNotice User.Nick, "Unknown Command or missing parameters", "MemoServ"
End Select
End Sub

Private Function IsRegistered(Nick As String) As Boolean
If FS.FileExists(App.Path & "\users\" & Nick & ".usr") Then IsRegistered = True
End Function

Private Sub SendMemo(Target As String, memo As String, Source As String)
Dim Count As String, Rand As Long
Rand = GetRand
With FS.OpenTextFile(App.Path & "\memos\" & Rand & ".memo", ForWriting, True)
    .WriteLine Target
    .WriteLine Source
    .WriteLine memo
    .WriteLine "0"
End With
Memos.Add Target, Source, memo
Memos(Memos.Count).Read = False
If NickInUse(Target) And NickToObject(Target).Identified Then SendNotice Target, "You have a new Memo, type '/msg MemoServ list'", "MemoServ"
End Sub

