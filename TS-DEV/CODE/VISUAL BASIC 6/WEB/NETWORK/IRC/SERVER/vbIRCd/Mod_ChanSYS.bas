Attribute VB_Name = "Mod_ChanSYS"
' vbIRCd - Software/Code is an IRCd(Internet Relay Chat Daemon) used to host IRC sessions
' Copyright (C) 2001  Nathan Martin
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' To Contact the author e-mail TRON at tron@ircd-net.org
' * Note: There is no post mail contact information due to that it can be abused...
'
'
'--------Channels Buffers---------
Type b_ChannelSys           '* = NEW!
    iChan As New Collection '1Channel's Name
    iCUsers As New Collection '2Channel Users
    iCUsersM As New Collection '3*Channel Users' Modes
    iCCreator As New Collection '4Channel Creator
    iCBans As New Collection '5Channel Bans
    iCBansSN As New Collection '6*Channel Bans Set by Nick
    iCBansTS As New Collection '7*Channel Bans' Time Stamp
    iCExcep As New Collection '8*Channel Exceptions
    iCExcepSN As New Collection '9*Channel Exception Set by Nick
    iCExcepTS As New Collection '10*Channel Exceptions' Time Stamp
    iCLinked As New Collection '11*Channel Link
    iCModes As New Collection '12Channel Flag Modes
    iCKey As New Collection '13Channel Key
    iCTopic As New Collection '14Channel Topic
    iCULimit As New Collection '15Channel User Limit
    iCCDate As New Collection '16Channel Creation Date
    iCTopicD As New Collection '17Channel Topic Date
    iCTopicC As New Collection '18Channel Topic Changer
End Type
Global iChanSys As b_ChannelSys
'--------End Channels Buffers-----

Sub uINVITE(Channel As String, User As String, Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim X As Integer
Dim Y As Integer
Dim Q As Integer
With iChanSys
    
    
End With
End Sub


Sub uJOIN(User As String, Index As Integer, Text As String)
On Error Resume Next
Dim sChan As String
Dim sKey As String
Dim sUsers As String
Dim sUsersM As String
Dim sDate As String
Dim TmpNumber, TmpNumber2 As Integer
Dim TmpText, TmpText2 As String
Dim DFY As Boolean
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
    
    With iChanSys
    X = InStr(1, Text, " ")
    Y = InStr(X + 1, Text, " ")
    sChan = Mid$(Text, 1, X - 1)
    If Not Y = 0 Then sKey = Mid$(Text, X + 1, Y - 1 - X)
    If Not Left$(sChan, 1) = "#" Then
        SendData Index, ":" & sServer & " 476 " & User & " " & sChan & " :" & sConvertText(ERR_476, Index) & CRLF
        Exit Sub
    End If
    
    For X = 1 To .iChan.Count
        If LCase(.iChan(X)) = LCase(sChan) Then
            Q = X
            Exit For
        End If
    Next X
    
    If Not Q = 0 Then
        TmpText = Replace(.iCUsers(X), " ", "", 1, Len(.iCUsers(X)), vbTextCompare)
        TmpNumber = Len(.iCUsers(X)) - Len(TmpText)
        TmpText2 = Replace(.iCExcep(Q), " ", "", 1, Len(.iCExcep(Q)), vbTextCompare)
        TmpNumber2 = Len(.iCExcep(Q)) - Len(TmpText)
            
        X = InStr(1, .iCUsers(Q), "@" & iUser(Index) & " ")
        If Not X = 0 Then Exit Sub
        X = InStr(1, .iCUsers(Q), "%" & iUser(Index) & " ")
        If Not X = 0 Then Exit Sub
        X = InStr(1, .iCUsers(Q), "+" & iUser(Index) & " ")
        If Not X = 0 Then Exit Sub
        X = InStr(1, .iCUsers(Q), " " & iUser(Index) & " ")
        If Not X = 0 Then Exit Sub
        X = InStr(1, .iCUsers(Q), iUser(Index) & " ")
        If Not X = 0 And X = 1 Then Exit Sub
        
        If Not .iCKey(Q) = "" Then If Not LCase(.iCKey(Q)) = LCase(sKey) Then SendData Index, ":" & sServer & " 475 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_475, Index) & CRLF: Exit Sub
        If Not .iCULimit(Q) = 0 Then
            If TmpNumber + 1 > .iCULimit(Q) Then
                If .iCLinked(Q) = "" Then
                    SendData Index, ":" & sServer & " 471 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_471, Index) & CRLF: Exit Sub
                Else
                    
                End If
            End If
        End If
        
        TmpLoad = .iCBans(Q)
        TmpLoad2 = .iCExcep(Q)
        DFY = False
        For Y = 1 To TmpNumber
            X = InStr(1, TmpLoad, " ")
            TmpText = Mid$(TmpLoad, 1, X - 1)
            TmpLoad = Mid$(TmpLoad, X + 1)
            If LCase(iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index)) Like LCase(TmpText) Then DFY = True
            If LCase(iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)) Like LCase(TmpText) Then DFY = True
            If DFY = True Then
                For Z = 1 To TmpNumber2
                    X = InStr(1, TmpLoad2, " ")
                    TmpText2 = Mid$(TmpLoad2, 1, X - 1)
                    TmpLoad2 = Mid$(TmpLoad2, X + 1)
                    If LCase(iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index)) Like LCase(TmpText2) Then DFY = False: Exit For
                    If LCase(iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)) Like LCase(TmpText2) Then DFY = False: Exit For
                Next Z
                If DFY = True Then
                    SendData Index, ":" & sServer & " 474 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_474, Index) & CRLF
                    Exit Sub
                End If
            End If
        Next Y
        
        If Not InStr(1, .iCModes(Q), "R") = 0 Then If InStr(1, iModes(Index), "r") = 0 Then SendData Index, ":" & sServer & " 477 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_477, Index) & CRLF: Exit Sub
        If Not InStr(1, .iCModes(Q), "A") = 0 Then If iUserLevel(Index) > 0 Then SendData Index, ":" & sServer & " 520 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_520, Index) & CRLF: Exit Sub
        If Not InStr(1, .iCModes(Q), "O") = 0 Then If iUserLevel(Index) > 2 Then SendData Index, ":" & sServer & " 519 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_519, Index) & CRLF: Exit Sub
        If Not InStr(1, .iCModes(Q), "H") = 0 Then If Not InStr(1, iModes(Index), "I") = 0 Then SendData Index, ":" & sServer & " 459 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_459, Index) & CRLF: Exit Sub
        
        Q = cChanChange(.iChan(Q), .iCModes(Q), iUser(Index) & " " & .iCUsers(Q), iUser(Index) & "$ " & .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepTS(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        
        If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :" & sChan & " - " & .iCCDate(Q) & " - " & .iCModes(Q) & CRLF
        JOINNotify .iChan(Q), Index
        
        SendData2 Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " JOIN :" & .iChan(Q) & CRLF
        If Not .iCTopic(Q) = "" Then
            SendData2 Index, ":" & sServer & " 332 " & iUser(Index) & " " & .iChan(Q) & " :" & .iCTopic(Q) & CRLF & _
                             ":" & sServer & " 333 " & iUser(Index) & " " & .iChan(Q) & " " & .iCTopicC(Q) & " " & .iCTopicD(Q) & CRLF
        Else
            SendData2 Index, ":" & sServer & " 331 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(RPL_331, Index) & CRLF
        End If
        SendData2 Index, ":" & sServer & " 353 " & iUser(Index) & " = " & .iChan(Q) & " :" & .iCUsers(Q) & CRLF & _
                        ":" & sServer & " 366 " & iUser(Index) & " " & sChan & " :" & sConvertText(RPL_366, Index) & CRLF
        iChan(Index) = iChan(Index) & .iChan(Q) & " "
    Else
        If sNameCheck(sChan) = True And sFOCN = 1 Then
            SendData Index, ":" & sServer & " 403 " & iUser(Index) & " " & sChan & " :Cannot create channel that uses words not allowed to be used in network channels" & CRLF
            Exit Sub
        End If
        Select Case iWhoCCC
            Case 1: If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 553 " & iUser(Index) & " " & sChan & " :Channel Creation is not allowed-  You don't have high enough access." & CRLF: Exit Sub
            Case 2: If iUserLevel(Index) < 2 Then SendData Index, ":" & sServer & " 553 " & iUser(Index) & " " & sChan & " :Channel Creation is not allowed-  You don't have high enough access." & CRLF: Exit Sub
            Case 3: If iUserLevel(Index) < 3 Then SendData Index, ":" & sServer & " 553 " & iUser(Index) & " " & sChan & " :Channel Creation is not allowed-  You don't have high enough access." & CRLF: Exit Sub
            Case 4: If iUserLevel(Index) < 4 Then SendData Index, ":" & sServer & " 553 " & iUser(Index) & " " & sChan & " :Channel Creation is not allowed-  You don't have high enough access." & CRLF: Exit Sub
        End Select
        .iChan.Add sChan
        .iCCDate.Add GetTime
        .iCBans.Add ""
        .iCBansSN.Add ""
        .iCBansTS.Add ""
        .iCExcep.Add ""
        .iCExcepSN.Add ""
        .iCExcepTS.Add ""
        .iCKey.Add ""
        .iCModes.Add iSMOCC
        .iCTopic.Add ""
        .iCTopicC.Add ""
        .iCTopicD.Add ""
        .iCCreator.Add iUser(Index)
        .iCULimit.Add "0"
        .iCUsers.Add "@" & iUser(Index) & " "
        .iCUsersM.Add iUser(Index) & "$oq "
        .iCLinked.Add ""
        frmMain.lbl_CC = frmMain.lbl_CC + 1
        For X = 1 To .iChan.Count
            If LCase(sChan) = LCase(.iChan(X)) Then
                Q = X
                If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :" & sChan & " - " & .iCCDate(Q) & " - " & .iCModes(Q) & CRLF
                Exit For
            End If
        Next X
        
        If Q = 0 Then
            LogIt "Faild 2 Find " & sChan & " N JOIN Buffer 4 " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)
            If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: New Channel not found in Buffer!  Error has been logged." & CRLF
            Exit Sub
        End If
        SendData Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " JOIN :" & sChan & CRLF & _
                        ":" & sServer & " 353 " & iUser(Index) & " = " & sChan & " :@" & iUser(Index) & CRLF & _
                        ":" & sServer & " 366 " & iUser(Index) & " " & sChan & " :" & sConvertText(RPL_366, Index) & CRLF & _
                        ":" & sServer & " 331 " & iUser(Index) & " " & sChan & " :" & sConvertText(RPL_331, Index) & CRLF & _
                        ":" & sServer & " MODE " & sChan & " +q " & iUser(Index) & CRLF

        iChan(Index) = iChan(Index) & "@" & sChan & " "
    End If
    frmMain.lbl_CC = .iChan.Count
    End With
End Sub

Sub uPART(User As String, Index As Integer, Text As String)
On Error Resume Next
Dim sChan As String
Dim sPartMsg As String
Dim sUsers As String
Dim sUsersM As String
Dim sDate As String
Dim TmpText As String
Dim TmpText2 As String
Dim sNick As String
Dim DFY As Boolean
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
    sNick = iUser(Index)
    With iChanSys
    X = InStr(1, Text, " ")
    Y = InStr(1, Text, " :")
    sChan = Mid$(Text, 1, X - 1)
    
    If Not Y = 0 Then sPartMsg = Mid$(Text, Y + 2, Len(Text) - 2 - Y)
    If Not Left$(sChan, 1) = "#" Then
        SendData Index, ":" & sServer & " 461 " & User & " PART :" & sConvertText(ERR_461, Index) & CRLF
        Exit Sub
    End If
    
    For X = 1 To .iChan.Count
        If LCase(sChan) = LCase(.iChan(X)) Then
            Q = X
            DFY = True
            Exit For
        End If
    Next X
    
    If DFY = True Then
        sUsers = .iCUsers(Q)
        sUsersM = .iCUsersM(Q)
        
        X = InStr(1, LCase(sUsers), LCase(sNick & " "))
        If X = 1 Then
            TmpText2 = Mid$(sUsers, X)
        Else
            TmpText2 = Mid$(sUsers, X - 1)
        End If
        
        If X = 0 Then
            SendData Index, ":" & sServer & " 442 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_442, Index) & CRLF
            Exit Sub
        Else
            Select Case Left$(TmpText2, 1)
                Case "@": sUsers = ReplaceStr(sUsers, "@" & sNick & " ", "")
                Case "%": sUsers = ReplaceStr(sUsers, "%" & sNick & " ", "")
                Case "+": sUsers = ReplaceStr(sUsers, "+" & sNick & " ", "")
                Case " ": sUsers = ReplaceStr(sUsers, " " & sNick & " ", " ")
                Case Else: sUsers = Mid$(TmpText2, Len(sNick) + 2)
            End Select
            
            Z = InStr(1, iChan(Index), sChan)
            If Z = 1 Then
                iChan(Index) = Mid$(iChan(Index), Len(sChan) + 2)
            ElseIf Z = 2 Then
                    iChan(Index) = Mid$(iChan(Index), Len(sChan) + 3)
                Else
                    TmpText = Mid$(iChan(Index), Z - 2, 2)
                    
                    Select Case TmpText
                        Case "~*": iChan(Index) = ReplaceStr(iChan(Index), "~*" & sChan & " ", "")
                        Case "~@": iChan(Index) = ReplaceStr(iChan(Index), "~@" & sChan & " ", "")
                        Case "~%": iChan(Index) = ReplaceStr(iChan(Index), "~%" & sChan & " ", "")
                        Case "~+": iChan(Index) = ReplaceStr(iChan(Index), "~+" & sChan & " ", "")
                        Case " *": iChan(Index) = ReplaceStr(iChan(Index), "*" & sChan & " ", "")
                        Case " @": iChan(Index) = ReplaceStr(iChan(Index), "@" & sChan & " ", "")
                        Case " %": iChan(Index) = ReplaceStr(iChan(Index), "%" & sChan & " ", "")
                        Case " +": iChan(Index) = ReplaceStr(iChan(Index), "+" & sChan & " ", "")
                        Case Else: iChan(Index) = ReplaceStr(iChan(Index), " " & sChan & " ", " ")
                    End Select
            End If
            
            TmpFlags = GetUserCFlags(.iChan(Q), sNick)
            sUsersM = ReplaceStr(" " & sUsersM, " " & sNick & "$" & TmpFlags & " ", " ")
            sUsersM = Mid$(sUsersM, 2)
            
            TmpText2 = Replace(.iCUsers(Q), " ", "", 1, Len(.iCUsers(Q)), vbTextCompare)
            If Len(.iCUsers(Q)) - Len(TmpText2) = 1 Then cRemoveChannel .iChan(Q): Exit Sub
        End If
        
        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        
        If Q = 0 Then
            LogIt "Faild 2 Find " & sChan & " N PART Buffer 4 " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)
            If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Channel not found in Buffer!  Error has been logged." & CRLF
            Exit Sub
        End If
        
        PARTNotify .iChan(Q), sPartMsg, Index
        SendData Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " PART " & .iChan(Q) & CRLF
        
        'If .iCUsers(Q) = "" Then
        '    cRemoveChannel .iChan(Q)
        'End If
        frmMain.lbl_CC = .iChan.Count
    Else
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & sChan & " No such nick/channel" & CRLF
    End If
    
    End With
End Sub

Sub uQUIT(UserHostMask As String, Channels As String, QuitMsg As String)
On Error Resume Next
Dim sChan As String
Dim sPartMsg As String
Dim sUsers As String
Dim sUsersM As String
Dim sNick As String
Dim sDate As String
Dim TmpText As String
Dim TmpText2 As String
Dim TmdData As String
Dim NotToUsers As String
Dim DFY As Boolean
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
    Z = InStr(1, UserHostMask, "!")
    sNick = Mid$(UserHostMask, 1, Z - 1)
    sChans = Channels
    
ReScan:
    If sChans = "" Then Exit Sub
    X = InStr(1, sChans, " ")
    sChan = Mid$(sChans, 1, X - 1)
    sChans = Mid$(sChans, X + 1)
    Select Case Left$(sChan, 1)
        Case "~": sChan = Mid$(sChan, 2)
        Case "@": sChan = Mid$(sChan, 2)
        Case "%": sChan = Mid$(sChan, 2)
        Case "+": sChan = Mid$(sChan, 2)
        Case "^": sChan = Mid$(sChan, 2)
        Case "*": sChan = Mid$(sChan, 2)
    End Select
    
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(.iChan(X)) = LCase(sChan) Then
            Q = X
            DFY = True
            Exit For
        End If
    Next X
    
    If DFY = True Then
        sUsers = .iCUsers(Q)
        sUsersM = .iCUsersM(Q)
        
        TmpText2 = Replace(.iCUsers(Q), " ", "", 1, Len(.iCUsers(Q)), vbTextCompare)
        If Len(.iCUsers(Q)) - Len(TmpText2) = 1 Then cRemoveChannel .iChan(Q): GoTo ReScan
        
        X = InStr(1, LCase(sUsers), LCase(sNick & " "))
        If X = 1 Then
            TmpText2 = Mid$(sUsers, X)
        Else
            TmpText2 = Mid$(sUsers, X - 1)
        End If
        If X = 0 Then
            LogIt "Faild Find 4 " & UserHostMask & " QUIT 4 " & iChan(Q)
            Exit Sub
        Else
            Select Case Left$(TmpText2, 1)
                Case "@": sUsers = ReplaceStr(sUsers, "@" & sNick & " ", "")
                Case "%": sUsers = ReplaceStr(sUsers, "%" & sNick & " ", "")
                Case "+": sUsers = ReplaceStr(sUsers, "+" & sNick & " ", "")
                Case " ": sUsers = ReplaceStr(sUsers, " " & sNick & " ", " ")
                Case Else: sUsers = Mid$(TmpText2, Len(sNick) + 2)
            End Select
            
            TmpFlags = GetUserCFlags(.iChan(Q), sNick)
            sUsersM = ReplaceStr(" " & sUsersM, " " & sNick & "$" & TmpFlags & " ", " ")
            sUsersM = Mid$(sUsersM, 2)
        End If
        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        NotToUsers = QuitNotify(.iChan(Q), QuitMsg, UserHostMask, NotToUsers)

        'If .iCUsers(Q) = "" Then
        '    cRemoveChannel .iChan(Q)
        'End If
    Else
        LogIt "Faild 2 Find " & sChan & " 4 :" & UserHostMask & " QUIT"
    End If
GoTo ReScan
    End With
End Sub

Sub uNickChange(UserHostMask As String, Channels As String, NewNick As String)
On Error Resume Next
Dim sChan As String
Dim Index As Integer
Dim sPartMsg As String
Dim sUsers As String
Dim sDate As String
Dim sUsersM As String
Dim sNick As String
Dim NotToUsers As String
Dim TmpText As String
Dim DFY As Boolean
Dim sUCS As String
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
    sChans = Channels
    X = InStr(1, UserHostMask, "!")
    sNick = Mid$(UserHostMask, 1, X - 1)
    For X = 1 To iUserMax
        If LCase(sNick) = iUser(X) Then
            Index = X
            Exit For
        End If
    Next X
    
ReScan:
    If sChans = "" Then Exit Sub
    X = InStr(1, sChans, " ")
    sChan = Mid$(sChans, 1, X - 1)
    sChans = Mid$(sChans, X + 1)
    If Left$(sChan, 1) = "~" Then sChan = Mid$(sChan, 2)
    If Left$(sChan, 1) = "@" Then sChan = Mid$(sChan, 2)
    If Left$(sChan, 1) = "+" Then sChan = Mid$(sChan, 2)
    If Left$(sChan, 1) = "^" Then sChan = Mid$(sChan, 2)
    If Left$(sChan, 1) = "*" Then sChan = Mid$(sChan, 2)
    
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(.iChan(X)) = LCase(sChan) Then
            Q = X
            DFY = True
            Exit For
        End If
    Next X
    
    If DFY = True Then
        sUsers = .iCUsers(Q)
        sUsersM = .iCUsersM(Q)
        
        X = InStr(1, LCase(sUsers), LCase(sNick & " "))
        If X = 1 Then
            TmpText = Mid$(sUsers, X)
        Else
            TmpText = Mid$(sUsers, X - 1)
        End If
        
        If X = 0 Then
            If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Your not in Channel Members Record!  Error has been logged." & CRLF
            LogIt "Faild NickFind N " & .iChan(Q) & " 4 :" & UserHostMask & " NICK " & NewNick
            Exit Sub
        Else
            Select Case Left$(TmpText, 1)
                Case "@": sUsers = ReplaceStr(sUsers, "@" & sNick & " ", "@" & NewNick & " ")
                Case "%": sUsers = ReplaceStr(sUsers, "%" & sNick & " ", "%" & NewNick & " ")
                Case "+": sUsers = ReplaceStr(sUsers, "+" & sNick & " ", "+" & NewNick & " ")
                Case " ": sUsers = ReplaceStr(sUsers, " " & sNick & " ", " " & NewNick & " ")
                Case Else: sUsers = Mid$(sUsers, Len(sNick) + 2): sUsers = NewNick & " " & sUsers
            End Select
            TmpFlags = GetUserCFlags(.iChan(Q), sNick)
            sUsersM = ReplaceStr(" " & sUsersM, " " & sNick & "$" & TmpFlags & " ", " " & NewNick & "$" & TmpFlags & " ")
            sUsersM = Mid$(sUsersM, 2)
        End If
        
        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        
        NotToUsers = NCNotify(.iChan(Q), NewNick, UserHostMask, NotToUsers)
    Else
        LogIt "Faild 2 Find " & sChan & " 4 :" & UserHostMask & " NICK " & NewNick
        If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Channel not found in Buffer!  Error has been logged." & CRLF
    End If
GoTo ReScan
    End With
End Sub

Sub cMODE(Index As Integer, Chan As String, Optional Modes As String, Optional iValue As String)
On Error Resume Next
Dim TmpText As String
Dim TmpLoad As String
Dim TmpFlag As String
Dim TmpNumber As Integer
Dim flag As String
Dim sUsers As String
Dim sUsersM As String
Dim cUserLevel As Integer
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
Dim Y As Integer
Dim DUI As Boolean
Dim AddedFlags As Boolean
Dim SFlags As String
Dim SFlagValue As String
Dim AddFlag As Boolean
    AddFlag = True
    AddedFlags = True
    iValue = iValue & " "
    
    With iChanSys
        For X = 1 To .iChan.Count
            If LCase(Chan) = LCase(.iChan(X)) Then
                Q = X
                Exit For
            End If
        Next X
         
    TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
    X = InStr(1, TmpText, "h")
    If Not X = 0 Then cUserLevel = 1
    X = InStr(1, TmpText, "o")
    If Not X = 0 Then cUserLevel = 2
    
    If Not Q = 0 Then
        If Modes = "" Then
            SendData Index, ":" & sServer & " 324 " & iUser(Index) & " " & .iChan(Q) & " +" & .iCModes(Q) & CRLF & _
                            ":" & sServer & " 329 " & iUser(Index) & " " & .iChan(Q) & " " & .iCCDate(Q) & CRLF
            Exit Sub
        End If
        If TmpText = "!" Then SendData Index, ":" & sServer & " 442 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_442, Index) & CRLF: Exit Sub
        
ReScan:
        DUI = False
        flag = Mid$(Modes, 1, 1)
        Modes = Mid$(Modes, 2)
        X = InStr(1, "-+lvhopsmntikrRcaqOALQbSeKVfHGCuzN", flag)
        If X = 0 Then SendData Index, ":" & sServer & " 472 " & iUser(Index) & " " & flag & " :" & sConvertText(ERR_472, Index) & CRLF: GoTo ReScan
        
        Z = InStr(1, "kelohvqabfL", flag)
        If Not Z = 0 Then DUI = True
        
        If DUI = True Then
            X = InStr(1, iValue, " ")
            TmpLoad = Mid$(iValue, 1, X - 1)
            iValue = Mid$(iValue, X + 1)
        End If
        
        If AddFlag = True Then
            If cUserLevel = 0 Then
                Select Case flag
                    Case "+"
                    Case "b": If Not TmpLoad = "" Then SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
                    Case "I": If Not TmpLoad = "" Then SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
                    Case "e": If Not TmpLoad = "" Then SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
                    Case Else
                        SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
                End Select
            End If
            If cUserLevel = 1 Then
                '+vmntibe
                Select Case flag
                    Case "+"
                    Case "-"
                    Case "v"
                    Case "m"
                    Case "n"
                    Case "t"
                    Case "i"
                    Case "b"
                    Case "e"
                    Case Else
                        SendData Index, ":" & sServer & " 460 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_460 & " " & flag, Index) & CRLF: GoTo ReScan
                End Select
            End If
            
            Select Case flag
                Case "O": If iUserLevel(Index) > 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** Only IRCops can set that mode" & CRLF: GoTo ReScan
                Case "A": If iUserLevel(Index) > 2 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** Only admins can set that mode" & CRLF: GoTo ReScan
                Case "H": If iUserLevel(Index) > 2 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** Only admins can set that mode" & CRLF: GoTo ReScan
            End Select
            
            Select Case flag
                'lvhopsmntikrRcaqOALQbSeKVfHGCuzN  <-- Known channel modes :)
                Case "l" '<number of max users>  Channel may hold at most <number> of users
                    If TmpLoad = "" Then GoTo ReScan
                    TmpNumber = Mid$(TmpLoad, 1, 4)
                    If TmpNumber = 0 Then
                        Modes = Modes & "-l"
                        GoTo ReScan
                    End If
                    
                    'X = InStr(1, .iCModes(Q), flag)
                    'If Not X = 0 Then GoTo ReScan
                    TmpText = .iCModes(Q) & flag
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), TmpLoad, .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "v" '<nickname>  Gives Voice to the user (May talk if chan is +m)
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "h")
                        If Not X = 0 Then TmpFlag = "%"
                        X = InStr(1, TmpText, "o")
                        If Not X = 0 Then TmpFlag = "@"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        If TmpFlag = "" Then sUsers = Mid$(ReplaceStr(" " & sUsers, " " & TmpFlag & TmpLoad & " ", " " & "+" & iUser(Y) & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "+v")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "h" '<nickname>  Gives HalfOp status to the user
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "v")
                        If Not X = 0 Then TmpFlag = "+"
                        X = InStr(1, TmpText, "o")
                        If Not X = 0 Then TmpFlag = "@"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        If Not TmpFlag = "@" Then sUsers = Mid$(ReplaceStr(" " & sUsers, " " & TmpFlag & TmpLoad & " ", " " & "%" & TmpLoad & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "+h")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "o" '<nickname>  Gives Operator status to the user
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "v")
                        If Not X = 0 Then TmpFlag = "+"
                        X = InStr(1, TmpText, "h")
                        If Not X = 0 Then TmpFlag = "%"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        sUsers = Mid$(ReplaceStr(" " & sUsers, " " & TmpFlag & TmpLoad & " ", " " & "@" & TmpLoad & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "+o")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "p" 'Private channel
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    X = InStr(1, .iCModes(Q), "s")
                    TmpText = .iCModes(Q)
                    If Not X = 0 Then TmpText = ReplaceStr(TmpText, "s", ""): SFlags = SFlags & "-s": AddedFlags = False
                    Q = cChanChange(.iChan(Q), TmpText & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "s" 'Secret channel
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    X = InStr(1, .iCModes(Q), "p")
                    TmpText = .iCModes(Q)
                    If Not X = 0 Then TmpText = ReplaceStr(TmpText, "p", ""): SFlags = SFlags & "-p": AddedFlags = False
                    Q = cChanChange(.iChan(Q), TmpText & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "m" 'Moderated channel, Only users with mode +voh can speak.
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "n" 'No messages from outside channel
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "t" 'Only Channel Operators may set the topic
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "i" 'Invite-only allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "I" 'Invited Users
                    If TmpLoad = "" Then
                        SendData Index, ":" & sServer & " 347 " & iUser(Index) & " " & .iChan(Q) & " :End of Channel Invite List" & CRLF
                        ':<Server> 346 <Nick> <Chan> <user!ident@host> <SetByNick> <SetTS>
                        ':<Server> 347 <Nick> <Chan> :End of Channel Invite List
                    Else
                        
                        
                    End If
                    
                Case "k" '<key>  Needs the Channel Key to join the channel
                    If TmpLoad = "" Then GoTo ReScan
                    
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then SendData Index, ":" & sServer & " 467 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_467, Index) & CRLF: GoTo ReScan
                    TmpText = .iCModes(Q) & flag
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), TmpLoad, .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "r" 'Channel is Registered
                    
                Case "R" 'Requires a Registered nickname to join the channel
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "c" 'Blocks messages with ANSI colour (ColourBlock).
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    X = InStr(1, .iCModes(Q), "S")
                    TmpText = .iCModes(Q)
                    If Not X = 0 Then TmpText = ReplaceStr(TmpText, "S", ""): SFlags = SFlags & "-S": AddedFlags = False
                    Q = cChanChange(.iChan(Q), TmpText & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                
                Case "a" '<nickname>  Gives protection to the user (No kick/drop)
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
                        If InStr(1, TmpText, "q") = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** Protected users can only be set by the channel owner." & CRLF: GoTo ReScan
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        sUsersM = .iCUsersM(Q)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "+a")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "q" '<nickname>  Channel owner
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
                        If InStr(1, TmpText, "q") = 0 Then SendData Index, ":" & sServer & " 468 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_468, Index) & CRLF: GoTo ReScan
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        sUsersM = .iCUsersM(Q)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "+q")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "O" 'IRC Operators Only
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "A" 'IRC Admins Only
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "L" '<chan2>  If +l is full, the next user will auto-join <chan2>
                    If TmpLoad = "" Then GoTo ReScan
                    
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then SendData Index, ":" & sServer & " 469 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_469, Index) & CRLF: GoTo ReScan
                    TmpText = .iCModes(Q) & flag
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), TmpLoad)
                    If AddedFlags = True Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "Q" 'No Kicking
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "b" '<nick!user@host>  Bans the nick!user@host from the channel
                    If TmpLoad = "" Then
                        TmpText = Replace(.iCBans(Q), " ", "", 1, Len(.iCBans(Q)), vbTextCompare)
                        TmpNumber = Len(.iCBans(Q)) - Len(TmpText)
                        TmpText = .iCBans(Q)
                        DFY = False
                        
                        For Z = 1 To TmpNumber
                            X = InStr(1, TmpText, " ")
                            TmpLoad = Mid$(TmpText, 1, X - 1)
                            TmpText = Mid$(TmpText, X + 1)
                            
                            X = InStr(1, " " & LCase(.iCBansTS(Q)), " " & LCase(TmpLoad) & "$")
                            Y = InStr(X, .iCBansTS(Q), " ")
                            sUsers = Mid$(.iCBansTS(Q), X, Y)
                            X = InStr(1, sUsers, "$")
                            sUsers = Mid$(sUsers, X + 1, Len(sUsers) - X - 1)
                            
                            X = InStr(1, " " & LCase(.iCBansSN(Q)), " " & LCase(TmpLoad) & "$")
                            Y = InStr(X, .iCBansSN(Q), " ")
                            sUsersM = Mid$(.iCBansSN(Q), X, Y)
                            X = InStr(1, sUsersM, "$")
                            sUsersM = Mid$(sUsersM, X + 1, Len(sUsersM) - X - 1)
                            
                            SendData Index, ":" & sServer & " 367 " & iUser(Index) & " " & .iChan(Q) & " " & TmpLoad & " " & sUsersM & " " & sUsers & CRLF
                        Next Z
                        SendData Index, ":" & sServer & " 368 " & iUser(Index) & " " & .iChan(Q) & " :End of Channel Ban List" & CRLF
                        ':<Server> 367 <Nick> <Chan> <user!ident@host> <BanSetByNick> <BanSetTS>
                        ':<Server> 368 <Nick> <Chan> :End of channel ban list
                    Else
                        If TmpLoad = "*@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*!" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*!*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*@" Then TmpLoad = "*!*@*"
                        If TmpLoad = "@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "!@" Then TmpLoad = "*!*@*"
                        If Not TmpLoad Like "*!*@*" Then GoTo ReScan
                        X = InStr(1, " " & .iCBans(Q), " " & TmpLoad & " ")
                        If Not X = 0 Then GoTo ReScan
                        
                        TmpText = Replace(.iCBans(Q), " ", "", 1, Len(.iCBans(Q)), vbTextCompare)
                        TmpNumber = Len(.iCBans(Q)) - Len(TmpText) + 1
                        If TmpNumber > 60 Then SendData Index, ":" & sServer & " 478 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_478, Index) & CRLF: GoTo ReScan
                        
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q) & TmpLoad & " ", .iCBansTS(Q) & TmpLoad & "$" & GetTime & " ", .iCBansSN(Q) & TmpLoad & "$" & iUser(Index) & " ", .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    End If
                    
                Case "S" 'Strip all incoming colours away
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    X = InStr(1, .iCModes(Q), "c")
                    TmpText = .iCModes(Q)
                    If Not X = 0 Then TmpText = ReplaceStr(TmpText, "c", ""): SFlags = SFlags & "-c": AddedFlags = False
                    Q = cChanChange(.iChan(Q), TmpText & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "e" 'Exception ban
                    If TmpLoad = "" Then
                        TmpText = Replace(.iCExcep(Q), " ", "", 1, Len(.iCExcep(Q)), vbTextCompare)
                        TmpNumber = Len(.iCExcep(Q)) - Len(TmpText)
                        TmpText = .iCExcep(Q)
                        DFY = False
                        
                        For Z = 1 To TmpNumber
                            X = InStr(1, TmpText, " ")
                            TmpLoad = Mid$(TmpText, 1, X - 1)
                            TmpText = Mid$(TmpText, X + 1)
                            
                            X = InStr(1, " " & LCase(.iCExcepTS(Q)), " " & LCase(TmpLoad) & "$")
                            Y = InStr(X, .iCExcepTS(Q), " ")
                            sUsers = Mid$(.iCExcepTS(Q), X, Y)
                            X = InStr(1, sUsers, "$")
                            sUsers = Mid$(sUsers, X + 1, Len(sUsers) - X - 1)
                            
                            X = InStr(1, " " & LCase(.iCExcepSN(Q)), " " & LCase(TmpLoad) & "$")
                            Y = InStr(X, .iCExcepSN(Q), " ")
                            sUsersM = Mid$(.iCExcepSN(Q), X, Y)
                            X = InStr(1, sUsersM, "$")
                            sUsersM = Mid$(sUsersM, X + 1, Len(sUsersM) - X - 1)
                            
                            SendData Index, ":" & sServer & " 348 " & iUser(Index) & " " & .iChan(Q) & " " & TmpLoad & " " & sUsersM & " " & sUsers & CRLF
                        Next Z
                        SendData Index, ":" & sServer & " 349 " & iUser(Index) & " " & .iChan(Q) & " :End of Channel Exception List" & CRLF
                        ':<Server> 348 <Nick> <Chan> <user!ident@host> <BanSetByNick> <BanSetTS>
                        ':<Server> 349 <Nick> <Chan> :End of Channel Exception List
                    Else
                        If TmpLoad = "*@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*!" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*!*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*@" Then TmpLoad = "*!*@*"
                        If TmpLoad = "@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "*@*" Then TmpLoad = "*!*@*"
                        If TmpLoad = "!@" Then TmpLoad = "*!*@*"
                        If Not TmpLoad Like "*!*@*" Then GoTo ReScan
                        X = InStr(1, " " & .iCExcep(Q), " " & TmpLoad & " ")
                        If Not X = 0 Then GoTo ReScan
                        
                        TmpText = Replace(.iCExcep(Q), " ", "", 1, Len(.iCExcep(Q)), vbTextCompare)
                        TmpNumber = Len(.iCExcep(Q)) - Len(TmpText) + 1
                        If TmpNumber > 60 Then SendData Index, ":" & sServer & " 478 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_478, Index) & CRLF: GoTo ReScan
                        
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q) & TmpLoad & " ", .iCExcepTS(Q) & TmpLoad & "$" & GetTime & " ", .iCExcepSN(Q) & TmpLoad & "$" & iUser(Index) & " ", .iCLinked(Q))
                        If AddedFlags = True Then
                            SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = True
                            SFlags = SFlags & "+" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    End If
                    
                Case "K" '/KNOCK is not allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "V" '/INVITE is not allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "f" '[*]<lines>:<seconds>  Flood protection
                    
                Case "H" 'No +I uMode Users
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "G" 'Swear Filter
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                
                Case "C" 'No CTCPs
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "u" '"Auditorium". Makes /NAMES and /WHO #channel only show Operators.
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "z" 'Only Clients on a Secure Connection (SSL) can join.
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "N" 'No Nickname changes are permitted in the channel.
                    X = InStr(1, .iCModes(Q), flag)
                    If Not X = 0 Then GoTo ReScan
                    Q = cChanChange(.iChan(Q), .iCModes(Q) & flag, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = True Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = True
                        SFlags = SFlags & "+" & flag
                    End If
                    
                Case "-"
                AddFlag = False
            End Select
        Else
            If cUserLevel = 0 Then SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
            Select Case flag
                Case "l" '<number of max users>  Channel may hold at most <number> of users
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpLoad = .iCULimit(Q)
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), "0", .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "v" '<nickname>  Gives Voice to the user (May talk if chan is +m)
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "h")
                        If Not X = 0 Then TmpFlag = "%"
                        X = InStr(1, TmpText, "o")
                        If Not X = 0 Then TmpFlag = "@"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        If TmpFlag = "" Then sUsers = Mid$(ReplaceStr(" " & sUsers, " " & "+" & TmpLoad & " ", " " & TmpFlag & TmpLoad & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "-v")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = False Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = False
                            SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    End If
                    
                Case "h" '<nickname>  Gives HalfOp status to the user
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "v")
                        If Not X = 0 Then TmpFlag = "+"
                        X = InStr(1, TmpText, "o")
                        If Not X = 0 Then TmpFlag = "@"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        If Not TmpFlag = "@" Then sUsers = Mid$(ReplaceStr(" " & sUsers, " " & "%" & TmpLoad & " ", " " & TmpFlag & TmpLoad & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "-h")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = False Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = False
                            SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    End If
                    
                Case "o" '<nickname>  Gives Operator status to the user
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If X = 0 Then GoTo ReScan
                        X = InStr(1, TmpText, "v")
                        If Not X = 0 Then TmpFlag = "+"
                        X = InStr(1, TmpText, "h")
                        If Not X = 0 Then TmpFlag = "%"
                        sUsers = .iCUsers(Q)
                        sUsersM = .iCUsersM(Q)
                        sUsers = Mid$(ReplaceStr(" " & sUsers, " " & "@" & TmpLoad & " ", " " & TmpFlag & TmpLoad & " "), 2)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "-o")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = False Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = False
                            SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    End If
                    
                Case "p" 'Private channel
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "s" 'Secret channel
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "m" 'Moderated channel, Only users with mode +voh can speak.
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "n" 'No messages from outside channel
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "t" 'Only Channel Operators may set the topic
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "i" 'Invite-only allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "I" 'Invited Users
                    
                Case "k" '<key>  Needs the Channel Key to join the channel
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpLoad = .iCKey(Q)
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), "", .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "r" 'Channel is Registered
                    
                Case "R" 'Requires a Registered nickname to join the channel
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "c" 'Blocks messages with ANSI colour (ColourBlock).
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "a" '<nickname>  Gives protection to the user (No kick/drop)
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
                        If InStr(1, TmpText, "q") = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** Protected users can only be set by the channel owner." & CRLF: GoTo ReScan
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        sUsersM = .iCUsersM(Q)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "-a")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = False Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = False
                            SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "q" '<nickname>  Channel owner
                    If TmpLoad = "" Then GoTo ReScan
                    For X = 1 To iUserMax
                        If LCase(TmpLoad) = LCase(iUser(X)) Then
                            TmpLoad = iUser(X)
                            Y = X
                        End If
                    Next X
                    
                    If Not Y = 0 Then
                        TmpFlag = ""
                        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
                        If InStr(1, TmpText, "q") = 0 Then SendData Index, ":" & sServer & " 468 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_468, Index) & CRLF: GoTo ReScan
                        TmpText = GetUserCFlags(.iChan(Q), TmpLoad)
                        If TmpText = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & iUser(Y) & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: GoTo ReScan
                        X = InStr(1, TmpText, flag)
                        If Not X = 0 Then GoTo ReScan
                        sUsersM = .iCUsersM(Q)
                        sUsersM = SetUserCFlags(.iChan(Q), TmpLoad, "-q")
                        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                        If AddedFlags = False Then
                            SFlags = SFlags & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        Else
                            AddedFlags = False
                            SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                        End If
                    Else
                        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpLoad & " " & sConvertText(ERR_401, Index) & CRLF
                    End If
                    
                Case "O" 'IRC Operators Only
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "A" 'IRC Admins Only
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "L" '<chan2>  If +l is full, the next user will auto-join <chan2>
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpLoad = .iCLinked(Q)
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), "")
                    If AddedFlags = False Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "Q" 'No Kicking
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "b" '<nick!user@host>  Bans the nick!user@host from the channel
                    If TmpLoad = "" Then GoTo ReScan
                    X = InStr(1, " " & LCase(.iCBans(Q)), " " & LCase(TmpLoad) & " ")
                    If X = 0 Then GoTo ReScan
                    Y = InStr(X, .iCBans(Q), " ")
                    TmpLoad = Mid$(.iCBans(Q), X, Y - X)
                    TmpText = Mid$(.iCBans(Q), 1, X - 1) & Mid$(.iCBans(Q), Y + 1)
                    X = InStr(1, " " & LCase(.iCBansTS(Q)), " " & LCase(TmpLoad) & "$")
                    Y = InStr(X, .iCBansTS(Q), " ")
                    sUsers = Mid$(.iCBansTS(Q), 1, X - 1) & Mid$(.iCBansTS(Q), Y + 1)
                    X = InStr(1, " " & LCase(.iCBansSN(Q)), " " & LCase(TmpLoad) & "$")
                    Y = InStr(X, .iCBansSN(Q), " ")
                    sUsersM = Mid$(.iCBansSN(Q), 1, X - 1) & Mid$(.iCBansSN(Q), Y + 1)
                    
                    Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), TmpText, sUsers, sUsersM, .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                    
                Case "S" 'Strip all incoming colours away
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "e" 'Exception ban
                    If TmpLoad = "" Then GoTo ReScan
                    X = InStr(1, " " & LCase(.iCExcep(Q)), " " & LCase(TmpLoad) & " ")
                    If X = 0 Then GoTo ReScan
                    Y = InStr(X, .iCExcep(Q), " ")
                    TmpLoad = Mid$(.iCExcep(Q), X, Y - X)
                    TmpText = Mid$(.iCExcep(Q), 1, X - 1) & Mid$(.iCExcep(Q), Y + 1)
                    X = InStr(1, " " & LCase(.iCExcepTS(Q)), " " & LCase(TmpLoad) & "$")
                    Y = InStr(X, .iCExcepTS(Q), " ")
                    sUsers = Mid$(.iCExcepTS(Q), 1, X - 1) & Mid$(.iCExcepTS(Q), Y + 1)
                    X = InStr(1, " " & LCase(.iCExcepSN(Q)), " " & LCase(TmpLoad) & "$")
                    Y = InStr(X, .iCExcepSN(Q), " ")
                    sUsersM = Mid$(.iCExcepSN(Q), 1, X - 1) & Mid$(.iCExcepSN(Q), Y + 1)
                    
                    Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), TmpText, sUsers, sUsersM, .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag:  SFlagValue = SFlagValue & TmpLoad & " "
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag: SFlagValue = SFlagValue & TmpLoad & " "
                    End If
                    
                Case "K" '/KNOCK is not allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "V" '/INVITE is not allowed
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "f" '[*]<lines>:<seconds>  Flood protection
                    
                Case "H" 'No +I uMode Users
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "G" 'Swear Filter
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "C" 'No CTCPs
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "u" '"Auditorium". Makes /NAMES and /WHO #channel only show Operators.
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "z" 'Only Clients on a Secure Connection (SSL) can join.
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "N" 'No Nickname changes are permitted in the channel.
                    X = InStr(1, .iCModes(Q), flag)
                    If X = 0 Then GoTo ReScan
                    TmpText = ReplaceStr(.iCModes(Q), flag, "")
                    Q = cChanChange(.iChan(Q), TmpText, .iCUsers(Q), .iCUsersM(Q), .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
                    If AddedFlags = False Then
                        SFlags = SFlags & flag
                    Else
                        AddedFlags = False
                        SFlags = SFlags & "-" & flag
                    End If
                    
                Case "+"
                    AddFlag = True
            End Select
        End If
        
        If Not Modes = "" Then GoTo ReScan
        If Not SFlags = "" Then
            If Not Left$(SFlags, 1) = "-" Then SFlags = "+" & SFlags
            If Not SFlagValue = "" Then
                cMODENotify .iChan(Q), SFlags & " " & SFlagValue, Index
            Else
                cMODENotify .iChan(Q), SFlags, Index
            End If
        End If
    Else
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " " & Chan & " :" & sConvertText(ERR_401, Index) & CRLF
    End If
    End With
End Sub

Sub cRemoveChannel(Channel As String)
On Error Resume Next
Dim X As Integer
    With iChanSys
    
    If Channel = "ALL" Then
        For X = 1 To iChanSys.iChan.Count
            .iCBans.Remove X
            .iCBansSN.Remove X
            .iCBansTS.Remove X
            .iCExcep.Remove X
            .iCExcepSN.Remove X
            .iCExcepTS.Remove X
            .iCCDate.Remove X
            .iChan.Remove X
            .iCKey.Remove X
            .iCModes.Remove X
            .iCTopic.Remove X
            .iCTopicC.Remove X
            .iCTopicD.Remove X
            .iCCreator.Remove X
            .iCULimit.Remove X
            .iCUsers.Remove X
            .iCUsersM.Remove X
            .iCLinked.Remove X
        Next X
        frmMain.lbl_CC = 0
    Else
        For X = 1 To iChanSys.iChan.Count
            If LCase(Channel) = LCase(iChanSys.iChan(X)) Then
            .iCBans.Remove X
            .iCBansSN.Remove X
            .iCBansTS.Remove X
            .iCExcep.Remove X
            .iCExcepSN.Remove X
            .iCExcepTS.Remove X
            .iCCDate.Remove X
            .iChan.Remove X
            .iCKey.Remove X
            .iCModes.Remove X
            .iCTopic.Remove X
            .iCTopicC.Remove X
            .iCTopicD.Remove X
            .iCCreator.Remove X
            .iCULimit.Remove X
            .iCUsers.Remove X
            .iCUsersM.Remove X
            .iCLinked.Remove X
                frmMain.lbl_CC = .iChan.Count
                Exit Sub
            End If
        Next X
    End If
    End With
End Sub

Private Function cChanChange(Channel As String, Modes As String, Users As String, UsersM As String, Topic As String, TopicChanger As String, ChanLimit As String, key As String, Bans As String, BansTS As String, BansSN As String, Except As String, ExceptTS As String, ExceptSN As String, Link As String) As Integer
On Error Resume Next
Dim ChanName, ChanDate, TopicDate, ChanCreator As String
Dim X As Integer
Dim Q As Integer
    With iChanSys
    
    For X = 1 To iChanSys.iChan.Count
        If LCase(Channel) = LCase(iChanSys.iChan(X)) Then
            ChanCreator = .iCCreator(X)
            If Topic = .iCTopic(X) Then TopicDate = .iCTopicD(X) Else:  TopicDate = GetTime
            ChanDate = .iCCDate(X)
            ChanName = .iChan(X)
            ChanCreator = .iCCreator(X)
            Q = X
            cRemoveChannel .iChan(X)
            Exit For
        End If
    Next X
        .iChan.Add ChanName
        .iCBans.Add Bans
        .iCBansSN.Add BansSN
        .iCBansTS.Add BansTS
        .iCExcep.Add Except
        .iCExcepSN.Add ExceptSN
        .iCExcepTS.Add ExceptTS
        .iCCDate.Add ChanDate
        .iCCreator.Add ChanCreator
        .iCKey.Add key
        .iCModes.Add Modes
        .iCTopic.Add Topic
        .iCTopicC.Add TopicChanger
        .iCTopicD.Add TopicDate
        .iCULimit.Add ChanLimit
        .iCUsers.Add Users
        .iCUsersM.Add UsersM
        .iCLinked.Add Link
        cChanChange = .iChan.Count
    End With
End Function

Private Sub PARTNotify(Channel As String, PartMSG As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not Index = X Then
                If PartMSG = "" Then
                    SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " PART :" & .iChan(Q) & CRLF
                Else
                    SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " PART :" & .iChan(Q) & " :" & PartMSG & CRLF
                End If
            End If
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub

Private Sub cMODENotify(Channel As String, Modes As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " MODE " & .iChan(Q) & " " & Modes & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub


Private Sub TOPICNotify(Channel As String, Topic As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not Index = X Then
                SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " TOPIC " & .iChan(Q) & " :" & Topic & CRLF
            End If
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub

Private Function QuitNotify(Channel As String, PartMSG As String, UserHostMask As String, Optional NotToUsers As String) As String
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then QuitNotify = NotToUsers: Exit Function
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    Q = InStr(1, " " & LCase(NotToUsers), " " & LCase(TmpUser) & " ")
    If Not Q = 0 Then GoTo ReScan
    NotToUsers = NotToUsers & TmpUser & " "
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            SendData X, ":" & UserHostMask & " QUIT :" & PartMSG & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Function

Private Function NCNotify(Channel As String, NewNick As String, UserHostMask As String, Optional NotToUsers As String) As String
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then NCNotify = NotToUsers: Exit Function
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    Q = InStr(1, " " & LCase(NotToUsers), " " & LCase(TmpUser) & " ")
    If Not Q = 0 Then GoTo ReScan
    NotToUsers = NotToUsers & TmpUser & " "
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not NewNick = iUser(X) Then SendData X, ":" & UserHostMask & " NICK " & NewNick & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Function

Private Sub JOINNotify(Channel As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not Index = X Then SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " JOIN :" & .iChan(Q) & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub

Sub sChanMessage(Channel As String, Message As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim TmpText As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    Q = -1
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    If Q = -1 Then
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & Channel & " No such nick/channel" & CRLF
        Exit Sub
    End If
    If Len(Message) = 0 Then
        SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :No text to send" & CRLF
        Exit Sub
    End If
    
    X = InStr(1, .iCModes(Q), "n")
    If Not X = 0 Then
        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
        If TmpText = "!" Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_404, Index) & CRLF: Exit Sub
    End If
    X = InStr(1, .iCModes(Q), "m")
    If Not X = 0 Then
        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
        X = InStr(1, TmpText, "v")
        If Not X = 0 Then cUserLevel = 1
        X = InStr(1, TmpText, "h")
        If Not X = 0 Then cUserLevel = 2
        X = InStr(1, TmpText, "o")
        If Not X = 0 Then cUserLevel = 3
        If cUserLevel = 0 Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_404B, Index) & CRLF: Exit Sub
    End If
    
    X = InStr(1, .iCModes(Q), "c")
    If Not X = 0 Then If SysCFCC(Message) = True Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :Color is not permitted in this channel" & CRLF: Exit Sub
    ' :linux.ircd-net.org 404 TRON TRON :Colour is not permitted in this channel (#TRON)
    X = InStr(1, .iCModes(Q), "S")
    If Not X = 0 Then Message = SysFOCC(Message, True)
    'X = InStr(1, .iCModes(Q), "G")
    'If Not X = 0 Then Message = SysFilter(Message, True)
    ' Code has been stripped for above to even work, sorry...
    
    TmpBuffer = .iCUsers(Q) & " "
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not Index = X Then SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " PRIVMSG " & .iChan(Q) & " :" & Message & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub

Sub sChanNotice(Channel As String, Message As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim cUserLevel As Integer
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    Q = -1
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    If Q = -1 Then
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & Channel & " No such nick/channel" & CRLF
        Exit Sub
    End If
    If Len(Message) = 0 Then
        SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :No text to send" & CRLF
        Exit Sub
    End If
    
    X = InStr(1, .iCModes(Q), "n")
    If Not X = 0 Then
        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
        If TmpText = "!" Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_404, Index) & CRLF: Exit Sub
    End If
    X = InStr(1, .iCModes(Q), "m")
    If Not X = 0 Then
        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
        X = InStr(1, TmpText, "v")
        If Not X = 0 Then cUserLevel = 1
        X = InStr(1, TmpText, "h")
        If Not X = 0 Then cUserLevel = 2
        X = InStr(1, TmpText, "o")
        If Not X = 0 Then cUserLevel = 3
        If cUserLevel = 0 Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_404B, Index) & CRLF: Exit Sub
    End If
    
    X = InStr(1, .iCModes(Q), "c")
    If Not X = 0 Then If SysCFCC(Message) = True Then SendData Index, ":" & sServer & " 404 " & iUser(Index) & " " & .iChan(Q) & " :Color is not permitted in this channel" & CRLF: Exit Sub
    ' :linux.ircd-net.org 404 TRON TRON :Colour is not permitted in this channel (#TRON)
    X = InStr(1, .iCModes(Q), "S")
    If Not X = 0 Then Message = SysFOCC(Message, True)
    'X = InStr(1, .iCModes(Q), "G")
    'If Not X = 0 Then Message = SysFilter(Message, True)
    ' Code has been stripped for above to even work, sorry...
    
    TmpBuffer = .iCUsers(Q) & " "
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            If Not Index = X Then SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " NOTICE " & .iChan(Q) & " :" & Message & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub

Sub sChanTOPIC(Index As Integer, Channel As String, Topic As String)
On Error Resume Next
Dim sChan As String
Dim sUsers As String
Dim sDate As String
Dim sNick As String
Dim TmpText As String
Dim cUserLevel As Integer
Dim DFY As Boolean
Dim sUCS As String
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
With iChanSys
    sNick = iUser(Index)
    sChan = Channel
    Topic = Mid$(Topic, 1, 307)
    
    If Not Left$(sChan, 1) = "#" Then
        SendData Index, ":" & sServer & " 461 " & User & " TOPIC :" & sConvertText(ERR_461, Index) & CRLF
        Exit Sub
    End If
    
    For X = 1 To .iChan.Count
        If LCase(sChan) = LCase(.iChan(X)) Then
            Q = X
            DFY = True
            Exit For
        End If
    Next X
    
    If Left(Topic, 1) = ":" Then
        Topic = Mid$(Topic, 2)
    Else
        If Not .iCTopic(Q) = "" Then
            SendData2 Index, ":" & sServer & " 332 " & iUser(Index) & " " & .iChan(Q) & " :" & .iCTopic(Q) & CRLF & _
                             ":" & sServer & " 333 " & iUser(Index) & " " & .iChan(Q) & " " & .iCTopicC(Q) & " " & .iCTopicD(Q) & CRLF
        Else
            SendData2 Index, ":" & sServer & " 331 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(RPL_331, Index) & CRLF
        End If
        Exit Sub
    End If
    
    If DFY = True Then
        sUsers = .iCUsers(Q)
        
        X = InStr(1, LCase(sUsers), LCase(sNick & " "))
        If X = 0 Then
            SendData Index, ":" & sServer & " 442 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_442, Index) & CRLF
            Exit Sub
        End If
        
        TmpText = GetUserCFlags(.iChan(Q), iUser(Index))
        X = InStr(1, TmpText, "v")
        If Not X = 0 Then cUserLevel = 1
        X = InStr(1, TmpText, "h")
        If Not X = 0 Then cUserLevel = 2
        X = InStr(1, TmpText, "o")
        If Not X = 0 Then cUserLevel = 3
        
        X = InStr(1, .iCModes(Q), "t")
        If Not X = 0 Then
            If Not cUserLevel = 3 Then
                If Not cUserLevel = 2 Then
                    SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF
                    Exit Sub
                End If
            End If
        End If
        
        Q = cChanChange(.iChan(Q), .iCModes(Q), .iCUsers(Q), .iCUsersM(Q), Topic, iUser(Index), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        
        If Q = 0 Then
            LogIt "Faild 2 Find " & sChan & " N TOPIC Buffer 4 " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)
            If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Channel not found in Buffer!  Error has been logged." & CRLF
            Exit Sub
        End If
        
        TOPICNotify .iChan(Q), Topic, Index
        SendData Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " TOPIC " & .iChan(Q) & " :" & Topic & CRLF
    Else
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & sChan & " No such nick/channel" & CRLF
    End If
End With
End Sub

Sub SendNAMES(Index As Integer, Text As String)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpChan As String
    With iChanSys
    
    X = InStr(1, Text, " ")
    TmpChan = Mid$(Text, 1, X - 1)
    
    For X = 1 To .iChan.Count
        If LCase(.iChan(X)) = LCase(TmpChan) Then
            Q = X
            Exit For
        End If
    Next X
    
    If Not Q = 0 Then
        SendData Index, ":" & sServer & " 353 " & iUser(Index) & " = " & .iChan(Q) & " :" & .iCUsers(Q) & CRLF & _
                        ":" & sServer & " 366 " & iUser(Index) & " " & TmpChan & " :" & sConvertText(RPL_366, Index) & CRLF
    Else
        SendData Index, ":" & sServer & " 403 " & iUser(Index) & " :" & sConvertText(ERR_403, Index) & CRLF
    End If
    End With
    'get 39 Names Limition working for next Release...
    
    ' :ircdnet.dyndns.org 353 TRON = #IRCdNet :TRON @DreX
    ' :ircdnet.dyndns.org 366 TRON #ircdnet :End of /NAMES list.
End Sub

Sub SendLIST(Index As Integer, Text As String)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim TmpText As String
Dim TmpChan As String
Dim TmpData As String
Dim TmpNumber As String
Dim TmpModes As String
Dim DUI As Boolean
    With iChanSys
    frmMain.lbl_CC = .iChan.Count
    
    X = InStr(1, Text, " ")
    TmpChan = Mid$(Text, 1, X - 2)
    SendData Index, ":" & sServer & " 321 " & iUser(Index) & " Channel :Users  Name" & CRLF
    
    If TmpChan = "" Then
        For X = 1 To .iChan.Count
            DUI = False
            If Not .iCModes(X) = "" Then
                TmpModes = "[+" & .iCModes(X) & "]"
            Else
                TmpModes = ""
            End If
            TmpText = Replace(.iCUsers(X), " ", "", 1, Len(.iCUsers(X)), vbTextCompare)
            TmpNumber = Len(.iCUsers(X)) - Len(TmpText)
            If Not InStr(1, .iCModes(X), "s") = 0 Then DUI = True
            If Not InStr(1, .iCModes(X), "p") = 0 Then DUI = True
            If DUI = True Then
                TmpData = GetUserCFlags(.iChan(X), iUser(Index))
                If Not TmpData = "!" Then DUI = False
                'MsgBox "'" & TmpData & "'"
            End If
            If DUI = False Then SendData Index, ":" & sServer & " 322 " & iUser(Index) & " " & .iChan(X) & " " & TmpNumber & " :" & TmpModes & " " & .iCTopic(X) & CRLF
        Next X
    Else
        For X = 1 To .iChan.Count
            If .iChan(X) Like "*" & TmpChan & "*" Then
                DUI = False
                If Not .iCModes(X) = "" Then
                    TmpModes = "[+" & .iCModes(X) & "]"
                Else
                    TmpModes = ""
                End If
                TmpText = Replace(.iCUsers(X), " ", "", 1, Len(.iCUsers(X)), vbTextCompare)
                TmpNumber = Len(.iCUsers(X)) - Len(TmpText)
                If Not InStr(1, .iCModes(X), "s") = 0 Then DUI = True
                If Not InStr(1, .iCModes(X), "p") = 0 Then DUI = True
                If DUI = True Then: If Not GetUserCFlags(.iChan(X), iUser(Index)) = "!" Then DUI = False
                If DUI = False Then SendData Index, ":" & sServer & " 322 " & iUser(Index) & " " & .iChan(X) & " " & TmpNumber & " :" & TmpModes & " " & .iCTopic(X) & CRLF
            End If
        Next X
    End If
        SendData Index, ":" & sServer & " 323 " & iUser(Index) & " :" & sConvertText(RPL_323, Index) & CRLF
    End With
    
    ':ircdnet.dyndns.org 321 TRON Channel :Users  Name
    ':ircdnet.dyndns.org 322 TRON #IRCOps 1 :
    ':ircdnet.dyndns.org 323 TRON :End of /LIST
End Sub

Private Function GetUserCFlags(sChan As String, sUser As String) As String
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim sUsers As String
Dim X As Integer
Dim Q As Integer
With iChanSys
    For X = 1 To .iChan.Count
        If LCase(sChan) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    
    If Not Q = 0 Then
        sUsers = " " & .iCUsersM(Q)
        X = InStr(1, sUsers, " " & sUser & "$")
        If X = 0 Then GetUserCFlags = "!": Exit Function
        Q = InStr(X + 1, sUsers, " ")
        GetUserCFlags = Mid$(sUsers, X + 2 + Len(sUser), Q - 2 - X - Len(sUser))
    Else
        GetUserCFlags = "-"
    End If
    
End With
End Function

Private Function SetUserCFlags(sChan As String, sUser As String, Modes As String) As String
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim sUsers As String
Dim TmpFlags As String
Dim OldFlags As String
Dim flag As String
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
Dim AddFlag As Boolean

With iChanSys

    For X = 1 To .iChan.Count
        If LCase(sChan) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    
    If Q = 0 Then Exit Function
        sUsers = " " & .iCUsersM(Q)
        X = InStr(1, sUsers, " " & sUser & "$")
        Z = InStr(X + 1, sUsers, " ")
        TmpFlags = Mid$(sUsers, X + 2 + Len(sUser), Z - 2 - X - Len(sUser))
        'sUser = Mid$(sUsers, X + 1, Z - 2 - X)
        OldFlags = TmpFlags
        AddFlag = True
ReScan:
        If Modes = "" Then GoTo ScanDone
        flag = Mid$(Modes, 1, 1)
        Modes = Mid$(Modes, 2)
        
        Select Case flag
        Case "+": AddFlag = True
        Case "-": AddFlag = False
        Case Else
            If AddFlag = True Then
                X = InStr(1, TmpFlags, flag)
                If Not X = 0 Then GoTo ReScan
                TmpFlags = TmpFlags & flag
            Else
                X = InStr(1, TmpFlags, flag)
                If X = 0 Then GoTo ReScan
                If X = 1 Then
                    TmpFlags = Mid$(TmpFlags, 2)
                Else
                    TmpText = Mid$(TmpFlags, 1, X - 1)
                    TmpData = Mid$(TmpFlags, X + 1)
                    TmpFlags = TmpText & TmpData
                End If
            End If
        End Select
        GoTo ReScan
ScanDone:
        TmpText = ReplaceStr(sUsers, " " & sUser & "$" & OldFlags & " ", " " & sUser & "$" & TmpFlags & " ")
        SetUserCFlags = Mid$(TmpText, 2)
    
End With
End Function



Sub uKICK(Channel As String, User As String, Reason As String, Index As Integer)
On Error Resume Next
Dim sChan As String
Dim sPartMsg As String
Dim sUsers As String
Dim sUsersM As String
Dim sDate As String
Dim TmpText As String
Dim TmpText2 As String
Dim cUserLevel As Integer
Dim sNick As String
Dim DFY As Boolean
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
    'sNick = User
    sChan = Channel
    With iChanSys
    If Reason = "" Then Reason = iUser(Index)
    If Not Left$(sChan, 1) = "#" Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " KICK :" & sConvertText(ERR_461, Index) & CRLF
        Exit Sub
    End If
    If User = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " KICK :" & sConvertText(ERR_461, Index) & CRLF: Exit Sub
    
    For X = 1 To iUserMax
        If LCase(User) = LCase(iUser(X)) Then
            User = iUser(X): DFY = True
            sNick = iUser(X)
            Y = X
            Exit For
        End If
    Next X
    
    If sNick = "" Then SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & User & " No such nick/channel" & CRLF: Exit Sub
    
    For X = 1 To .iChan.Count
        If LCase(sChan) = LCase(.iChan(X)) Then
            sChan = .iChan(Q)
            Q = X
            Exit For
        End If
    Next X
    
    If Not Q = 0 Then
        sUsers = .iCUsers(Q)
        sUsersM = .iCUsersM(Q)
        
        
        If GetUserCFlags(.iChan(Q), User) = "!" Then SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & User & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF: Exit Sub
        If Not InStr(1, GetUserCFlags(.iChan(Q), iUser(Index)), "h") = 0 Then cUserLevel = 1
        If Not InStr(1, GetUserCFlags(.iChan(Q), iUser(Index)), "o") = 0 Then cUserLevel = 2
        If cUserLevel = 0 Then SendData Index, ":" & sServer & " 482 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_482, Index) & CRLF: Exit Sub
        If cUserLevel = 1 Then
            If Not InStr(1, GetUserCFlags(.iChan(Q), User), "o") = 0 Then
                SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** You cannot kick channel operators on " & .iChan(Q) & " if you only are halfop" & CRLF
                Exit Sub
            End If
        End If
        If Not InStr(1, GetUserCFlags(.iChan(Q), User), "a") = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** You cannot kick " & User & " from " & .iChan(Q) & " because " & User & " is channel protected" & CRLF: Exit Sub
        
        X = InStr(1, LCase(sUsers), LCase(sNick & " "))
        If X = 1 Then
            TmpText2 = Mid$(sUsers, X)
        Else
            TmpText2 = Mid$(sUsers, X - 1)
        End If
        
        If X = 0 Then
            SendData Index, ":" & sServer & " 441 " & iUser(Index) & " " & sNick & " " & .iChan(Q) & " :" & sConvertText(ERR_441, Index) & CRLF
            Exit Sub
        Else
            Select Case Left$(TmpText2, 1)
                Case "@": sUsers = ReplaceStr(sUsers, "@" & sNick & " ", "")
                Case "%": sUsers = ReplaceStr(sUsers, "%" & sNick & " ", "")
                Case "+": sUsers = ReplaceStr(sUsers, "+" & sNick & " ", "")
                Case " ": sUsers = ReplaceStr(sUsers, " " & sNick & " ", " ")
                Case Else: sUsers = Mid$(TmpText2, Len(sNick) + 2)
            End Select
            
            Z = InStr(1, iChan(Index), sChan)
            If Z = 1 Then
                iChan(Index) = Mid$(iChan(Index), Len(sChan) + 2)
            ElseIf Z = 2 Then
                    iChan(Index) = Mid$(iChan(Index), Len(sChan) + 3)
                Else
                    TmpText = Mid$(iChan(Index), Z - 2, 2)
                    
                    Select Case TmpText
                        Case "~*": iChan(Index) = ReplaceStr(iChan(Index), "~*" & sChan & " ", "")
                        Case "~@": iChan(Index) = ReplaceStr(iChan(Index), "~@" & sChan & " ", "")
                        Case "~%": iChan(Index) = ReplaceStr(iChan(Index), "~%" & sChan & " ", "")
                        Case "~+": iChan(Index) = ReplaceStr(iChan(Index), "~+" & sChan & " ", "")
                        Case " *": iChan(Index) = ReplaceStr(iChan(Index), "*" & sChan & " ", "")
                        Case " @": iChan(Index) = ReplaceStr(iChan(Index), "@" & sChan & " ", "")
                        Case " %": iChan(Index) = ReplaceStr(iChan(Index), "%" & sChan & " ", "")
                        Case " +": iChan(Index) = ReplaceStr(iChan(Index), "+" & sChan & " ", "")
                        Case Else: iChan(Index) = ReplaceStr(iChan(Index), " " & sChan & " ", " ")
                    End Select
            End If
            
            TmpFlags = GetUserCFlags(.iChan(Q), sNick)
            sUsersM = ReplaceStr(" " & sUsersM, " " & sNick & "$" & TmpFlags & " ", " ")
            sUsersM = Mid$(sUsersM, 2)
            
            TmpText2 = Replace(.iCUsers(Q), " ", "", 1, Len(.iCUsers(Q)), vbTextCompare)
            If Len(.iCUsers(Q)) - Len(TmpText2) = 1 Then
                SendData2 Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " KICK " & sChan & " " & User & " :" & Reason & CRLF
                cRemoveChannel .iChan(Q): Exit Sub
            End If
        End If
        
        Q = cChanChange(.iChan(Q), .iCModes(Q), sUsers, sUsersM, .iCTopic(Q), .iCTopicC(Q), .iCULimit(Q), .iCKey(Q), .iCBans(Q), .iCBansTS(Q), .iCBansSN(Q), .iCExcep(Q), .iCExcepTS(Q), .iCExcepSN(Q), .iCLinked(Q))
        For X = 1 To .iChan.Count
            If LCase(.iChan(X)) = LCase(sChan) Then Q = X: Exit For
        Next X
        
        If Q = 0 Then
            LogIt "Faild 2 Find " & sChan & " N PART Buffer 4 " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)
            If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Channel not found in Buffer!  Error has been logged." & CRLF
            Exit Sub
        End If
        
        KICKNotify .iChan(Q), User, Reason, Index
        SendData2 Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " KICK " & sChan & " " & User & " :" & Reason & CRLF
        SendData2 Y, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " KICK " & sChan & " " & User & " :" & Reason & CRLF
    Else
        SendData2 Index, ":" & sServer & " 401 " & iUser(Index) & " :" & sChan & " No such nick/channel" & CRLF
    End If
    
    End With
End Sub

Private Sub KICKNotify(Channel As String, User As String, KickMSG As String, Index As Integer)
On Error Resume Next
Dim TmpBuffer As String
Dim TmpUser As String
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    With iChanSys
    For X = 1 To .iChan.Count
        If LCase(Channel) = LCase(.iChan(X)) Then
            Q = X
            Exit For
        End If
    Next X
    TmpBuffer = .iCUsers(Q) & " "
    
ReScan:
    X = InStr(1, TmpBuffer, " ")
    If X = 0 Then Exit Sub
    TmpUser = Mid$(TmpBuffer, 1, X - 1)
    TmpBuffer = Mid$(TmpBuffer, X + 1)
    If Left$(TmpUser, 1) = "@" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "%" Then TmpUser = Mid$(TmpUser, 2)
    If Left$(TmpUser, 1) = "+" Then TmpUser = Mid$(TmpUser, 2)
    
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) And iPeerFree(X) = False Then
            SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " KICK " & sChan & " " & User & " :" & Reason & CRLF
            Exit For
        End If
    Next X
    
    GoTo ReScan
    End With
End Sub
