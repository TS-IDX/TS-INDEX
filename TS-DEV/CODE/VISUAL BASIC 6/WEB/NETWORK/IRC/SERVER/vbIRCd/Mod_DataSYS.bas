Attribute VB_Name = "Mod_DataSYS"
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
'


Public Sub SYS(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Z As Integer
Dim Y As Integer
Dim Q As Integer
Dim iX As Double
Dim TmpText As String
Dim TmpText2 As String
Dim TmpLoad As String
Dim TmpMsg As String
Dim TmpFlag As String
Dim TmpData As String
Dim DFY As Boolean
Dim DUI As Boolean ' Dead User Info
With frmMain
    
    iPing(Index) = 0 ' Let's reset ping idle so we won't have to ping alive connections
    sBReceived = sBReceived + Len(Text)
    sTReceived = sTReceived + 1
Text = Text & " "
    If Left$(LCase(Text), 2 + Len(iUser(Index))) = ":" & LCase(iUser(Index)) & " " Then Text = Mid$(Text, 2 + Len(iUser(Index)))
    If Left$(Text, 1) = ":" Then Exit Sub
ReScan:
    If Left$(Text, 1) = " " Then Text = Mid$(Text, 2): GoTo ReScan
    If Left$(Text, 1) = Chr$(34) Then
        Text = Mid$(Text, 2)
        Text = Mid$(Text, 1, Len(Text) - 3)
        Text = Text & " ": GoTo ReScan
    End If
    If Text = "" Then Exit Sub
    
    If Not Left$(Text, 5) = "PONG " Then If iFC(Index) > iFloodCMDs Then KillUser Index, sServer, "Command Flooding": Exit Sub
    If Not Left$(Text, 5) = "PONG " Then If Not iFloodCMDs = 0 And iUserLevel(Index) = 0 Then iFC(Index) = iFC(Index) + 1
    If Not Left$(Text, 5) = "PONG " Then If Not iFloodCMDs = 0 And iFC(Index) = iFloodCMDs Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** WARNING: DO NOT SEND ONE MORE COMMAND FOR ONE MINUTE OR ELSE YOU WILL BE DISCONNECTED FOR FLOODING!!!" & CRLF
    
    'Text = Mid$(Text, 1, Len(Text) - 2)
    
    If Left$(UCase(Text), 5) = "PASS " Then
        xPASS = xPASS + 1
        xPASS2 = xPASS2 + Len(Text)
        TmpText = Mid$(Text, 1, Len(Text) - 1)
        TmpText = Mid$(TmpText, 6)
        If sACC = 1 Then
            If LCase(TmpText) = LCase(iConnPass) Then
                iAAC(Index) = True
                LogIt "Pass-Accepted -> '" & TmpText & "'"
            Else
                SendData Index, ":" & sServer & " 464 AUTH :Password incorrect" & CRLF
                KillUser Index, sServer, "Password incorrect"
            End If
        End If
        If sACC = 2 Then iTP(Index) = TmpText
        Exit Sub
    End If
    
    If sACC = 1 Then
        If iAAC(Index) = False Then
            KillUser Index, sServer, "Connection access password incorrect"
            Exit Sub
        End If
    End If
    
    If Left$(UCase(Text), 5) = "NICK " Then
        xNICK = xNICK + 1
        xNICK2 = xNICK2 + Len(Text)
        TmpText = Mid$(Text, 1, Len(Text) - 1)
        TmpText = Mid$(TmpText, 6)
        Q = InStr(1, TmpText, " ")
        If Not Q = 0 Then TmpText = Mid$(TmpText, 1, Q - 1)
        If Left$(TmpText, 1) = ":" Then TmpText = Mid$(TmpText, 2)
        If TmpText = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " NICK :Not enough parameters" & CRLF: Exit Sub
        For Q = 1 To iUserMax
            If LCase(iUser(Q)) = LCase(TmpText) Then
                If Not Index = Q Then
                    SendData Index, ":" & sServer & " 433 " & iUser(Index) & " :Nickname already in use" & CRLF
                    Exit Sub
                Else
                    If TmpText = iUser(Index) Then Exit Sub
                    Exit For
                End If
            End If
        Next Q
        
            If sNickValid(TmpText) = True Then
                If sNameCheck(TmpText) = True And sFONN = 1 Then
                    SendData Index, ":" & sServer & " 432 " & iUser(Index) & " " & TmpText & " :Erroneus nickname -That type of nick is not allowed to be used" & CRLF
                    Exit Sub
                End If
                
                If iUserLevel(Index) = 0 Then
                    For X = 1 To sQline.Count
                        If LCase(TmpText) Like LCase(sQline(X)) Then
                            SendData Index, ":" & sServer & " 432 " & iUser(Index) & " " & TmpText & " :Erroneus nickname -Nick is reserved(Reason: " & sQlineR(X) & ")" & CRLF & _
                                            ":" & sServer & " NOTICE " & iUser(Index) & " :*** ERROR: Cannot use '" & TmpText & "' nickname cause it's Q:lined(Reason: " & sQlineR(X) & ")" & CRLF
                            Exit Sub
                        End If
                    Next X
                End If
            Else
                SendData Index, ":" & sServer & " 432 " & iUser(Index) & " " & TmpText & " :Erroneus nickname" & CRLF
                Exit Sub
            End If
            
        If iUser(Index) = "" And Not iName(Index) = "" Then
            iUser(Index) = TmpText
            If iHost(Index) = "" Then Exit Sub
            
            For X = 1 To sKline.Count
                If LCase(iName(Index) & "@" & iRHost(Index)) Like LCase(sKline(X)) Then
                    DUI = False
                    For Q = 1 To sEline.Count
                        If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                            DUI = True
                            Exit For
                        End If
                    Next Q
                    If DUI = False Then
                        SendData Index, "ERROR :You are banned from this server, Reason: " & sKlineR(X) & CRLF
                        KillUser Index, sServer, "(" & sKlineR(X) & "(KLINED))", , True, True
                        Exit Sub
                    End If
                    Exit For
                End If
            Next X
            
            For X = 1 To sAKill.Count
                If LCase(iName(Index) & "@" & iRHost(Index)) Like LCase(sAKill(X)) Then
                    DUI = False
                    For Q = 1 To sEline.Count
                        If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                            DUI = True
                            Exit For
                        End If
                    Next Q
                    If DUI = False Then
                        SendData Index, "ERROR :You are banned from this network, Reason: " & sAKillR(X) & CRLF
                        KillUser Index, sServer, "(" & sAKillR(X) & "(AKILLED))", , True, True
                        Exit Sub
                    End If
                    Exit For
                End If
            Next X
            
            For Q = 1 To iUserMax
                If iPeerFree(Q) = False Then
                    X = InStr(1, iModes(Q), "c")
                    If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has connected to server on port " & .Win(Index).LocalPort & CRLF
                End If
            Next Q
            
            SendData Index, _
                ":" & sServer & " 001 " & iUser(Index) & " :Welcome to the " & iNetName & " IRC Network " & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & CRLF & _
                ":" & sServer & " 002 " & iUser(Index) & " :Your Host is " & sServer & ", Running vbIRCd(IRCServ Clone) " & sVersion & " " & sRelease & CRLF & _
                ":" & sServer & " 003 " & iUser(Index) & " :This server was created " & sUDT & CRLF & _
                ":" & sServer & " 004 " & iUser(Index) & " :" & sServer & " vbIRCd-" & sVersion & " oOiwghskSaHANTcCfrxebWqBFI1dvtGz lvhopsmntikrRcaqOALQbSeKVfHGCuzN" & CRLF & _
                ":" & sServer & " 005 " & iUser(Index) & " :NOQUIT ISON USERS MODES=13 MAXCHANNELS=" & iChanMax & " MAXBANS=60 NICKLEN=30 TOPICLEN=307 KICKLEN=307 CHANTYPES=# PREFIX=(ohv)@%+ :are available on this server" & CRLF
            SendLUSERS Index
            SendMOTD Index
            iSignOn(Index) = GetTime
            
            .lbl_UC = .lbl_UC - 1
            .lbl_CU = .lbl_CU + 1
            If .lbl_HU = .lbl_CU - 1 Then .lbl_HU = .lbl_CU
            .lbl_CGU = .lbl_CGU + 1
            If .lbl_HGU = .lbl_CGU - 1 Then .lbl_HGU = .lbl_HGU + 1
            If iForceCloak = 1 Then uMode Index, iUser(Index), "+x"
        Else
            If uCanNICK(iUser(Index), Index) = False Then Exit Sub
            TmpText2 = iUser(Index)
            iUser(Index) = TmpText
            SendData Index, ":" & TmpText2 & "!" & iName(Index) & "@" & iRHost(Index) & " NICK :" & iUser(Index) & CRLF
            uNickChange TmpText2 & "!" & iName(Index) & "@" & iHost(Index), iChan(Index), iUser(Index)
        End If
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "SERVER " Then
        xSERVER = xSERVER + 1
        xSERVER2 = xSERVER2 + Len(Text)
        TmpText = Mid$(Text, 8)
        If iUser(Index) = "" And iName(Index) = "" Then
            SendData Index, "ERROR :IRC Serv doesn't support linking IRCDs yet..." & CRLF
            UserClosed Index, "ERROR: Server " & TmpText & "(" & .Win(Index).PeerAddress & ") tried to link to this non-linkable ircd."
        Else
            SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Sorry, but your IRC software doesn't appear to support changing servers." & CRLF
        End If
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "USER " Then
        xUSER = xUSER + 1
        xUSER2 = xUSER2 + Len(Text)
        Text = Mid$(Text, 1, Len(Text) - 1)
        TmpText = Text
        TmpText = Mid$(TmpText, 6)
        If iName(Index) = "" Then
            X = InStr(1, TmpText, " ")
            iName(Index) = Mid$(TmpText, 1, X - 1)
            TmpText = Mid$(Text, X + 1)
            X = InStr(1, TmpText, " ")
            If iGODNS = 0 Then
                iHost(Index) = Mid$(TmpText, 1, X - 1)
                If Left$(iHost(Index), 1) = Chr$(34) Then iHost(Index) = Mid$(iHost(Index), 2)
                If Right$(iHost(Index), 1) = Chr$(34) Then iHost(Index) = Mid$(iHost(Index), 1, Len(iHost(Index)) - 1)
            End If
            
            TmpText = Mid$(TmpText, X + 1)
            X = InStr(1, TmpText, " ")
            TmpText = Mid$(TmpText, X + 1)
            X = InStr(1, TmpText, " ")
            iRealName(Index) = Mid$(TmpText, X + 1)
            
            If Left$(iRealName(Index), 1) = ":" Then iRealName(Index) = Mid$(iRealName(Index), 2)
            If iHost(Index) = "" Then iHost(Index) = .Win(Index).PeerAddress
            iIP(Index) = .Win(Index).PeerAddress
            If iName(Index) = "" Then DUI = True
            If iHost(Index) = "" Then DUI = True
            If iRealName(Index) = "" Then DUI = True
            If DUI = True Then
                SendData Index, ":" & sServer & " 461 " & iUser(Index) & " USER :Not enough parameters" & CRLF
                Exit Sub
            End If
            X = InStr(1, iRHost(Index), ".")
            If X = 0 Then iRHost(Index) = frmMain.Win(Index).PeerAddress
            iHost(Index) = iRHost(Index)
            
            If sNickValid(iName(Index)) = True Then
                If sNameCheck(iName(Index)) = True And sFONN = 1 Then
                    SendData Index, ":" & sServer & " 455 " & iUser(Index) & " " & iName(Index) & " :Erroneus nickname -That type of UserID is not allowed to be used" & CRLF
                    KillUser Index, sServer, "Valid User Name/ID with forbidden chars", , True, True
                    Exit Sub
                End If
            Else
                SendData Index, ":" & sServer & " 455 " & iUser(Index) & " " & iName(Index) & " :Erroneus UserID" & CRLF
                KillUser Index, sServer, "Invalid User Name/ID", , True, True
                Exit Sub
            End If
            
            If sACC = 2 Then
                For X = 0 To .List_SO.ListCount - 1
                    If LCase(iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) Like .List_SOA.List(X)) And LCase(iName(Index) = .List_SO.List(X)) Then
                        If LCase(iTP(Index) = .List_SOP.List(X)) Then
                            DUI = True
                            iAAC(Index) = True
                            Exit For
                        Else
                            Exit For
                        End If
                    End If
                Next X
                
                If DUI = False Then
                    KillUser Index, sServer, "Faild to login with correct user information", , True, True
                    Exit Sub
                End If
            End If
            
            If iUser(Index) = "" Then Exit Sub
            
            For X = 1 To sKline.Count
                If LCase(iName(Index) & "@" & iRHost(Index)) Like LCase(sKline(X)) Then
                    DUI = False
                    For Q = 1 To sEline.Count
                        If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                            DUI = True
                            Exit For
                        End If
                    Next Q
                    If DUI = False Then
                        SendData Index, "ERROR :You are banned from this server, Reason: " & sKlineR(X) & CRLF
                        KillUser Index, sServer, "(" & sKlineR(X) & "(KLINED))", , True, True
                        Exit Sub
                    End If
                    Exit For
                End If
            Next X
            
            For X = 1 To sAKill.Count
                If LCase(iName(Index) & "@" & iRHost(Index)) Like LCase(sAKill(X)) Then
                    DUI = False
                    For Q = 1 To sEline.Count
                        If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                            DUI = True
                            Exit For
                        End If
                    Next Q
                    If DUI = False Then
                        SendData Index, "ERROR :You are banned from this network, Reason: " & sAKillR(X) & CRLF
                        KillUser Index, sServer, "(" & sAKillR(X) & "(AKILLED))", , True, True
                        Exit Sub
                    End If
                    Exit For
                End If
            Next X
            
            For Q = 1 To iUserMax
                If iPeerFree(Q) = False Then
                    X = InStr(1, iModes(Q), "c")
                    If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has connected to server on port " & .Win(Index).LocalPort & CRLF
                End If
            Next Q
            
            SendData Index, _
                ":" & sServer & " 001 " & iUser(Index) & " :Welcome to the " & iNetName & " IRC Network " & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & CRLF & _
                ":" & sServer & " 002 " & iUser(Index) & " :Your Host is " & sServer & ", Running vbIRCd(IRCServ Clone) " & sVersion & " " & sRelease & CRLF & _
                ":" & sServer & " 003 " & iUser(Index) & " :This server was created " & sUDT & CRLF & _
                ":" & sServer & " 004 " & iUser(Index) & " :" & sServer & " vbIRCd-" & sVersion & " oOiwghskSaHANTcCfrxebWqBFI1dvtGz lvhopsmntikrRcaqOALQbSeKVfHGCuzN" & CRLF & _
                ":" & sServer & " 005 " & iUser(Index) & " :NOQUIT ISON USERS MODES=13 MAXCHANNELS=" & iChanMax & " MAXBANS=60 NICKLEN=30 TOPICLEN=307 KICKLEN=307 CHANTYPES=# PREFIX=(ohv)@%+ :are available on this server" & CRLF
            SendLUSERS Index
            SendMOTD Index
            iSignOn(Index) = GetTime
            
            .lbl_UC = .lbl_UC - 1
            .lbl_CU = .lbl_CU + 1
            If .lbl_HU = .lbl_CU - 1 Then .lbl_HU = .lbl_CU
            .lbl_CGU = .lbl_CGU + 1
            If .lbl_HGU = .lbl_CGU - 1 Then .lbl_HGU = .lbl_HGU + 1
            If iForceCloak = 1 Then uMode Index, iUser(Index), "+x"
        Else
                
            If iUserLevel(Index) = 0 Then TmpOp = "Normal User"
            If iUserLevel(Index) = 1 Then TmpOp = "Local IRC Operator"
            If iUserLevel(Index) = 2 Then TmpOp = "Global IRC Operator"
            If iUserLevel(Index) = 3 Then TmpOp = "Server Administrator"
            If iUserLevel(Index) = 4 Then TmpOp = "Services Administrator"
            If iUserLevel(Index) = 5 Then TmpOp = "Co Administrator"
            If iUserLevel(Index) = 6 Then TmpOp = "Technical Administrator"
            If iUserLevel(Index) = 7 Then TmpOp = "Network Administrator"
            
            TmpText = Mid$(Text, 6)
            If Not TmpText = "" Then
                For X = 1 To iUserMax
                    If LCase(TmpText) = LCase(iUser(X)) Then
                        Q = X
                        Exit For
                    End If
                Next X
            End If
            
            If iUserLevel(Q) = 0 Then TmpOp2 = "Normal User"
            If iUserLevel(Q) = 1 Then TmpOp2 = "Local IRC Operator"
            If iUserLevel(Q) = 2 Then TmpOp2 = "Global IRC Operator"
            If iUserLevel(Q) = 3 Then TmpOp2 = "Server Administrator"
            If iUserLevel(Q) = 4 Then TmpOp2 = "Services Administrator"
            If iUserLevel(Q) = 5 Then TmpOp2 = "Co Administrator"
            If iUserLevel(Q) = 6 Then TmpOp2 = "Technical Administrator"
            If iUserLevel(Q) = 7 Then TmpOp2 = "Network Administrator"
            iX = iSignOn(Index)
            
            If TmpText = "" Or UCase(TmpText) = UCase(iUser(Index)) Then
                SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :Your IRC user information:" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Nickname:    " & iUser(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Address:     " & iRHost(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     IP:          " & iIP(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     User ID:     " & iName(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Real Name:   " & iRealName(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     On Port:     " & .Win(Index).LocalPort & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Mode Flags:  " & iModes(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Hostmask:    " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     User Level:  " & TmpOp & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     MSG Count:   " & iFM(Index) & "(of " & iFloodMSGs & " allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     CMD Count:   " & iFC(Index) & "(of " & iFloodCMDs & " allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Bad Passes:  " & iFP(Index) & "(of 3 allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Signed On:   " & sUnixDate(iX) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Idle Time:   " & sCTS(iIdle(Index)) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                                ":" & sServer & " 374 " & iUser(Index) & " :End of your IRC user information" & CRLF
            Else
                If Not iUserLevel(Index) = 0 Then
                If Q = 0 Then SendData Index, ":" & sServer & " 374 " & iUser(Index) & " :User " & TmpText & " doesn't exist, looking for a ghost?" & CRLF: Exit Sub
                iX = iSignOn(Q)
                
                SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :" & iUser(Q) & "'s IRC user information:" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Nickname:    " & iUser(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Address:     " & iRHost(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     IP:          " & iIP(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     User ID:     " & iName(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Real Name:   " & iRealName(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     On Port:     " & .Win(Q).LocalPort & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Mode Flags:  " & iModes(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Hostmask:    " & iUser(Q) & "!" & iName(Q) & "@" & iHost(Q) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     User Level:  " & TmpOp2 & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     MSG Count:   " & iFM(Q) & "(of " & iFloodMSGs & " allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     CMD Count:   " & iFC(Q) & "(of " & iFloodCMDs & " allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Bad Passes:  " & iFP(Q) & "(of 3 allowed)" & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Signed On:   " & sUnixDate(iX) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " :     Idle Time:   " & sCTS(iIdle(Q)) & CRLF & _
                                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                                ":" & sServer & " 374 " & iUser(Index) & " :End of " & iUser(Q) & "'s IRC user information" & CRLF
                Else
                    If Q = 0 Then SendData Index, ":" & sServer & " 374 " & iUser(Index) & " :User " & TmpText & " doesn't exist, looking for a ghost?" & CRLF: Exit Sub
                    SendData Index, ":" & sServer & " 374 " & iUser(Index) & " :Trying to Spy on " & TmpText & " are we?  What a peep'n tom you are!" & CRLF
                End If
            End If
        End If
        Exit Sub
    End If
    
    If sACC = 2 Then
        If iAAC(Index) = False Then
            KillUser Index, sServer, "Connection access Denied for client", , True & CRLF
            Exit Sub
        End If
    End If
    
    If iRHost(Index) = "" Then
        If uHaveMode(Index, "D") = True Then SendData2 Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: No Host '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & CRLF
        SendData Index, ":" & sServer & " 451 " & iUser(Index) & " :" & sConvertText(ERR_451, Index) & CRLF
        KillUser Index, sServer, "User not registered(ERROR: No Host '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & ")", , True
        Exit Sub
    End If
    
    If iUser(Index) = "" Then
        If uHaveMode(Index, "D") = True Then SendData2 Index, ":" & sServer & " NOTICE AUTH :ERROR: No Nick! '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & CRLF
        SendData Index, ":" & sServer & " 451 " & iUser(Index) & " :" & sConvertText(ERR_451, Index) & CRLF
        KillUser Index, sServer, "User not registered(ERROR: No Nick! '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & ")", , True
        Exit Sub
    End If
    
    If iName(Index) = "" Then
        If uHaveMode(Index, "D") = True Then SendData2 Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: No UserID/Ident '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & CRLF
        SendData Index, ":" & sServer & " 451 " & iUser(Index) & " :" & sConvertText(ERR_451, Index) & CRLF
        KillUser Index, sServer, "User not registered(ERROR: No UserID/Ident '" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & "'" & ")", , True
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "MOTD " Then
        xMOTD = xMOTD + 1
        xMOTD2 = xMOTD2 + Len(Text)
        SendMOTD Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "LICENSE " Then
        'xMOTD = xMOTD + 1
        'xMOTD2 = xMOTD2 + Len(Text)
        SendLICENSE Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "LUSERS " Then
        xLUSERS = xLUSERS + 1
        xLUSERS2 = xLUSERS2 + Len(Text)
        SendLUSERS Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "ISON " Then
        xISON = xISON + 1
        xISON2 = xISON2 + Len(Text)
        TmpText = Mid$(Text, 6)
        sCISON Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 9) = "USERHOST " Then
        xUSERHOST = xUSERHOST + 1
        xUSERHOST2 = xUSERHOST2 + Len(Text)
        TmpText = Mid$(Text, 10)
        sCUserHost Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "REHASH " Then
        xREHASH = xREHASH + 1
        xREHASH2 = xREHASH2 + Len(Text)
        X = InStr(8, Text, " ")
        TmpText = Mid$(Text, 8, X - 8)
        sCRehash TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "STATS " Then
        xSTATS = xSTATS + 1
        xSTATS2 = xSTATS2 + Len(Text)
        SendStats Index, Text
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "INFO " Then
        xINFO = xINFO + 1
        xINFO2 = xINFO2 + Len(Text)
        SendINFO (Index)
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "MISC " Then
        SendMISC Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "LOGO " Then
        SendLOGO Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "SUMMON " Then
        xSUMMON = xSUMMON + 1
        xSUMMON2 = xSUMMON2 + Len(Text)
        SendData Index, ":" & sServer & " 445 " & iUser(Index) & " :" & sConvertText(ERR_445, Index) & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "ADMIN " Then
        xADMIN = xADMIN + 1
        xADMIN2 = xADMIN2 + Len(Text)
        SendData Index, ":" & sServer & " 256 " & TmpGet & " " & sServer & " :" & sConvertText(RPL_256, Index) & CRLF & _
                        ":" & sServer & " 257 " & iUser(Index) & " :" & iServAdminLine1 & CRLF & _
                        ":" & sServer & " 258 " & iUser(Index) & " :" & iServAdminLine2 & CRLF & _
                        ":" & sServer & " 259 " & iUser(Index) & " :" & iServAdminLine3 & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "VERSION " Then
        xVERSION = xVERSION + 1
        xVERSION2 = xVERSION2 + Len(Text)
        SendData Index, ":" & sServer & " 351 " & iUser(Index) & " vbIRCd/v" & App.Major & "." & App.Minor & "." & App.Revision & "(" & sRelease & ") " & sServer & " :Open Source" & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "PRIVMSG " Then
        If iFM(Index) > iFloodMSGs Then KillUser Index, sServer, "Message Flooding": Exit Sub
        If Not iFloodMSGs = 0 And iUserLevel(Index) = 0 Then iFM(Index) = iFM(Index) + 1: If Not iFloodCMDs = 0 Then iFC(Index) = iFC(Index) - 1
        If Not iFloodMSGs = 0 And iFM(Index) = iFloodMSGs Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** WARNING: DO NOT SEND ONE MORE PRIVMSG OR NOTICE FOR ONE MINUTE OR ELSE YOU WILL BE DISCONNECTED FOR FLOODING!!!" & CRLF
        xPRIVMSG = xPRIVMSG + 1
        xPRIVMSG2 = xPRIVMSG2 + Len(Text)
        SendPrivMSG Index, Text
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "NOTICE " Then
        If iFM(Index) > iFloodMSGs Then KillUser Index, sServer, "Message Flooding": Exit Sub
        If Not iFloodMSGs = 0 And iUserLevel(Index) = 0 Then iFM(Index) = iFM(Index) + 1: If Not iFloodCMDs = 0 Then iFC(Index) = iFC(Index) - 1
        If Not iFloodMSGs = 0 And iFM(Index) = iFloodMSGs Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :*** WARNING: DO NOT SEND ONE MORE PRIVMSG OR NOTICE FOR ONE MINUTE OR ELSE YOU WILL BE DISCONNECTED FOR FLOODING!!!" & CRLF
        xNOTICE = xNOTICE + 1
        xNOTICE2 = xNOTICE2 + Len(Text)
        SendNOTICE Index, Text
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "KILL " Then
        xKILL = xKILL + 1
        xKILL2 = xKILL2 + Len(Text)
        TmpText = Mid$(Text, 6)
        uKILL TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 4) = "DIE " Then
        xDIE = xDIE + 1
        xDIE2 = xDIE2 + Len(Text)
        TmpText = Mid$(Text, 5)
        sDIE TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "KLINE " Then
        xKLINE = xKLINE + 1
        xKLINE2 = xKLINE2 + Len(Text)
        TmpText = Mid$(Text, 7)
        sCKline TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "WHOIS " Then
        xWHOIS = xWHOIS + 1
        xWHOIS2 = xWHOIS2 + Len(Text)
        TmpText = Mid$(Text, 7)
        sCWhoIs TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 4) = "WHO " Then
        'xWHO = xWHO + 1
        'xWHO2 = xWHO2 + Len(Text)
        TmpText = Mid$(Text, 5)
        sCWho TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "WHOWAS " Then
        'xWHOWAS = xWHOWAS + 1
        'xWHOWAS2 = xWHOWAS2 + Len(Text)
        TmpText = Mid$(Text, 8)
        sCWhoWas TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "INVITE " Then
        'xINVITE = xINVITE + 1
        'xINVITE2 = xINVITE2 + Len(Text)
        TmpText = Mid$(Text, 8)
        uINVITE TmpData, TmpLoad, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "AWAY " Then
        xAWAY = xAWAY + 1
        xAWAY2 = xAWAY2 + Len(Text)
        TmpText = Mid$(Text, 6)
        sCAway TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "AKILL " Then
        xAKILL = xAKILL + 1
        xAKILL2 = xAKILL2 + Len(Text)
        TmpText = Mid$(Text, 7)
        sCAKill TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "RAKILL " Then
        xRAKILL = xRAKILL + 1
        xRAKILL2 = xRAKILL2 + Len(Text)
        TmpText = Mid$(Text, 8)
        sCRAKILL TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "UNKLINE " Then
        xUNKLINE = xUNKLINE + 1
        xUNKLINE2 = xUNKLINE2 + Len(Text)
        TmpText = Mid$(Text, 9)
        sCUNKLINE TmpText, Index
        Exit Sub
    End If
        
    If Left$(UCase(Text), 6) = "LINKS " Then
        xLINKS = xLINKS + 1
        xLINKS2 = xLINKS2 + Len(Text)
        TmpText = Mid$(Text, 7)
        SendLINKS Index, TmpText
        Exit Sub
    End If
    
    '[>--------------------------------------<]
    
    If Left$(UCase(Text), 6) = "NAMES " Then
        xNAMES = xNAMES + 1
        xNAMES2 = xNAMES2 + Len(Text)
        TmpText = Mid$(Text, 7)
        SendNAMES Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "LIST " Then
        xNAMES = xNAMES + 1
        xNAMES2 = xNAMES2 + Len(Text)
        TmpText = Mid$(Text, 6)
        SendLIST Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "PING " Then
        xPING = xPING + 1
        xPING2 = xPING2 + Len(Text)
        TmpText = Mid$(Text, 6)
        'Ping Reply Code goes here :P
        Exit Sub
    End If
    '[>--------------------------------------<]
        
    If Left$(UCase(Text), 8) = "RESTART " Then
        xRESTART = xRESTART + 1
        xRESTART2 = xRESTART2 + Len(Text)
        TmpText = Mid$(Text, 9)
        sRestart TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "PONG " Then
        xPONG = xPONG + 1
        xPONG2 = xPONG2 + Len(Text)
        TmpText = Mid$(Text, 1, Len(Text) - 1)
        TmpText = Mid$(TmpText, 7)
        If LCase(TmpText) = LCase(sServer) Then iPing(Index) = 0
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "KICK " Then
        'xkick = xkick + 1
        'xkick2 = xkick2 + Len(Text)
        TmpText = Mid$(Text, 6)
        X = InStr(1, TmpText, " ")
        TmpLoad = Mid$(TmpText, 1, X - 1)
        TmpText = Mid$(TmpText, X + 1)
        X = InStr(1, TmpText, " ")
        TmpData = Mid$(TmpText, 1, X - 1)
        TmpText = Mid$(TmpText, X + 2)
        If TmpText = "" Then TmpText = iUser(Index) & " "
        uKICK TmpLoad, TmpData, Mid$(TmpText, 1, Len(TmpText) - 1), Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "QUIT " Then
        xQUIT = xQUIT + 1
        xQUIT2 = xQUIT2 + Len(Text)
        TmpText = Mid$(Text, 1, Len(Text) - 1)
        TmpText = Mid$(TmpText, 7)
        If TmpText = "" Then TmpText = "Client Exiting"
        
        .Win(Index).Disconnect
        iPeerFree(Index) = True
        
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "c")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has Quit IRC (Reason: " & TmpText & ")" & CRLF
            End If
        Next Q
        
        'For Y = 1 To iUserMax
        '    If iPeerFree(Y) = False Then
        '        SendData Y, ":" & iUser(Index) & "!" & iName(Index) & "@" & irHost(Index) & " QUIT :QUIT: " & TmpText & CRLF
        '    End If
        'Next Y
        
        UserClosed Index, "Quit: " & TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "USERS " Then
        xUSERS = xUSERS + 1
        xUSERS2 = xUSERS2 + Len(Text)
        SendData Index, ":" & sServer & " 392 " & iUser(Index) & " :UserID(Nick)  -  Port  -  Hostmask" & CRLF & _
                        ":" & sServer & " 393 " & iUser(Index) & " : " & CRLF
        
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                If iUserLevel(Index) = 0 Then
                    If InStr(1, iModes(Q), "i") = 0 Then
                        SendData Index, ":" & sServer & " 393 " & iUser(Index) & " :" & iName(Q) & "(" & iUser(Q) & ") - " & .Win(Q).LocalPort & " - " & iUser(Q) & "!" & iName(Q) & "@" & iHost(Q) & CRLF
                    End If
                Else
                    SendData Index, ":" & sServer & " 393 " & iUser(Index) & " :" & iName(Q) & "(" & iUser(Q) & ") - " & .Win(Q).LocalPort & " - " & iUser(Q) & "!" & iName(Q) & "@" & iRHost(Q) & CRLF
                End If
            End If
        Next Q
        SendData Index, ":" & sServer & " 393 " & iUser(Index) & " : " & CRLF & _
                        ":" & sServer & " 394 " & iUser(Index) & " :End of /USERS" & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "JOIN " Then
        xJOIN = xJOIN + 1
        xJOIN2 = xJOIN2 + Len(Text)
        TmpText = Mid(Text, 6)
        uJOIN iUser(Index), Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "PART " Then
        xPART = xPART + 1
        xPART2 = xPART2 + Len(Text)
        TmpText = Mid(Text, 6)
        uPART iUser(Index), Index, TmpText
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "MODE " Then
        xMODE = xMODE + 1
        xMODE2 = xMODE2 + Len(Text)
        Q = InStr(6, Text, " ")
        TmpText = Mid$(Text, 6, Q - 6)
        Y = InStr(Q + 1, Text, " ")
        TmpFlag = Mid$(Text, Q + 1, Y - Q)
        TmpData = Mid$(Text, Y + 1)
        TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
        If TmpText = "" Then
            SendData Index, ":" & sServer & " 461 " & iUser(Index) & " MODE :Not enough parameters" & CRLF
            Exit Sub
        End If
        If Not TmpFlag = "" Then TmpFlag = Mid$(TmpFlag, 1, Len(TmpFlag) - 1)
        If Left(TmpText, 1) = "#" Then
            cMODE Index, TmpText, TmpFlag, TmpData
        Else
            uMode Index, TmpText, TmpFlag
        End If
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "TIME " Then
        xTIME = xTIME + 1
        xTIME2 = xTIME2 + Len(Text)
        SendTime (Index)
        Exit Sub
    End If
    
    If Left$(UCase(Text), 6) = "TOPIC " Then
        xTOPIC = xTOPIC + 1
        xTOPIC2 = xTOPIC2 + Len(Text)
        Text = Mid$(Text, 7)
        X = InStr(1, Text, " ")
        TmpText = Mid$(Text, 1, X - 1)
        TmpData = Mid$(Text, X + 1, Len(Text) - 1 - X)
        
        sChanTOPIC Index, TmpText, TmpData
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "SETHOST " Then
        xSETHOST = xSETHOST + 1
        xSETHOST2 = xSETHOST2 + Len(Text)
        TmpText = Mid$(Text, 9)
        sCSetHost TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 9) = "SETIDENT " Then
        xSETIDENT = xSETIDENT + 1
        xSETIDENT2 = xSETIDENT2 + Len(Text)
        TmpText = Mid$(Text, 10)
        sCSetIdent TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "SETNAME " Then
        xSETNAME = xSETNAME + 1
        xSETNAME2 = xSETNAME2 + Len(Text)
        TmpText = Mid$(Text, 9)
        sCSetName TmpText, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "CHGHOST " Then
        xCHGHOST = xCHGHOST + 1
        xCHGHOST2 = xCHGHOST2 + Len(Text)
        X = InStr(9, Text, " ")
        TmpText = Mid$(Text, 9, X - 9)
        TmpData = Mid$(Text, X + 1)
        sCCHGHost TmpText, TmpData, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 9) = "CHGIDENT " Then
        xCHGIDENT = xCHGIDENT + 1
        xCHGIDENT2 = xCHGIDENT2 + Len(Text)
        X = InStr(10, Text, " ")
        TmpText = Mid$(Text, 10, X - 10)
        TmpData = Mid$(Text, X + 1)
        sCCHGIdent TmpText, TmpData, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 8) = "CHGNAME " Then
        xCHGNAME = xCHGNAME + 1
        xCHGNAME2 = xCHGNAME2 + Len(Text)
        X = InStr(9, Text, " ")
        TmpText = Mid$(Text, 9, X - 9)
        TmpData = Mid$(Text, X + 1)
        sCCHGName TmpText, TmpData, Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "SHUN " Then
        xSHUN = xSHUN + 1
        xSHUN2 = xSHUN2 + Len(Text)
        'Code It!
        SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :SHUN, Coming soon to a working version near you! =)" & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "SAJOIN " Then
        xSAJOIN = xSAJOIN + 1
        xSAJOIN2 = xSAJOIN2 + Len(Text)
        SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Only lamers try to use this one ;p" & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "SAPART " Then
        xSAPART = xSAPART + 1
        xSAPART2 = xSAPART2 + Len(Text)
        SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Only lamers try to use this one ;p" & CRLF
        Exit Sub
    End If
    
    If Left$(UCase(Text), 7) = "STATUS " Then
        xSTATUS = xSTATUS + 1
        xSTATUS2 = xSTATUS2 + Len(Text)
        SendStatus Index
        Exit Sub
    End If
    
    If Left$(UCase(Text), 5) = "OPER " Then
        xOPER = xOPER + 1
        xOPER2 = xOPER2 + Len(Text)
        Q = InStr(6, Text, " ")
        Y = InStr(Q + 1, Text, " ")
        TmpText = Mid$(Text, 6, Q - 6)
        TmpData = Mid$(Text, Q + 1, Y - Q - 1)
        If TmpText = "" Then DUI = True
        If TmpData = "" Then DUI = True
        If DUI = True Then
            SendData Index, ":" & sServer & " 461 " & iUser(Index) & " OPER :Not enough parameters" & CRLF
            Exit Sub
        End If
        uLoginOper TmpText, TmpData, Index
        Exit Sub
    End If
    
    'LogIt "Unknown CMD-> '" & Text & "'"
    Y = InStr(1, Text, " ")
    If Not Y = 0 Then
        TmpLoad = Mid$(Text, 1, Y - 1)
    Else
        TmpLoad = Text
    End If
    SendData Index, ":" & sServer & " 421 " & iUser(Index) & " :" & UCase(TmpLoad) & " Unknown command" & CRLF
    '":" & sServer & " 005 " & iUser(Index) & " : " & CRLF
    '":" & sServer & " 371 " & iUser(Index) & " : " & CRLF  <-- Any Info Line Code
    'Chr$(34) is "
    ':[Nick]![ID]@[Host] QUIT :QUIT: [Quit Message]
    End With
End Sub


