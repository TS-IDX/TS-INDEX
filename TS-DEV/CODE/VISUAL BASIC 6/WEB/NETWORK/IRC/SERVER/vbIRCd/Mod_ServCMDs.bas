Attribute VB_Name = "Mod_ServCMDs"
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
' This Module has the following Subs:
' uMode            [User Modes]
' SendData         [Socket SendOut]
' SendData2        [Socket SendOut]
' SendTime         [Time Command]
' SendLUSERS       [Server Status]
' SendStats        [Stats Report]
' SendINFO         [Server InfoOut]
' SendNOTICE       [NOTICE MsgOut]
' SendPrivMSG      [Private MsgOut]
' KillUser         [KillUser Operation]
' UserClosed       [Close User Socket]
' uKILL            [KIll Message]
' sDIE             [Shutdown Msg]
' sRestart         [Restart Msg]
' sCAKILL          [Set AKill Ban]
' sCKline          [Set Kline Ban]
' sCRAKILL         [Del AKILL ban]
' sCUnKline        [Del Kline Ban]
' sCWhoIs          [Whois Query]
' sCAway           [Away Setting]
' SendLINKS        [Send Network Links]
' sCISON           [IS User ON Request]
' sCUserHost       [UserHost Request]

Sub uMode(Index As Integer, User As String, Optional Modes As String, Optional OverRide As Boolean)
On Error Resume Next
Dim iAdd As Boolean
Dim Z As Integer
Dim SFlags As String
Dim TmpFlag As String
Dim TmpText As String
Dim TmpData As String
Dim iAdded As String
Dim iRemoved As String
Dim Y As Integer
Dim iPass As Boolean
    iAdd = True
    
If Not Modes = "" Then
    If UCase(User) = UCase(iUser(Index)) Then
    
DoAgain:
        TmpFlag = Mid$(Modes, 1, 1)
        iPass = False
        
        If TmpFlag = "+" Then
            iAdd = True
            Modes = Mid$(Modes, 2)
            GoTo DoAgain
        Else
            If TmpFlag = "-" Then
                iAdd = False
                Modes = Mid$(Modes, 2)
                GoTo DoAgain
            End If
        End If
        
        Z = InStr(1, "DoOiwghskSaHANTcCfrxebWqBFI1dvtG", TmpFlag)
        If Z = 0 Then iPass = False Else: iPass = True
        
        If iPass = False Then
            SendData Index, ":" & sServer & " 501 " & iUser(Index) & " " & TmpFlag & " :" & sConvertText(ERR_501, Index) & CRLF
            Modes = Mid$(Modes, 2)
            If Modes = "" Then GoTo Done
            GoTo DoAgain
        End If
        
        If iAdd = True Then
            Y = InStr(1, iModes(Index), TmpFlag)
            If Y = 0 Then
                iPass = True
                Z = InStr(1, "AaOoghNTCcfrebWqFI1", TmpFlag)
                If Not Z = 0 Then iPass = False
                If Not iUserLevel(Index) = 0 Then iPass = True
                If TmpFlag = "r" Then iPass = False
                If OverRide = True Then iPass = True
                'AaOoghNTCcfrebWqFI1
                If iPass = True Then
                    If TmpFlag = "i" Then frmMain.lbl_IU = frmMain.lbl_IU + 1
                    If TmpFlag = "D" Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :WARNING: You have just actived Debug mode on your self which means the server will send you data when using some commands and this could become annoying quick!  =D" & CRLF
                    If TmpFlag = "x" Then iHost(Index) = sCloakHost(Index)
                    iModes(Index) = iModes(Index) & TmpFlag
                    iAdded = iAdded & TmpFlag
                End If
            End If
            Modes = Mid$(Modes, 2)
            If Modes = "" Then GoTo Done
            GoTo DoAgain
        Else
            Y = InStr(1, iModes(Index), TmpFlag)
            If Not Y = 0 Then
                iPass = True
                If TmpFlag = "r" Then iPass = False
                If OverRide = True Then iPass = True
                If iPass = True Then
                    If Y = 1 Then
                        iModes(Index) = Mid$(iModes(Index), 2)
                        iRemoved = iRemoved & TmpFlag
                    Else
                        TmpText = Mid$(iModes(Index), 1, Y - 1)
                        TmpData = Mid$(iModes(Index), Y + 1)
                        iModes(Index) = TmpText & TmpData
                        iRemoved = iRemoved & TmpFlag
                    End If
                    If TmpFlag = "i" Then frmMain.lbl_IU = frmMain.lbl_IU - 1
                    If TmpFlag = "D" Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :You now have been deactived from Debug mode info reports, so now your free from annoying debug messages  =)" & CRLF
                    If TmpFlag = "x" Then iHost(Index) = iRHost(Index)
                End If
            End If
            Modes = Mid$(Modes, 2)
            If Modes = "" Then GoTo Done
            GoTo DoAgain
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF
    End If
Else
    SendData2 Index, ":" & sServer & " 221 " & iUser(Index) & " :+" & iModes(Index) & CRLF
End If
    
    Exit Sub
Done:
    If Not iAdded = "" Then
        If Not iRemoved = "" Then
            SFlags = "+" & iAdded & "-" & iRemoved
        Else
            SFlags = "+" & iAdded
        End If
    Else
        If Not iRemoved = "" Then
            SFlags = "-" & iRemoved
        End If
    End If
    If Not SFlags = "" Then SendData2 Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & " MODE " & iUser(Index) & " :" & SFlags & CRLF
End Sub

Sub SendData(Index As Integer, Data As String)
On Error Resume Next
    sBSent = sBSent + Len(Data)
    sTSent = sTSent + 1
    frmMain.Win(Index).Write Data, Len(Data)
End Sub

Sub SendData2(Index As Integer, Data As String)
On Error Resume Next
    sBSent = sBSent + Len(Data)
    sTSent = sTSent + 1
    frmMain.Win(Index).Write Data, Len(Data)
End Sub

Public Sub SendMOTD(Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim IDL As String
Dim iNews As String
Dim X
iNews = frmMain.txt_CMOTD.Text
    If iNews = "" Then SendData Index, ":" & sServer & " 422 " & iUser(Index) & " :" & sConvertText(ERR_422, Index) & CRLF: Exit Sub
    
    SendData Index, ":" & sServer & " 375 " & iUser(Index) & " :" & sServer & " Message Of The Day" & CRLF
ReSend:

X = InStr(1, iNews, CRLF)
    If X = 0 Then
        IDL = iNews
        iNews = ""
    Else
        IDL = Mid$(iNews, 1, X - 1)
        iNews = Mid$(iNews, X + 2)
    End If
    SendData Index, ":" & sServer & " 372 " & iUser(Index) & " :- " & sConvertText(IDL, Index) & CRLF
    
    If X = 0 Then
        SendData Index, ":" & sServer & " 376 " & iUser(Index) & " :End of /MOTD Command" & CRLF
    Else
        GoTo ReSend
    End If
    
    '":" & sServer & " 375 " & iUser(Index) & " :" & sServer & " Message Of The Day" & CRLF
    '":" & sServer & " 372 " & iUser(Index) & " :- " & iDL & CRLF
    '":" & sServer & " 376 " & iUser(Index) & " :End of /MOTD Command" & CRLF
End Sub

Sub SendLUSERS(Index As Integer)
On Error Resume Next
    frmMain.lbl_CC = iChanSys.iChan.Count
    With frmMain
        SendData Index, _
        ":" & sServer & " 251 " & iUser(Index) & " :There are " & .lbl_CGU - .lbl_IU & " users and " & .lbl_IU & " invisible on " & .lbl_CS + 1 & " servers" & CRLF & _
        ":" & sServer & " 252 " & iUser(Index) & " " & .lbl_CO & " :" & RPL_252 & CRLF & _
        ":" & sServer & " 253 " & iUser(Index) & " " & .lbl_UC & " :" & RPL_253 & CRLF & _
        ":" & sServer & " 254 " & iUser(Index) & " " & .lbl_CC & " :" & RPL_254 & CRLF & _
        ":" & sServer & " 255 " & iUser(Index) & " :I have " & .lbl_CU & " clients and " & .lbl_CS & " servers" & CRLF & _
        ":" & sServer & " 265 " & iUser(Index) & " :Current Local Users: " & .lbl_CU & "  Max: " & .lbl_HU & CRLF & _
        ":" & sServer & " 266 " & iUser(Index) & " :Current Global Users: " & .lbl_CGU & "  Max: " & .lbl_HGU & CRLF
    End With
End Sub

Sub SendTime(Index As Integer)
On Error Resume Next
Dim TmpData As String

    TmpData = Format(Now, "Long Time") & "   " & Format(Now, "Long Date")
    SendData Index, ":" & sServer & " 391 " & iUser(Index) & " :Local Server Time and Date: " & TmpData & CRLF & _
                    ":" & sServer & " 391 " & iUser(Index) & " :Unix Time:  " & sUnixDate(GetTime) & " - RAW: " & GetTime & CRLF
End Sub

Sub SendStatus(Index As Integer)
On Error Resume Next
    SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :STATUS Report of " & sServer & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " :Bytes Sent,Receive: " & sBSent & "," & sBReceived & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " :Times Sent,Receive: " & sTSent & "," & sTReceived & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " :Server Up:          " & frmMain.lbl_UT & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " :Connect Count:      " & sCCount & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " :Conn Acc,Ref,Kill:  " & sCAccepted & "," & sCRefused & "," & sCKilled & CRLF & _
                    ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                    ":" & sServer & " 374 " & iUser(Index) & " :End of /STATUS report" & CRLF
End Sub

Sub SendStats(Index As Integer, Text As String)
On Error Resume Next
Dim TmpFlag As String
Dim X As Integer
Dim Y As Integer
Dim Q As Integer
Dim Z As Integer
Dim TmpData As String
Dim TmpText As String
Dim TmpLoad As String
Dim DFY As Boolean

    TmpFlag = Mid$(Text, 7, 1)
    DFY = False
    If TmpFlag = "" Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " STATS :Not enough parameters" & CRLF
        DFY = True
    End If
    
    If UCase(TmpFlag) = "U" Then
        SendData Index, ":" & sServer & " 242 " & iUser(Index) & " :Server Up " & frmMain.lbl_UT & CRLF & _
                        ":" & sServer & " 250 " & iUser(Index) & " :" & sConvertText(RPL_250, Index) & CRLF & _
                        ":" & sServer & " 219 " & iUser(Index) & " " & TmpFlag & " :" & RPL_219 & CRLF
        DFY = True
    End If
    
    If TmpFlag = "Q" Then
        For X = 1 To sQline.Count
            SendData Index, ":" & sServer & " 217 " & iUser(Index) & " Q <NULL> " & Replace$(sQlineR(X), " ", "_", , , vbTextCompare) & " " & sQline(X) & CRLF
        Next X
        ':linux.ircd-net.org 217 RedFile Q <NULL> Reserved_for_services *C*h*a*n*S*e*r*v* 0 -1
    End If
    
    If UCase(TmpFlag) = "O" Then
        For X = 0 To frmMain.List_SO.ListCount - 1
            SendData Index, ":" & sServer & " 243 " & iUser(Index) & " :O " & frmMain.List_SOA.List(X) & " * " & frmMain.List_SO.List(X) & " " & frmMain.List_SOM.List(X) & " 1" & CRLF
        Next X
        SendData Index, ":" & sServer & " 219 " & iUser(Index) & " " & TmpFlag & " :" & RPL_219 & CRLF
        DFY = True
    End If
    
    If UCase(TmpFlag) = "M" Then
        SendData Index, ":" & sServer & " 212 " & iUser(Index) & " :MODE " & xMODE & " " & xMODE2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :USER " & xUSER & " " & xUSER2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :OPER " & xOPER & " " & xOPER2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :MOTD " & xMOTD & " " & xMOTD2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NICK " & xNICK & " " & xNICK2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :USERS " & xUSERS & " " & xUSERS2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :STATS " & xSTATS & " " & xSTATS2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :ADMIN " & xADMIN & " " & xADMIN2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :INFO " & xINFO & " " & xINFO2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :VERION " & xVERSION & " " & xVERSION2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :JOIN " & xJOIN & " " & xJOIN2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :PART " & xPART & " " & xPART2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NAMES " & xNAMES & " " & xNAMES2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :TIME " & xTIME & " " & xTIME2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :QUIT " & xQUIT & " " & xQUIT2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NOTICE " & xNOTICE & " " & xNOTICE2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :PRIVMSG " & xPRIVMSG & " " & xPRIVMSG2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :ISON " & xISON & " " & xISON2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :USERHOST " & xUSERHOST & " " & xUSERHOST2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :PING " & xPING & " " & xPING2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :TOPIC " & xTOPIC & " " & xTOPIC2 & CRLF
        
        SendData Index, ":" & sServer & " 212 " & iUser(Index) & " :SETHOST " & xSETHOST & " " & xSETHOST2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :SETIDENT " & xSETIDENT & " " & xSETIDENT2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :SETNAME " & xSETNAME & " " & xSETNAME2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :CHGHOST " & xCHGHOST & " " & xCHGHOST2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :CHGIDENT " & xCHGIDENT & " " & xCHGIDENT2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :CHGNAME " & xCHGNAME & " " & xCHGNAME2 & CRLF
        
        SendData Index, ":" & sServer & " 212 " & iUser(Index) & " :PONG " & xPONG & " " & xPONG2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :LUSERS " & xLUSERS & " " & xLUSERS2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :AWAY " & xAWAY & " " & xAWAY2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :WHOIS " & xWHOIS & " " & xWHOIS2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :AKILL " & xAKILL & " " & xAKILL2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :RAKILL " & xRAKILL & " " & xRAKILL2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :KLINE " & xKLINE & " " & xKLINE2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :UNKLINE " & xUNKLINE & " " & xUNKLINE2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :DIE " & xDIE & " " & xDIE2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :RESTART " & xRESTART & " " & xRESTART2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :PASS " & xPASS & " " & xPASS2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :SUMMON " & xSUMMON & " " & xSUMMON2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :REHASH " & xREHASH & " " & xREHASH2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NICKSERV " & xNICKSERV & " " & xnikcserv2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :CHANSERV " & xCHANSERV & " " & xCHANSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :MEMOSERV " & xMEMOSERV & " " & xMEMOSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :OPERSERV " & xOPERSERV & " " & xOPERSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :STATSERV " & xSTATSERV & " " & xSTATSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :INFOSERV " & xINFOSERV & " " & xINFOSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NETSERV " & xNETSERV & " " & xNETSERV2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :SUPPORT " & xSUPPORT & " " & xSUPPORT2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :MAINCHAN " & xMAINCHAN & " " & xMAINCHAN2 & CRLF & _
                        ":" & sServer & " 212 " & iUser(Index) & " :NETWORK " & xNETWORK & " " & xNETWORK2 & CRLF & _
                        ":" & sServer & " 219 " & iUser(Index) & " " & TmpFlag & " :" & RPL_219 & CRLF
        DFY = True
    End If
    
    If UCase(TmpFlag) = "K" Then
        For X = 1 To sEline.Count
            Y = InStr(1, sEline(X), "@")
            TmpText = Mid$(sEline(X), Y + 1)
            TmpData = Mid$(sEline(X), 1, Y - 1)
            SendData Index, ":" & sServer & " 216 " & iUser(Index) & " :E " & TmpText & " " & Replace$(sElineR(X), " ", "_", , , vbTextCompare) & " " & TmpData & " 0 -1" & CRLF
            'K * Kahn_is_potato! *UnRegged 0 -1
            'K <host> <Reason> <username> <port> <class>
        Next X
        For X = 1 To sKline.Count
            Y = InStr(1, sKline(X), "@")
            TmpText = Mid$(sKline(X), Y + 1)
            TmpData = Mid$(sKline(X), 1, Y - 1)
            SendData Index, ":" & sServer & " 216 " & iUser(Index) & " :K " & TmpText & " " & Replace$(sKlineR(X), " ", "_", , , vbTextCompare) & " " & TmpData & " 0 -1" & CRLF
            'K * Kahn_is_potato! *UnRegged 0 -1
            'K <host> <Reason> <username> <port> <class>
        Next X
        For X = 1 To sAKill.Count
            Y = InStr(1, sAKill(X), "@")
            TmpText = Mid$(sAKill(X), Y + 1)
            TmpData = Mid$(sAKill(X), 1, Y - 1)
            SendData Index, ":" & sServer & " 216 " & iUser(Index) & " :A " & TmpText & " " & Replace$(sAKillR(X), " ", "_", , , vbTextCompare) & " " & TmpData & " 0 -1" & CRLF
            'A * Kahn_is_potato! *UnRegged 0 -1
            'A <host> <Reason> <username> <port> <class>
        Next X
    End If
    
    If DFY = False Then
        SendData Index, ":" & sServer & " 219 " & iUser(Index) & " " & TmpFlag & " :" & RPL_219 & CRLF
    End If
End Sub

Sub SendINFO(Index As Integer)
On Error Resume Next
SendData Index, ":" & sServer & " 351 " & iUser(Index) & " vbIRCd/v" & App.Major & "." & App.Minor & "." & App.Revision & "(" & sRelease & ") " & sServer & " :Open Source" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :Server Up " & frmMain.lbl_UT & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :Built and Developed by TRON 2001 Network®, The SDT (Server Development Team)" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     -Develop Class-    -Nick-     -Real Name-     -E-mail Address-" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     Main Programmer:    TRON       Nathan Martin   TRON@ircd-net.org" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :       Co Programmer:    eternal    N/A             eternal@observers.net" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     Software Tester:    Killmor    N/A             Killmor@t2n.dyndns.org" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     Some Ideas from:    Ransom     Sean Ouimet     souimet@telusplanet.net" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     Idea Coding of:     Zack       Zack Lantz      ripper@chemfreak.com" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :     Code Helper:        SpyMaster  Paul Heinlein   SpyMaster@ircplus.com" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :This version of IRC Serv, but stripped down is under the GPL" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :(General Public License) which you can read from the LICENSE.txt File" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :that came with vbIRCd or you can have it displayed by typing this:" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :/license.  Note that this is NOT IRC Serv, so don't ask for the code" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :of it since it's not under the GPL like this code is.  The only thing" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :that this code/software has to do anything with IRC Serv is that it did" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :come from the IRC Serv project which is not the full code of IRC Serv," & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :but stripped down to a more basic version/clone. So there for this is" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :not the real IRC Serv software." & CRLF & _
                ":" & sServer & " 374 " & iUser(Index) & " :End of /INFO List" & CRLF
End Sub

Sub SendPrivMSG(Index As Integer, Text As String)
On Error Resume Next
Dim TmpMsg As String
Dim TmpText As String
Dim Q As Integer
Dim X As Integer
Dim Z As Integer
Dim DUI As Boolean

    Text = Mid$(Text, 1, Len(Text) - 1)
    Y = InStr(10, Text, " ")
    TmpText = Mid$(Text, 9, Y - 9)
    TmpMsg = Mid$(Text, Y + 2)
    X = "-128"
    If TmpText = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " PRIVMSG :Not enough parameters" & CRLF: Exit Sub
    If TmpMsg = "" Then SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :" & sConvertText(ERR_412, Index) & CRLF: Exit Sub
            
    TmpMsg = SysFOCC(TmpMsg)
    iIdle(Index) = 0
            
    If Left$(TmpText, 1) = "#" Then
        sChanMessage TmpText, TmpMsg, Index
        Exit Sub
    Else
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                If LCase(TmpText) = LCase(iUser(Q)) Then
                    X = Q
                    Exit For
                End If
            End If
        Next Q
    End If
            
    If X = "-128" Then
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpText & " No such nick/channel" & CRLF
    Else
        'Q = InStr(1, iModes(X), "G")
        'If Q = 0 Then TmpMsg = SysFilter(TmpMsg) Else: TmpMsg = SysFilter(TmpMsg, True)
        ' Code has been stripped out for above to work, sorry...
        If Not TmpText <> "" Then
            SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :No text to send" & CRLF
        Else
            SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " PRIVMSG " & TmpText & " :" & TmpMsg & CRLF
            If Not iAway(X) = "" Then SendData Index, ":" & sServer & " 301 " & iUser(Index) & " " & iUser(X) & " :" & iAway(X) & CRLF
        End If
    End If
End Sub

Sub SendNOTICE(Index As Integer, Text As String)
On Error Resume Next
Dim TmpText As String
Dim TmpMsg As String
Dim Q As Integer
Dim X As Integer
Dim Z As Integer
Dim DUI As Boolean

    Text = Mid$(Text, 1, Len(Text) - 1)
    Y = InStr(10, Text, " ")
    TmpText = Mid$(Text, 8, Y - 8)
    TmpMsg = Mid$(Text, Y + 2)
    X = "-128"
    
    If TmpText = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " PRIVMSG :Not enough parameters" & CRLF: Exit Sub
    If TmpMsg = "" Then SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :" & sConvertText(ERR_412, Index) & CRLF: Exit Sub
    
    TmpMsg = SysFOCC(TmpMsg)
            
    If Left$(TmpText, 1) = "#" Then
        sChanNotice TmpText, TmpMsg, Index
        Exit Sub
    Else
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                If LCase(TmpText) = LCase(iUser(Q)) Then
                    X = Q
                    Exit For
                End If
            End If
        Next Q
    End If
    
    If X = "-128" Then
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpText & " No such nick/channel" & CRLF
    Else
        'Q = InStr(1, iModes(X), "G")
        'If Q = 0 Then TmpMsg = SysFilter(TmpMsg) Else: TmpMsg = SysFilter(TmpMsg, True)
        ' Code has been stripped out for above to work, sorry...
        If Not TmpText <> "" Then
            SendData Index, ":" & sServer & " 412 " & iUser(Index) & " :No text to send" & CRLF
        Else
            SendData X, ":" & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " NOTICE " & TmpText & " :" & TmpMsg & CRLF
        End If
    End If
End Sub

Sub uLoginOper(TmpText As String, TmpData As String, Index As Integer)
On Error Resume Next
Dim Y As Integer
Dim Q As Integer
Dim X As Integer
Dim TmpData2 As String
Dim TmpFlag As String
Dim TmpOp As String
        For Y = 0 To frmMain.List_SO.ListCount - 1
            If UCase(TmpText) = UCase(frmMain.List_SO.List(Y)) Then
                If TmpData = frmMain.List_SOP.List(Y) Then
                    If iName(Index) & "@" & iRHost(Index) Like frmMain.List_SOA.List(Y) Then
                        DFY = True
                        TmpData2 = frmMain.List_SOM.List(Y)
                        Exit For
                    Else
                        DFY = False
                        Exit For
                    End If
                Else
                    SendData Index, ":" & sServer & " 464 " & iUser(Index) & " :Password incorrect" & CRLF
                    Exit Sub
                    Exit For
                End If
            End If
        Next Y
        
        If DFY = True Then
            TmpText = ""
            TmpData2 = TmpData2 & "s"
ReScan:
            If Not TmpData2 = "" Then
                TmpFlag = Mid$(TmpData2, 1, 1)
                Q = InStr(1, iModes(Index), TmpFlag)
                If Q = 0 And TmpFlag = "a" Then TmpFlag = TmpFlag & "q"
                If Q = 0 Then TmpText = TmpText & TmpFlag
                TmpData2 = Mid$(TmpData2, 2)
                GoTo ReScan
            End If
                X = InStr(1, TmpText, "o")
                If Not X = 0 Then TmpOp = "Local IRC Operator": iUserLevel(Index) = 1
                X = InStr(1, TmpText, "O")
                If Not X = 0 Then TmpOp = "Global IRC Operator": iUserLevel(Index) = 2
                X = InStr(1, TmpText, "A")
                If Not X = 0 Then TmpOp = "Server Administrator": iUserLevel(Index) = 3
                X = InStr(1, TmpText, "a")
                If Not X = 0 Then TmpOp = "Services Administrator": iUserLevel(Index) = 4
                X = InStr(1, TmpText, "C")
                If Not X = 0 Then TmpOp = "Co Administrator": iUserLevel(Index) = 5
                X = InStr(1, TmpText, "T")
                If Not X = 0 Then TmpOp = "Technical Administrator": iUserLevel(Index) = 6
                X = InStr(1, TmpText, "N")
                If Not X = 0 Then TmpOp = "Network Administrator": iUserLevel(Index) = 7
            
            If Not TmpOp = "" Then
                iModes(Index) = iModes(Index) & TmpText
                SendData Index, ":" & iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index) & " MODE " & iUser(Index) & " :+" & TmpText & CRLF & _
                                ":" & sServer & " 381 " & iUser(Index) & " :You are now an IRC operator" & CRLF
                frmMain.lbl_CO = frmMain.lbl_CO + 1
                
                For Q = 1 To iUserMax
                    If iPeerFree(Q) = False Then
                        X = InStr(1, iModes(Q), "s")
                        If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") is now a " & TmpOp & CRLF
                    End If
                Next Q
            Else
                SendData Index, ":" & sServer & " 381 " & iUser(Index) & " :You are now an IRC operator" & CRLF
            End If
        Else
            SendData Index, ":" & sServer & " 491 " & iUser(Index) & " :No O-lines for your host" & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "O")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to perform OPER Login" & CRLF
            End If
        Next Q
        End If
End Sub

Sub uKILL(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 2)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If TmpData = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " KILL :Not enough parameters" & CRLF
        Exit Sub
    End If
    
ReDO_Loop:
    If Left$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 2)
    Else
        GoTo Loop_Done
    End If
    GoTo ReDO_Loop
Loop_Done:

ReDO_Loop2:
    If Right$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 1, Len(TmpText) - 1)
    Else
        GoTo Loop_Done2
    End If
    GoTo ReDO_Loop2
Loop_Done2:

    DUI = False
    If iUserLevel(Index) > 0 Then DUI = True
    
    If DUI = True Then
        Q = 0
        For X = 1 To iUserMax
            If LCase(TmpText) = LCase(iUser(X)) Then
                Q = X
                Exit For
            End If
        Next X
        If Not Q = 0 Then
            KillUser Q, iUser(Index), TmpData, Index
        Else
            SendData2 Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpText & " No such nick/channel" & CRLF
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "O")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to perform a User KILL" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sDIE(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " DIE :Not enough parameters" & CRLF
        Exit Sub
    End If
    
ReDO_Loop:
    If Left$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 2)
    Else
        GoTo Loop_Done
    End If
    GoTo ReDO_Loop
Loop_Done:

ReDO_Loop2:
    If Right$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 1, Len(TmpText) - 1)
    Else
        GoTo Loop_Done2
    End If
    GoTo ReDO_Loop2
Loop_Done2:

    If iUserLevel(Index) = 0 Then DUI = False Else: DUI = True
    
    If DUI = True Then
        If Not iDPass = TmpText Then
            SendData Index, ":" & sServer & " 464 " & iUser(Index) & " :Password incorrect" & CRLF
        Else
            frmMain.cSysRD False, TmpData, Index
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to perform a Server DIE" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sRestart(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " RESTART :Not enough parameters" & CRLF
        Exit Sub
    End If
    
ReDO_Loop:
    If Left$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 2)
    Else
        GoTo Loop_Done
    End If
    GoTo ReDO_Loop
Loop_Done:

ReDO_Loop2:
    If Right$(TmpText, 1) = " " Then
        TmpText = Mid$(TmpText, 1, Len(TmpText) - 1)
    Else
        GoTo Loop_Done2
    End If
    GoTo ReDO_Loop2
Loop_Done2:

    If iUserLevel(Index) = 0 Then DUI = False Else: DUI = True
    
    If DUI = True Then
        If Not iRPass = TmpText Then
            SendData Index, ":" & sServer & " 464 " & iUser(Index) & " :Password incorrect" & CRLF
        Else
            frmMain.cSysRD True, TmpData, Index
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "O")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to perform a Server RESTART" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sCAKill(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If TmpData = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " AKILL :Not enough parameters" & CRLF
        Exit Sub
    End If

    DUI = False
    If iUserLevel(Index) > 0 Then DUI = True
    
    If DUI = True Then
        DUI = True
        X = InStr(1, TmpText, "!")
        If Not X = 0 Then DUI = False
        X = InStr(1, TmpText, "@")
        If X = 0 Then DUI = False
        If DUI = False Then
            SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :You must enter a ban in this style ONLY:  user@host   Note: wildcards '*' are allowed to be used." & CRLF
            Exit Sub
        End If
        sAKill.Add TmpText
        sAKillR.Add TmpData
        
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- IRC Operator " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") set a Network Wide Ban(AKILL) for " & TmpText & " (Reason: " & TmpData & ")" & CRLF
            End If
        Next Q
        ScanFBU
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to set a Network Wide Ban(AKILL)" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sCRAKILL(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " RAKILL :Not enough parameters" & CRLF
        Exit Sub
    End If

    DUI = False
    If iUserLevel(Index) > 0 Then DUI = True
    
    If DUI = True Then
        DUI = False
        For X = 1 To sAKill.Count
            If LCase(TmpText) = LCase(sAKill(X)) Then
                sAKill.Remove X
                sAKillR.Remove X
                DUI = True
                Exit For
            End If
        Next X
        
        If DUI = True Then
            For Q = 1 To iUserMax
                If iPeerFree(Q) = False Then
                    X = InStr(1, iModes(Q), "s")
                    If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- IRC Operator " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has removed a Network Wide Ban(AKILL) for " & TmpText & CRLF
                End If
            Next Q
            
        Else
            SendData2 Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: No such ban set(" & TmpText & ", Not Found)" & CRLF
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to remove a set Network Wide Ban(AKILL)" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sCKline(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If TmpData = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " KLINE :Not enough parameters" & CRLF
        Exit Sub
    End If

    DUI = False
    If iUserLevel(Index) > 0 Then DUI = True
    
    If DUI = True Then
        DUI = True
        X = InStr(1, TmpText, "!")
        If Not X = 0 Then DUI = False
        X = InStr(1, TmpText, "@")
        If X = 0 Then DUI = False
        If DUI = False Then
            SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :You must enter a ban in this style ONLY:  user@host   Note: wildcards '*' are allowed to be used." & CRLF
            Exit Sub
        End If
        sKline.Add TmpText
        sKlineR.Add TmpData
        
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- IRC Operator " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") set a Local Server Ban(KLine) for " & TmpText & " (Reason: " & TmpData & ")" & CRLF
            End If
        Next Q
        
        ScanFBU
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to set a Local Server Ban(KLINE)" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sCUNKLINE(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpData As String
Dim DUI As Boolean

    X = InStr(1, Text, " ")
    TmpText = Mid$(Text, 1, X - 1)
    TmpData = Mid$(Text, X + 1)
    TmpData = Mid$(TmpData, 1, Len(TmpData) - 1)
    If TmpText = "" Then DUI = True
    If DUI = True Then
        SendData Index, ":" & sServer & " 461 " & iUser(Index) & " UNKLINE :Not enough parameters" & CRLF
        Exit Sub
    End If

    DUI = False
    If iUserLevel(Index) > 0 Then DUI = True
    
    If DUI = True Then
        DUI = False
        For X = 1 To sKline.Count
            If LCase(TmpText) = LCase(sKline(X)) Then
                sKline.Remove X
                sKlineR.Remove X
                DUI = True
                Exit For
            End If
        Next X
        
        If DUI = True Then
            For Q = 1 To iUserMax
                If iPeerFree(Q) = False Then
                    X = InStr(1, iModes(Q), "s")
                    If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- IRC Operator " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has removed a Local Server Ban(KLINE) for " & TmpText & CRLF
                End If
            Next Q
            
        Else
            SendData2 Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: No such ban set(" & TmpText & ", Not Found)" & CRLF
        End If
    Else
        SendData2 Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "s")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") faild a temp to remove a set Local Server Ban(KLINE)" & CRLF
            End If
        Next Q
    End If
End Sub

Sub sCRehash(Text As String, Index As Integer)
On Error Resume Next
Dim Q As Integer
Dim X As Integer
Dim DUI As Boolean
DUI = False
    If Left$(Text, 1) = ":" Then Text = Mid$(Text, 2)
    If iUserLevel(Index) = 0 Then X = 0 Else: X = 1
    If X = 0 Then
        SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & sConvertText(ERR_481, Index) & CRLF
    Else
        If LCase(Text) = LCase(sServer) Then DUI = True
        If Text = "" Then DUI = True
        If DUI = False Then
            'Net Code will soon be placed here for Remote Rehashing...
            SendData Index, ":" & sServer & " 402 " & iUser(Index) & " " & Text & " :" & sConvertText(ERR_402, Index) & CRLF
        Else
            SendData Index, ":" & sServer & " 382 " & iUser(Index) & " IRC_Serv.Conf :Rehashing..." & CRLF
            LoadConf
            LoadMOTD
            For Q = 1 To iUserMax
                If iPeerFree(Q) = False Then
                    X = InStr(1, iModes(Q), "s")
                    If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** Notice -- " & iUser(Index) & " is rehashing Server config file" & CRLF
                End If
            Next Q
            frmMain.lbl_CC = iChanSys.iChan.Count
            SendData Index, ":" & sServer & " 382 " & iUser(Index) & " IRC_Serv.Conf :Rehashed!" & CRLF
        End If
    End If
End Sub

Sub sCWhoIs(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Z As Integer
Dim Q As Integer
Dim TmpText As String
Dim TmpLoad As String
    Q = -2
    X = InStr(1, Text, " ")
    TmpText = Mid(Text, 1, X - 1)
    For X = 1 To iUserMax
        If LCase(TmpText) = LCase(iUser(X)) Then
            Q = X
            Exit For
        End If
    Next X
    
    If Q = -2 Then
        SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & TmpText & " No such nick/channel" & CRLF
    Else
        X = InStr(1, iModes(Q), "W")
        If Not X = 0 Then SendData Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") did a /whois on you." & CRLF
        SendData Index, ":" & sServer & " 311 " & iUser(Index) & " " & iUser(Q) & " " & iName(Q) & " " & iHost(Q) & " * :" & iRealName(Q) & CRLF
        If Not iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 378 " & iUser(Index) & " " & iUser(Q) & " is using modes +" & iModes(Q) & CRLF
        If Not iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 379 " & iUser(Index) & " " & iUser(Q) & " :is connecting from *@" & iRHost(Q) & CRLF
        X = InStr(1, iModes(Q), "r")
        If Not X = 0 Then SendData Index, ":" & sServer & " 307 " & iUser(Index) & " " & iUser(Q) & " :" & sConvertText(RPL_307, Index) & CRLF
        If Not iChan(Q) = "" Then SendData Index, ":" & sServer & " 319 " & iUser(Index) & " " & iUser(Q) & " :" & iChan(Q) & CRLF
        SendData Index, ":" & sServer & " 312 " & iUser(Index) & " " & iUser(Q) & " " & sServer & " :" & iServDSC & CRLF
        If Not iAway(Q) = "" Then SendData Index, ":" & sServer & " 301 " & iUser(Index) & " " & iUser(Q) & " :" & iAway(Q) & CRLF
        If iUserLevel(Q) = 7 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Network Administrator on " & iNetName & CRLF
        If iUserLevel(Q) = 6 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Tech Administrator on " & iNetName & CRLF
        If iUserLevel(Q) = 5 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Co-Administrator on " & iNetName & CRLF
        If iUserLevel(Q) = 4 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Services Administrator on " & iNetName & CRLF
        If iUserLevel(Q) = 3 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Server Administrator on " & iNetName & CRLF
        If iUserLevel(Q) = 2 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Global IRC Operator on " & iNetName & CRLF
        If iUserLevel(Q) = 1 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Local IRC Operator on " & iNetName & CRLF
        
        X = InStr(1, iModes(Q), "h")
        If Not X = 0 Then SendData Index, ":" & sServer & " 310 " & iUser(Index) & " " & iUser(Q) & " :is available for help." & CRLF
        X = InStr(1, iModes(Q), "1")
        If Not X = 0 Then SendData Index, ":" & sServer & " 313 " & iUser(Index) & " " & iUser(Q) & " :is a Coder on " & iNetName & CRLF
        
        X = InStr(1, iModes(Q), "B")
        If Not X = 0 Then SendData Index, ":" & sServer & " 335 " & iUser(Index) & " " & iUser(Q) & " :is a Bot on " & iNetName & CRLF
        
        SendData Index, ":" & sServer & " 317 " & iUser(Index) & " " & iUser(Q) & " " & iIdle(Q) & " " & iSignOn(Q) & " :seconds idle, signon time" & vbCrLf & _
                        ":" & sServer & " 318 " & iUser(Index) & " " & TmpText & " :" & sConvertText(RPL_318, Index) & CRLF
    End If
    ':jolt.horizonws.org 301 RedFile RedFile :booooooo
End Sub

Sub sCAway(Text As String, Index As Integer)
On Error Resume Next
Dim X As Integer
Dim Z As Integer
    If Left$(Text, 1) = ":" Then
        iAway(Index) = Mid$(Text, 2)
        SendData Index, ":" & sServer & " 306 " & iUser(Index) & " :You have been marked as being away" & CRLF
    Else
        iAway(Index) = ""
        SendData Index, ":" & sServer & " 305 " & iUser(Index) & " :You are no longer marked as being away" & CRLF
    End If
End Sub

Sub KillUser(Index As Integer, Op As String, Reason As String, Optional OpIndex As Integer, Optional QuietKill As Boolean, Optional UserNotReg As Boolean)
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim Y As Integer
Dim Q As Integer
    'frmMain.Win(Index).Disconnect
    'iPeerFree(Index) = True
    sCKilled = sCKilled + 1
    
    If Not OpIndex = 0 Then
        SendData Index, ":" & iUser(OpIndex) & "!" & iName(OpIndex) & "@" & iHost(OpIndex) & " KILL " & iUser(Index) & " :" & Reason & CRLF & _
                        "ERROR :Closing Link: " & iUser(Index) & "[" & iRHost(Index) & "] ([" & sServer & "] Local kill by " & Op & " " & Reason & ")" & CRLF
    ElseIf Op = sServer Or Op = "Administrator@Server.Console" Then
        SendData Index, ":" & sServer & " KILL " & iUser(Index) & " :" & Reason & CRLF & _
                        "ERROR :Closing Link: " & iUser(Index) & "[" & iRHost(Index) & "] ([" & sServer & "] Local kill by " & Op & " " & Reason & ")" & CRLF
    End If
    
    If QuietKill = False Then
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "k")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") killed by " & Op & "(" & Reason & ")" & CRLF
            End If
        Next Q
    End If
        'For Y = 1 To iUserMax
        '    If iPeerFree(Y) = False Then
        '        SendData Y, ":" & iUser(Index) & "!" & iName(Index) & "@" & irHost(Index) & " QUIT :Killed by " & Op & "(" & reason & ")" & CRLF
        '    End If
        'Next Y
    
    UserClosed Index, "KILL by " & Op & "(" & Reason & ")", UserNotReg
    ' USE THIS FOR HELP:
    ' :jolt.horizonws.org NOTICE TRON :*** Notice -- Received KILL message for TRON!TRON@t2n.dyndns.org from TRON Path: t2n.dyndns.org!TRON (testing!)
    ' :TRON!TRON@t2n.dyndns.org KILL TRON :t2n.dyndns.org!TRON (testing!)
    ' ERROR :Closing Link: TRON[ci102897-b.nash1.tn.home.com] ([jolt.horizonws.org] Local kill by TRON (testing!))
End Sub

Sub UserClosed(Index As Integer, Optional CloseMsg As String, Optional UserNotReg As Boolean)
On Error Resume Next
Dim Q As Integer
Dim GUI As Boolean
If UserNotReg = True Then GUI = False Else: GUI = True
    If CloseMsg = "" Then CloseMsg = "User has been dropped"
    With frmMain
        If Not CloseMsg = "0" Then
            uQUIT iUser(Index) & "!" & iName(Index) & "@" & iHost(Index), iChan(Index), CloseMsg
        End If
        .Win(Index).Disconnect
        Q = InStr(1, iModes(Index), "i")
        If Not Q = 0 Then frmMain.lbl_IU = frmMain.lbl_IU - 1
        If iUser(Index) = "" Then GUI = False
        If iRHost(Index) = "" Then GUI = False
        If iName(Index) = "" Then GUI = False
        iUser(Index) = ""
        iHost(Index) = ""
        iRHost(Index) = ""
        iRealName(Index) = ""
        iName(Index) = ""
        iIP(Index) = ""
        iModes(Index) = ""
        iChan(Index) = ""
        iAway(Index) = ""
        If Not iUserLevel(Index) = 0 Then .lbl_CO = .lbl_CO - 1
        iUserLevel(Index) = 0
        iSignOn(Index) = 0
        iTP(Index) = ""
        iFP(Index) = 0
        iFC(Index) = 0
        iFM(Index) = 0
        iKT(Index) = 0
        iAAC(Index) = False
        If GUI = True Then .lbl_CU = .lbl_CU - 1
        If GUI = True Then .lbl_CGU = .lbl_CGU - 1
        If GUI = False Then .lbl_UC = .lbl_UC - 1
        Unload .Win(Index)
        iPeerFree(Index) = True
        iPing(Index) = 0
        iIdle(Index) = 0
        iHolted(Index) = False
        iHoldData(Index) = ""
        iKILL(Index) = False
        iHolted(Index) = False
    End With
End Sub

Sub SendLINKS(Index As Integer, Text As String)
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim TmpText As String
    
    SendData Index, ":" & sServer & " 364 " & iUser(Index) & " " & sServer & " " & sServer & " :0 " & iServDSC & CRLF & _
                    ":" & sServer & " 365 " & iUser(Index) & " * :" & RPL_365 & CRLF
End Sub

Sub sCISON(Index As Integer, Nicks As String)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim TmpUser As String
Dim sDone As String

ReDo:
    If Left$(Nicks, 1) = " " Then Nicks = Mid$(Nicks, 2): GoTo ReDo
    If Right$(Nicks, 1) = " " Then Nicks = Mid$(Nicks, 1, Len(Nicks) - 1): GoTo ReDo
    If Nicks = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " :" & sConvertText(ERR_461, Index) & CRLF

ReScan:
    X = InStr(1, Nicks, " ")
    If Nicks = "" Then
        SendData Index, ":" & sServer & " 303 " & iUser(Index) & " :" & sDone & CRLF
        Exit Sub
    End If
    If Not X = 0 Then
        TmpUser = Mid$(Nicks, 1, X - 1)
        Nicks = Mid$(Nicks, X + 1)
    Else
        TmpUser = Nicks
        Nicks = ""
    End If
    
    If TmpUser = "" Then GoTo ReScan
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) Then
            sDone = sDone & iUser(X) & " "
            Exit For
        End If
    Next X
    GoTo ReScan
End Sub

Sub sCUserHost(Index As Integer, Nicks As String)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim TmpUser As String
Dim sDone As String

ReDo:
    If Left$(Nicks, 1) = " " Then Nicks = Mid$(Nicks, 2): GoTo ReDo
    If Right$(Nicks, 1) = " " Then Nicks = Mid$(Nicks, 1, Len(Nicks) - 1): GoTo ReDo
    If Nicks = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " :" & sConvertText(ERR_461, Index) & CRLF

ReScan:
    X = InStr(1, Nicks, " ")
    If Nicks = "" Then
        SendData Index, ":" & sServer & " 302 " & iUser(Index) & " :" & sDone & CRLF
        Exit Sub
    End If
    If Not X = 0 Then
        TmpUser = Mid$(Nicks, 1, X - 1)
        Nicks = Mid$(Nicks, X + 1)
    Else
        TmpUser = Nicks
        Nicks = ""
    End If
    
    If TmpUser = "" Then GoTo ReScan
    For X = 1 To iUserMax
        If LCase(TmpUser) = LCase(iUser(X)) Then
            If iUserLevel(Index) = 0 Then
                If Not iUserLevel(X) = 0 Then
                    If iAway(X) = "" Then
                        sDone = sDone & iUser(X) & "*=+" & iName(X) & "@" & iHost(X) & " "
                    Else
                        sDone = sDone & iUser(X) & "*=-" & iName(X) & "@" & iHost(X) & " "
                    End If
                Else
                    If iAway(X) = "" Then
                        sDone = sDone & iUser(X) & "=+" & iName(X) & "@" & iHost(X) & " "
                    Else
                        sDone = sDone & iUser(X) & "=-" & iName(X) & "@" & iHost(X) & " "
                    End If
                End If
            Else
                If Not iUserLevel(X) = 0 Then
                    If iAway(X) = "" Then
                        sDone = sDone & iUser(X) & "*=+" & iName(X) & "@" & iRHost(X) & " "
                    Else
                        sDone = sDone & iUser(X) & "*=-" & iName(X) & "@" & iRHost(X) & " "
                    End If
                Else
                    If iAway(X) = "" Then
                        sDone = sDone & iUser(X) & "=+" & iName(X) & "@" & iRHost(X) & " "
                    Else
                        sDone = sDone & iUser(X) & "=-" & iName(X) & "@" & iRHost(X) & " "
                    End If
                End If
            Exit For
        End If
        End If
    Next X
    GoTo ReScan
End Sub

Sub sCSetHost(Host As String, Index As Integer)
On Error Resume Next
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
ReScan:
    If Left$(Host, 1) = " " Then
        Host = Mid$(Host, 2)
        GoTo ReScan
    End If
    If Right$(Host, 1) = " " Then Host = Mid$(Host, 1, Len(Host) - 1): GoTo ReScan
    If Host = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " SETHOST :" & ERR_461 & CRLF: Exit Sub
    If sHostValid(Host) = False Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Illegal characters for host, please use A-Z a-z 0-9 . and - Only": Exit Sub
    uMode Index, iUser(Index), "+x"
    iHost(Index) = Host
    SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Your Nick!Ident@Host is now " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " - To disable it type /mode " & iUser(Index) & " -x" & CRLF
    
End Sub

Sub sCSetIdent(Ident As String, Index As Integer)
On Error Resume Next
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
ReScan:
    If Left$(Ident, 1) = " " Then Ident = Mid$(Ident, 2): GoTo ReScan
    If Right$(Ident, 1) = " " Then Ident = Mid$(Ident, 1, Len(Ident) - 1): GoTo ReScan
    If Ident = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " SETIDENT :" & ERR_461 & CRLF: Exit Sub
    iName(Index) = Ident
    SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Your Nick!Ident@Host is now " & iUser(Index) & "!" & iName(Index) & "@" & iHost(Index) & " - To change ident back, do it manually by /SetIdent <OldIdent>" & CRLF
End Sub

Sub sCSetName(Name As String, Index As Integer)
On Error Resume Next
    Name = Mid$(Name, 1, Len(Name) - 1)
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
    If Name = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " SETNAME :" & ERR_461 & CRLF: Exit Sub
    iRealName(Index) = Name
    SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Your 'real name' is now set to: `" & iRealName(Index) & "' - You have to set it manually to undo it" & CRLF
End Sub

Sub sCCHGHost(Nick As String, Host As String, Index As Integer)
On Error Resume Next
Dim Q As Integer
Dim X As Integer
Dim Y As Integer
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
ReScan:
    If Left$(Host, 1) = " " Then
        Host = Mid$(Host, 2)
        GoTo ReScan
    End If
    If Right$(Host, 1) = " " Then Host = Mid$(Host, 1, Len(Host) - 1): GoTo ReScan
    If Host = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGHOST :" & ERR_461 & CRLF: Exit Sub
    If Nick = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGHOST :" & ERR_461 & CRLF: Exit Sub
    If sHostValid(Host) = False Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :Illegal characters for host, please use A-Z a-z 0-9 . and - Only": Exit Sub
    For X = 1 To iUserMax
        If LCase(iUser(X)) = LCase(Nick) Then
            Q = X
            Exit For
        End If
    Next X
    If Q = 0 Then SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & Nick & " No such nick/channel" & CRLF: Exit Sub
    uMode Index, iUser(Q), "+x"
    iHost(Q) = Host
    
    For X = 1 To iUserMax
        Y = InStr(1, iModes(X), "e")
        If Not Y = 0 And Not iUserLevel(X) = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(X) & " :" & iUser(Index) & " changed the dynamic hostname of " & iUser(Q) & " (" & iName(Q) & "@" & iHost(Q) & ") to: '" & Host & "'" & CRLF
    Next X
    
    '[15:50] -irc.ircd-net.org- TRON changed the dynamic hostname of TRON (TRON@ci102897-b.nash1.tn.home.com) to be t2n.dyndns.org
End Sub

Sub sCCHGIdent(Nick As String, Ident As String, Index As Integer)
On Error Resume Next
Dim Q As Integer
Dim X As Integer
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
ReScan:
    If Left$(Ident, 1) = " " Then Ident = Mid$(Ident, 2): GoTo ReScan
    If Right$(Ident, 1) = " " Then Ident = Mid$(Ident, 1, Len(Ident) - 1): GoTo ReScan
    If Ident = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGIDENT :" & ERR_461 & CRLF: Exit Sub
    If Nick = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGIDENT :" & ERR_461 & CRLF: Exit Sub
    For X = 1 To iUserMax
        If LCase(iUser(X)) = LCase(Nick) Then
            Q = X
            Exit For
        End If
    Next X
    If Q = 0 Then SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & Nick & " No such nick/channel" & CRLF: Exit Sub
    iName(Q) = Ident
    
    For X = 1 To iUserMax
        Y = InStr(1, iModes(X), "e")
        If Not Y = 0 And Not iUserLevel(X) = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :" & iUser(Index) & " changed the Ident of " & iUser(Q) & " (" & iName(Q) & "@" & iHost(Q) & ") to: '" & Ident & "'" & CRLF
    Next X
    '[15:50] -irc.ircd-net.org- TRON changed the Ident of TRON (TRON@ci102897-b.nash1.tn.home.com) to be TRON
End Sub

Sub sCCHGName(Nick As String, Name As String, Index As Integer)
On Error Resume Next
Dim Q As Integer
Dim X As Integer
    Name = Mid$(Name, 1, Len(Name) - 1)
    If iUserLevel(Index) = 0 Then SendData Index, ":" & sServer & " 481 " & iUser(Index) & " :" & ERR_481 & CRLF: Exit Sub
    If Name = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGNAME :" & ERR_461 & CRLF: Exit Sub
    If Nick = "" Then SendData Index, ":" & sServer & " 461 " & iUser(Index) & " CHGNAME :" & ERR_461 & CRLF: Exit Sub
    For X = 1 To iUserMax
        If LCase(iUser(X)) = LCase(Nick) Then
            Q = X
            Exit For
        End If
    Next X
    If Q = 0 Then SendData Index, ":" & sServer & " 401 " & iUser(Index) & " :" & Nick & " No such nick/channel" & CRLF: Exit Sub
    iRealName(Q) = Name
    
    For X = 1 To iUserMax
        Y = InStr(1, iModes(X), "e")
        If Not Y = 0 And Not iUserLevel(X) = 0 Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :" & iUser(Index) & " changed the GECOS of " & iUser(Q) & " (" & iName(Q) & "@" & iHost(Q) & ") to: '" & Name & "'" & CRLF
    Next X
End Sub
