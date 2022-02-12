Attribute VB_Name = "Mod_ServCMDs2"
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


Sub SendMISC(Index As Integer)
On Error Resume Next
SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :Misc Information list on hidden commands:" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /STATUS" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Shows you info on data processing" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /USERS" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Lists all locally connected users" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /USER [Nick]" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Will give you allot of information on your self and" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :      if '[Nick]' field is used it will give you allot of" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :      info on that user, but that part is limited to IRC Ops" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /SETHOST <NewHost>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Sets your host to the '<NewHost>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /SETIDENT <NewIdent>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Sets your ident to the '<NewIdent>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /SETNAME <NewName>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Sets your real name to the '<NewName>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /CHGHOST <Nick> <NewHost>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Changes '<Nick>'s host to '<NewHost>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /CHGIDENT <Nick> <NewIdent>" & CRLF
SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :    - Changes '<Nick>'s ident to '<NewIdent>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /CHGNAME <Nick> <NewName>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - Changes '<Nick>'s real name to '<NewName>' value, IRC Ops ONLY" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /SAJOIN <Nick> <Channel>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - This command has been disabled to provent abuse" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /SAPART <Nick> <Channel>" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - This command has been disabled to provent abuse" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    /4 Char Command <-- Figure it out, Hint: A Product Icon." & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :    - It's an easter egg that displays some ASCII Art about IRC Serv" & CRLF & _
                ":" & sServer & " 374 " & iUser(Index) & " :End of /MISC info" & CRLF
End Sub

Sub SendLOGO(Index As Integer)
On Error Resume Next
SendData Index, ":" & sServer & " 371 " & iUser(Index) & " :7#####  ####     #####     ####   ######  ####    #    #" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7  #    #   #   #         #    #  #       #   #   #    #" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7  #    #   #   #         #       #       #   #   #    #       12!!    !!!" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7  #    ####    #          ####   ####    ####    #    #        12!      !" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7  #    #  #    #              #  #       #  #    #    #        12!    !!!" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7  #    #   #   #         #    #  #       #   #    #  #    9# #  12!    !  " & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " :7#####  #    #   #####     ####   ######  #    #    ##      9#  12!!! * !!!" & CRLF & _
                ":" & sServer & " 371 " & iUser(Index) & " : " & CRLF & _
                ":" & sServer & " 374 " & iUser(Index) & " :End of /LOGO ASCII Art of IRC Serv [Easter Egg]" & CRLF
End Sub

Function uCanNICK(User As String, Index As Integer) As Boolean
On Error Resume Next
Dim sChan As String
Dim sPartMsg As String
Dim sUsers As String
Dim sDate As String
Dim sUsersM As String
Dim sNick As String
Dim NotToUsers As String
Dim TmpNumber, TmpNumber2 As Integer
Dim TmpText, TmpText2 As String
Dim TmpLoad, TmpLoad2 As String
Dim DFY As Boolean
Dim sUCS As String
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim Q As Integer
DFY = False
uCanNICK = True
    sChans = iChan(Index)
    sNick = iUser(Index)
    
ReScan:
    If sChans = "" Then Exit Function
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
            Exit For
        End If
    Next X
    
    If Not Q = 0 Then
        TmpText = Replace(.iCBans(Q), " ", "", 1, Len(.iCBans(Q)), vbTextCompare)
        TmpNumber = Len(.iCBans(Q)) - Len(TmpText)
        TmpLoad = .iCBans(Q)
        TmpText2 = Replace(.iCExcep(Q), " ", "", 1, Len(.iCExcep(Q)), vbTextCompare)
        TmpNumber2 = Len(.iCExcep(Q)) - Len(TmpText)
        TmpLoad2 = .iCExcep(Q)
        DFY = False
        For Y = 1 To TmpNumber
            X = InStr(1, TmpLoad, " ")
            TmpText = Mid$(TmpLoad, 1, X - 1)
            TmpLoad = Mid$(TmpLoad, X + 1)
            If LCase(iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index)) Like LCase(TmpText) Then DFY = True
            If LCase(iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)) Like LCase(TmpText) Then DFY = True
            If DFY = True Then
                For Q = 1 To TmpNumber2
                    X = InStr(1, TmpLoad2, " ")
                    TmpText2 = Mid$(TmpLoad2, 1, X - 1)
                    TmpLoad2 = Mid$(TmpLoad2, X + 1)
                    If LCase(iUser(Index) & "!" & iName(Index) & "@" & iRHost(Index)) Like LCase(TmpText2) Then DFY = False: Exit For
                    If LCase(iUser(Index) & "!" & iName(Index) & "@" & iHost(Index)) Like LCase(TmpText2) Then DFY = False: Exit For
                Next Q
                
                If DFY = True Then
                    SendData Index, ":" & sServer & " 437 " & iUser(Index) & " " & .iChan(Q) & " :" & sConvertText(ERR_437, Index) & CRLF
                    uCanNICK = False
                    Exit Function
                End If
            End If
        Next Y
        X = InStr(1, .iCModes(Q), "N")
        If Not X = 0 Then
            uCanNICK = False
            SendData Index, ":" & sServer & " 447 " & iUser(Index) & " :" & sConvertText(ERR_447, Index) & " " & .iChan(Q) & " (+N)" & CRLF
            Exit Function
        End If
    Else
        LogIt "Faild 2 Find " & sChan & " 4 :" & UserHostMask & " NICK " & NewNick
        If uHaveMode(Index, "D") = True Then SendData Index, ":" & sServer & " NOTICE " & iUser(Index) & " :ERROR: Channel not found in Buffer!  Error has been logged." & CRLF
    End If
GoTo ReScan
    End With
End Function

Sub sCWho(Text As String, Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim Y As Integer
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
With iChanSys

If Text = "" Then
    For Q = 1 To iUserMax
        If iPeerFree(Q) = False Then
            If Not iUserLevel(Index) = 0 Or InStr(1, iModes(Q), "i") = 0 Then
                X = InStr(1, iChan(Q), " ")
                If Not X = 0 Then
                    If iAway(Q) = "" Then TmpData = "H" Else: TmpData = "G"
                    If Not iUserLevel(Q) = 0 Then TmpData = TmpData & "*"
                    If Not InStr(1, iModes(Q), "r") = 0 Then TmpData = TmpData & "r"
                    TmpText = Mid$(iChan(Q), 1, X - 1)
                    If Left$(TmpText, 1) = "~" Then TmpText = Mid$(TmpText, 2)
                    If Left$(TmpText, 1) = "@" Then TmpText = Mid$(TmpText, 2): TmpData = TmpData & "@"
                    If Left$(TmpText, 1) = "%" Then TmpText = Mid$(TmpText, 2): TmpData = TmpData & "%"
                    If Left$(TmpText, 1) = "+" Then TmpText = Mid$(TmpText, 2): TmpData = TmpData & "+"
                    SendData Index, ":" & sServer & " 352 " & iUser(Index) & " " & TmpText & " " & iName(Q) & " " & iHost(Q) & " " & sServer & " " & iUser(Q) & " " & TmpData & " :0 " & iRealName(Q) & CRLF
                Else
                    If iAway(Q) = "" Then TmpData = "H" Else: TmpData = "G"
                    If Not iUserLevel(Q) = 0 Then TmpData = TmpData & "*"
                    If Not InStr(1, iModes(Q), "r") = 0 Then TmpData = TmpData & "r"
                    SendData Index, ":" & sServer & " 352 " & iUser(Index) & " * " & iName(Q) & " " & iHost(Q) & " " & sServer & " " & iUser(Q) & " " & TmpData & " :0 " & iRealName(Q) & CRLF
                End If
            End If
        End If
    Next Q
Else
    
    
End If
    SendData Index, ":" & sServer & " 315 " & iUser(Index) & " " & Text & " :End of /WHO list" & CRLF
'352  RPL_WHOREPLY - "<channel> <user> <host> <server> <nick> <H|G|r>[*][@|%|+] :<hopcount> <real name>"
'315  RPL_ENDOFWHO - "<name> :End of /WHO list"
End With
End Sub

Sub sCWhoWas(Text As String, Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim TmpData As String
Dim Y As Integer
Dim X As Integer
Dim Q As Integer
Dim Z As Integer


    
'314  RPL_WHOWASUSER - "<nick> <user> <host> * :<real name>"
'369  RPL_ENDOFWHOWAS - "<nick> :End of WHOWAS"

End Sub

Public Sub SendLICENSE(Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim IDL As String
Dim iNews As String
Dim X
iNews = frmMain.txt_License.Text

    SendData Index, ":" & sServer & " 375 " & iUser(Index) & " :" & sServer & " vbIRCd LICENSE Information" & CRLF
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
        SendData Index, ":" & sServer & " 376 " & iUser(Index) & " :End of /LICENSE Command" & CRLF
    Else
        GoTo ReSend
    End If
    
    '":" & sServer & " 375 " & iUser(Index) & " :" & sServer & " Message Of The Day" & CRLF
    '":" & sServer & " 372 " & iUser(Index) & " :- " & iDL & CRLF
    '":" & sServer & " 376 " & iUser(Index) & " :End of /MOTD Command" & CRLF
End Sub
