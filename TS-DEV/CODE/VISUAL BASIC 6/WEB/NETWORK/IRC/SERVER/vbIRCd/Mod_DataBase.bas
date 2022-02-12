Attribute VB_Name = "Mod_DataBase"
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



' OK, I forgot to remove these lines from the release since they
' were apart of the built-in services of IRC Serv, but ohh well
' I'll keep them in here since they were on here on the release. :)
'
' So, the Services Info Buffers are not in use by vbIRCd at this time
' but could be used in the near future if some one decides to continue
' the built-in services for vbIRCd then they are welcome to.
' I just thought you may want to know. :)  -TRON

' Start of Services Information Buffers
Private Type iNSDB
    Nick As New Collection
    Pass As New Collection
    Email As New Collection
    Modes As New Collection
    LastHost As New Collection
    RegTS As New Collection
    LastOnTS As New Collection
    Hosts As New Collection
    RealName As New Collection
End Type
Private Type iMSDB
    Nick As New Collection
    Msg As New Collection
    From As New Collection
    MsgTS As New Collection
End Type

Global rUsers As iNSDB
Global rUserM As iMSDB
' End of Services Information Buffers


Public Sub ACTB(Text As String)
    frmMain.txt_Buffer.SelStart = Len(frmMain.txt_Buffer)
    frmMain.txt_Buffer.SelText = Text + CRLF
    frmMain.txt_Buffer.SelStart = Len(frmMain.txt_Buffer)
End Sub

Public Sub LoadConf()
On Error Resume Next
Dim X As Integer
Dim Z As Integer
Dim Q As Integer
Dim TmpLoad As String
Dim TmpText As String
Dim TmpData As String
Dim TmpString As String
Dim TmpGet As String
Dim Text As String
Dim DCF As Boolean
With frmMain

    If Not Dir(SysFile, vbNormal) = "ircd.conf" Then
        ' Do Nothing but report file not there and Exit process                -TRON
        DisplayLog "ERROR: " & SysFile & " - Couldn't Find Configuration File", True
        Exit Sub
    End If

.List_Ports.Clear
.List_SO.Clear
.List_SOP.Clear
.List_SOM.Clear

Do Until sEline.Count = 0
    sEline.Remove (sEline.Count) 'Clear out Eline Hosts
Loop
Do Until sElineR.Count = 0
    sElineR.Remove (sElineR.Count) 'Clear out Eline Reasons
Loop
'Do Until sAKill.Count = 0
'    sAKill.Remove (sAKill.Count) 'Clear out AKILL Hosts
'Loop
'Do Until sAKillR.Count = 0
'    sAKillR.Remove (sAKillR.Count) 'Clear out AKILL Reasons
'Loop
Do Until sKline.Count = 0
    sKline.Remove (sKline.Count) 'Clear out Kline Hosts
Loop
Do Until sKlineR.Count = 0
    sKlineR.Remove (sKlineR.Count) 'Clear out Kline Reasons
Loop

    .txt_Buffer.Text = ""
    Open SysFile For Input As #1
        .txt_Buffer.Text = StrConv(InputB(LOF(1), 1), vbUnicode) & CRLF & CRLF
    Close #1
ReLoad:

    If .txt_Buffer = "" Then 'ircd.conf is done loading, now let's do a few checks and report results to display log =)
            If iServer = "" Then ' Check if Server name is set
                DisplayLog "ERROR: Server Name not set in M:line of ircd.conf", True
                Exit Sub
            Else
                DisplayLog "Server Name set to '" & iServer & "'"
            End If
            
            DisplayLog "Server Description set to '" & iServDSC & "' [it's OK if nothing]" ' Report server description
            DisplayLog "Die / Restart Password: '" & iDPass & "' / '" & iRPass & "'"
            For X = 0 To .List_Ports.ListCount - 1
                DisplayLog "Bind to Port '" & .List_Ports.List(X) & "' listed"
            Next X
            DisplayLog sQline.Count & " Q-lines Found and Set"
            DisplayLog sKline.Count & " K-lines Found and Set"
            DisplayLog sEline.Count & " Ban Exceptions Found and Set"
            For X = 0 To .List_SO.ListCount - 1
                DisplayLog "IRC Oper '" & .List_SO.List(X) & "' Added with Flags '" & .List_SOM.List(X) & "'"
            Next X
            DisplayLog .List_SO.ListCount & " IRC Operators Found and Set"
            
        If .C_SSC.Caption = "Start" Then ' Check to make sure Server is off
            frmMain.C_SSC_Click ' Since Server is not up, let's auto-start it
        Else
            ScanFBU 'Since Server is up, let's scan for banned users >:D
        End If
        Exit Sub
    End If
    
    Z = InStr(1, .txt_Buffer, CRLF)
    If Z = 0 Then
        TmpLoad = .txt_Buffer
        .txt_Buffer = ""
    Else
        TmpLoad = Mid$(.txt_Buffer, 1, Z - 1)
    End If
    
        If Left$(TmpLoad, 1) = "#" Then
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If TmpLoad = "" Then
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "M:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            iServer = Mid$(TmpText, 1, X - 1) 'Set Server Name
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            '<Bind-IP-String> = Mid$(TmpText, 1, X - 1) 'Set Bind IP [Not Supported]
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            iServDSC = Mid$(TmpText, 1, X - 1) ' Set Server Description :)
            TmpText = Mid$(TmpText, X + 1)
            
            .List_Ports.AddItem TmpText 'Add default port to ports to bind to when starting up
            
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If

        If Left$(TmpLoad, 2) = "A:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            iServAdminLine1 = Mid$(TmpText, 1, X - 1) 'Set the First Admin Line
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            iServAdminLine2 = Mid$(TmpText, 1, X - 1) 'Set the Second Admin Line
            TmpText = Mid$(TmpText, X + 1)
            
            iServAdminLine3 = TmpText 'Set the Last Admin Line
            
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_CINFO " Then
            'iSupportEmail = Mid$(TmpLoad, 12, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_PING " Then
            'iPing = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_CCM " Then
            'iSMOCC = Mid$(TmpLoad, 10, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_MMFLOODC " Then
            'iFloodMSGs = Mid$(TmpLoad, 15, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_MCFLLODC " Then
            'iFloodCMDs = Mid$(TmpLoad, 15, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_LCCT " Then
            'iWhoCCC = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_SMUC " Then
            'iChanMax = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_PASS " Then
            'iConnPass = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_ACC " Then
            'sACC = Mid$(TmpLoad, 10, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_DNS " Then
            'iGODNS = Mid$(TmpLoad, 10, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "X:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            iDPass = Mid$(TmpText, 1, X - 1) 'Set the DIE password
            TmpText = Mid$(TmpText, X + 1)
            
            iRPass = TmpText 'Set the Restart password
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_FONN " Then
            'sFONN = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_FOCN " Then
            'sFOCN = Mid$(TmpLoad, 11, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_FOS " Then
            'sFOS = Mid$(TmpLoad, 10, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "K:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            TmpData = Mid$(TmpText, 1, X - 1) 'Get Banned Host Address
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            TmpGet = Mid$(TmpText, 1, X - 1) 'Get Ban Reason
            TmpText = Mid$(TmpText, X + 1) 'Get Banned Ident
            
            X = InStr(1, TmpText, " ")
            If Not X = 0 Then TmpText = Mid$(TmpText, 1, X - 1) 'Check for any spaces and stop at the first one if any
            
            sKline.Add TmpText & "@" & TmpData 'Add banned hostmask to list
            sKlineR.Add TmpGet 'Add Reason for the ban to list
            
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "E:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            TmpData = Mid$(TmpText, 1, X - 1) 'Get Excepted Host Address
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            TmpGet = Mid$(TmpText, 1, X - 1) 'Get Except Reason
            TmpText = Mid$(TmpText, X + 1) 'Get Excepted Ident
            
            X = InStr(1, TmpText, " ")
            If Not X = 0 Then TmpText = Mid$(TmpText, 1, X - 1) 'Check for any spaces and stop at the first one if any
            
            sEline.Add TmpText & "@" & TmpData 'Add Excepted hostmask to list
            sElineR.Add TmpGet 'Add Reason for the Except to list
            
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "SERV_LIMIT " Then
            '.iConnMax = Mid$(TmpLoad, 12, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "O:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            .List_SOA.AddItem Mid$(TmpText, 1, X - 1) 'Set Oper Hostmask
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            .List_SOP.AddItem Mid$(TmpText, 1, X - 1) 'Set Oper pass
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            .List_SO.AddItem Mid$(TmpText, 1, X - 1) ' Set Oper ID
            TmpText = Mid$(TmpText, X + 1)
            
            X = InStr(1, TmpText, ":")
            .List_SOM.AddItem Mid$(TmpText, 1, X - 1) ' Set Oper Flags
            
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
    '-------------->New Conf Settings<-------------
        If Left$(TmpLoad, 13) = "SERV_NETNAME " Then
            'iNetName = Mid$(TmpLoad, 14, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 14) = "SERV_MAINCHAN " Then
            'iMainChan = Mid$(TmpLoad, 15, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 14) = "SERV_HELPCHAN " Then
            'iHelpChan = Mid$(TmpLoad, 15, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 14) = "SERV_HIDHPREX " Then
            'iHiddenPrefix = Mid$(TmpLoad, 15, Z - 1)
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
        
        If Left$(TmpLoad, 2) = "P:" Then
            X = InStrRev(TmpLoad, ":")
            TmpText = Mid$(TmpLoad, X + 1)
            DCF = False
            For X = 0 To .List_Ports.ListCount - 1
                If .List_Ports.List(X) = TmpText Then
                    DCF = True
                    Exit For
                End If
            Next X
            If Not DCF = True Then .List_Ports.AddItem TmpText
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
    '----------------------------------------------
    
        If Left$(TmpLoad, 2) = "Q:" Then
            TmpText = Mid$(TmpLoad, 3)
            X = InStr(1, TmpText, ":")
            TmpText = Mid$(TmpText, X + 1)
            X = InStr(1, TmpText, ":")
            sQlineR.Add Mid$(TmpText, 1, X - 1)  'Set the Q:line'd Reason
            TmpText = Mid$(TmpText, X + 1)
            
            sQline.Add TmpText  'Set the Q:line'd Nickname
            .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
        GoTo ReLoad
        End If
    
    .txt_Buffer = Mid$(.txt_Buffer, Z + 2)
    GoTo ReLoad
    
    End With
End Sub

Public Sub LoadMOTD()
On Error Resume Next
    With frmMain
    
    If GetAttr(MOTDFile) < vbVolume > 0 Then
        'Open MOTDFile For Output As #1
        'Print #1, "MOTD Needs to be edited"
        'DoEvents
        'Close
        '
        'lets just leave it empty and exit out OK? ;)  -TRON
        .txt_CMOTD = "": Exit Sub
    End If
    
    .txt_CMOTD.Text = ""
    Open MOTDFile For Input As #1
        .txt_CMOTD.Text = StrConv(InputB(LOF(1), 1), vbUnicode)
    Close #1
    .txt_CMOTD = Mid$(.txt_CMOTD, 1, Len(.txt_CMOTD) - 2)
    End With
End Sub

Public Sub SaveMOTD()
On Error Resume Next
Dim X
    ' This is not in use since IRCd's aren't really suppose to be
    ' saving MOTDs, but just reading them on /rehash -MOTD or startup.
    '
    ' This is just more code that was left laying around from
    ' IRC Serv code ;)  -TRON
    Open MOTDFile For Output As #1
    Print #1, frmMain.txt_CMOTD.Text
    DoEvents
    Close
End Sub
