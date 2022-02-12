Attribute VB_Name = "Modoule"
Option Explicit

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


'---SysTray Code---!
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Global TrayI As NOTIFYICONDATA
'---SysTray Code---.

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const BM_SETSTYLE = &HF4
Private Const BS_SOLID = 0

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800
Public Const CBSTYLE_FLAT = &H800
Public Const LB_SETHORIZONTALEXTENT = WM_USER + 21

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Global CRLF As String ' Cairrage return/Line feed
Global CR As String
Global LF As String
' Set Max of Users can possibliy use ircd, and this HAS to be left Punblic Const at all times!  -TRON
Public Const iUserMax = "15900" 'WARNING: 19000 Recommanded MAX

'----------Users Buffers----------
Global iPeerFree(1 To iUserMax) As Boolean 'Place Keeper
Global iUser(1 To iUserMax) As String 'User's Nick
Global iName(1 To iUserMax) As String 'User's ID
Global iHost(1 To iUserMax) As String 'User's Dynamic HostMask
Global iRHost(1 To iUserMax) As String 'User's Real HostMask
Global iRealName(1 To iUserMax) As String 'User's Real Name
Global iModes(1 To iUserMax) As String 'User's Flag Modes
Global iIP(1 To iUserMax) As String 'User's IP Address
Global iPing(1 To iUserMax) As Integer 'User's Ping Idle
Global iIdle(1 To iUserMax) As Long 'User's Idle Time
Global iSignOn(1 To iUserMax) As Long 'User's Sign On Date/Time in Unix Time
Global iUserLevel(1 To iUserMax) As Integer  'User Level(0 - Normal, 1 - Local IRC Op & etc...)
Global iAAC(1 To iUserMax) As Boolean  'Accepted Access Connection?
Global iTP(1 To iUserMax) As String 'Temp. Password
Global iChan(1 To iUserMax) As String 'Channels
Global iAway(1 To iUserMax) As String 'Away
Global iFP(1 To iUserMax) As Integer 'Faild Passwords
Global iKT(1 To iUserMax) As Integer 'Time until Killed for not identify to nickserv
Global iFC(1 To iUserMax) As Integer 'Flood Commands
Global iFM(1 To iUserMax) As Integer 'Flood Messages
Global iHoldData(1 To iUserMax) As String 'Hold Data until IdentScan is DONE
Global iHolted(1 To iUserMax) As Boolean 'If User is Holted for IdentScan
'--------End Users Buffers--------
'--------Server Info Buffers------
Global sServer As String 'Server's Name <- Used in code for server name ONLY, don't use iServer.  Cause sServer is dynamicly changed during server restarts/start
Global iServer As String 'Server's Name from ircd.conf
Global iNetName As String  ' Network Name
Global iServDSC As String  ' Server's Description
Global iServAdminLine1 As String ' \
Global iServAdminLine2 As String '  `,-Server Admin Lines
Global iServAdminLine3 As String '  /
Global iRPass As String  ' Restart Pass
Global iDPass As String  ' Die Pass
Global iSupportEmail As String  ' E-mail of Support Team
Global iHiddenPrefix As String  ' Host Hidden Prefix
Global iFloodCMDs As Integer  ' Max Commands Per Minute
Global iFloodMSGs As Integer  ' Max Messages/Notices Per Minute
Global iConnPass As String  ' Server password
Global iServPing As Integer  ' How long until next ping in secs
Global iMainChan As String  ' Main Channel of the Server
Global iHelpChan As String  ' The Help Support Channel
Global iChanMax As Integer  ' Max Channels user can be on
Global iConnMax As Integer ' Max of Connections allowed
Global iSMOCC As String   ' Set Modes on Channel Creation
Global iForceCloak As Integer '[1|0] Force User umode +x (xxx.xxx.xxx.mpx-#### or mpx-####.host.isp.net) kinda hostmask
Global iGODNS As Integer  '[1|0] Use Hostmask from what client sends(Not recommanded to enable)
Global iWhoCCC As Integer  ' Who Can Create Channels
' 0 = Anyone
' 1 = Local Ops & Up Only
' 2 = Global Ops & Up Only
' 3 = Services Ops & Up Only
' 4 = Server Admins Only
Global sVersion As String 'Server's Version
Global sRelease As String 'Server's Release Info
Global sUDT As String 'Server Up Date Time
Global sCCount As Long 'Connect Estbalish Count
Global sCRefused As Long 'Connections Refused
Global sCAccepted As Long 'Connections Accepted
Global sCKilled As Long 'Connections Killed
Global sBSent As Long 'Bytes Sent
Global sBReceived As Long  'Bytes Received
Global sTSent As Long 'Times Sent Data
Global sTReceived As Long 'Times Received Data
'--------End Server Info Buffers--
'--------Security Buffers---------
Global sFONN As Integer  'Disallow listed words in NickNames in filter
Global sFOCN As Integer  'Disallow listed words in Channel Names in filter
Global sFOS As Integer 'Filter Options Setting...
'[>-----------------------------<]
Global sAKill As New Collection  'Auto-Kill Bans
Global sAKillR As New Collection 'Auto-Kill Ban Reasons
Global sKline As New Collection  'K-line Bans
Global sKlineR As New Collection 'K-line Ban Reasons

Global sEline As New Collection  'E-line Hostmask
Global sElineR As New Collection 'E-line Exception Ban Reasons
Global sZline As New Collection  'Z-line IPs
Global sZlineR As New Collection 'Z-line Ban Reasons

Global sQline As New Collection  'Q-line Nicks
Global sQlineR As New Collection 'K-line Nick Reasons

Global sIlineP As New Collection 'I-line Password
Global sIlineH As New Collection 'I-line Hostmask
Global sIlineI As New Collection 'I-line IP Address
'---------End of Buffers----------

Global xMODE As Integer
Global xMODE2 As Integer
Global xUSER As Integer
Global xUSER2 As Integer
Global xOPER As Integer
Global xOPER2 As Integer
Global xMOTD As Integer
Global xMOTD2 As Integer
Global xNICK As Integer
Global xNICK2 As Integer
Global xUSERS As Integer
Global xUSERS2 As Integer
Global xSTATS As Integer
Global xSTATS2 As Integer
Global xADMIN As Integer
Global xADMIN2 As Integer
Global xINFO As Integer
Global xINFO2 As Integer
Global xVERSION As Integer
Global xVERSION2 As Integer
Global xREHASH As Integer
Global xREHASH2 As Integer
Global xKILL As Integer
Global xKILL2 As Integer
Global xJOIN As Integer
Global xJOIN2 As Integer
Global xPART As Integer
Global xPART2 As Integer
Global xTIME As Integer
Global xTIME2 As Integer
Global xQUIT As Integer
Global xQUIT2 As Integer
Global xNOTICE As Integer
Global xNOTICE2 As Integer
Global xPRIVMSG As Integer
Global xPRIVMSG2 As Integer
Global xPONG As Integer
Global xPONG2 As Integer
Global xLUSERS As Integer
Global xLUSERS2 As Integer
Global xPASS As Integer
Global xPASS2 As Integer
Global xDIE As Integer
Global xDIE2 As Integer
Global xRESTART As Integer
Global xRESTART2 As Integer
Global xAKILL As Integer
Global xAKILL2 As Integer
Global xRAKILL As Integer
Global xRAKILL2 As Integer
Global xKLINE As Integer
Global xKLINE2 As Integer
Global xUNKLINE As Integer
Global xUNKLINE2 As Integer
Global xSUMMON As Integer
Global xSUMMON2 As Integer
Global xWHOIS As Integer
Global xWHOIS2 As Integer
Global xAWAY As Integer
Global xAWAY2 As Integer
Global xNICKSERV As Integer
Global xNICKSERV2 As Integer
Global xCHANSERV As Integer
Global xCHANSERV2 As Integer
Global xMEMOSERV As Integer
Global xMEMOSERV2 As Integer
Global xOPERSERV As Integer
Global xOPERSERV2 As Integer
Global xSTATSERV As Integer
Global xSTATSERV2 As Integer
Global xINFOSERV As Integer
Global xINFOSERV2 As Integer
Global xNETSERV As Integer
Global xNETSERV2 As Integer
Global xLINKS As Integer
Global xLINKS2 As Integer
Global xSUPPORT As Integer
Global xSUPPORT2 As Integer
Global xMAINCHAN As Integer
Global xMAINCHAN2 As Integer
Global xNETWORK As Integer
Global xNETWORK2 As Integer
Global xRULES As Integer
Global xRULES2 As Integer
Global xNAMES As Integer
Global xNAMES2 As Integer
Global xISON As Integer
Global xISON2 As Integer
Global xUSERHOST As Integer
Global xUSERHOST2 As Integer
Global xPING As Integer
Global xPING2 As Integer
Global xTOPIC As Integer
Global xTOPIC2 As Integer
Global xSETHOST As Integer
Global xSETHOST2 As Integer
Global xSETIDENT As Integer
Global xSETIDENT2 As Integer
Global xSETNAME As Integer
Global xSETNAME2 As Integer
Global xCHGHOST As Integer
Global xCHGHOST2 As Integer
Global xCHGIDENT As Integer
Global xCHGIDENT2 As Integer
Global xCHGNAME As Integer
Global xCHGNAME2 As Integer

'New Command Stats m/M Info Buffers
Global xSTATUS As Integer
Global xSTATUS2 As Integer
Global xSAJOIN As Integer
Global xSAJOIN2 As Integer
Global xSAPART As Integer
Global xSAPART2 As Integer
Global xSERVER As Integer
Global xSERVER2 As Integer
'End of New Command Info Buffers

'Customizable Reply and Error Strings
Global Const RPL_219 = "End of /STATS report"
Global Const RPL_232 = ":- $Data" 'Rules
Global Const RPL_250 = "Highest connection count: %HighConC% (%MLUsers% clients)"
Global Const RPL_251 = "There are %d users and %d invisible on %d servers"
Global Const RPL_252 = "operator(s) online"
Global Const RPL_253 = "unknown connection(s)"
Global Const RPL_254 = "channels formed"
Global Const RPL_255 = "I have %d clients and %d servers"
Global Const RPL_256 = "Administrative info about %Server%"
Global Const RPL_265 = "Current Local Users: %d  Max: %d"
Global Const RPL_266 = "Current Global Users: %d  Max: %d"
Global Const RPL_294 = "Your help-request has been forwarded to Help Operators"
Global Const RPL_295 = "Your address has been ignored from forwarding"
Global Const RPL_305 = "You are no longer marked as being away"
Global Const RPL_306 = "You have been marked as being away"
Global Const RPL_307 = "is a registered nick"
Global Const RPL_308 = "- %server% Server Rules -"
Global Const RPL_309 = "End of /RULES command."
Global Const RPL_310 = "is available for help."
Global Const RPL_313 = "is a %IRCopStats% on %Network%" ' Shall not be editable :P
Global Const RPL_315 = "End of /WHO list."
Global Const RPL_317 = "seconds idle, signon time"
Global Const RPL_318 = "End of /WHOIS list."
Global Const RPL_323 = "End of /LIST"
Global Const RPL_331 = "No topic is set."
Global Const RPL_335 = "is a Bot on %Network%"
Global Const RPL_343 = ":Just NewFlahs - Used for ADS!!! :D" ' <-- May or may not be supported, I haven't decided...
Global Const RPL_347 = "End of Channel Invite List"
Global Const RPL_349 = "End of Channel Exception List"
Global Const RPL_365 = "End of /LINKS list."
Global Const RPL_366 = "End of /NAMES list."
Global Const RPL_368 = "End of Channel Ban List"
Global Const RPL_369 = "End of WHOWAS"
Global Const RPL_373 = "Server INFO"
Global Const RPL_374 = "End of /INFO list."
Global Const RPL_375 = "- %s Message of the Day - "
Global Const RPL_376 = "End of /MOTD command."
Global Const RPL_378 = "is connecting from *@%s"
Global Const RPL_379 = "is using modes %s"
Global Const RPL_381 = "You are now an IRC Operator"
Global Const RPL_382 = "$ConfFile :Rehashing"
Global Const RPL_387 = ":End of Channel Owner List"
Global Const RPL_389 = ":End of Protected User List"
Global Const RPL_392 = ":UserID   Terminal  Host"
Global Const RPL_394 = ":End of Users"

Global Const ERR_401 = "No such nick/channel"
Global Const ERR_402 = "No such server"
Global Const ERR_403 = "No such channel"
Global Const ERR_404 = "No external channel messages"
Global Const ERR_404B = "You need voice (+v)"
Global Const ERR_405 = "You have joined too many channels"
Global Const ERR_406 = "There was no such nickname"
Global Const ERR_412 = "No text to send"
Global Const ERR_421 = "Unknown command"
Global Const ERR_422 = "MOTD File is missing"
Global Const ERR_423 = "No administrative info available"
Global Const ERR_425 = "OPERMOTD File is missing"
Global Const ERR_431 = "No nickname given"
Global Const ERR_432 = "Erroneus Nickname"
Global Const ERR_433 = "Nickname is already in use."
Global Const ERR_434 = "RULES File is missing"
Global Const ERR_436 = "Nickname collision KILL"
Global Const ERR_437 = "Cannot change nickname while banned on channel"
Global Const ERR_438 = "Nick change too fast. Please wait %d seconds"
Global Const ERR_439 = "Message target change too fast. Please wait %d seconds"
Global Const ERR_440 = "Services are currently down. Please try again later."
Global Const ERR_441 = "They aren't on that channel"
Global Const ERR_442 = "You're not on that channel"
Global Const ERR_443 = "is already on channel"
Global Const ERR_445 = "SUMMON has been disabled"
Global Const ERR_446 = "USERS has been disabled"
Global Const ERR_447 = "Can not change nickname while on"
Global Const ERR_451 = "You have not registered"
Global Const ERR_455 = "Your username contained the invalid character(s). Please use only the characters 0-9 a-z A-Z _ - or . in your username. Your username is the part before the @ in your email address. You will now be disconnected from server."
Global Const ERR_459 = "Cannot join channel (+H)"
Global Const ERR_460 = "Halfops cannot set mode"
Global Const ERR_461 = "Not enough parameters"
Global Const ERR_463 = "Your host isn't among the privileged"
Global Const ERR_464 = "Password Incorrect"
Global Const ERR_465 = "You are banned from this server.  Mail %s for more information"
Global Const ERR_467 = "Channel key already set"
Global Const ERR_468 = "Only servers can change that mode"
Global Const ERR_469 = "Channel link already set"
Global Const ERR_470 = "[Link] %s has become full, so you are automatically being transferred to the linked channel %s"
Global Const ERR_471 = "Cannot join channel (+l)"
Global Const ERR_472 = "is unknown mode char to me"
Global Const ERR_473 = "Cannot join channel (+i)"
Global Const ERR_474 = "Cannot join channel (+b)"
Global Const ERR_475 = "Cannot join channel (+k)"
Global Const ERR_476 = "Bad Channel Mask"
Global Const ERR_477 = "You need a registered nick to join that channel."
Global Const ERR_478 = "Channel ban/ignore list is full"
Global Const ERR_479 = "%s :Sorry, the channel has an invalid channel link set."
Global Const ERR_480 = "Cannot knock on %s (%s)"
Global Const ERR_481 = "Permission Denied- You do not have the correct IRC operator privileges"
Global Const ERR_482 = "You're not channel operator"
Global Const ERR_483 = "You cant kill a server!"
Global Const ERR_484 = "Cannot kick protected user %s."
Global Const ERR_485 = "Cannot kill protected user %s."
Global Const ERR_486 = "%s is currently disabled, please try again later."
Global Const ERR_492 = "No O-lines for your host"
Global Const ERR_501 = "Unknown MODE flag"
Global Const ERR_502 = "Cant change mode for other users"
Global Const ERR_518 = "Cannot invite (+I) at channel"
Global Const ERR_519 = "Cannot join channel (Admin only)"
Global Const ERR_520 = "Cannot join channel (IRCops only)"

Global sNickServ As String
Global sChanServ As String
Global sMemoServ As String
Global sOperServ As String
Global sNetServ As String

Global iSec(1 To 2) As Integer
Global iMin(1 To 2) As Integer
Global iHour As Integer
Global iDays As Integer
Global iEDITED(2) As Integer
Global SysFile As String ' ircd.conf file
Global MOTDFile As String ' ircd.motd file
Global iKILL(iUserMax) As Boolean
Global iFLU As Boolean
Global sACC As Integer ' Allow Connected Client

Dim RefArray(1 To 36) As String * 1

Public Function CButton(Button As CommandButton) As Long
    SendMessage Button.hwnd, BM_SETSTYLE, BS_SOLID, 1
End Function

Public Sub subAlwaysOnTop(bolIf As Boolean, objWindow As Object)

'Example: Write this into a form:
'   subAlwaysOnTop True, Me
'The bolIf value if it will be on top (false to disable)
Dim lngFlag As Long
    If bolIf Then
        lngFlag = HWND_TOPMOST
    Else
        lngFlag = HWND_NOTOPMOST
    End If
    Call SetWindowPos(objWindow.hwnd, lngFlag, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE))
End Sub

Function SysFOCC(Text As String, Optional OverRide As Boolean) As String
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
Dim TmpDone As String
Dim TmpText As String
Dim TmpCheck As String
    
    If OverRide = False Then
        SysFOCC = Text
        Exit Function
    End If
    
ReScan:

    X = InStr(1, Text, Chr$(3))
    If X = 0 Then
        SysFOCC = TmpDone & Text
        Exit Function
    End If
    
    TmpCheck = Mid$(Text, X, 3)
    If TmpCheck Like Chr$(3) & "##" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 3)
        GoTo ReScan
    End If
    
    TmpCheck = Mid$(Text, X, 5)
    If TmpCheck Like Chr$(3) & "##,#" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 5)
        GoTo ReScan
    End If
    
    TmpCheck = Mid$(Text, X, 6)
    If TmpCheck Like Chr$(3) & "##,##" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 6)
        GoTo ReScan
    End If
    
    TmpCheck = Mid$(Text, X, 5)
    If TmpCheck Like Chr$(3) & "#,##" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 5)
        GoTo ReScan
    End If
    
    TmpCheck = Mid$(Text, X, 4)
    If TmpCheck Like Chr$(3) & "#,#" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 4)
        GoTo ReScan
    End If
    
    TmpCheck = Mid$(Text, X, 2)
    If TmpCheck Like Chr$(3) & "#" Then
        TmpText = Mid$(Text, 1, X - 1)
        TmpDone = TmpDone & TmpText
        Text = Mid$(Text, X + 2)
        GoTo ReScan
    End If
    
    TmpText = Mid$(Text, 1, X - 1)
    TmpDone = TmpDone & TmpText
    Text = Mid$(Text, X + 1)
    GoTo ReScan
    
End Function

Function SysCFCC(Text As String) As Boolean
On Error Resume Next
Dim X As Integer
    X = InStr(1, Text, Chr$(3))
    If X = 0 Then
        SysCFCC = False
    Else
        SysCFCC = True
    End If
End Function

Function sCTS(ByVal sglTime As Long) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        sCTS = Format(sglTime, "0") & "sec"
    Case 60 To 3599 ' Minutes Seconds
        sCTS = Format(Int(sglTime / 60), "#0") & "min " & Format(sglTime Mod 60, "0") & "sec"
    Case Else       ' Hours Minutes
        sCTS = Format(Int(sglTime / 3600), "#0") & "hr " & Format(sglTime / 60 Mod 60, "0") & "min " & Format(sglTime Mod 60, "0") & "sec"
End Select

End Function

Function sNickValid(Nick As String) As Boolean
On Error Resume Next
Dim X As Integer
Dim TmpNick As String
Dim TmpChar As String
Dim TmpText As String
    TmpNick = Nick
    
    TmpText = Mid$(TmpNick, 1, 1)
    X = InStr(1, "1234567890", TmpText)
    If Not X = 0 Then GoTo NotValid

ReScan:
    If Not Len(TmpNick) = 0 Then
        TmpChar = Mid$(TmpNick, 1, 1)
        TmpNick = Mid$(TmpNick, 2)
        For X = 65 To 125
            If Chr$(X) = TmpChar Then
                GoTo ReScan
                Exit For
            End If
        Next X
        For X = 45 To 57
            If Chr$(X) = TmpChar Then
                If X = "46" Then GoTo NotValid
                If X = "47" Then GoTo NotValid
                GoTo ReScan
                Exit For
            End If
        Next X
        
NotValid:
        sNickValid = False
        Exit Function
    End If
    sNickValid = True
    
    'Valid Chars:   |   }   {   a-z    `  _  ^  ]  [  A-Z   0-9
    '              124 125 123 97-122 96 95 94 93 91 65-90 48-57
    '        48-57 65-125, '92' = Not Valid
    'AWAY Send ':ServerName 301 <Nick> <AwayNick> :Msg' CRLF
End Function


Function sHostValid(Host As String) As Boolean
On Error Resume Next
Dim X As Integer
Dim TmpHost As String
Dim TmpChar As String
Dim TmpText As String
    TmpHost = Host

ReScan:
    If Not Len(TmpHost) = 0 Then
        TmpChar = Mid$(TmpHost, 1, 1)
        TmpHost = Mid$(TmpHost, 2)
        For X = 65 To 122
            If Chr$(X) = TmpChar Then
                If X = "90" Then GoTo NotValid
                If X = "91" Then GoTo NotValid
                If X = "92" Then GoTo NotValid
                If X = "93" Then GoTo NotValid
                If X = "94" Then GoTo NotValid
                If X = "95" Then GoTo NotValid
                If X = "96" Then GoTo NotValid
                GoTo ReScan
                Exit For
            End If
        Next X
        For X = 45 To 57
            If Chr$(X) = TmpChar Then
                If X = "47" Then GoTo NotValid
                GoTo ReScan
                Exit For
            End If
        Next X
        
NotValid:
        sHostValid = False
        Exit Function
    End If
    sHostValid = True
    
    'Valid Chars:   a-z    A-Z   0-9
    '              97-122 65-90 48-57
    '        47, 90-96 = Not Valid
End Function


Function sConvertText(Text As String, Index As Integer) As String
On Error Resume Next
Dim TmpData As String
Dim TmpText As String
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
Dim TmpReplace As String
    Z = 1
ReScan:
    
    If Z > Len(Text) Or Z = Len(Text) Then
        sConvertText = Text
        Exit Function
    End If
    X = InStr(Z, Text, "%")
    Q = InStr(X + 1, Text, "%")
    If X = 0 Or Q = 0 Then
        sConvertText = Text
        Exit Function
    End If
        TmpText = Mid$(Text, X, Q - X + 1)
            Select Case LCase(TmpText)
            
                Case "%server%": TmpReplace = sServer
                'Case "%serverdesc%": TmpReplace = iServDSC
                Case "%serverip%": TmpReplace = frmMain.Win(Index).LocalAddress
                'Case "%support%": TmpReplace = iSupportEmail
                Case "%clusers%": TmpReplace = frmMain.lbl_CU
                Case "%mlusers%": TmpReplace = frmMain.lbl_HU
                Case "%cgusers%": TmpReplace = frmMain.lbl_CGU
                Case "%mgusers%": TmpReplace = frmMain.lbl_HGU
                Case "%iusers%": TmpReplace = frmMain.lbl_IU
                Case "%nusers%": TmpReplace = frmMain.lbl_CU - frmMain.lbl_IU
                Case "%fchans%": TmpReplace = frmMain.lbl_CC
                Case "%ircops%": TmpReplace = frmMain.lbl_CO
                Case "%lservers%": TmpReplace = "0"
                Case "%servers%": TmpReplace = frmMain.lbl_CS
                Case "%nick%": If Index = 0 Then TmpReplace = "John" Else: TmpReplace = iUser(Index)
                Case "%user%": If Index = 0 Then TmpReplace = "Doe" Else: TmpReplace = iName(Index)
                Case "%host%": If Index = 0 Then TmpReplace = "BillGates.DualT3.ISP.Microsoft.com" Else: TmpReplace = iRHost(Index)
                Case "%name%": If Index = 0 Then TmpReplace = "Mr. Lamer ;)" Else: TmpReplace = iRealName(Index)
                Case "%port%": If Index = 0 Then TmpReplace = "6667" Else: TmpReplace = frmMain.Win(Index).LocalPort
                Case "%modes%": If Index = 0 Then TmpReplace = "dGi" Else: TmpReplace = iModes(Index)
                Case "%flags%": If Index = 0 Then TmpReplace = "dGi" Else: TmpReplace = iModes(Index)
                Case "%idle%": If Index = 0 Then TmpReplace = "3450" Else: TmpReplace = iIdle(Index)
                Case "%cidle%": If Index = 0 Then TmpReplace = sCTS("3450") Else: TmpReplace = sCTS(iIdle(Index))
                Case "%ip%": If Index = 0 Then TmpReplace = "207.46.130.45" Else: TmpReplace = frmMain.Win(Index).PeerAddress
                'Case "%network%": TmpReplace = iNetName
                'Case "%admin%": TmpReplace = iServAdmin
                'Case "%adminmail%": TmpReplace = iServAdminE
                'Case "%conlimit%": TmpReplace = frmMain.iConnMax
                Case "%version%": TmpReplace = sVersion
                'Case "%chanlimit%": TmpReplace = iChanMax
                'Case "%helpchan%": TmpReplace = iHelpChan
                'Case "%mainchan%": TmpReplace = iMainChan
                Case "%macrover%": TmpReplace = "v1.0" 'MacroKeys Version :)
                Case "%uptime%": TmpReplace = "[" & frmMain.lbl_UT & "]"
                'Case "%cpuspeed%": TmpReplace = iSys_CPUspeed
                'Case "%cputype%": TmpReplace = iSys_CPUtype   ' The Code required for these items does not come with vbIRCd
                'Case "%osversion%": TmpReplace = iSys_OSVerMajor & "." & iSys_OSVerMiner
                'Case "%os%": TmpReplace = iSys_OSName
                'Case "%ix86%": TmpReplace = iSys_ix86type
                'Case "%unixtime%": TmpReplace = GetTime
                Case "%time%": TmpReplace = Format(Now, "hh:mm:ss")
                Case "%date%": TmpReplace = Format(Now, "MM/DD/YYYY")
                Case "%month%": TmpReplace = Format(Now, "mmmm")
                Case "%day%": TmpReplace = Format(Now, "dddd")
                Case "%year%": TmpReplace = Year(Now)
                Case "%month#%": TmpReplace = Month(Now)
                Case "%day#%": TmpReplace = Day(Now)
                Case "%pubhost%": If Index = 0 Then TmpReplace = "mpx-666666.DualT3.ISP.Microsoft.com" Else: TmpReplace = iRHost(Index)
                Case "%highconc%": TmpReplace = sCCount
                Case "%concount%": TmpReplace = 0
                'Case "%%": TmpReplace = ""
                Case Else
                    TmpReplace = TmpText 'Lets just put back the invalid
                                         'key and not cause annoyince.
                                         
                    'TmpReplace = "[Key '" & TmpText & "' is not a Valid MacroKey]"
            End Select
    
    
    If Not X = 1 Then
        TmpData = Mid$(Text, 1, X - 1)
        TmpText = Mid$(Text, Q + 1)
        Text = TmpData & TmpReplace & TmpText
    Else
        TmpText = Mid$(Text, Q + 1)
        Text = TmpReplace & TmpText
    End If
    
    Z = X + Len(TmpReplace)
    Z = 1
    GoTo ReScan
    
End Function

Sub ScanFBU()
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim Q As Integer
Dim DUI As Boolean
Dim TmpText As String
    
    For X = 1 To iUserMax
        For Y = 1 To sAKill.Count
            DUI = False
            If iPeerFree(X) = False And LCase(iName(X) & "@" & iRHost(X)) Like LCase(sAKill(Y)) Then
                For Q = 1 To sEline.Count
                    If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                        DUI = True
                        Exit For
                    End If
                Next Q
                
                If DUI = False Then KillUser X, sServer, "AKILLED(Reason: " & sAKillR(Y) & ")"
                Exit For
            End If
        Next Y
    Next X
    
    
    For X = 1 To iUserMax
        For Y = 1 To sKline.Count
            DUI = False
            If iPeerFree(X) = False And LCase(iName(X) & "@" & iRHost(X)) Like LCase(sKline(Y)) Then
                For Q = 1 To sEline.Count
                    If iName(X) & "@" & iRHost(X) Like sEline(Q) Then
                        DUI = True
                        Exit For
                    End If
                Next Q
                
                If DUI = False Then KillUser X, sServer, "KLINED(Reason: " & sKlineR(Y) & ")"
                Exit For
            End If
        Next Y
    Next X
End Sub

Function sNameCheck(Text As String) As Boolean
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim Z As Integer
Dim TmpDone As String
Dim TmpText As String
Dim TmpCheck As String
Dim sFaid As Boolean

ReScan:
    sFaid = False
    ' This sub has been disabled since word filtering is not currently supported.
    
    'TmpCheck = Text
        'For Q = 1 To sFilter.Count
        '    If LCase(TmpCheck) Like "*" & LCase(sFilter(Q)) & "*" Then
        '        sFaid = True
        '        For Z = 1 To sFilterE.Count
        '            If LCase(TmpCheck) Like "*" & LCase(sFilterE(Z)) & "*" Then sFaid = False
        '        Next Z
        '        Exit For
        '    End If
        'Next Q
        If sFaid = True Then
            sNameCheck = True
        Else
            sNameCheck = False
        End If
End Function

Sub LogIt(Text As String)
Dim LogFile As String
Dim strData As String
Dim sDate As String
Dim sTime As String
    sTime = Format(Now, "hh:mm:ss")
    sDate = Format(Now, "MM/DD/YYYY")
    LogFile = App.Path & "\IRCdLog.txt"
    
    If Dir(LogFile, vbNormal) = "" Then
        Open LogFile For Output As #1
            Print #1, "LOG STARTED " & sDate
            DoEvents
        Close #1
    End If
    
    strData = ""
    Open LogFile For Input As #1
        strData = StrConv(InputB(LOF(1), 1), vbUnicode)
    Close #1
    strData = Mid$(strData, 1, Len(strData) - 2)
    strData = strData & vbCrLf & "[" & sVersion & ": " & sDate & " - " & sTime & "] " & Text
    
    Open LogFile For Output As #1
        Print #1, strData
        DoEvents
    Close #1
End Sub

Function uHaveMode(Index As Integer, ModeChar As String) As Boolean
Dim X As Integer
    X = InStr(1, iModes(Index), ModeChar)
    If Not X = 0 Then uHaveMode = True Else: uHaveMode = False
End Function

Sub IdentScan(Index As Integer)
On Error Resume Next
Dim X As Integer
    SendData2 Index, ":" & sServer & " NOTICE AUTH :*** Looking up your hostname..." & CRLF
    iHost(Index) = frmMain.Win(Index).PeerName
    X = InStr(1, iHost(Index), ".")
    If X = 0 Then iHost(Index) = ""
    If iHost(Index) = "" Then iHost(Index) = frmMain.Win(Index).PeerAddress: SendData2 Index, ":" & sServer & " NOTICE AUTH :*** Could not find your hostname, using IP address instead" & CRLF Else SendData2 Index, ":" & sServer & " NOTICE AUTH :*** Found your hostname" & CRLF
    SendData2 Index, ":" & sServer & " NOTICE AUTH :*** Now processing incoming socket data..." & CRLF
    SendData Index, ":" & sServer & " NOTICE AUTH :*** If you need assistance with a connection problem, please email " & iSupportEmail & " with the name and version of the client you are using, and the server you tried to connect to: " & sServer & CRLF
    SYS iHoldData(Index), Index
    iHolted(Index) = False
    iHoldData(Index) = ""
    iRHost(Index) = iHost(Index)
End Sub

Function sCloakHost(Index As Integer) As String
On Error Resume Next
Dim X As Integer
Dim Y As Integer
Dim iHostID As Long
Dim sHost As String
    'txt_NHHP = Cloak Prefix
    sHost = iRHost(Index)
    For X = 1 To Len(sHost)
        iHostID = iHostID + Asc(Mid$(sHost, X, 1))
    Next X
    
    iHostID = (((iHostID * Len(sHost) / 4) * 6) / 3) * 2
    
    
    If sHost Like "[0-9]*.*.*.*" Then
        X = InStrRev(sHost, ".")
        If Not X = 0 Then iHost(Index) = Mid$(sHost, 1, Len(sHost) - (Len(sHost) - X)) & iHiddenPrefix & "-" & iHostID
    Else
        X = InStr(1, sHost, ".")
        If Not X = 0 Then iHost(Index) = iHiddenPrefix & "-" & iHostID & Mid$(sHost, X)
    End If
    sCloakHost = iHost(Index)
End Function

Function ReplaceStr(sData As String, sFind As String, sReplace As String) As String
On Error Resume Next
Dim X As Integer
Dim TmpText As String
Dim TmpLoad As String
    X = InStr(1, sData, sFind)
    TmpText = Mid$(sData, 1, X - 1)
    TmpLoad = Mid$(sData, X + Len(sFind))
    ReplaceStr = TmpText & sReplace & TmpLoad
End Function

Sub DisplayLog(Text As String, Optional LogAsERROR As Boolean)
On Error Resume Next
With frmMain.txt_Log
    .SelStart = Len(.Text)
    .SelText = "[" & Format(Now, "MM/DD/YYYY") & " " & Format(Now, "hh:mm") & "] " & Text & vbCrLf
    .SelStart = Len(.Text)
    If LogAsERROR = True Then frmMain.lbl_ToolBar(1).BackColor = &HFF
End With
End Sub

