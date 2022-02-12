VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "sIRC"
   ClientHeight    =   7545
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSelectServer 
         Caption         =   "Select Server..."
         Shortcut        =   ^S
      End
      Begin VB.Menu hy00 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuReConnect 
         Caption         =   "ReConnect"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu hy01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuColors 
         Caption         =   "Colors..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "Commands"
      Begin VB.Menu mnuJoin 
         Caption         =   "Join channel"
      End
      Begin VB.Menu mnuPart 
         Caption         =   "Part channel"
      End
      Begin VB.Menu mnuQuery 
         Caption         =   "Query user"
      End
      Begin VB.Menu msgSendNotice 
         Caption         =   "Send notice"
      End
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois user"
      End
      Begin VB.Menu mnuSendCTCP 
         Caption         =   "Send CTCP"
         Begin VB.Menu mnuSendCTCPPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu mnuSendCTCPTime 
            Caption         =   "Time"
         End
         Begin VB.Menu mnuSendCTCPVersion 
            Caption         =   "Version"
         End
      End
      Begin VB.Menu mnuSetAway 
         Caption         =   "Set Away"
         Begin VB.Menu mnuSetAwayOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuSetAwayOff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu mnuInvite 
         Caption         =   "Invite usr"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban user"
      End
      Begin VB.Menu mnuKick 
         Caption         =   "Kick user"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change nick"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit IRC"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About sIRC..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Sub MDIForm_Load()
    LoadColor
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mnuExit_Click
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuColors_Click()
frmColor.Show
End Sub

Private Sub mnuConnect_Click()
    MyNick = frmOption.txtNick
    Email = frmOption.txtMail
    RealName = frmOption.txtRealName
    Port = frmOption.txtPort
    
    If Len(frmOption.txtServer) = 0 Then Exit Sub
    Connect frmOption.txtServer, frmOption.txtPort
    frmStatus.Caption = "Status (" & frmOption.txtServer & ":" & frmOption.txtPort & ")"
    frmStatus.rtfStatus.SelColor = vbWhite
    frmStatus.rtfStatus.SelText = " *** Connecting to Server " & vbCrLf
End Sub

Private Sub mnuContents_Click()
'Open "sirc.chm" For Output As #1
'Open "sirc.chm" For Binary As #1
Dim r As Long
r = StartDoc("sirc.chm")
End Sub
Function StartDoc(DocName As String) As Long
Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
'change "Open" to "Explore" to bring up file explorer
StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "", 1)
'end function
End Function

Private Sub mnuDisconnect_Click()
Disconnect
End Sub

Private Sub mnuExit_Click()
    Disconnect
    End
End Sub

Private Sub mnuOption_Click()
    frmOption.Show 1
End Sub

Private Sub mnuJoin_Click()
frmJoin.Show 1
End Sub

Private Sub mnuOptions_Click()
    frmOption.Show 1
End Sub

Private Sub mnuSelectServer_Click()
frmServerSelect.Show
End Sub
