VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current Users"
   ClientHeight    =   4035
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   100
      Left            =   6240
      Top             =   3060
   End
   Begin VB.Frame fraOpts 
      Caption         =   "User Options"
      Height          =   1575
      Left            =   3660
      TabIndex        =   11
      Top             =   2400
      Width           =   3855
      Begin VB.CommandButton cmdOline 
         Caption         =   "Create O-Line"
         Height          =   375
         Left            =   2580
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdKline 
         Caption         =   "K-line"
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdMessage 
         Caption         =   "Send Message"
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdRawMsg 
         Caption         =   "Raw Message"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdAssignNick 
         Caption         =   "Assign Nick"
         Height          =   375
         Left            =   60
         TabIndex        =   14
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdKill 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info..."
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "User Statistics"
      Height          =   1575
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Width           =   3555
      Begin VB.TextBox txtflood 
         Height          =   255
         Left            =   1380
         TabIndex        =   10
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox txtSignedOn 
         Height          =   285
         Left            =   1380
         TabIndex        =   9
         Top             =   900
         Width           =   1995
      End
      Begin VB.TextBox txtBR 
         Height          =   285
         Left            =   1380
         TabIndex        =   8
         Top             =   540
         Width           =   1995
      End
      Begin VB.TextBox txtBS 
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "Flood %:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Signed On"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Bytes Recieved"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bytes Sent"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu mnuUserDisconnect 
         Caption         =   "Disconnect..."
      End
      Begin VB.Menu mnuUserSendMessage 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuUserLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserAssignNick 
         Caption         =   "Assign Nick"
      End
      Begin VB.Menu mnuUserLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserRawMsg 
         Caption         =   "Raw message"
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim CurUser As Long

Private Sub cmdAssignNick_Click()
ChangeNick CurUser, InputBox("Enter a new Nick")
Command1_Click
End Sub

Private Sub cmdInfo_Click()
frmUserInfo.strUser = Users(CurUser).Nick
frmUserInfo.Show 1, Me
End Sub

Private Sub cmdKill_Click()
If CurUser <= 3 Then
    MsgBox "Can't kill a service"
    Exit Sub
End If
Dim User As clsUser, Comment As String
Comment = InputBox("Enter a comment")
Set User = Users(CurUser)
User.Killed = True
SendWsock CurUser, ":Server!~admin@" & ServerName & " KILL " & User.Nick & " :" & Comment, True
SendWsock CurUser, "ERROR :Closing Link: " & User.Nick & "[" & frmMain.wsock(User.Index).RemoteHostIP & ".] " & ServerName & " (" & Comment & ")", True
SendQuit User.Index, "Killed by Server Admin (" & Comment & ")", True
Dim NN As Long
NN = GetRand
Load frmMain.tmrKill(NN)
frmMain.tmrKill(NN).Tag = User.Index
frmMain.tmrKill(NN).Enabled = True
End Sub

Private Sub cmdKline_Click()
Klines.Add frmMain.wsock(CurUser).RemoteHostIP, frmMain.wsock(CurUser).RemoteHostIP
End Sub

Private Sub cmdMessage_Click()
SendNotice Users(CurUser).Nick, InputBox("Enter a message"), "ServerAdmin"
End Sub

Private Sub cmdOline_Click()
On Error Resume Next
If Not IsRegistered(Users(CurUser).Nick) Then
    MsgBox "Nickname not registered!", vbCritical
    Exit Sub
End If
If HasOline(Users(CurUser).Nick, Users(CurUser).GetMask) Then
    MsgBox "O-Line already exists!", vbCritical
    Exit Sub
End If
If Users(CurUser).DNS = "" Then
    MsgBox "No DNS found, O-Line will be temporary!", vbCritical
    Dim CurOline As Long
    CurOline = GetFreeOLine
    With Olines(CurOline)
        .UserName = Users(CurUser).Nick
        .Password = "Generic"
        .Mask = Users(CurUser).GetMask
        .InUse = True
    End With
    SendNotice Users(CurUser).Nick, "An O-line has been created for you. type '/oper " & Users(CurUser).Nick & " Generic' to use it", "Dill.mine.nu"
    Exit Sub
End If
CurOline = GetFreeOLine
With Olines(CurOline)
    .UserName = Users(CurUser).Nick
    .Password = "Generic"
    .Mask = Users(CurUser).GetMask
    .InUse = True
End With
SaveOlines
MsgBox "O-Line successfully created", vbInformation
SendNotice Users(CurUser).Nick, "An O-line has been created for you. type '/oper " & Users(CurUser).Nick & " Generic' to use it", "Dill.mine.nu"
End Sub

Private Sub cmdRawMsg_Click()
SendWsock CurUser, InputBox("Enter the Raw Command")
End Sub

Private Sub Command1_Click()
Dim Item As ListItem, User As clsUser, i As Long
lvwUsers.ListItems.Clear
For i = 1 To UBound(Users)
    If Not Users(i) Is Nothing Then
        Set User = Users(i)
        Set Item = lvwUsers.ListItems.Add(, User.Nick, User.Nick)
        Item.SubItems(1) = User.Email
        Item.SubItems(2) = User.Name
        Item.SubItems(3) = User.DNS
        Item.SubItems(4) = User.GetOnChans
        Item.SubItems(5) = User.GetModes
    End If
Next i
End Sub

Private Sub Form_Load()
With lvwUsers.ColumnHeaders
    .Add , , "Nick"
    .Add , , "Email"
    .Add , , "Name"
    .Add , , "Dns"
    .Add , , "On Channels"
    .Add , , "Modes"
End With
End Sub

Private Sub lvwUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim User As clsUser
Set User = NickToObject(Item.Text)
CurUser = User.Index
txtBR = User.BR
txtBS = User.BS
txtSignedOn = SecsToMins2(UnixTime - User.SignOn)
txtflood = GetPercent(2500, User.MsgsSent)
End Sub

Private Sub lvwUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu mnuUser, , x, y, mnuUserDisconnect
End Sub

Private Sub mnuUserAssignNick_Click()
cmdAssignNick_Click
End Sub

Private Sub mnuUserDisconnect_Click()
cmdKill_Click
End Sub

Private Sub mnuUserRawMsg_Click()
cmdRawMsg_Click
End Sub

Private Sub mnuUserSendMessage_Click()
cmdMessage_Click
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub tmrRefresh_Timer()
On Error Resume Next
Dim Item As ListItem, User As clsUser, i As Long
If UserCount <> lvwUsers.ListItems.Count Then
    lvwUsers.ListItems.Clear
    For i = 1 To UBound(Users)
        If Not Users(i) Is Nothing Then
            Set User = Users(i)
            Set Item = lvwUsers.ListItems.Add(, User.Nick, User.Nick)
            Item.SubItems(1) = User.Email
            Item.SubItems(2) = User.Name
            Item.SubItems(3) = User.DNS
            Item.SubItems(4) = User.GetOnChans
            Item.SubItems(5) = User.GetModes
        End If
    Next i
End If
If CurUser = 0 Then Exit Sub
txtBR = Users(CurUser).BR
txtBS = Users(CurUser).BS
txtSignedOn = SecsToMins2(UnixTime - Users(CurUser).SignOn)
txtflood = GetPercent(2500, Users(CurUser).MsgsSent)
End Sub

