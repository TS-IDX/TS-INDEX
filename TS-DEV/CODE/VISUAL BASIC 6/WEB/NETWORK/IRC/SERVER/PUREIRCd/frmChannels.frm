VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmChannels 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Current Channels"
   ClientHeight    =   3735
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrRefreshUser 
      Interval        =   100
      Left            =   5160
      Top             =   3300
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Frame fraChanUsers 
      Caption         =   "User Options"
      Height          =   1155
      Left            =   4200
      TabIndex        =   4
      Top             =   2100
      Width           =   3975
      Begin VB.CommandButton cmdBan 
         Caption         =   "Ban"
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdKick 
         Caption         =   "Kick"
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeVoice 
         Caption         =   "DeVoice"
         Height          =   375
         Left            =   1380
         TabIndex        =   11
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdVoice 
         Caption         =   "Voice"
         Height          =   375
         Left            =   1380
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeOp 
         Caption         =   "DeOp"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "Op"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraChan 
      Caption         =   "Channel Options"
      Height          =   1155
      Left            =   60
      TabIndex        =   2
      Top             =   2100
      Width           =   4035
      Begin VB.CommandButton cmdClearChan 
         Caption         =   "Clear Channel"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdTopic 
         Caption         =   "Change Topic"
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtTopic 
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Width           =   3915
      End
   End
   Begin MSComctlLib.ListView lvwChannels 
      Height          =   2055
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
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
      Left            =   6960
      TabIndex        =   0
      Top             =   3300
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwChanUsers 
      Height          =   2055
      Left            =   5040
      TabIndex        =   3
      Top             =   60
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
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
End
Attribute VB_Name = "frmChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim CurChan As clsChannel, CurUser As clsUser

Private Sub cmdBan_Click()
BanUser CurChan, "*!*@" & frmMain.wsock(CurUser.Index).RemoteHostIP, "ChanServ"
cmdRefresh_Click
End Sub

Private Sub cmdClearChan_Click()
Dim Item As Variant
For Each Item In CurChan.All
    KickUser "ChanServ", CurChan.Name, CStr(Item), "Clear command used by ChanServ", True
Next
cmdRefresh_Click
End Sub

Private Sub cmdDeOp_Click()
DeOpUser CurChan, CurUser.Nick, "ChanServ"
cmdRefresh_Click
End Sub

Private Sub cmdDeVoice_Click()
DeVoiceUser CurChan, CurUser.Nick, "ChanServ"
cmdRefresh_Click
End Sub

Private Sub cmdKick_Click()
KickUser "ChanServ", CurChan.Name, CurUser.Nick, "Kicked by ChanServ", True
cmdRefresh_Click
End Sub

Private Sub cmdOp_Click()
OpUser CurChan, CurUser.Nick, "ChanServ"
cmdRefresh_Click
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
Dim Item As ListItem, User As clsUser, i As Long
lvwChannels.ListItems.Clear
For i = 1 To UBound(Channels)
    If Not Channels(i) Is Nothing Then
        Set Item = lvwChannels.ListItems.Add(, Channels(i).Name, Channels(i).Name)
        Item.SubItems(1) = Channels(i).Topic
        Item.SubItems(2) = Channels(i).All.Count
    End If
    If Not Channels(i) Is Nothing Then Set CurChan = Channels(i)
Next i
If CurChan.All.Count <> lvwChanUsers.ListItems.Count Then
    lvwChanUsers.ListItems.Clear
    For i = 1 To CurChan.All.Count
        If Not Users(i) Is Nothing Then
            Set User = NickToObject(CurChan.All(i))
            Set Item = lvwChanUsers.ListItems.Add(, User.Nick, User.Nick)
            Item.SubItems(1) = User.GetChanModes(CurChan.Name)
        End If
    Next i
End If
End Sub

Private Sub cmdTopic_Click()
SetTopic CurChan.Name, InputBox("Enter a new Topic"), "ChanServ"
cmdRefresh_Click
End Sub

Private Sub cmdVoice_Click()
VoiceUser CurChan, CurUser.Nick, "ChanServ"
cmdRefresh_Click
End Sub

Private Sub Form_Load()
lvwChannels.ColumnHeaders.Add , , "Name"
lvwChannels.ColumnHeaders.Add , , "Topic"
lvwChannels.ColumnHeaders.Add , , "UserCount"
lvwChanUsers.ColumnHeaders.Add , , "Nick"
lvwChanUsers.ColumnHeaders.Add , , "Channel Modes"
End Sub

Private Sub lvwChannels_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim User As clsUser, i As Long
Set CurChan = ChanToObject(Item.Text)
If CurChan.All.Count <> lvwChanUsers.ListItems.Count Then
    lvwChanUsers.ListItems.Clear
    For i = 1 To CurChan.All.Count
        If Not Users(i) Is Nothing Then
            Set User = NickToObject(CurChan.All(i))
            Set Item = lvwChanUsers.ListItems.Add(, User.Nick, User.Nick)
            Item.SubItems(1) = User.GetChanModes(CurChan.Name)
        End If
    Next i
End If
End Sub

Private Sub lvwChanUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set CurUser = NickToObject(Item.Text)
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub tmrRefreshUser_Timer()
On Error Resume Next
If Not ChanCount = lvwChannels.ListItems.Count Then
    lvwChannels.ListItems.Clear
    Dim i As Long, Item As ListItem
    For i = 1 To UBound(Channels)
        If Not Channels(i) Is Nothing Then
            Set Item = lvwChannels.ListItems.Add(, Channels(i).Name, Channels(i).Name)
            Item.SubItems(1) = Channels(i).Topic
            Item.SubItems(2) = Channels(i).All.Count
        End If
        If Not Channels(i) Is Nothing Then Set CurChan = Channels(i)
    Next i
End If
txtTopic = CurChan.Topic
End Sub
