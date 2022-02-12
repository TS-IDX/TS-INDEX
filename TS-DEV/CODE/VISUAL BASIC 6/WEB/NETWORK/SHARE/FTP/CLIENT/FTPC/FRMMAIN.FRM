VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTPC Wrapper"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCDS 
      Caption         =   "C&lose Data Socket"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   6060
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   5580
      Width           =   3495
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "&List File"
      Height          =   495
      Left            =   4140
      TabIndex        =   4
      Top             =   5520
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   5040
      Width           =   5235
   End
   Begin VB.CommandButton cmdDat 
      Caption         =   "&Open Data Socket"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   5520
      Width           =   1155
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   4980
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   4875
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents objFTPC As FTPClient
Attribute objFTPC.VB_VarHelpID = -1

Private Sub cmdCDS_Click()
objFTPC.CloseDataSocket
End Sub

Private Sub cmdCon_Click()
objFTPC.Connect
End Sub

Private Sub cmdDat_Click()
objFTPC.OpenDataSocket
End Sub

Private Sub cmdLst_Click()
objFTPC.ListFilesInfo ""
End Sub

Private Sub Form_Load()
Set objFTPC = New FTPClient
objFTPC.UserName = "anonymous"
objFTPC.PassiveMode = True
objFTPC.Password = ""
objFTPC.RemoteHost = "127.0.0.1"
objFTPC.RemotePort = 21
objFTPC.HostType = HT_WINDOWS_NT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
objFTPC.Disconnect
objFTPC.Dispose
Set objFTPC = Nothing
End Sub

Private Sub objFTPC_IncomingData(ByVal bytesTotal As Long)
Dim Data As String
objFTPC.GetData Data, bytesTotal
AppendText Data
End Sub

Private Sub objFTPC_IncomingReply(ByVal Data As String)
AppendText Data
End Sub

Private Sub AppendText(Text As String)
Text1.SelStart = Len(Text1.Text)
Text1.SelText = Text & vbCrLf
End Sub

Private Sub objFTPC_TransferComplete()
Dim I As Long
For I = 1 To objFTPC.FileCount
    Combo1.AddItem objFTPC.FileInfo(I, FI_CREATED_TIME)
Next
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text2.Text = "connect" Then
        Call cmdCon_Click
    ElseIf Text2.Text = "initdata" Then
        Call cmdDat_Click
    Else
    objFTPC.Execute Text2.Text
    End If
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End If
End Sub
