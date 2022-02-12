VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User information"
   ClientHeight    =   1335
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   4260
      TabIndex        =   10
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   780
      TabIndex        =   9
      Top             =   660
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   780
      TabIndex        =   7
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   780
      TabIndex        =   5
      Top             =   60
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3075
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "IP"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   660
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Email"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "On Channels"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   975
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public strUser As String

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
SendNotice strUser, InputBox("Enter a message"), "" & ServerName & ""
End Sub

Private Sub Form_Load()
strUser = frmMain.lvwUsers.SelectedItem.Text
Text1 = NickToObject(strUser).Name
Text2 = NickToObject(strUser).Email
Text3 = frmMain.wsock(NickToObject(strUser).Index).RemoteHostIP
Dim i As Long
For i = 1 To NickToObject(strUser).Onchannels.Count
    Combo1.AddItem NickToObject(strUser).Onchannels(i)
Next i
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
