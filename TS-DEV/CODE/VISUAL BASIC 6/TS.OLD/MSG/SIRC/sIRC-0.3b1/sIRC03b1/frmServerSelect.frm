VERSION 5.00
Begin VB.Form frmServerSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Servers"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton anonymous 
      Caption         =   "Anonymous"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton ccc 
      Caption         =   "CCC"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton buha 
      Caption         =   "BuHa"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "IRC Servers:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmServerSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub anonymous_Click()
frmOption.txtServer.Text = "irc.anonnet.org"
frmOption.txtPort.Text = "6667"
frmOption.txtNick.Text = "sirc_user"
frmOption.txtMail.Text = "anonymous@thomasschilb.de"
frmOption.txtRealName.Text = "Thomas Schilb"
Unload Me
End Sub

Private Sub buha_Click()
frmOption.txtServer.Text = "irc.euirc.net"
frmOption.txtPort.Text = "6667"
frmOption.txtNick.Text = "sirc_user"
frmOption.txtMail.Text = "anonymous@thomasschilb.de"
frmOption.txtRealName.Text = "Thomas Schilb"
Unload Me
End Sub

Private Sub ccc_Click()
frmOption.txtServer.Text = "irc.hackint.org"
frmOption.txtPort.Text = "6667"
frmOption.txtNick.Text = "sirc_user"
frmOption.txtMail.Text = "anonymous@thomasschilb.de"
frmOption.txtRealName.Text = "Thomas Schilb"
Unload Me
End Sub
