VERSION 5.00
Begin VB.Form frmJoin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Request"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton JoinCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton JoinOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox JoinText 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Enter channel name:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub JoinOK_Click()
frmSocket.Socket.SendData "JOIN #" & JoinText.Text & vbCrLf
Unload Me
End Sub

