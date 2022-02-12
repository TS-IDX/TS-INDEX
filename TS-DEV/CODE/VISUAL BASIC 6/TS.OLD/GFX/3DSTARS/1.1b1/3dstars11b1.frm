VERSION 5.00
Begin VB.Form frm3dstars 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " 3D Stars"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6900
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "3dstars11b1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "thomas.schilb{at}live.de"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "www.thomasschilb.tk"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2019 TSC. All Rights Reserved"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "3D Stars v1.1 Beta1 (720p in a form)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frm3dstars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' 3D Stars v1.1 Beta1 (720p in a form)
'
' Copyright © 2019 TSC. All Rights Reserved
'
'
' www.thomasschilb.tk
' thomas.schilb@live.de
'
Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Private Sub Form_Activate()
Speed = -5
K = 2038
Zoom = 256
Timer1.Interval = 1
For i = 0 To 100
X(i) = Int(Rnd * 1024) - 512
Y(i) = Int(Rnd * 1024) - 512
Z(i) = Int(Rnd * 512) - 256
Next i
End Sub

Private Sub Timer1_Timer()
For i = 0 To 100
Circle (tmpX(i), tmpY(i)), 5, BackColor
Z(i) = Z(i) + Speed
If Z(i) > 255 Then Z(i) = -255
If Z(i) < -255 Then Z(i) = 255
tmpZ(i) = Z(i) + Zoom
tmpX(i) = (X(i) * K / tmpZ(i)) + (frm3dstars.Width / 15)
tmpY(i) = (Y(i) * K / tmpZ(i)) + (frm3dstars.Height / 5)
Radius = 1
StarColor = 256 - Z(i)
Circle (tmpX(i), tmpY(i)), 5, RGB(StarColor, StarColor, StarColor)
Next i
End Sub
