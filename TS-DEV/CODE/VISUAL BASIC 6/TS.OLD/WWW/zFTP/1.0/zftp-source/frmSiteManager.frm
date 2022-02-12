VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSiteManager 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Site Manager"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7920
   Icon            =   "frmSiteManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstSites 
      Height          =   4575
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteManager.frx":058A
            Key             =   "fld"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteManager.frx":09DC
            Key             =   "ofld"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteManager.frx":0E2E
            Key             =   "site"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSiteManager.frx":1280
            Key             =   "bmark"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Anonymous"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         MaskColor       =   &H00000000&
         TabIndex        =   16
         Top             =   4200
         Width           =   1455
      End
      Begin VB.OptionButton OptNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   15
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "21"
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Login Type :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sitename:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label cmdExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label cmdConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label cmdNew 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmSiteManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ************************************************************
'    ..........................................................
'     Application Name: TechFTP 1.00
'                       Complete FTP Client like CuteFTP.
'     Developer/Programmer: Muktak Joshi (muktak@techprotean.com) - http://www.techprotean.com/muktak
'    ..........................................................
'     Module Name: frmSiteManager
'     Module File: frmSiteManager.frm
'     Module Type: Form
'     Module Description:
'    ..........................................................
'
'    ..........................................................
'   ************************************************************

Option Explicit


'........................................
'Name: cmdConnect_Click
'Object: cmdConnect
'Event: Click
'Description:
'........................................
Private Sub cmdConnect_Click()
frmMain.txtUsername.Text = txtUsername.Text
frmMain.txtPassword.Text = txtPassword.Text
frmMain.txtPort.Text = txtPort.Text
frmMain.cboHost.Text = txtServer.Text
Call frmMain.cmdConnect_Click
Unload Me
End Sub


'........................................
'Name: cmdConnect_MouseMove
'Object: cmdConnect
'Event: MouseMove
'Description:
'........................................
Private Sub cmdConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ChangeButState Me, 0
    ChangeButState cmdConnect, 1

End Sub




'........................................
'Name: cmdExit_Click
'Object: cmdExit
'Event: Click
'Description:
'........................................
Private Sub cmdExit_Click()
    Unload Me
End Sub


'........................................
'Name: cmdExit_MouseMove
'Object: cmdExit
'Event: MouseMove
'Description:
'........................................
Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ChangeButState Me, 0
    ChangeButState cmdExit, 1

End Sub


'........................................
'Name: cmdNew_Click
'Object: cmdNew
'Event: Click
'Description:
'........................................
Private Sub cmdNew_Click()
lstSites.ListItems.Add , , "New Site", "site"
lstSites.SelectedItem = lstSites.ListItems.Item(lstSites.ListItems.Count)
lstSites.SelectedItem.Tag = "New Site,,,,,"
lstSites.SetFocus
End Sub


'........................................
'Name: cmdNew_MouseMove
'Object: cmdNew
'Event: MouseMove
'Description:
'........................................
Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ChangeButState Me, 0
    ChangeButState cmdNew, 1

End Sub





'........................................
'Name: Form_Load
'Object: Form
'Event: Load
'Description:
'........................................
Private Sub Form_Load()
If Dir(App.path + "\sts.dat") = "" Then Exit Sub
Dim cval As String
Dim filenum
filenum = FreeFile
Open App.path + "\sts.dat" For Input As filenum
While Not EOF(filenum)
Line Input #filenum, cval
lstSites.ListItems.Add , , GetLeftWord(cval, ",", False)
lstSites.ListItems.Item(lstSites.ListItems.Count).Tag = cval
Wend
Close filenum
End Sub


'........................................
'Name: Form_MouseMove
'Object: Form
'Event: MouseMove
'Description:
'........................................
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ChangeButState Me, 0
End Sub
'........................................
'Name: SaveDetails
'Description:
'........................................
Private Sub SaveDetails()

lstSites.SelectedItem.Tag = txtLabel.Text & "," & txtServer.Text & "," & txtUsername.Text & "," & txtPassword.Text & "," & txtPort.Text & "," & OptNormal.Value
End Sub


'........................................
'Name: Form_Unload
'Object: Form
'Event: Unload
'Description:
'........................................
Private Sub Form_Unload(Cancel As Integer)
Dim x As Integer
Dim filenum
filenum = FreeFile
Open App.path + "\sts.dat" For Output As filenum
For x = 1 To lstSites.ListItems.Count
Print #filenum, lstSites.ListItems.Item(x).Tag
Next
Close filenum
End Sub


'........................................
'Name: lstSites_ItemClick
'Object: lstSites
'Event: ItemClick
'Description:
'........................................
Private Sub lstSites_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim curdetails

curdetails = Split(Item.Tag, ",")
txtLabel.Text = curdetails(0)
txtServer.Text = curdetails(1)
txtUsername.Text = curdetails(2)
txtPassword.Text = curdetails(3)
txtPort.Text = curdetails(4)
If curdetails(5) = "True" Then
Option1.Value = False
OptNormal.Value = True
Else
Option1.Value = True
OptNormal.Value = False
End If
End Sub


'........................................
'Name: Option1_Click
'Object: Option1
'Event: Click
'Description:
'........................................
Private Sub Option1_Click()
Call SaveDetails
txtUsername.Text = "Anonymous"
End Sub


'........................................
'Name: OptNormal_Click
'Object: OptNormal
'Event: Click
'Description:
'........................................
Private Sub OptNormal_Click()
Call SaveDetails
End Sub


'........................................
'Name: txtLabel_Change
'Object: txtLabel
'Event: Change
'Description:
'........................................
Private Sub txtLabel_Change()
lstSites.SelectedItem.Text = txtLabel.Text
Call SaveDetails
End Sub


'........................................
'Name: txtPassword_Change
'Object: txtPassword
'Event: Change
'Description:
'........................................
Private Sub txtPassword_Change()
Call SaveDetails
End Sub


'........................................
'Name: txtPort_Change
'Object: txtPort
'Event: Change
'Description:
'........................................
Private Sub txtPort_Change()
Call SaveDetails
End Sub


'........................................
'Name: txtServer_Change
'Object: txtServer
'Event: Change
'Description:
'........................................
Private Sub txtServer_Change()
Call SaveDetails
End Sub


'........................................
'Name: txtUsername_Change
'Object: txtUsername
'Event: Change
'Description:
'........................................
Private Sub txtUsername_Change()
Call SaveDetails
End Sub

