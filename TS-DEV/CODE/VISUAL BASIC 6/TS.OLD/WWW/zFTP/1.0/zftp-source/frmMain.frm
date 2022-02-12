VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "zFTP 1.0"
   ClientHeight    =   9345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3360
      Top             =   5040
   End
   Begin MSWinsockLib.Winsock sckData 
      Index           =   0
      Left            =   4080
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckConn 
      Index           =   0
      Left            =   4680
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame FRemote 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   6240
      TabIndex        =   1
      Top             =   3240
      Width           =   5895
      Begin MSComctlLib.ImageCombo cboRemote 
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImgFileSys"
      End
      Begin VB.CommandButton cmdRemoteUp 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.ListView lstRemote 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgFileSys"
         SmallIcons      =   "ImgFileSys"
         ColHdrIcons     =   "ImgFileSys"
         ForeColor       =   16777215
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Attrib"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLookInRemote 
         BackColor       =   &H00000000&
         Caption         =   "Remote:"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   6015
      Begin VB.CommandButton cmdLocalUP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.ImageList ImgGeneral 
         Left            =   2160
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":014A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgFileSys 
         Left            =   1200
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":059C
               Key             =   "fld"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":09EE
               Key             =   "ofld"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0E40
               Key             =   "fdd"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1292
               Key             =   "cd"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16E4
               Key             =   "hdd"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B36
               Key             =   "mycomp"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F88
               Key             =   "fil"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgTools 
         Left            =   480
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23DA
               Key             =   "conn"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":282C
               Key             =   "graph"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2C7E
               Key             =   "site"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":30D0
               Key             =   "options"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3522
               Key             =   "stop"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3974
               Key             =   "disc"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3DC6
               Key             =   "ul"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4218
               Key             =   "dl"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":466A
               Key             =   "fo"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboLookInLocal 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImageList       =   "ImgFileSys"
      End
      Begin MSComctlLib.ListView lstLocal 
         Height          =   2655
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgFileSys"
         SmallIcons      =   "ImgFileSys"
         ColHdrIcons     =   "ImgFileSys"
         ForeColor       =   16777215
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLookinLocal 
         BackStyle       =   0  'Transparent
         Caption         =   "Local:"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstQueue 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Local"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTools 
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImgTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            ImageKey        =   "conn"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Disconnect"
            ImageKey        =   "disc"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Upload Select File(s)"
            ImageKey        =   "ul"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Download Select File(s)"
            ImageKey        =   "dl"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop Transfer"
            ImageKey        =   "stop"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "File/Folder Operations"
            ImageKey        =   "fo"
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Site Manager"
            ImageKey        =   "site"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FQuick 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   12015
      Begin VB.TextBox cboHost 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9360
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6960
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label cmdConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connect "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10200
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8640
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1320
      Width           =   12015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© 2017 // thomas.schilb@live.de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   9000
      Width           =   12015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   12120
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   24
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "zFTP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   12120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSM 
         Caption         =   "&Site Manager"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileQC 
         Caption         =   "&Quick Connect"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFileS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRC 
         Caption         =   "&Reconnect"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileDC 
         Caption         =   "&Disconnect"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuFileS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEX 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGS 
         Caption         =   "Group S&elect"
      End
      Begin VB.Menu mnuEditGD 
         Caption         =   "Group &Deselect"
      End
      Begin VB.Menu mnuEditIS 
         Caption         =   "&Invert Selection"
      End
      Begin VB.Menu mnuEditSA 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditS2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSettings 
         Caption         =   "Settings"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSL 
         Caption         =   "S&hort Listing"
      End
      Begin VB.Menu mnuViewLL 
         Caption         =   "&Long Listing"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFD 
         Caption         =   "File Details"
         Begin VB.Menu mnuViewFD1 
            Caption         =   "show file size"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewFD2 
            Caption         =   "show file date"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewFD3 
            Caption         =   "show file attribute"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRef 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewQCB 
         Caption         =   "Quick Connect Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About zFTP"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopTr 
         Caption         =   "Upload"
      End
      Begin VB.Menu mnuPopTras 
         Caption         =   "Upload as.."
      End
      Begin VB.Menu mnuPopATQ 
         Caption         =   "Add to Queue"
      End
      Begin VB.Menu mnuPopView 
         Caption         =   "View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopExec 
         Caption         =   "Execute"
      End
      Begin VB.Menu mnuPopRen 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuPopDel 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuPopS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCD 
         Caption         =   "Change Directory"
      End
      Begin VB.Menu mnuPopMD 
         Caption         =   "Create Directory"
      End
      Begin VB.Menu mnuPopS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopSL 
         Caption         =   "Short Listing"
      End
      Begin VB.Menu mnuPopLL 
         Caption         =   "Long Listing"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPopS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopRef 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuQueue 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuQueueTr 
         Caption         =   "Transfer Queue"
      End
      Begin VB.Menu mnuQueueRm 
         Caption         =   "Remove from  Queue"
      End
      Begin VB.Menu mnuQueueRmall 
         Caption         =   "Remove all items"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' zFTP v1.0
'
Option Explicit
Dim CurPath As String
Dim CurRemDir As String
Dim PortStr As String
Dim Code As Integer
Dim LastCommandSent As String
Dim RootDir As String

Dim DownloadFile As String
Dim DownloadDir As String
Dim DownloadRemDir As String
Dim DownloadRemFile As String

Dim Contents As String
Dim FileSize As Long
Dim SelList As Object
Dim RMDown As Boolean
Dim LastPort As String
Dim ReceivedBytes
Dim OverWriteFlag As Boolean
Dim SettShows As Boolean


'........................................
'Name: cboLookInLocal_Click
'Object: cboLookInLocal
'Event: Click
'Description:
'........................................
Private Sub cboLookInLocal_Click()
On Error Resume Next

    CurPath = cboLookInLocal.SelectedItem.Tag
    LoadLocalList (cboLookInLocal.SelectedItem.Tag)
    LoadLocalElements cboLookInLocal.SelectedItem.Tag

End Sub


'........................................
'Name: ChangeStatus
'Description:
'........................................
Private Sub ChangeStatus(ByVal Status As String, Direction As Integer, SType As Integer)
On Error Resume Next

    If Right(txtStatus.Text, 1) = vbTab Then txtStatus.Text = Left$(txtStatus.Text, Len(txtStatus.Text) - 2)
    If Direction = 0 Then
        If SType = 1 Then
            txtStatus.Text = txtStatus.Text & "STATUS:>" & vbTab & Status & vbCrLf
        Else
            txtStatus.Text = txtStatus.Text & "COMMAND:>" & vbTab & Status & vbCrLf
        End If
    Else
        txtStatus.Text = txtStatus.Text & vbTab & vbTab & Replace(Status, vbCrLf, vbCrLf & vbTab & vbTab)   '& vbCrLf
    End If
End Sub




'........................................
'Name: cboRemote_Change
'Object: cboRemote
'Event: Change
'Description:
'........................................
Private Sub cboRemote_Change()
On Error Resume Next

    CurRemDir = cboRemote.Text
End Sub


'........................................
'Name: cboRemote_Click
'Object: cboRemote
'Event: Click
'Description:
'........................................
Private Sub cboRemote_Click()
On Error Resume Next


    SendData 0, "CWD " & cboRemote.Text
    CurRemDir = cboRemote.Text
    ChangeStatus "CWD " & cboRemote.Text, 0, 2

End Sub


'........................................
'Name: cmdConnect_Click
'Object: cmdConnect
'Event: Click
'Description:
'........................................
Public Sub cmdConnect_Click()
On Error Resume Next
If cboHost.Text = "" Then Exit Sub
    If cmdConnect.Caption = "Connect" Then
        Connect 0, cboHost.Text, txtPort.Text
        ChangeStatus "Connecting to " & cboHost.Text, 0, 1
        ChangeUIState 1
    Else
        SendData 0, "QUIT"
        DoEvents
        sckConn(0).Close
        ChangeUIState 0
        lstRemote.ListItems.Clear
        cboRemote.ComboItems.Clear
        LastCommandSent = ""
        LastPort = 1026
    End If
End Sub


'........................................
'Name: cmdConnect_MouseMove
'Object: cmdConnect
'Event: MouseMove
'Description:
'........................................
Private Sub cmdConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    ChangeButState Me, 0
    ChangeButState cmdConnect, 1

End Sub
'........................................
'Name: ChangeUIState
'Description:
'........................................
Private Sub ChangeUIState(State As Integer)
On Error Resume Next

    If State = 1 Then
        mnuFileDC.Enabled = True
        
        mnuViewRef.Enabled = True
        txtPassword.Enabled = False
        txtUsername.Enabled = False
        cboHost.Enabled = False
        txtPort.Enabled = False
        mnuFileRC.Caption = "Reconnect"
        cmdConnect.Caption = "Disconnect"

    Else
        mnuFileDC.Enabled = False
        
        mnuViewRef.Enabled = False
        txtPassword.Enabled = True
        txtUsername.Enabled = True
        cboHost.Enabled = True
        txtPort.Enabled = True
        mnuFileRC.Caption = "Connect"
        cmdConnect.Caption = "Connect"
    End If
End Sub


'........................................
'Name: SendData
'Description:
'........................................
Private Sub SendData(ByVal Index As Integer, ByVal Data As String)

    On Error Resume Next
    sckConn(Index).SendData Data & vbCrLf

End Sub
'........................................
'Name: Connect
'Description:
'........................................
Private Sub Connect(Index As Integer, Server As String, Port As String)
On Error Resume Next

    sckConn(Index).Close
    sckConn(Index).Connect Server, Port
End Sub


'........................................
'Name: cmdLocalUP_Click
'Object: cmdLocalUP
'Event: Click
'Description:
'........................................
Private Sub cmdLocalUP_Click()
On Error Resume Next

    LoadLocalList (cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 2).Tag)
    LoadLocalElements (cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 2).Tag)
    If Right(CurPath, 1) = "\" Then CurPath = Left$(CurPath, Len(CurPath) - 1)
    Call GetRightWord(CurPath, "\", True)
    If Right(CurPath, 1) = "\" Then CurPath = Left$(CurPath, Len(CurPath) - 1)
    cboLookInLocal.Text = GetRightWord(CurPath, "\", False)
    CurPath = CurPath & "\"

End Sub


'........................................
'Name: cmdRemoteUp_Click
'Object: cmdRemoteUp
'Event: Click
'Description:
'........................................
Private Sub cmdRemoteUp_Click()
On Error Resume Next

    Dim path As String
    SendData 0, "CDUP"
    ChangeStatus "CDUP", 0, 2
    DoEvents
    SendData 0, "pwd"
    ChangeStatus "PWD", 0, 2

End Sub


'........................................
'Name: SaveFile
'Description:
'........................................
Private Sub SaveFile(Content As String)
    On Error Resume Next
    Dim filenum
    Dim FSize As Long

    filenum = FreeFile

    DoEvents
    If InStr(1, AsciiList, GetRightWord(DownloadFile, ".", False), vbTextCompare) = 0 Then
        If LenB(Contents) >= (2 * FileSize) Then

            Open CurPath & DownloadFile For Binary Access Write As #filenum
            Put #filenum, 1, Contents
            Contents = ""
            Content = ""
            ReceivedBytes = 0
            lstQueue.ListItems.Remove (1)
            Timer2.Enabled = True
        End If
    Else
        Open CurPath & DownloadFile For Append As filenum
        Print #filenum, Contents
        Contents = ""
        Content = ""
        ReceivedBytes = 0
        lstQueue.ListItems.Remove (1)
        Timer2.Enabled = True
    End If

    Close filenum
    DoEvents
    LoadLocalList CurPath
    DoEvents
    StartDataSck
End Sub




'........................................
'Name: FileExists
'Description:
'........................................
Private Function FileExists(FileName As String) As Boolean
On Error Resume Next

Dim x As Integer
If SelList.Name = "lstLocal" Then
For x = 1 To lstRemote.ListItems.Count
If lstRemote.ListItems.Item(x) = FileName Then
FileExists = True
Exit Function
End If
Next
Else
For x = 1 To lstLocal.ListItems.Count
If lstLocal.ListItems.Item(x) = FileName Then
FileExists = True
Exit Function
End If
Next
End If
End Function


'........................................
'Name: Form_Activate
'Object: Form
'Event: Activate
'Description:
'........................................
Private Sub Form_Activate()
On Error Resume Next

    LoadLocalList CurPath
    LoadLocalElements CurPath
    Set SelList = lstLocal
    
    If SettShows = False Then
    frmSiteManager.Show vbModal, frmMain
    SettShows = True
    End If
    
End Sub


'........................................
'Name: Form_GotFocus
'Object: Form
'Event: GotFocus
'Description:
'........................................
Private Sub Form_GotFocus()
On Error Resume Next

    LoadLocalList CurPath
    LoadLocalElements CurPath
End Sub


'........................................
'Name: Form_Load
'Object: Form
'Event: Load
'Description:
'........................................
Private Sub Form_Load()
On Error Resume Next

SettShows = False
    LoadLocalElements (App.path + "\")
    cboLookInLocal.SelectedItem = cboLookInLocal.ComboItems.Item(2)
    LoadLocalList (App.path + "\")
    CurPath = App.path & "\"
    LastPort = 1026
    lstLocal.ColumnHeaders.Item(1).Width = lstLocal.Width / 1.8
    lstLocal.ColumnHeaders.Item(2).Width = (lstLocal.Width - lstLocal.Width / 1.8) / 2
    lstLocal.ColumnHeaders.Item(3).Width = (lstLocal.Width - lstLocal.Width / 1.8) / 2

    lstRemote.ColumnHeaders.Item(1).Width = lstRemote.Width / 1.8
    lstRemote.ColumnHeaders.Item(2).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3
    lstRemote.ColumnHeaders.Item(3).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3
    lstRemote.ColumnHeaders.Item(4).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3

    lstQueue.ColumnHeaders.Item(1).Width = lstQueue.Width / 3.5
    lstQueue.ColumnHeaders.Item(4).Width = lstQueue.Width / 3.5
    lstQueue.ColumnHeaders.Item(2).Width = (lstQueue.Width - 2 * lstQueue.ColumnHeaders.Item(1).Width) / 3
    lstQueue.ColumnHeaders.Item(3).Width = lstQueue.ColumnHeaders.Item(2).Width
    lstQueue.ColumnHeaders.Item(5).Width = lstQueue.ColumnHeaders.Item(2).Width


    mnuFileQC_Click
    ChangeUIState (0)
    Set cmdRemoteUp.Picture = ImgGeneral.ListImages.Item(1).Picture
    Set cmdLocalUP.Picture = ImgGeneral.ListImages.Item(1).Picture
    'sckData(0).LocalPort = 1026


    AsciiList = "asp,bat,c,cfm,cgi,cpp,dhtml,h,html,htm,mak,php,pl,shtml,txt"

End Sub


'........................................
'Name: Form_MouseMove
'Object: Form
'Event: MouseMove
'Description:
'........................................
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    ChangeButState Me, 0

End Sub


'........................................
'Name: FQuick_MouseMove
'Object: FQuick
'Event: MouseMove
'Description:
'........................................
Private Sub FQuick_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    ChangeButState Me, 0
End Sub


'........................................
'Name: lstLocal_Click
'Object: lstLocal
'Event: Click
'Description:
'........................................
Private Sub lstLocal_Click()
On Error Resume Next

    Set SelList = lstLocal
End Sub


'........................................
'Name: lstLocal_DblClick
'Object: lstLocal
'Event: DblClick
'Description:
'........................................
Private Sub lstLocal_DblClick()

    On Error Resume Next
    If lstLocal.SelectedItem.Tag = "fld" Then
        cboLookInLocal.Text = lstLocal.SelectedItem.Text
        CurPath = CurPath & lstLocal.SelectedItem.Text & "\"
        LoadLocalElements CurPath
        LoadLocalList CurPath
    Else
        Set SelList = lstLocal
        mnuPopTr.Caption = "Upload"
        mnuPopTr_Click

    End If
End Sub


'........................................
'Name: lstLocal_GotFocus
'Object: lstLocal
'Event: GotFocus
'Description:
'........................................


'........................................
'Name: lstLocal_ItemClick
'Object: lstLocal
'Event: ItemClick
'Description:
'........................................
Private Sub lstLocal_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If lstLocal.ListItems.Count <> 0 Then
        If RMDown = True Then
            mnuPopTr.Caption = "Upload"
            mnuPopTras.Caption = "Upload as.."
            mnuPopExec.Enabled = True

            If InStr(1, lstLocal.SelectedItem.Text, ".", vbTextCompare) = 0 Then
                mnuPopView.Enabled = False
            Else
                mnuPopView.Enabled = True
            End If
            If InStr(1, SelList.SelectedItem.Text, ".") = 0 Then
                mnuPopTras.Enabled = False
            Else
                mnuPopTras.Enabled = True
            End If
            Me.PopupMenu mnuPop
            RMDown = False
        End If
    End If
End Sub


'........................................
'Name: lstLocal_MouseDown
'Object: lstLocal
'Event: MouseDown
'Description:
'........................................
Private Sub lstLocal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 2 Then
        Set SelList = lstLocal

        RMDown = True



    End If
End Sub


'........................................
'Name: lstLocal_Validate
'Object: lstLocal
'Event: Validate
'Description:
'........................................
Private Sub lstLocal_Validate(Cancel As Boolean)
On Error Resume Next
    LoadLocalList CurPath
    LoadLocalElements CurPath
End Sub


'........................................
'Name: lstQueue_ItemClick
'Object: lstQueue
'Event: ItemClick
'Description:
'........................................
Private Sub lstQueue_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    RMDown = False
    Me.PopupMenu mnuQueue
End Sub


'........................................
'Name: lstQueue_MouseDown
'Object: lstQueue
'Event: MouseDown
'Description:
'........................................
Private Sub lstQueue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 2 Then
        RMDown = True
    End If
End Sub


'........................................
'Name: lstRemote_Click
'Object: lstRemote
'Event: Click
'Description:
'........................................
Private Sub lstRemote_Click()
On Error Resume Next
    Set SelList = lstRemote
End Sub


'........................................
'Name: lstRemote_DblClick
'Object: lstRemote
'Event: DblClick
'Description:
'........................................
Private Sub lstRemote_DblClick()
    On Error Resume Next
    Dim x As Integer
    Dim found As Boolean
    If InStr(1, lstRemote.SelectedItem.Text, ".") = 0 Then
        SendData 0, "cwd " & lstRemote.SelectedItem.Text '& "/"
        ChangeStatus "CWD " & lstRemote.SelectedItem.Text, 0, 2
        For x = 1 To lstRemote.ListItems.Count
            If lstRemote.ListItems.Item(x) = cboRemote.Text & lstRemote.SelectedItem.Text & "\" Then found = True
        Next
        If found = False Then
            cboRemote.ComboItems.Add , , cboRemote.Text & lstRemote.SelectedItem.Text & "\", "fld", "fld"
            cboRemote.Text = cboRemote.Text & lstRemote.SelectedItem.Text & "\"
        End If
    Else

        Set SelList = lstRemote
        mnuPopTr.Caption = "Download"
        mnuPopTr_Click
    End If

End Sub
'........................................
'Name: TransferDir
'Description:
'........................................
Private Sub TransferDir(ByVal Direction As Integer)
On Error Resume Next
    If Direction = 0 Then
        MsgBox "UP"
    Else

    End If
End Sub
'........................................
'Name: Transfer
'Description:
'........................................
Private Sub Transfer(ByVal Direction As Integer)
    On Error GoTo errh
    
If FileExists(DownloadFile) Then
If MsgBox("Overwrite Existing File/Folder ?", vbYesNo + vbQuestion) = vbNo Then
lstQueue.ListItems.Remove (1)
Exit Sub
End If
End If


    If InStr(1, AsciiList, GetRightWord(DownloadFile, ".", False), vbTextCompare) <> 0 Then
        ChangeStatus "TYPE A", 0, 2
        SendData 0, "TYPE A"
        DoEvents
    Else
        ChangeStatus "TYPE I", 0, 2
        SendData 0, "TYPE I"
        DoEvents
    End If
    CreatePortStr
    SendData 0, "PORT " & PortStr '& CStr(232 + (-1000 + LastPort))
    DoEvents
    ChangeStatus "PORT " & PortStr, 0, 2 '& CStr(232 + (-1000 + LastPort)), 0, 2

    CurPath = DownloadDir
    If CurRemDir <> DownloadRemDir Then
        If Len(CurRemDir) > Len(DownloadRemDir) Then
            CurRemDir = Replace(CurRemDir, DownloadRemDir, "")
            While InStr(1, CurRemDir, "\") <> 0
                SendData 0, "CDUP"
                ChangeStatus "CDUP", 0, 2
                SendData 0, "PWD"
                ChangeStatus "PWD", 0, 2
                CurRemDir = Replace(CurRemDir, "\", "", , 1)
            Wend
        Else
            SendData 0, "CWD " & Right(DownloadRemDir, Len(DownloadRemDir) - Len(CurRemDir))
            ChangeStatus "CWD " & Right(DownloadRemDir, Len(DownloadRemDir) - Len(CurRemDir)), 0, 2
            CurRemDir = DownloadRemDir
        End If
    End If

    If Direction = 0 Then
        ChangeStatus "Receiving: " & DownloadRemDir & DownloadRemFile & " --> " & DownloadDir & DownloadFile, 0, 1
        LastCommandSent = "RETR"
        SendData 0, "RETR " & DownloadRemFile
        DoEvents
        ChangeStatus "RETR " & DownloadRemFile, 0, 2
    Else
        ChangeStatus "Sending : " & DownloadFile & " --> " & CurRemDir & DownloadRemFile, 0, 1
        LastCommandSent = "STOR"
        SendData 0, "STOR " & DownloadRemFile
        DoEvents
        ChangeStatus "STOR " & DownloadRemFile, 0, 2
        sckData(0).SendData ReadFile()

    End If

errh:
    If Err.Number = 40006 Then
        Timer2.Enabled = False
        cmdConnect_Click
        DoEvents

    End If
End Sub
'........................................
'Name: ReadFile
'Description:
'........................................
Private Function ReadFile() As String
On Error Resume Next

    Dim sSend As String, sBuf As String
    Dim ifreefile As Integer
    Dim lTotal As Long
    Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long

    ifreefile = FreeFile

    If InStr(1, AsciiList, GetRightWord(DownloadFile, ".", False), vbTextCompare) = 0 Then
        Open CurPath & DownloadFile For Binary Access Read As #ifreefile


        lLen = LOF(ifreefile)

        ' Loop through the file, loading it up in chunks of 64k:
        Do While lRead < lLen
            lThisRead = 65536
            If lThisRead + lRead > lLen Then
                lThisRead = lLen - lRead
            End If
            If Not lThisRead = lLastRead Then
                sBuf = Space$(lThisRead)
            End If
            Get #ifreefile, , sBuf
            lRead = lRead + lThisRead
            sSend = sSend & sBuf
        Loop
        lTotal = lLen
    Else
        Open CurPath & DownloadFile For Input As ifreefile
        While Not EOF(ifreefile)
            Line Input #ifreefile, sBuf
            sSend = sSend & sBuf & vbCrLf
        Wend

    End If
    Close ifreefile
    ReadFile = sSend
End Function


'........................................
'Name: lstRemote_GotFocus
'Object: lstRemote
'Event: GotFocus
'Description:
'........................................
Private Sub lstRemote_GotFocus()
On Error Resume Next
    LoadLocalList CurPath
    LoadLocalElements CurPath
End Sub


'........................................
'Name: lstRemote_ItemClick
'Object: lstRemote
'Event: ItemClick
'Description:
'........................................
Private Sub lstRemote_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If lstRemote.ListItems.Count <> 0 Then
        If RMDown = True Then
            mnuPopTr.Caption = "Dowload"
            mnuPopTras.Caption = "Download as.."
            mnuPopExec.Enabled = False

            If InStr(1, lstRemote.SelectedItem.Text, ".", vbTextCompare) = 0 Then
                mnuPopView.Enabled = False
                mnuPopTras.Enabled = False
            Else
                mnuPopView.Enabled = True
                mnuPopTras.Enabled = True
            End If

            Me.PopupMenu mnuPop
            RMDown = False
        End If


    End If

End Sub


'........................................
'Name: lstRemote_MouseDown
'Object: lstRemote
'Event: MouseDown
'Description:
'........................................
Private Sub lstRemote_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = 2 Then
        Set SelList = lstRemote
        RMDown = True

    End If
End Sub


'........................................
'Name: lstRemote_Validate
'Object: lstRemote
'Event: Validate
'Description:
'........................................
Private Sub lstRemote_Validate(Cancel As Boolean)
On Error Resume Next
    LoadLocalList CurPath
    LoadLocalElements CurPath
End Sub


'........................................
'Name: mnuEditCopy_Click
'Object: mnuEditCopy
'Event: Click
'Description:
'........................................
Private Sub mnuEditCopy_Click()
On Error Resume Next
    SendKeys "^c"
End Sub


'........................................
'Name: mnuEditCut_Click
'Object: mnuEditCut
'Event: Click
'Description:
'........................................
Private Sub mnuEditCut_Click()
On Error Resume Next
    SendKeys "^x"
End Sub


'........................................
'Name: mnuEditGD_Click
'Object: mnuEditGD
'Event: Click
'Description:
'........................................
Private Sub mnuEditGD_Click()
On Error Resume Next
    Dim pattern As String
    Dim x As Integer
    pattern = InputBox("Selection Pattern :", "Group Selection", "*.*")
    If pattern = "" Then Exit Sub
    For x = 1 To SelList.ListItems.Count
        If pattern = "*.*" Then
            SelList.ListItems.Item(x).Selected = True
        Else
            If GetRightWord(SelList.ListItems.Item(x), ".", False) = GetRightWord(pattern, ".", False) Then
                SelList.ListItems.Item(x).Selected = False

            End If
        End If
    Next
    SelList.SetFocus
End Sub


'........................................
'Name: mnuEditGS_Click
'Object: mnuEditGS
'Event: Click
'Description:
'........................................
Private Sub mnuEditGS_Click()
On Error Resume Next
    Dim pattern As String
    Dim x As Integer
    pattern = InputBox("Selection Pattern :", "Group Selection", "*.*")
    If pattern = "" Then Exit Sub
    For x = 1 To SelList.ListItems.Count
        If pattern = "*.*" Then
            SelList.ListItems.Item(x).Selected = True
        Else
            If GetRightWord(SelList.ListItems.Item(x), ".", False) = GetRightWord(pattern, ".", False) Then
                SelList.ListItems.Item(x).Selected = True
            Else
                SelList.ListItems.Item(x).Selected = False
            End If
        End If
    Next
    SelList.SetFocus
End Sub


'........................................
'Name: mnuEditIS_Click
'Object: mnuEditIS
'Event: Click
'Description:
'........................................
Private Sub mnuEditIS_Click()
On Error Resume Next
    Dim x As Integer
    For x = 1 To SelList.ListItems.Count

        SelList.ListItems.Item(x).Selected = Not SelList.ListItems.Item(x).Selected

    Next
End Sub


'........................................
'Name: mnuEditPaste_Click
'Object: mnuEditPaste
'Event: Click
'Description:
'........................................
Private Sub mnuEditPaste_Click()
On Error Resume Next
    SendKeys "^v"
End Sub




'........................................
'Name: mnuEditSA_Click
'Object: mnuEditSA
'Event: Click
'Description:
'........................................
Private Sub mnuEditSA_Click()
On Error Resume Next
    Dim x As Integer
    For x = 1 To SelList.ListItems.Count
        SelList.ListItems.Item(x).Selected = True
    Next
End Sub


'........................................
'Name: mnuEditSettings_Click
'Object: mnuEditSettings
'Event: Click
'Description:
'........................................
Private Sub mnuEditSettings_Click()
On Error Resume Next
    frmSettings.Show vbModal, frmMain
End Sub


'........................................
'Name: mnuFileDC_Click
'Object: mnuFileDC
'Event: Click
'Description:
'........................................
Private Sub mnuFileDC_Click()
On Error Resume Next
        SendData 0, "QUIT"
        DoEvents
        sckConn(0).Close
        ChangeUIState 0
        lstRemote.ListItems.Clear
        cboRemote.ComboItems.Clear
        LastCommandSent = ""
        LastPort = 1026
End Sub


'........................................
'Name: mnuFileEX_Click
'Object: mnuFileEX
'Event: Click
'Description:
'........................................
Private Sub mnuFileEX_Click()
On Error Resume Next
    End
End Sub


'........................................
'Name: mnuFileQC_Click
'Object: mnuFileQC
'Event: Click
'Description:
'........................................
Private Sub mnuFileQC_Click()
On Error Resume Next
    If mnuFileQC.Checked = False Then
        mnuFileQC.Checked = True
        txtStatus.Top = txtStatus.Top + FQuick.Height
        txtStatus.Height = 1695 - FQuick.Height
        FQuick.Visible = True
        FQuick.Top = tbTools.Top + tbTools.Height
    Else
        mnuFileQC.Checked = False
        FQuick.Visible = False
        txtStatus.Top = tbTools.Top + tbTools.Height - 10
        txtStatus.Height = 1695

    End If

End Sub


'........................................
'Name: mnuFileRC_Click
'Object: mnuFileRC
'Event: Click
'Description:
'........................................
Private Sub mnuFileRC_Click()
On Error Resume Next
    If mnuFileRC.Caption = "Connect" Then
        cmdConnect_Click
    Else
        sckConn(0).Close
        DoEvents
        cmdConnect_Click
    End If
End Sub


'........................................
'Name: mnuFileSM_Click
'Object: mnuFileSM
'Event: Click
'Description:
'........................................
Private Sub mnuFileSM_Click()
On Error Resume Next
    frmSiteManager.Show vbModal, frmMain
End Sub


'........................................
'Name: LoadLocalElements
'Description:
'........................................
Private Sub LoadLocalElements(ByVal path As String)
On Error Resume Next
    Dim fsys As New FileSystemObject
    Dim drv As Drive
    Dim fld As Folder
    Dim fil As File
    Dim fils As Files
    Dim img As String

    cboLookInLocal.ComboItems.Clear

    For Each drv In fsys.Drives
        Select Case drv.DriveType
            Case CDRom
                img = "cd"
            Case Removable
                img = "fdd"
            Case Fixed
                img = "hdd"
        End Select
        cboLookInLocal.ComboItems.Add , , drv.DriveLetter & ":", img, img
        cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count).Tag = drv.DriveLetter & ":\"
        If Left$(path, 1) = drv.DriveLetter Then
            Call GetLeftWord(path, "\", True)

            While InStr(1, path, "\") > 0
                cboLookInLocal.ComboItems.Add , , GetLeftWord(path, "\", True), "fld", "fld"
                cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count).Tag = cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 1).Tag & cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count).Text & "\"
                cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count).Indentation = cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 1).Indentation + 1
            Wend
        End If
    Next

    'cboLookInLocal.SelectedItem = cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count)

End Sub


'........................................
'Name: LoadLocalList
'Description:
'........................................
Private Sub LoadLocalList(ByVal path As String)
On Error Resume Next
    'on error Resume Next
    Dim fsys As New FileSystemObject
    Dim fld As Folder
    Dim sfld As Folder
    Dim fil As File
    Dim fils As Files
    lstLocal.ListItems.Clear
    Set fld = fsys.GetFolder(path)
    For Each sfld In fld.SubFolders

        lstLocal.ListItems.Add , , sfld.Name, , "fld"
        lstLocal.ListItems.Item(lstLocal.ListItems.Count).Tag = "fld"
        If sfld.Attributes = Directory Then
            lstLocal.ListItems.Item(lstLocal.ListItems.Count).SubItems(1) = sfld.Size
        Else
            lstLocal.ListItems.Item(lstLocal.ListItems.Count).SubItems(1) = 0
        End If
        lstLocal.ListItems.Item(lstLocal.ListItems.Count).SubItems(2) = sfld.DateLastModified

    Next
    Set fils = fsys.GetFolder(path).Files
    For Each fil In fils
        lstLocal.ListItems.Add , , fil.Name, , "fil"
        lstLocal.ListItems.Item(lstLocal.ListItems.Count).Tag = "fil"
        lstLocal.ListItems.Item(lstLocal.ListItems.Count).SubItems(1) = fil.Size
        lstLocal.ListItems.Item(lstLocal.ListItems.Count).SubItems(2) = fil.DateLastModified
    Next
End Sub



'........................................
'Name: mnuHelpAbout_Click
'Object: mnuHelpAbout
'Event: Click
'Description:
'........................................
Private Sub mnuHelpAbout_Click()
On Error Resume Next
MsgBox "zFTP 1.0 BETA 1 - Thomas Schilb (thomas.schilb@live.de)" & vbCrLf & vbCrLf & "For more info see: http://www.thomasschilb.tk"
End Sub


'........................................
'Name: mnuPopATQ_Click
'Object: mnuPopATQ
'Event: Click
'Description:
'........................................
Private Sub mnuPopATQ_Click()
On Error Resume Next
    mnuPopTr_Click
    Timer2.Enabled = False
End Sub


'........................................
'Name: mnuPopCD_Click
'Object: mnuPopCD
'Event: Click
'Description:
'........................................
Private Sub mnuPopCD_Click()
On Error Resume Next
    Dim dirpath As String

    dirpath = InputBox("Path :", "Change Directory", "Enter Full Path Here")
    If (dirpath = CurPath) Or (dirpath = CurRemDir) Or (dirpath = "") Then Exit Sub

    If SelList.Name = "lstLocal" Then
        CurPath = dirpath
        LoadLocalList (dirpath)
        LoadLocalElements (dirpath)
    Else
        SendData 0, "CWD " & dirpath
        DoEvents
        ChangeStatus "CWD " & dirpath, 0, 2
    End If
End Sub


'........................................
'Name: mnuPopDel_Click
'Object: mnuPopDel
'Event: Click
'Description:
'........................................
Private Sub mnuPopDel_Click()
    On Error Resume Next
    Dim fils As Files
    Dim fil As File
    Dim fsys As New FileSystemObject
    Dim x As Integer


    For x = 1 To SelList.ListItems.Count
        If SelList.ListItems.Item(x).Selected = True Then
            If SelList.Name = "lstLocal" Then


                If InStr(1, SelList.ListItems.Item(x).Text, ".") <> 0 Then
                    Kill CurPath & SelList.ListItems.Item(x).Text
                Else
                    Set fils = fsys.GetFolder(CurPath & SelList.ListItems.Item(x).Text & "\").Files
                    For Each fil In fils
                        Kill CurPath & SelList.ListItems.Item(x).Text & "\" & fil.Name
                    Next
                    RmDir CurPath & SelList.ListItems.Item(x).Text

                End If

            Else
                If InStr(1, SelList.ListItems.Item(x).Text, ".") <> 0 Then

                    ChangeStatus "DELE " & SelList.ListItems.Item(x).Text, 0, 2
                    SendData 0, "DELE " & SelList.ListItems.Item(x).Text
                    DoEvents
                Else
                    ChangeStatus "RMD " & SelList.ListItems.Item(x).Text, 0, 2
                    SendData 0, "RMD " & SelList.ListItems.Item(x).Text
                    DoEvents

                End If
            End If
        End If
    Next
    mnuPopRef_Click


End Sub


'........................................
'Name: mnuPopExec_Click
'Object: mnuPopExec
'Event: Click
'Description:
'........................................
Private Sub mnuPopExec_Click()
On Error Resume Next
    ShellExecute Me.hwnd, "Open", CurPath & SelList.SelectedItem.Text, "", "c:\", 1
End Sub


'........................................
'Name: mnuPopLL_Click
'Object: mnuPopLL
'Event: Click
'Description:
'........................................
Private Sub mnuPopLL_Click()
    On Error Resume Next
    If mnuPopLL.Checked = False Then
        mnuPopLL.Checked = Not mnuPopLL.Checked
        mnuViewLL.Checked = Not mnuViewLL.Checked

        mnuViewSL.Checked = Not mnuViewLL.Checked
        mnuPopSL.Checked = Not mnuPopLL.Checked
        SelList.View = lvwReport
    Else
        mnuPopSL_Click
    End If
End Sub


'........................................
'Name: mnuPopMD_Click
'Object: mnuPopMD
'Event: Click
'Description:
'........................................
Private Sub mnuPopMD_Click()
On Error Resume Next
    Dim dirpath As String

    dirpath = InputBox("Enter Directory name  :", "Change Directory", "Enter Full Path Here")
    If (dirpath = CurPath) Or (dirpath = CurRemDir) Or (dirpath = "") Then Exit Sub

    If SelList.Name = "lstLocal" Then
        FileSystem.MkDir CurPath & dirpath
    Else
        SendData 0, "MKD " & dirpath
        DoEvents
        ChangeStatus "MKD " & dirpath, 0, 2
        GetList 0
    End If
End Sub


'........................................
'Name: mnuPopRef_Click
'Object: mnuPopRef
'Event: Click
'Description:
'........................................
Private Sub mnuPopRef_Click()
On Error Resume Next
    If SelList.Name = "lstLocal" Then
        LoadLocalList (cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 1).Tag)
        LoadLocalElements (cboLookInLocal.ComboItems.Item(cboLookInLocal.ComboItems.Count - 1).Tag)

    Else
        GetList 0
    End If
End Sub


'........................................
'Name: mnuPopRen_Click
'Object: mnuPopRen
'Event: Click
'Description:
'........................................
Private Sub mnuPopRen_Click()
On Error Resume Next
    Dim rnto As String
    Dim fils As Files
    Dim fil As File
    Dim fsys As New FileSystemObject

    rnto = InputBox("Rename to :", "Rename")

    If (rnto = SelList.SelectedItem.Text) Or (rnto = "") Then Exit Sub
    If SelList.Name = "lstLocal" Then
        If InStr(1, SelList.SelectedItem.Text, ".") <> 0 Then
            FileCopy CurPath & SelList.SelectedItem.Text, CurPath & rnto
            Kill CurPath & SelList.SelectedItem.Text
        Else
            MkDir (CurPath & rnto)
            Set fils = fsys.GetFolder(CurPath & SelList.SelectedItem.Text & "\").Files
            For Each fil In fils
                FileCopy CurPath & SelList.SelectedItem.Text & "\" & fil.Name, CurPath & rnto & "\" & fil.Name
                Kill CurPath & SelList.SelectedItem.Text & "\" & fil.Name
            Next
            RmDir CurPath & SelList.SelectedItem.Text

        End If
        mnuPopRef_Click
    Else
        ChangeStatus "RNFR " & lstRemote.SelectedItem.Text, 0, 2
        SendData 0, "RNFR " & lstRemote.SelectedItem.Text
        DoEvents
        ChangeStatus "RNTO " & rnto, 0, 2
        SendData 0, "RNTO " & rnto
        DoEvents

    End If
errh:
    If Err.Number <> 0 Then
        MsgBox "An error occured while performing the operation", vbCritical
        Resume Next
    End If
End Sub


'........................................
'Name: mnuPopSL_Click
'Object: mnuPopSL
'Event: Click
'Description:
'........................................
Private Sub mnuPopSL_Click()
    On Error Resume Next
    If mnuPopSL.Checked = False Then
        mnuPopSL.Checked = Not mnuPopSL.Checked
        mnuViewSL.Checked = Not mnuViewSL.Checked

        mnuViewLL.Checked = Not mnuViewSL.Checked
        mnuPopLL.Checked = Not mnuPopSL.Checked
        SelList.View = lvwList
    Else
        mnuPopLL_Click
    End If

End Sub


'........................................
'Name: mnuPopTr_Click
'Object: mnuPopTr
'Event: Click
'Description:
'........................................
Private Sub mnuPopTr_Click()
On Error Resume Next
    Dim x As Integer

    If mnuPopTr.Caption = "Upload" Then

        For x = 1 To lstLocal.ListItems.Count
            If lstLocal.ListItems.Item(x).Selected = True Then
                lstQueue.ListItems.Add , , CurPath & lstLocal.ListItems.Item(x).Text
                With lstQueue.ListItems.Item(lstQueue.ListItems.Count)
                    .SubItems(1) = lstLocal.ListItems.Item(x).SubItems(1)
                    .SubItems(2) = "->"
                    .SubItems(3) = CurRemDir & lstLocal.ListItems.Item(x).Text
                    .SubItems(4) = ""
                End With
            End If
        Next
    Else
        For x = 1 To lstRemote.ListItems.Count
            If lstRemote.ListItems.Item(x).Selected = True Then
                lstQueue.ListItems.Add , , CurPath & lstRemote.ListItems.Item(x).Text
                With lstQueue.ListItems.Item(lstQueue.ListItems.Count)
                    .SubItems(1) = lstRemote.ListItems.Item(x).SubItems(1)
                    .SubItems(2) = "<-"
                    .SubItems(3) = CurRemDir & lstRemote.ListItems.Item(x).Text
                    .SubItems(4) = ""
                End With
            End If
        Next
    End If
    Timer2.Enabled = True
End Sub


'........................................
'Name: mnuPopTras_Click
'Object: mnuPopTras
'Event: Click
'Description:
'........................................
Private Sub mnuPopTras_Click()
On Error Resume Next
    Dim trfas As String
    trfas = InputBox("File Name :", "Transfer as..", SelList.SelectedItem.Text)
    Timer2.Enabled = False
    mnuPopTr_Click
    DoEvents
    If lstQueue.ListItems.Item(lstQueue.ListItems.Count).SubItems(2) = "<-" Then
        lstQueue.ListItems.Item(lstQueue.ListItems.Count).Text = Replace(lstQueue.ListItems.Item(lstQueue.ListItems.Count).Text, SelList.SelectedItem.Text, trfas)
    Else
        lstQueue.ListItems.Item(lstQueue.ListItems.Count).SubItems(3) = Replace(lstQueue.ListItems.Item(lstQueue.ListItems.Count).SubItems(3), SelList.SelectedItem.Text, trfas)
    End If
    Timer2.Enabled = True
End Sub


'........................................
'Name: mnuPopView_Click
'Object: mnuPopView
'Event: Click
'Description:
'........................................
Private Sub mnuPopView_Click()
On Error Resume Next
    If SelList.Name = "lstLocal" Then

    Else

    End If
End Sub



'........................................
'Name: mnuQueueRm_Click
'Object: mnuQueueRm
'Event: Click
'Description:
'........................................
Private Sub mnuQueueRm_Click()
On Error Resume Next
    lstQueue.ListItems.Remove (lstQueue.SelectedItem.Index)
End Sub


'........................................
'Name: mnuQueueRmall_Click
'Object: mnuQueueRmall
'Event: Click
'Description:
'........................................
Private Sub mnuQueueRmall_Click()
On Error Resume Next
    Dim x As Integer
    For x = 1 To lstQueue.ListItems.Count
        lstQueue.ListItems.Remove (1)
    Next
End Sub


'........................................
'Name: mnuQueueTr_Click
'Object: mnuQueueTr
'Event: Click
'Description:
'........................................
Private Sub mnuQueueTr_Click()
On Error Resume Next
    Timer2.Enabled = True
    Timer2_Timer
End Sub


'........................................
'Name: mnuToolsQClr_Click
'Object: mnuToolsQClr
'Event: Click
'Description:
'........................................
Private Sub mnuToolsQClr_Click()
On Error Resume Next
    Dim x As Integer
    For x = 1 To lstQueue.ListItems.Count
        lstQueue.ListItems.Remove (1)
    Next
End Sub


'........................................
'Name: mnuToolsQTr_Click
'Object: mnuToolsQTr
'Event: Click
'Description:
'........................................
Private Sub mnuToolsQTr_Click()
On Error Resume Next
    Timer2.Enabled = True
End Sub



'........................................
'Name: mnuToolsSCH_Click
'Object: mnuToolsSCH
'Event: Click
'Description:
'........................................


'........................................
'Name: mnuViewFD1_Click
'Object: mnuViewFD1
'Event: Click
'Description:
'........................................
Private Sub mnuViewFD1_Click()
    mnuViewFD1.Checked = Not (mnuViewFD1.Checked)

    If mnuViewFD1.Checked = True Then
        lstLocal.ColumnHeaders.Item(2).Width = (lstLocal.Width - lstLocal.Width / 1.8) / 2
        lstRemote.ColumnHeaders.Item(2).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3
    Else
        lstLocal.ColumnHeaders.Item(2).Width = 0
        lstRemote.ColumnHeaders.Item(2).Width = 0
    End If

End Sub
'........................................
'Name: mnuViewFD2_Click
'Object: mnuViewFD2
'Event: Click
'Description:
'........................................
Private Sub mnuViewFD2_Click()
On Error Resume Next
    mnuViewFD2.Checked = Not (mnuViewFD2.Checked)

    If mnuViewFD2.Checked = True Then
        lstLocal.ColumnHeaders.Item(3).Width = (lstLocal.Width - lstLocal.Width / 1.8) / 2
        lstRemote.ColumnHeaders.Item(3).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3
    Else
        lstLocal.ColumnHeaders.Item(3).Width = 0
        lstRemote.ColumnHeaders.Item(3).Width = 0
    End If

End Sub
'........................................
'Name: mnuViewFD3_Click
'Object: mnuViewFD3
'Event: Click
'Description:
'........................................
Private Sub mnuViewFD3_Click()
On Error Resume Next
    mnuViewFD3.Checked = Not (mnuViewFD3.Checked)

    If mnuViewFD3.Checked = True Then
        lstRemote.ColumnHeaders.Item(4).Width = (lstRemote.Width - lstRemote.Width / 1.8) / 3
    Else
        lstRemote.ColumnHeaders.Item(4).Width = 0
    End If

End Sub


'........................................
'Name: mnuViewLL_Click
'Object: mnuViewLL
'Event: Click
'Description:
'........................................
Private Sub mnuViewLL_Click()
On Error Resume Next
    mnuPopLL_Click
End Sub


'........................................
'Name: mnuViewQCB_Click
'Object: mnuViewQCB
'Event: Click
'Description:
'........................................
Private Sub mnuViewQCB_Click()
On Error Resume Next
    mnuFileQC_Click
    mnuViewQCB.Checked = FQuick.Visible
End Sub


'........................................
'Name: mnuViewRef_Click
'Object: mnuViewRef
'Event: Click
'Description:
'........................................
Private Sub mnuViewRef_Click()
On Error Resume Next
    mnuPopRef_Click
End Sub


'........................................
'Name: mnuViewSL_Click
'Object: mnuViewSL
'Event: Click
'Description:
'........................................
Private Sub mnuViewSL_Click()
On Error Resume Next
    mnuPopSL_Click
End Sub




'........................................
'Name: sckConn_Close
'Object: sckConn
'Event: Close
'Description:
'........................................
Private Sub sckConn_Close(Index As Integer)
On Error Resume Next
    ChangeStatus "Disconnected. " & Time, 0, 1
    cmdConnect_Click
End Sub


'........................................
'Name: sckConn_Connect
'Object: sckConn
'Event: Connect
'Description:
'........................................
Private Sub sckConn_Connect(Index As Integer)
On Error Resume Next
    ChangeStatus "Socket Connected. Waiting for Welcome Message...", 0, 1


End Sub
'........................................
'Name: CreatePortStr
'Description:
'........................................
Private Sub CreatePortStr()
On Error Resume Next
    Dim tpstr As String
    Dim bin1 As Long
    Dim bin2 As Long
    PortStr = ""
    PortStr = Replace(sckConn(0).LocalIP, ".", ",")
    tpstr = LastPort
    tpstr = CStr(Dec2Bin(CInt(tpstr)))
    bin1 = Bin2Dec(Right(tpstr, 8))
    bin2 = Bin2Dec(CInt(Left$(tpstr, Len(tpstr) - 8)))
    PortStr = PortStr & "," & bin2 & "," & bin1

End Sub
'........................................
'Name: sckConn_DataArrival
'Object: sckConn
'Event: DataArrival
'Description:
'........................................
Private Sub sckConn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim sData As String
    Dim x As Integer
    Dim found As Boolean



    sckConn(Index).GetData sData
    sData = GetLeftWord(sData, vbCrLf, False) & vbCrLf
    sData = Replace(sData, "-", " ")
    sData = Trim(sData)
    Code = CInt(GetLeftWord(sData, " ", True))
    ChangeStatus Code & " " & sData, 1, 1

    Select Case Code
        Case 250
            GetList Index
        Case 220
            ChangeStatus "Connected. Authenticating...", 0, 1
            SendData Index, "USER " & txtUsername.Text
            ChangeStatus "USER " & txtUsername.Text, 0, 2
        Case 331
            SendData Index, "PASS " & txtPassword.Text
            ChangeStatus "PASS " & String(Len(txtPassword.Text), "*"), 0, 2
        Case 230
            ChangeStatus "Login Successful", 0, 1
            ChangeStatus "pwd", 0, 2
            SendData Index, "pwd"
            DoEvents

        Case 350
            ChangeStatus "This site can resume broken downloads", 0, 1

        Case 257

            CurRemDir = Replace(GetLeftWord(sData, " ", True), Chr(34), "")
            CurRemDir = Replace(CurRemDir, "/", "\")
            If Right(CurRemDir, 1) <> "\" Then CurRemDir = CurRemDir & "\"
            cboRemote.Text = CurRemDir
            If LastCommandSent = "LIST" Then Exit Sub

            For x = 1 To cboRemote.ComboItems.Count
                If cboRemote.ComboItems.Item(x) = CurRemDir Then found = True
            Next
            If found = False Then cboRemote.ComboItems.Add , , CurRemDir, "fld", "fld"
            Call GetList(Index)

    End Select


End Sub
'........................................
'Name: GetList
'Description:
'........................................
Private Function GetList(Index As Integer)
On Error Resume Next
    lstRemote.ListItems.Clear
    StartDataSck


    ChangeStatus "Retrieving directory listing...", 0, 1

    ChangeStatus "PORT " & PortStr, 0, 2 '& CStr((232 + (-1000 + LastPort))), 0, 2
    SendData Index, "PORT " & PortStr '& CStr((232 + (-1000 + LastPort)))
    DoEvents
    ChangeStatus "LIST", 0, 2
    SendData Index, "LIST"
    DoEvents
    LastCommandSent = "LIST"

End Function
'........................................
'Name: StartDataSck
'Description:
'........................................
Private Sub StartDataSck()
On Error Resume Next
    sckData(0).Close
    LastPort = LastPort + 1
    sckData(0).LocalPort = LastPort
    sckData(0).Bind LastPort, sckConn(0).LocalIP
    sckData(0).Listen
    CreatePortStr

End Sub
'........................................
'Name: sckConn_Error
'Object: sckConn
'Event: Error
'Description:
'........................................
Private Sub sckConn_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
    If Number = 10061 Then
        ChangeStatus "Cannot connect to server.", 0, 1
    End If
End Sub




'........................................
'Name: sckData_ConnectionRequest
'Object: sckData
'Event: ConnectionRequest
'Description:
'........................................
Private Sub sckData_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
    sckData(Index).Close
    sckData(Index).Accept requestID

End Sub


'........................................
'Name: sckData_DataArrival
'Object: sckData
'Event: DataArrival
'Description:
'........................................
Private Sub sckData_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    ReceivedBytes = ReceivedBytes + bytesTotal
    If lstQueue.ListItems.Count > 0 Then
        lstQueue.ListItems.Item(1).SubItems(4) = CInt(ReceivedBytes / FileSize * 100) & "% done"
    End If
    Dim sData As String
    sckData(Index).GetData sData


    Select Case LastCommandSent
        Case "LIST"
            LoadList (sData)
            ReceivedBytes = 0

        Case "RETR"
            Timer2.Enabled = False
            
            If Contents = "" Then
            Kill CurPath & DownloadFile
            End If
            
            Contents = Contents & sData
            SaveFile Contents
            sData = ""




    End Select
End Sub
'........................................
'Name: LoadList
'Description:
'........................................
Private Sub LoadList(ByVal list As String)
On Error Resume Next
    Dim lstarr() As String
    Dim itmarr() As String
    Dim x As Integer
    Dim y As Integer

    Dim FName As String
    Dim fpermission As String
    Dim fdate As String
    Dim FSize As String


    Dim pos As Integer

    lstarr = Split(list, vbCrLf)
    'Debug.Print list

    For x = 0 To UBound(lstarr) - 1
        fpermission = GetLeftWord(lstarr(x), " ", False)
        lstarr(x) = Right(lstarr(x), Len(lstarr(x)) - Len(fpermission))
        lstarr(x) = RemoveSpaces(CStr(lstarr(x)))

        FSize = GetLeftWord(lstarr(x), " ", False)
        lstarr(x) = Right(lstarr(x), Len(lstarr(x)) - Len(FSize))
        lstarr(x) = RemoveSpaces(CStr(lstarr(x)))

        FSize = GetLeftWord(lstarr(x), " ", False)
        lstarr(x) = Right(lstarr(x), Len(lstarr(x)) - Len(FSize))
        lstarr(x) = RemoveSpaces(CStr(lstarr(x)))

        FSize = GetLeftWord(lstarr(x), " ", False)
        lstarr(x) = Right(lstarr(x), Len(lstarr(x)) - Len(FSize))
        lstarr(x) = RemoveSpaces(CStr(lstarr(x)))

        FSize = GetLeftWord(lstarr(x), " ", False)
        lstarr(x) = Right(lstarr(x), Len(lstarr(x)) - Len(FSize))
        lstarr(x) = RemoveSpaces(CStr(lstarr(x)))

        If InStr(1, lstarr(x), "->") = 0 Then
            FName = GetRightWord(CStr(lstarr(x)), " ", False)
        Else
            lstarr(x) = Replace(lstarr(x), " -> ", "->")
            FName = GetRightWord(CStr(lstarr(x)), " ", False)
            FName = GetLeftWord(CStr(lstarr(x)), "->", False)
            FName = GetRightWord(FName, " ", True)
        End If

        fdate = Left$(lstarr(x), Len(lstarr(x)) - Len(FName))

        If (FName <> ".") And (FName <> "..") Then
            If InStr(1, FName, ".") = 0 Then
                lstRemote.ListItems.Add , , Trim(FName), "fld", "fld"
            Else
                lstRemote.ListItems.Add , , Trim(FName), "fil", "fil"
            End If

            lstRemote.ListItems(lstRemote.ListItems.Count).SubItems(1) = Trim(FSize)
            lstRemote.ListItems(lstRemote.ListItems.Count).SubItems(2) = Trim(fdate)
            lstRemote.ListItems(lstRemote.ListItems.Count).SubItems(3) = Trim(fpermission)
        End If
    Next
    'If list <> "" Then StartDataSck
End Sub
'........................................
'Name: RemoveSpaces
'Description:
'........................................
Private Function RemoveSpaces(ByRef str As String) As String
On Error Resume Next
    While Left$(str, 1) = " "
        str = Right(str, Len(str) - 1)
    Wend
    RemoveSpaces = str
End Function


'........................................
'Name: sckData_SendComplete
'Object: sckData
'Event: SendComplete
'Description:
'........................................
Private Sub sckData_SendComplete(Index As Integer)
On Error Resume Next
    lstQueue.ListItems.Remove (1)
    Timer2.Enabled = True
    StartDataSck

End Sub




'........................................
'Name: sckData_SendProgress
'Object: sckData
'Event: SendProgress
'Description:
'........................................
Private Sub sckData_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
    lstQueue.ListItems.Item(1).SubItems(4) = CStr(Int(100 - ((bytesRemaining / CLng(lstQueue.ListItems.Item(1).SubItems(1))) * 100))) & "% done"
End Sub


Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
cmdConnect_Click
Case 2
cmdConnect_Click
Case 4
        Set SelList = lstLocal
        mnuPopTr.Caption = "Upload"
        mnuPopTr_Click

Case 5
        Set SelList = lstRemote
        mnuPopTr.Caption = "Download"
        mnuPopTr_Click

Case 6
SendData 0, "Abor"
Case 8
ShellExecute Me.hwnd, "Open", CurPath & SelList.SelectedItem.Text, "", "c:\", 1
Case 10
mnuFileSM_Click

End Select
End Sub

'........................................
'Name: Timer2_Timer
'Object: Timer2
'Event: Timer
'Description:
'........................................
Private Sub Timer2_Timer()

On Error Resume Next
    If lstQueue.ListItems.Count > 0 Then
        StartDataSck
        DownloadDir = lstQueue.ListItems.Item(1)
        DownloadFile = GetRightWord(DownloadDir, "\", True)
        DownloadRemDir = lstQueue.ListItems.Item(1).SubItems(3)
        DownloadRemFile = GetRightWord(DownloadRemDir, "\", True)
        FileSize = CLng(lstQueue.ListItems.Item(1).SubItems(1))

        Timer2.Enabled = False
        If (InStr(1, DownloadFile, ".") = 0) Or (InStr(1, DownloadRemFile, ".") = 0) Then
            If lstQueue.ListItems.Item(1).SubItems(2) = "<-" Then
                TransferDir 0
            Else
                TransferDir 1
            End If
        Else

            If lstQueue.ListItems.Item(1).SubItems(2) = "<-" Then
                Transfer 0
            Else
                Transfer 1
            End If

        End If
        lstQueue.ListItems.Item(1).SubItems(4) = "Transfering"
    End If
End Sub


'........................................
'Name: txtStatus_Change
'Object: txtStatus
'Event: Change
'Description:
'........................................
Private Sub txtStatus_Change()
On Error Resume Next
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

