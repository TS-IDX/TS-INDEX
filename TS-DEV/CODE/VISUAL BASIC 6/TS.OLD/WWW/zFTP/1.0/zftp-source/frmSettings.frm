VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvwOptions 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9763
      _Version        =   393217
      Style           =   4
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmSettings"
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
'     Module Name: frmSettings
'     Module File: frmSettings.frm
'     Module Type: Form
'     Module Description:
'    ..........................................................
'
'    ..........................................................
'   ************************************************************

Option Explicit


'........................................
'Name: Form_Load
'Object: Form
'Event: Load
'Description:
'........................................
Private Sub Form_Load()
    With tvwOptions
        .Nodes.Add , , "conn", "Connection"
        .Nodes.Item(.Nodes.Count).Expanded = True
        .Nodes.Add "conn", tvwChild, , "Login Options"
        .Nodes.Add "conn", tvwChild, , "Keep Alive"
        .Nodes.Add "conn", tvwChild, , "Proxy/Firewall"
        .Nodes.Add "conn", tvwChild, , "Autodial Options"
        .Nodes.Add "conn", tvwChild, , "Disconnect Options"
        .Nodes.Add "conn", tvwChild, , "Timeout/Retry"

        .Nodes.Add , , "ui", "User Interface"
        .Nodes.Item(.Nodes.Count).Expanded = True
        .Nodes.Add "ui", tvwChild, , "Mouse/Keyboard"
        .Nodes.Add "ui", tvwChild, , "Popup Messages"
        .Nodes.Add "ui", tvwChild, , "Other"

        .Nodes.Add , , "trf", "Transfers"
        .Nodes.Item(.Nodes.Count).Expanded = True
        .Nodes.Add "trf", tvwChild, , "Transfer Mode"
        .Nodes.Add "trf", tvwChild, , "Speed Limits"
        .Nodes.Add "trf", tvwChild, , "Filename Conversion"
        .Nodes.Add "trf", tvwChild, , "Other"

    End With
End Sub

