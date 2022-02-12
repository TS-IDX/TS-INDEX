VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "vbIRCd - Open Source IRCd built in VB6                    [Update2]"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Win 
      Index           =   0
      Left            =   1920
      Top             =   3960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   0   'False
      Timeout         =   12
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Win2 
      Index           =   0
      Left            =   1920
      Top             =   4440
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   0   'False
      Timeout         =   12
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox P_1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   5055
      TabIndex        =   11
      Top             =   480
      Width           =   5055
      Begin VB.PictureBox Picture2 
         Height          =   310
         Left            =   3000
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   32
         Top             =   1680
         Width           =   975
         Begin VB.CommandButton C_SSC 
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Server Status:"
         Height          =   1575
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   5055
         Begin VB.Label lbl_CGU 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Current:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4320
            TabIndex        =   28
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl_CS 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   27
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Servers linked:"
            Height          =   255
            Index           =   43
            Left            =   3120
            TabIndex        =   26
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lbl_CC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   25
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lbl_CO 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   24
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Channels:"
            Height          =   255
            Index           =   42
            Left            =   3480
            TabIndex        =   23
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Operators:"
            Height          =   255
            Index           =   41
            Left            =   3480
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl_HGU 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lbl_UT 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 Days, 0:00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   20
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Server up:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   44
            Left            =   315
            TabIndex        =   19
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Users:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   46
            Left            =   580
            TabIndex        =   18
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lbl_HU 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lbl_CU 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Current:      Highest:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1080
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Global users:"
            Height          =   255
            Index           =   45
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.TextBox txt_Date 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   915
         TabIndex        =   12
         Top             =   1680
         Width           =   975
         Begin VB.CommandButton C_Hide 
            Caption         =   "Tray It"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.CheckBox Ch_UT 
         Height          =   195
         Left            =   2280
         TabIndex        =   34
         Top             =   1725
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "Unix Time:"
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.PictureBox P_2 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   5055
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton C_Log 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4780
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Clear Log Display"
         Top             =   1660
         Width           =   255
      End
      Begin VB.TextBox txt_Log 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox P_3 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   4995
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.PictureBox picScroll 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1875
         Left            =   0
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label lbl_About 
         Caption         =   $"frmMain.frx":014A
         Height          =   1695
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   4815
      End
   End
   Begin VB.TextBox txt_Credits 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Text            =   "frmMain.frx":03AA
      ToolTipText     =   "DO NOT REMOVE THIS TEXT BOX!"
      Top             =   360
      Width           =   5295
   End
   Begin VB.TextBox txt_Buffer 
      Height          =   645
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt_License 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmMain.frx":061D
      ToolTipText     =   $"frmMain.frx":372F
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List_Ports 
      Height          =   645
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List_SOP 
      Height          =   255
      ItemData        =   "frmMain.frx":37C4
      Left            =   240
      List            =   "frmMain.frx":37C6
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List_SOM 
      Height          =   255
      ItemData        =   "frmMain.frx":37C8
      Left            =   240
      List            =   "frmMain.frx":37CA
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List_SOA 
      Height          =   255
      ItemData        =   "frmMain.frx":37CC
      Left            =   240
      List            =   "frmMain.frx":37CE
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer T_UT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   3480
   End
   Begin VB.PictureBox pichook 
      Height          =   315
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List_SO 
      Height          =   1230
      ItemData        =   "frmMain.frx":37D0
      Left            =   120
      List            =   "frmMain.frx":37D2
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt_CMOTD 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl_ToolBar 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   40
      ToolTipText     =   "vbIRCd Development Credits"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl_ToolBar 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   38
      ToolTipText     =   "About vbIRCd"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl_ToolBar 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Log Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "View Logged Server Information"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl_ToolBar 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Server Status Information and Controls"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image iIcon 
      Height          =   240
      Left            =   2040
      Picture         =   "frmMain.frx":37D4
      Top             =   3120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu_Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Menu_n1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Controls 
         Caption         =   "Controls"
         Begin VB.Menu Menu_Rehash 
            Caption         =   "Rehash"
         End
         Begin VB.Menu Menu_n3 
            Caption         =   "-"
         End
         Begin VB.Menu Menu_Start 
            Caption         =   "Start"
         End
         Begin VB.Menu Menu_Restart 
            Caption         =   "Restart"
            Enabled         =   0   'False
         End
         Begin VB.Menu Menu_Stop 
            Caption         =   "Stop"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu Menu_n2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' vbIRCd - Software/Code is an IRCd(Internet Relay Chat Daemon) used to host IRC sessions
' Copyright (C) 2001  Nathan Martin
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' To Contact the author e-mail TRON at tron@ircd-net.org
' * Note: There is no post mail contact information due to that it can be abused...

Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim ScrollText As String
Dim EndingFlag As Boolean


Public lbl_UC As Integer
Public lbl_IU As Integer
Dim sDownTime As Integer
Dim sReboot As Boolean
Dim sTickCount As Integer


Private Sub C_Hide_Click()
    Me.Hide
End Sub

Private Sub C_Log_Click()
    txt_Log = ""
    DisplayLog "Display Log was cleared..."
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim X As Integer
    CRLF = Chr$(13) & Chr$(10)
    CR = Chr$(13)
    LF = Chr(10)
    
    ' Set the Files we're going be working with   -TRON
    MOTDFile = App.Path & "\ircd.motd"
    SysFile = App.Path & "\ircd.conf"
    
    ' Some settings have been commented out since they are no longer
    ' needed cause of the ircd.conf file supports those settings
    
    iNetName = "vbIRCd-Net" ' Network Name
    'iServDSC = "vbIRCd Open Source Server" ' Server's Description
    'iServAdminLine1 = "The Server Master of " & iServDSC & ""
    'iServAdminLine2 = "John Joe" ' Server Admin
    'iServAdminLine3 = "JJoe@Mynet.org" ' Server Admin's E-mail
    'iRPass = "ChangeMe" ' Restart Pass
    'iDPass = "ChangeME" ' Die Pass
    iSupportEmail = "support@Mynet.org" ' E-mail of Support Team
    iHiddenPrefix = "mpx" ' Host Hidden Prefix
    iFloodCMDs = "160" ' Max Commands Per Minute
    iFloodMSGs = "120" ' Max Messages/Notices Per Minute
    iConnPass = "" ' Server password
    iServPing = "300" '[25 to 300] How long until next ping in secs
    iMainChan = "#" & iNetName ' Main Channel of the Server
    iHelpChan = "#Support" ' The Help Support Channel
    iChanMax = "20" '[1 o 25] Max Channels user can be on
    iConnMax = "2000" '[0 to 5000] Max of connections allowed
    iSMOCC = "tn" ' Set Modes on Channel Creation
    iForceCloak = 1 '[0|1] Force User umode +x (xxx.xxx.xxx.mpx-#### or mpx-####.host.isp.net) kinda hostmask
    iGODNS = 1 '[0|1] Use Hostmask from what client sends(Not recommanded to enable)
    iWhoCCC = 0 '[0|1|2|3|4] Who Can Create Channels
    ' 0 = Anyone
    ' 1 = Local Ops & Up Only
    ' 2 = Global Ops & Up Only
    ' 3 = Services Ops & Up Only
    ' 4 = Server Admins Only
    sACC = 0  '[0|1|2] Who can use Server
    ' 0 = Anyone
    ' 1 = Need Password to use Server
    ' 2 = Must be an IRC Op to use server.
    '     Set UserID/Ident as Your Oper ID and
    '     use normal pass field for oper pass to
    '     get onto the server.  -TRON
    
    ' Set Version
    sVersion = "v" & App.Major & "." & App.Minor & "." & App.Revision
    sRelease = "Beta" ' Yep, we're in beta since this is not a complet project.
    lbl_About = "                                vbIRCd v" & App.Major & "." & App.Minor & "." & App.Revision & " By TRON                                                             Copyright© 2001 Nathan Martin                                                                                                                                        vbIRCd: is an UnrealIRCd Style IRC Server developed using M$'s Visual Basic 6(SP4).  The Project was started on 03/22/2001 and it has been developed since then to make it campatible with the changes and growth of the Unreal IRCd project which it's founder is Stskeeps.  The vbIRCd project is also being developed by eternal."
    
    
    ' Set the buttons as C++ Style
    CButton C_SSC
    CButton C_Hide
    CButton C_Log
    
    ' Set when the server was loaded up, it's suppose to be when
    ' the server was compiled, but ohh well, this is VB, not C/C++ ;)   -TRON
    sUDT = Format(Now, "Long Date") & " at " & Format(Now, "Long Time")
    
    For X = 1 To iUserMax
        iPeerFree(X) = True
    Next X
    
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = pichook.hwnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = iIcon.Picture
    TrayI.szTip = "vbIRCd " & sVersion & " [" & sRelease & "] - Open Source" & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
    
    LoadConf  '<-- Now ready  :D
    LoadMOTD
    
    If Left$(LCase(Command$), 6) = "-tray" Then Me.Hide ' If vbIRCd loaded like "vbIRCd.exe -Tray" then Hide GUI into System Tray
End Sub

Public Sub C_SSC_Click()
On Error Resume Next
Dim X As Integer

    If C_SSC.Caption = "Start" Then
        sServer = iServer ' Set Server Name
        For X = 0 To List_Ports.ListCount - 1
            Win2(X).LocalPort = List_Ports.List(X)
            Win2(X).Listen
            If Not X = List_Ports.ListCount - 1 Then Load Win2(X + 1)
        Next X
        C_SSC.Caption = "Stop"
        T_UT.Enabled = True
        Menu_Start.Enabled = False
        Menu_Stop.Enabled = True
        Menu_Restart.Enabled = True
        DisplayLog "vbIRCd up and running..."
    Else
        For X = 0 To List_Ports.ListCount - 1
            Win2(X).Disconnect
            If Not X = 0 Then Unload Win2(X)
        Next X
        T_UT.Enabled = False
            For X = 1 To 2
                iSec(X) = 0
                iMin(X) = 0
                iHour = 0
                iDays = 0
            Next X
            
        For X = 1 To iUserMax
            If iPeerFree(X) = False Then
                UserClosed X
            End If
        Next X
        
        C_SSC.Caption = "Start"
        cRemoveChannel "ALL"
        lbl_CO = 0
        Menu_Start.Enabled = True
        Menu_Stop.Enabled = False
        Menu_Restart.Enabled = False
        
        If sReboot = True Then
            DisplayLog "vbIRCd restarting..."
            sReboot = False
            C_SSC_Click
        Else
            DisplayLog "vbIRCd has been shutdown"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim MsgRes As Long
    'Make sure that the user really want to shutdown
    'MsgRes = MsgBox("Do you really want to Shut Down the Server?", vbYesNo Or vbQuestion, "Are you Sure!")
    'If the user selects no, exit this sub
    'If MsgRes = vbYes Then
        TrayI.cbSize = Len(TrayI)
        TrayI.hwnd = pichook.hwnd
        TrayI.uId = 1&
        'Delete the icon
        Shell_NotifyIcon NIM_DELETE, TrayI
        Set frmMain = Nothing
        End
    'Else
    '    Cancel = 1
    'End If
End Sub

Private Sub lbl_ToolBar_Click(Index As Integer)
    lbl_ToolBar(0).BackColor = &HE0E0E0
    lbl_ToolBar(1).BackColor = &HE0E0E0
    lbl_ToolBar(2).BackColor = &HE0E0E0
    lbl_ToolBar(3).BackColor = &HE0E0E0
    P_1.Visible = False
    P_2.Visible = False
    P_3.Visible = False
    lbl_About.Visible = False
    picScroll.Visible = False
    EndingFlag = True
    
    Select Case Index
        Case 0
            lbl_ToolBar(Index).BackColor = &HFF8080
            P_1.Visible = True
            
        Case 1
            lbl_ToolBar(Index).BackColor = &HFF8080
            P_2.Visible = True
            txt_Log.SelStart = Len(txt_Log)
            
        Case 2
            lbl_ToolBar(Index).BackColor = &HFF8080
            P_3.Visible = True
            lbl_About.Visible = True
            
        Case 3
            lbl_ToolBar(Index).BackColor = &HFF8080
            P_3.Visible = True
            ScrollText = txt_Credits
            EndingFlag = False
            picScroll.Visible = True
            RunMain
            
    End Select
End Sub

Private Sub Menu_Exit_Click()
    Unload Me
End Sub

Private Sub Menu_Rehash_Click()
    DisplayLog "vbIRCd Rehashing ircd.conf file..."
    LoadConf
End Sub

Private Sub Menu_Restart_Click()
    Menu_Restart.Enabled = False
    cSysRD True, "Requested by Admin @ Console", 0
End Sub

Private Sub Menu_Show_Click()
    Me.Visible = True
End Sub

Private Sub Menu_Start_Click()
    C_SSC_Click
End Sub

Private Sub Menu_Stop_Click()
    C_SSC_Click
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then  'If the user dubbel-clicked on the icon
        If frmMain.Visible = False Then
            frmMain.Visible = True
            Me.SetFocus
        Else
            frmMain.Visible = False
        End If
    ElseIf Msg = WM_RBUTTONUP Then  'Right click
        PopupMenu Menu
    End If
End Sub

Private Sub Win_Disconnect(Index As Integer)
On Error Resume Next
Dim TmpText As String
Dim TmpChan As String
Dim Y As Integer
Dim Q As Integer
Dim X As Integer
    Win(Index).Disconnect
    iPeerFree(Index) = True
    
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                X = InStr(1, iModes(Q), "c")
                If Not X = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(Index) & " (" & iName(Index) & "@" & iRHost(Index) & ") has Quit IRC [Connection reset by peer]" & CRLF
            End If
        Next Q
    
        'For Y = 1 To iUserMax
        '    If iPeerFree(Y) = False Then
        '        SendData Y, ":" & iUser(Index) & "!" & iName(Index) & "@" & irHost(Index) & " QUIT :Connection Reset by Peer" & CRLF
        '    End If
        'Next Y
        
    UserClosed Index, "Connection reset by peer"
End Sub

Private Sub Win2_Accept(Index As Integer, SocketID As Integer)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
With frmMain
    If lbl_CU + 1 > iConnMax Then
        For X = 1 To iUserMax
            If iPeerFree(X) = True Then
                iPeerFree(X) = False
                Load Win(X)
                Win(X).LocalPort = Win(0).LocalPort
                Win(X).Accept = SocketID
                SendData X, ":" & sServer & " 465 AUTH :Too Many users on already" & CRLF
                iKILL(X) = True
                .lbl_UC = .lbl_UC + 1
                sCCount = sCCount + 1
                sCRefused = sCRefused + 1
                KillUser X, sServer, "Too Many users on already", , True, True
                Exit For
            End If
        Next X
        Exit Sub
    End If
    
    For X = 1 To iUserMax
        If iPeerFree(X) = True Then
            iPeerFree(X) = False
            Load Win(X)
            Win(X).LocalPort = Win(0).LocalPort
            Win(X).Accept = SocketID
            iPing(X) = 0
            iUserLevel(X) = 0
            iAAC(X) = False
            iName(X) = ""
            iHost(X) = ""
            iRHost(X) = ""
            iRealName(X) = ""
            iUser(X) = ""
            iIdle(X) = 0
            iAway(X) = ""
            .lbl_UC = .lbl_UC + 1
            iHolted(X) = True: IdentScan (X)
            sCCount = sCCount + 1
            sCAccepted = sCAccepted + 1
            Exit For
        End If
    Next X
End With
End Sub

Private Sub Win_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error Resume Next
Dim Data As String
Dim TmpData As String
Dim Z, X, Q As Integer
If iKILL(Index) = True Then Exit Sub
    Win(Index).Read Data, DataLength
GoTo ReformatData
ReLoad:
    If TmpData = "" Then If iName(Index) = "" Then UserClosed Index: Exit Sub
    If TmpData = " " Then If iName(Index) = "" Then UserClosed Index: Exit Sub
    If Not Len(TmpData) < 2 Then
        If iHolted(Index) = True Then
            iHoldData(Index) = iHoldData(Index) & TmpData
        Else
            SYS TmpData, Index
        End If
    End If
    GoTo ReformatData
Exit Sub
ReformatData:
If Data = "" Then Exit Sub
    X = InStr(1, Data, CRLF)
    If Not X = 0 Then
        TmpData = Mid$(Data, 1, X - 1)
        Data = Mid$(Data, X + 2)
    Else
        X = InStr(1, Data, LF)
        If Not X = 0 Then
            TmpData = Mid$(Data, 1, X - 1)
            Data = Mid$(Data, X + 1)
        Else
            X = InStr(1, Data, CR)
            If Not X = 0 Then
                TmpData = Mid$(Data, 1, X - 1)
                Data = Mid$(Data, X + 1)
            Else
                TmpData = Data
                Data = ""
            End If
        End If
    End If
GoTo ReLoad
End Sub

Private Sub Win_Write(Index As Integer)
On Error Resume Next
        If iKILL(Index) = True Then
            If Not Win(Index).State = 6 Then
                iKILL(Index) = False
                UserClosed Index, "0"
            End If
        End If
End Sub

Public Sub cSysRD(Reboot As Boolean, Text As String, Index As Integer)
On Error Resume Next
Dim Q As Integer
Dim TmpText As String
If Text = "" Then Text = "No Reason Given"
    If Not Index = 0 Then TmpText = iUser(Index) Else: TmpText = sServer
    If Reboot = True Then
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :IRC Operator " & TmpText & " has used the /RESTART command to reboot the server." & CRLF & _
                             ":" & sServer & " NOTICE " & iUser(Q) & " :Reason why from " & TmpText & ": " & Text & CRLF & _
                             ":" & sServer & " NOTICE " & iUser(Q) & " :OK ALL HERE WE GO!  WEEEEEEEEEEEEEEEEEEEEE..." & CRLF
            End If
        Next Q
        
        sReboot = True
    Else
        For Q = 1 To iUserMax
            If iPeerFree(Q) = False Then
                SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :IRC Operator " & TmpText & " has used the /DIE command to shut down the server." & CRLF & _
                             ":" & sServer & " NOTICE " & iUser(Q) & " :Reason why from " & TmpText & ": " & Text & CRLF & _
                             ":" & sServer & " NOTICE " & iUser(Q) & " :OK ALL WE'RE GOING DOWN!  AHHHHHHHHHHHHHHHHHHH!!!" & CRLF
            End If
        Next Q
    End If
    sDownTime = 1
End Sub

Private Sub Win2_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
    If ErrorCode = 24048 Then MsgBox "Port " & Win2(Index).LocalPort & " is already in use, please remove it from Port list " & vbCrLf & "under General Section of Server Configuration Dialog." & vbCrLf & vbCrLf & "Make sure to save settings and Restart IRC Serv Software.", vbCritical, "Bind ERROR"
    LogIt "ERROR! -(Port:" & Win2(Index).LocalPort & "-" & Index & ") " & ErrorCode & " :" & ErrorString
End Sub

Private Sub Win2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error Resume Next
Dim Data As String
    Win2(Index).Read Data, DataLength
    
End Sub

Private Sub T_UT_Timer()
On Error Resume Next
Dim X As Integer
Dim Q As Integer
Dim Y As Integer
Dim GUI As Boolean
Dim DUI As Boolean
    DUI = True
    GUI = False

    iSec(1) = iSec(1) + 1
    If iSec(1) = "10" Then
        iSec(2) = iSec(2) + 1
        iSec(1) = "0"
    End If
    
    If iSec(2) = "6" Then
        iMin(1) = iMin(1) + 1
        iSec(2) = "0"
        GUI = True
    End If
    
    If iMin(1) = "10" Then
        iMin(2) = iMin(2) + 1
        iMin(1) = "0"
    End If
    
    If iMin(2) = "6" Then
        iHour = iHour + 1
        iMin(2) = "0"
    End If
    
    If iHour = "25" Then
        iDays = iDays + 1
        iHour = "0"
    End If
    lbl_UT.Caption = iDays & " Days, " & iHour & ":" & iMin(2) & iMin(1) & ":" & iSec(2) & iSec(1)
    If Not sDownTime = 0 Then sDownTime = sDownTime + 1
    If sDownTime = 4 Then
        C_SSC_Click
        sDownTime = 0
    End If
    
    For X = 1 To iUserMax
        If iPeerFree(X) = False Then
            iIdle(X) = iIdle(X) + 1
            iPing(X) = iPing(X) + 1
            
            If iIdle(X) = 30 Then
                If iUser(X) = "" Then DUI = False
                If iName(X) = "" Then DUI = False
                If iRHost(X) = "" Then DUI = False
                If DUI = False Then KillUser X, sServer, "Connection never sent USER or SERVER identification", , True
            End If
            
            Y = InStr(1, iModes(X), "D")
            If GUI = True Then If Not iFloodMSGs = 0 And iFM(X) = iFloodMSGs Then SendData X, ":" & sServer & " NOTICE " & iUser(X) & " :*** GOOD NEWS: You are now free to send messages and notices again." & CRLF
            If GUI = True Then If Not iFloodCMDs = 0 And iFC(X) = iFloodCMDs Then SendData X, ":" & sServer & " NOTICE " & iUser(X) & " :*** GOOD NEWS: You are now free to send/use commands again." & CRLF
            If GUI = True Then iFM(X) = 0:   iFC(X) = 0
            If Not Y = 0 And GUI = True Then SendData X, ":" & sServer & " NOTICE " & iUser(X) & " :Flood Counters have been reset to 0" & CRLF
            If iPing(X) = iServPing Then SendData X, "PING :" & sServer & CRLF
    
            If iPing(X) = iServPing + 20 Then
                For Q = 1 To iUserMax
                    If iPeerFree(Q) = False Then
                        Y = InStr(1, iModes(Q), "c")
                        If Not Y = 0 Then SendData2 Q, ":" & sServer & " NOTICE " & iUser(Q) & " :*** NOTICE -- " & iUser(X) & " (" & iName(X) & "@" & iRHost(X) & ") has Quit IRC (Reason: Ping timeout)" & CRLF
                    End If
                Next Q
                Win(X).Disconnect
                uQUIT iUser(X) & "!" & iName(X) & "@" & iHost(X), iChan(X), "Ping timeout"
                UserClosed X
            End If
        End If
    Next X
    
    If Ch_UT = 1 Then txt_Date = GetTime
End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmMain.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
    End If
    
    DoEvents
Loop

End Sub
