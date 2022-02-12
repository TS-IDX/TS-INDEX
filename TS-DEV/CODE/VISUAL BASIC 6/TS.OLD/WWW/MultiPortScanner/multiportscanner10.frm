VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMultiportscanner 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Port Scanner 1.0"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   Icon            =   "multiportscanner10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4080
      Width           =   5295
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "IP Range"
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
      Height          =   2535
      Left            =   5640
      TabIndex        =   33
      Top             =   3360
      Width           =   5775
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
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
         NumItems        =   0
      End
   End
   Begin VB.Timer tmrscan 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton cmdstart 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Start Scan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Statistics"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   5295
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Progress:"
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
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbltime 
         BackColor       =   &H00000000&
         Caption         =   "00:00:00"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Rem. Time:"
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
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblspeed 
         BackColor       =   &H00000000&
         Caption         =   "0 ip/sec"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "Speed:"
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
         TabIndex        =   26
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbldone 
         BackColor       =   &H00000000&
         Caption         =   "0 of 0"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "IP's Done:"
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
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H00000000&
         Caption         =   "idle..."
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
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Status:"
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
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1800
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Speed Options"
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
      Height          =   1095
      Left            =   5640
      TabIndex        =   14
      Top             =   2160
      Width           =   5775
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   200
         TickStyle       =   3
         Value           =   200
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   100
         Min             =   1
         Max             =   10000
         SelStart        =   3000
         TickStyle       =   3
         Value           =   3000
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3000"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Timeout:"
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
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "200"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Threads:"
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
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Port List"
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
      Height          =   1935
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   5775
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   2040
         TabIndex        =   13
         Text            =   "500"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   840
         TabIndex        =   11
         Text            =   "1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Range to list"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   3000
         TabIndex        =   8
         Text            =   "80"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Add Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.ListBox lstports 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Height          =   990
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "to:"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "from:"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "IP Range"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Timer tmrtime 
         Interval        =   1000
         Left            =   3960
         Top             =   0
      End
      Begin VB.Timer tmrstat 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4440
         Top             =   0
      End
      Begin VB.TextBox txtstop 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   960
         TabIndex        =   4
         Text            =   "84.144.255.255"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtstart 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Left            =   960
         TabIndex        =   2
         Text            =   "84.144.1.1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Stop:"
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
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Start:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "© 2017-2021 // thomas_schilb@outlook.com"
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
      TabIndex        =   37
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   5400
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
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
      Left            =   4200
      TabIndex        =   36
      Top             =   5160
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   5400
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Multi Port Scanner"
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
      Height          =   540
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Width           =   3930
   End
End
Attribute VB_Name = "frmMultiportscanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lstcounter As Long
Public maxxx As Long
Public tick As Long
Public ipdone As Long
'sry for the bad style and the bugs

Public Function AddOne(IP As String) As String 'add one to ip address
On Error Resume Next
    ipdone = ipdone + 1
    Bar1.Value = Bar1.Value + 1

    Dim A As Long, B As Long, C As Long, D As Long
    Dim A1 As Long, B1 As Long, C1 As Long
    A1 = InStr(1, IP, ".")
    A = Mid(IP, 1, A1)
    IP = Mid(IP, A1 + 1, Len(IP) - A1)
    B1 = InStr(1, IP, ".")
    B = Mid(IP, 1, B1)
    IP = Mid(IP, B1 + 1, Len(IP) - B1)
    C1 = InStr(1, IP, ".")
    C = Mid(IP, 1, C1)
    IP = Mid(IP, C1 + 1, Len(IP) - C1)
    D = IP


    If D >= 255 Then

        If C >= 255 Then

            If B >= 255 Then

                If A >= 255 Then
                
                Else
                    A = A + 1
                End If

                B = 0
            Else
                B = B + 1
            End If

            C = 0
        Else
            C = C + 1
        End If

        D = 0
    Else
        D = D + 1
    End If

    AddOne = A & "." & B & "." & C & "." & D
End Function


Private Sub Command1_Click() 'add a new port to the port List
    If Text3.Text = "" Or IsNumeric(Text3.Text) = False Or Val(Text3.Text) > "65535" Or Text3.Text = "0" Then Exit Sub
    lstports.AddItem Text3.Text
    Text3.Text = ""
End Sub

Private Sub Command2_Click() 'range ports to the port list
On Error Resume Next
    If Val(Text4.Text) > "65535" Or Val(Text5.Text) > "65535" Then Exit Sub
    For i = Text4.Text To Text5.Text
    lstports.AddItem i
    Next i
End Sub

Private Sub Cmdstart_Click()
On Error Resume Next
    If lstports.ListCount = "0" Then MsgBox "add some ports to list!": Exit Sub

    cmdstart.Enabled = False
    reset
    disall

    'calculate the number of ip's to scan
    Dim bIP1() As String
    Dim bIP2() As String
    Dim lResult As Double
    bIP1 = Split(txtstart.Text, ".")
    bIP2 = Split(txtstop.Text, ".")
    lResult = (bIP2(3) - bIP1(3)) + (bIP2(2) - bIP1(2)) * 255 + (bIP2(1) - bIP1(1)) * 255 * 255 + (bIP2(0) - bIP1(0)) * 255 * 255 * 255
    maxxx = lResult
    Bar1.Max = maxxx

Select Case cmdstart.Caption
Case "Start Scan"
    cmdstart.Caption = "Stop Scan"
    For i = 1 To Slider1.Value
    Load ws(i) 'load the threads
    DoEvents
    lblstatus.Caption = "Loading Threads... [" & i & "]"
    Next i
    tmrscan.Enabled = True
    tmrstat.Enabled = True
    tmrtime.Enabled = True
    cmdstart.Enabled = True

Case "Stop Scan"
    On Error Resume Next
    cmdstart.Caption = "Start Scan"
    tmrscan.Enabled = False
    tmrstat.Enabled = False
    tmrtime.Enabled = False
    For z = 1 To Slider1.Value
    ws(z).Close
    DoEvents
    Unload ws(z)
    DoEvents
    lblstatus.Caption = "Killing Threads... [" & z & "]"
    Next z
    enall
    reset
    cmdstart.Enabled = True

End Select
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
reset
With ListView1
    .ColumnHeaders.Add , , "IP Address", 3000
    .ColumnHeaders.Add , , "Open Port", 1950
    .View = lvwReport
End With
End Sub

Private Sub Slider1_Change() 'shows the value of the slider1
Label6.Caption = Slider1.Value
End Sub

Private Sub Slider2_Change() 'shows the value of the slider2
Label8.Caption = Slider2.Value
tmrscan.Interval = Slider2.Value
End Sub

Private Sub reset() 'reset the stats
lblstatus.Caption = "idle..."
lbldone.Caption = "0 of 0"
lblspeed.Caption = "0 ip/sec"
lbltime.Caption = "00:00:00"
Bar1.Value = "0"
maxxx = "0"
tick = "0"
lstcounter = "0"
ipdone = "0"
End Sub

Private Sub disall() 'disable buttons, sliders and texts
txtstart.Enabled = False
txtstop.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Slider1.Enabled = False
Slider2.Enabled = False
End Sub

Private Sub enall() 'enable buttons, sliders and texts
txtstart.Enabled = True
txtstop.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Slider1.Enabled = True
Slider2.Enabled = True
End Sub

Private Sub tmrscan_Timer()
On Error Resume Next
For i = 1 To Slider1.Value
    ws(i).Close
    DoEvents 'spare the cpu
    ws(i).Connect txtstart.Text, lstports.List(lstcounter)
    DoEvents
    trynext
Next i
End Sub

Private Sub tmrstat_Timer() 'shows some stats
lblstatus.Caption = "Scanning " & txtstart.Text & " on Port " & lstports.List(lstcounter)
lbldone.Caption = ipdone & " of " & maxxx

If txtstart.Text = txtstop.Text Then 'stop the scan
On Error Resume Next
    cmdstart.Caption = "Start Scan"
    tmrscan.Enabled = False
    tmrstat.Enabled = False
    tmrtime.Enabled = False
    For z = 1 To Slider1.Value
    ws(z).Close
    DoEvents
    Unload ws(z)
    DoEvents
    lblstatus.Caption = "Killing Threads... [" & z & "]"
    Next z
    enall
    reset
    cmdstart.Enabled = True
    End If
End Sub

Private Sub tmrtime_Timer() 'calculate the speed and the rem. time
On Error Resume Next
Dim bps As String
Dim rt As String
tick = tick + 1
bps = ipdone / tick
rt = (maxxx - ipdone) / bps

lblspeed.Caption = Round(bps) & " ip/sec"
lbltime.Caption = FormatSeconds(rt)
End Sub

Private Sub ws_Close(Index As Integer)
trynext
End Sub

Private Sub ws_Connect(Index As Integer) 'add the host and the port to the listview
ListView1.ListItems.Add , , ws(Index).RemoteHost
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ws(Index).RemotePort
trynext
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
trynext
End Sub

Private Sub trynext() 'goto the next port or add one to the ip address
If lstcounter = lstports.ListCount - 1 Then
lstcounter = "0"
txtstart.Text = AddOne(txtstart.Text)
Else
lstcounter = lstcounter + 1
End If
End Sub

Public Function FormatSeconds(ByVal nSeconds As Long, _
  Optional ByVal sFormat As String = "hh:nn:ss") As String 'format the seconds in the hh:mm:ss format

  sFormat = Replace(sFormat, "mm", "nn")

  FormatSeconds = Format$(DateAdd("s", nSeconds, CDate("00:00:00")), sFormat)
End Function

Public Function Round( _
                ByVal Number As Double, _
                Optional ByVal NumDigitsAfterDecimal As Integer = 0 _
                ) As Double 'round the speed number
    Round = Int(Number * 10 ^ NumDigitsAfterDecimal + 0.5) _
                / 10 ^ NumDigitsAfterDecimal
End Function

