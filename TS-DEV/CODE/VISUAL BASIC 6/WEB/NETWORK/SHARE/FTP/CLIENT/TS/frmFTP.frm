VERSION 5.00
Begin VB.Form frmftp 
   BackColor       =   &H00404040&
   Caption         =   "tsFTP 0.0.1"
   ClientHeight    =   6480
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Inet1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   18
      Top             =   1800
      Width           =   1200
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   4200
      Width           =   5895
   End
   Begin VB.TextBox txtRemoteDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtRemoteFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtLocalFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   13
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change to remote dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get remote file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Put local file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LogOff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LogOn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox txtURL 
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   6120
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Password:"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "User Name:"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Server URL:"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'GlowFTP 1.0
'Written by: S.S. Ahmed
'Nov 2001
'**************************************************************

'**************************************************************
'FTP Routine
'Description: Uploads, downloads files from the remote FTP server
'**************************************************************

'General Description Area
Option Explicit


Private Sub About_Click()
frmAbout.Show

End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo errhandler
Select Case Index
    Case 0 'LogOn Button
        txtDir.Text = ""
        Call LogOn
        
    Case 1 'LogOff Button
        txtDir.Text = ""
        Inet1.Cancel
        Call SetCmdButtonState(False)
        lblStatus.Caption = "Logged Off"
            
    Case 2 'Refresh Button
        Call RefreshDirList
        
    Case 3 'End Button
        Inet1.Cancel
        End
End Select
Exit Sub
errhandler:
    MsgBox Err.Source & " " & Err.Number & " " & Err.Description

End Sub

Private Sub Command2_Click(Index As Integer)

Dim s As String

Select Case Index

    Case 0 'Put a local file
        If txtLocalFileName.Text = "" Then
            s = "You must enter the name of the " & vbCrLf
            s = s & "local file to upload."
            MsgBox (s)
            Exit Sub
            
        End If
        Call PutLocalFile
        
    Case 1 'Get remote file
        If txtRemoteFileName.Text = "" Then
            s = "You must specify the name of the " & vbCrLf
            s = s & "remote file to download."
            MsgBox (s)
            Exit Sub
        
        End If
        Call GetRemoteFile
        
    Case 2 'Change Directory
        If txtRemoteDir.Text = "" Then
            s = "Please enter the name of the new directory."
            MsgBox (s)
            Exit Sub
        End If
        Call ChangeDir

End Select


End Sub

Private Sub Form_Load()

'Set buttons initially for logged off state.
Call SetCmdButtonState(False)

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error GoTo errhandler
Dim Data1 As String, Data2 As String

Select Case State
    
    Case icResolvingHost
        lblStatus.Caption = "Looking up host computer IP address"

    Case icHostResolved
        lblStatus.Caption = "IP address found"
        
    Case icConnecting
        lblStatus.Caption = "Connecting to host computer"


    Case icConnected
        lblStatus.Caption = "Connected to host computer"

    Case icRequesting
        lblStatus.Caption = "Sending a request to host computer"
        
    Case icRequestSent
        lblStatus.Caption = "Request sent"

    Case icReceivingResponse
        lblStatus.Caption = "Receiving a response from host computer"

    Case icResponseReceived
        lblStatus.Caption = "Response received"

    Case icDisconnecting
        lblStatus.Caption = "Disconnecting from host computer"


    Case icDisconnected
        lblStatus.Caption = "Disconnected from host computer"

    Case icError
        lblStatus.Caption = "Error " & Inet1.ResponseCode & Inet1.ResponseInfo

    Case icResponseCompleted
        lblStatus.Caption = "Request completed successfully"
        'Loop until you get all chunks
        Do While True
            'Change datatype to icByteArray to receive data in binary
            Data1 = Inet1.GetChunk(512, icString)
            If Len(Data1) = 0 Then Exit Do
            DoEvents 'Transfer control to operating system
            Data2 = Data2 & Data1
        Loop
        txtDir.Text = Data2
End Select
Exit Sub
errhandler:
    MsgBox Err.Source & " " & Err.Number & " " & Err.Description
End Sub

Public Sub LogOn()

'Logs on to the FTP host specified in the txtURL text box
'and displays the directory.

On Error GoTo LogOnError

If txtURL.Text = "" Or txtPassword.Text = "" Then
    MsgBox ("You must specify a URL and Password")
    Exit Sub
End If

Command1(0).Enabled = False
Inet1.Protocol = icFTP
Inet1.URL = txtURL.Text
If txtUserName.Text = "" Then
    Inet1.UserName = "anonymous"
Else
    Inet1.UserName = txtUserName.Text
End If

Inet1.Password = txtPassword.Text
Inet1.Execute , "DIR"
Call SetCmdButtonState(True)

Exit Sub

LogOnError:
    If Err = 35754 Then
        MsgBox "Cannot connect to the remote host"
    Else
        MsgBox Err.Description
    End If
    Call SetCmdButtonState(False)
    Inet1.Cancel
    
End Sub

Public Sub PutLocalFile()

On Error GoTo PutLocalFileErr

'Puts the local file specified in the txtlocalfilename to the
'remote server.

Dim RemoteFileName As String, cmd As String

RemoteFileName = InputBox("Name for remote file?", "Get", txtLocalFileName.Text)

cmd = "PUT " & RemoteFileName & " " & txtLocalFileName.Text

If inetReady(True) Then
    Call SendCommand(cmd)
End If

Exit Sub

PutLocalFileErr:
    MsgBox "Error: " & Err.Description
    'Resume Next
End Sub

Public Sub GetRemoteFile()

'Retrieves the remote file specified in the textbox
'txtRemoteFileName. Stores the downloaded file using
'user specified name.

Dim LocalFileName As String, cmd As String

On Error GoTo GetRemoteFileErr

LocalFileName = InputBox("Name for local file?", "GET", txtRemoteFileName.Text)

cmd = "GET " & txtRemoteFileName.Text & " " & LocalFileName
If inetReady(True) Then
    Call SendCommand(cmd)
End If
Exit Sub
GetRemoteFileErr:
    MsgBox "Error: " & Err.Description
    'Resume Next
End Sub

Public Sub ChangeDir()
On Error GoTo errhandler
Dim cmd As String

'Changes to the remote directory specified in the txtremotedir
' text box; then displays its directory.

cmd = "CD " & txtRemoteDir.Text
If inetReady(True) Then
    Call SendCommand(cmd)
    Do
        DoEvents 'Transfer control to the operating system
    Loop Until inetReady(False)
    txtDir.Text = ""
    Call SendCommand("DIR")
End If
Exit Sub
errhandler:
    MsgBox Err.Source & " " & Err.Number & " " & Err.Description
End Sub

Public Sub SetCmdButtonState(LoggedOn As Boolean)
On Error GoTo errhandler
Dim x As Integer

'Enables and disables the program's command button for the
'logged on and logged off situations.

If LoggedOn Then
'Logged on state.
    Command1(0).Enabled = False 'LogOn Button
    Command1(1).Enabled = True 'Logoff button
    Command1(2).Enabled = True 'Refresh
    For x = 0 To Command2.Count - 1
        Command2(x).Enabled = True
    Next
Else
'Logged off state
    Command1(0).Enabled = True 'logon
    Command1(1).Enabled = False 'logoff
    Command1(2).Enabled = False 'refresh
    For x = 0 To Command2.Count - 1
        Command2(x).Enabled = False
    Next
End If
Exit Sub
errhandler:
MsgBox Err.Source & " " & Err.Number & " " & Err.Description
End Sub

Public Sub RefreshDirList()
'Refreshed the remote directory listing
On Error GoTo errhandler

If inetReady(True) Then
    txtDir.Text = ""
    Call SendCommand("DIR")
End If
Exit Sub
errhandler:
    MsgBox Err.Source & " " & Err.Number & " " & Err.Description
End Sub

Public Sub SendCommand(cmd As String)

'Sends the specified command to the FTP Server.

On Error GoTo SendCommandErr

Inet1.Execute , cmd
Exit Sub

SendCommandErr:
    MsgBox "Error: " & Err.Description
    Resume Next

End Sub

Public Function inetReady(Message As Boolean) As Boolean

'Returns true if inet1 is ready to execute a  new command.
'If the control is busy and the message argument is true, displays
'an error message.
On Error GoTo errhandler
Dim msg As String

If Inet1.StillExecuting Then
    If Message Then
        msg = "The program has not finished "
        msg = msg & "executing your last request." & vbCrLf
        msg = msg & "Please wait then try again later."
        MsgBox (msg)
    
    End If
    inetReady = False
Else
    inetReady = True
End If
Exit Function
errhandler:
    MsgBox Err.Source & " " & Err.Number & " " & Err.Description
End Function
