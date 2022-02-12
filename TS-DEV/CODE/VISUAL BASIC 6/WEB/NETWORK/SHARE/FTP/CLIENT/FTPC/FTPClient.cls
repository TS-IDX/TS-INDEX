VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const minmsTimeOut As Long = 3000

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private frmSck As frmSckWrapper
Private WithEvents SckData As Winsock
Attribute SckData.VB_VarHelpID = -1
Private WithEvents SckCmd As Winsock
Attribute SckCmd.VB_VarHelpID = -1

Public Enum enmReplyCode
RESTART_MARKER_REPLY_101 = 101
SERVICE_READY_SOON_120 = 120
TRANSFER_STARTING_125 = 125
INITIATING_DATA_CONNECTION_150 = 150
COMMAND_OK_200 = 200
COMMAND_UNSUPPORTED_202 = 202
SYSTEM_STATUS_211 = 211
DIRECTORY_STATUS_212 = 212
FILE_STATUS_213 = 213
HELP_MESSAGE_214 = 214
NAME_SYSTEM_TYPE_215 = 215
SERVICE_READY_220 = 220
CLOSING_CONTROL_CONNECTION_221 = 221
DATA_OPEN_NO_TRANSFER_225 = 225
CLOSING_DATA_CONNECTION_226 = 226
ENTERING_PASSIVE_MODE_227 = 227
USER_LOGGED_IN_230 = 230
FILE_ACTION_COMPLETED_250 = 250
PATH_CREATED_257 = 257
USER_OK_NEED_PASSWORD_331 = 331
LOGIN_NEED_ACCOUNT_332 = 332
FILE_ACTION_PENDING_350 = 350
SERVICE_SHUTTING_DOWN_421 = 421
DATA_INITIATION_FAILED_425 = 425
TRANSFER_ABORTION_426 = 426
FILE_UNAVAILABLE_450 = 450
ACTION_ABORTED_LOCAL_ERROR_451 = 451
FILE_ACTION_FAILED_SPACELESS_452 = 452
SYNTAX_ERROR_COMMAND_500 = 500
SYNTAX_ERROR_PARAMETER_501 = 501
COMMAND_NOT_IMPLEMENTED_502 = 502
BAD_SEQUENCE_COMMAND_503 = 503
COMMAND_NOT_IMPLEMENTED_PARAMETER_504 = 504
NOT_LOGGED_IN_530 = 530
ACCOUNT_REQUIRED_532 = 532
FILE_NO_ACCESS_550 = 550
ACTION_ABORTED_UNKNOWN_PAGE_TYPE_551 = 551
FILE_ACTION_FAILED_SPACE_EXCEEDED_552 = 552
FILE_ACTION_FAILED_NAME_DISALLOWED_553 = 553
End Enum

Public Enum enmFileInfo
    FI_DUMMY = 0
    FI_ATTRIB = 1
    FI_NAME
    FI_SIZE
    FI_CREATED_DATE
    FI_CREATED_TIME
    FI_PATH
End Enum

Private Type udtFILEDATA
    IsFolder As Boolean
    Size As Long
    CreatedDate As String
    CreatedTime As String
    FileName As String
End Type

Public Enum enmHostType
    HT_AUTODETECT = 0
    HT_WINDOWS_NT = 1
    HT_UNIX_STANDARD
End Enum

Public Enum enmCommandConnectionState
    CS_DISCONNECTED = 0
    CS_DISCONNECTING = 1
    CS_CONNECTING
    CS_CONNECTED
End Enum

Public Enum enmDataConnectionState
    DS_DISCONNECTED = 0
    DS_DISCONNECTING = 1
    DS_INITIATING
    DS_INITIATED
    DS_TRANSFERRING
    DS_TRANSFERRED_AND_DISCONNECTED
End Enum

Public Enum enmDataType
    DT_ASCII = 0
    DT_BINARY = 1
End Enum

Dim myRemHost As String
Dim myRemPort As Long
Dim myLocalPort As Long
Dim myUser As String
Dim myPass As String
Dim myAcc As String
Dim myPassive As Boolean
Dim myReplyCode As Long
Dim myType As enmDataType
Dim myHostType As enmHostType
Dim myCmdState As enmCommandConnectionState
Dim myDatState As enmDataConnectionState
Dim myCmdSckTimeOut As Long
Dim myDatSckTimeOut As Long
Dim myReplyTimeOut As Long

Private ListFileFlag As Byte, FILEDATA() As String, FileEntryCount As Long
Private sDirListing As String

Public Event CommandStateChanged()
Public Event CommandSocketTimeOut()
Public Event IncomingReply(ByVal Data As String)

Public Event DataStateChanged()
Public Event DataSocketTimeOut()
Public Event IncomingData(ByVal bytesTotal As Long)
Public Event NeedData()
Public Event TransferComplete()

Public Property Get CommandState() As enmCommandConnectionState
CommandState = myCmdState
End Property
Friend Property Let CommandState(eCS As enmCommandConnectionState)
If myCmdState <> eCS Then
    myCmdState = eCS
    RaiseEvent CommandStateChanged
End If
End Property

Public Property Get DataState() As enmDataConnectionState
DataState = myDatState
End Property
Friend Property Let DataState(eDS As enmDataConnectionState)
If myDatState <> eDS Then
    myDatState = eDS
    RaiseEvent DataStateChanged
End If
End Property

Public Property Get RemoteHost() As String
RemoteHost = myRemHost
End Property
Public Property Let RemoteHost(sRH As String)
If CommandState = CS_DISCONNECTED Then myRemHost = sRH
End Property

Public Property Get RemotePort() As Long
RemotePort = myRemPort
End Property
Public Property Let RemotePort(lRP As Long)
If CommandState = CS_DISCONNECTED Then myRemPort = lRP
End Property

Public Property Get LocalPort() As Long
LocalPort = myLocalPort
End Property
Public Property Let LocalPort(lLP As Long)
If CommandState = CS_DISCONNECTED Then myLocalPort = lLP
End Property

Public Property Get UserName() As String
UserName = myUser
End Property
Public Property Let UserName(sUN As String)
If CommandState = CS_DISCONNECTED Then myUser = sUN
End Property

Public Property Get Password() As String
Password = myPass
End Property
Public Property Let Password(sPW As String)
If CommandState = CS_DISCONNECTED Then myPass = sPW
End Property

Public Property Get Account() As String
Account = myAcc
End Property
Public Property Let Account(sACC As String)
If CommandState = CS_DISCONNECTED Then myAcc = sACC
End Property

Public Property Get PassiveMode() As Boolean
PassiveMode = myPassive
End Property
Public Property Let PassiveMode(blnPM As Boolean)
If DataState = DS_DISCONNECTED Then myPassive = blnPM
End Property

Public Property Get DataType() As enmDataType
DataType = myType
End Property
Public Property Let DataType(eDT As enmDataType)
Dim TempType As enmDataType
TempType = myType
myType = eDT
If CommandState = CS_CONNECTED Then
    If myType = DT_ASCII Then
        SckCmd.SendData "TYPE A" & vbCrLf
    ElseIf myType = DT_BINARY Then
        SckCmd.SendData "TYPE I" & vbCrLf
    End If
Else
    myType = TempType
End If
End Property

Public Property Get HostType() As enmHostType
HostType = myHostType
End Property
Public Property Let HostType(eHT As enmHostType)
myHostType = eHT
End Property

Public Property Get CommandSocketmsTimeOut() As Long
CommandSocketmsTimeOut = myCmdSckTimeOut
End Property
Public Property Let CommandSocketmsTimeOut(lCST As Long)
If lCST > minmsTimeOut Then myCmdSckTimeOut = lCST
End Property

Public Property Get DataSocketmsTimeOut() As Long
DataSocketmsTimeOut = myDatSckTimeOut
End Property
Public Property Let DataSocketmsTimeOut(lDST As Long)
If lDST > minmsTimeOut Then myDatSckTimeOut = lDST
End Property

Public Property Get ReplyTimeOut() As Long
ReplyTimeOut = myReplyTimeOut
End Property
Public Property Let ReplyTimeOut(LRTO As Long)
If LRTO > minmsTimeOut Then myReplyTimeOut = LRTO
End Property

Sub Dispose()
Call CloseDataSocket
Set SckCmd = Nothing
Set SckData = Nothing
Unload frmSck
Set frmSck = Nothing
End Sub

Function Connect() As Boolean
Connect = False
'If Disconnect = False Then Exit Function
SckCmd.RemoteHost = myRemHost
SckCmd.RemotePort = myRemPort
If myLocalPort = 0 Then myLocalPort = SckCmd.LocalPort
SckCmd.LocalPort = myLocalPort
CommandState = CS_CONNECTING
SckCmd.Connect
If WaitCommandSocketState(sckConnected, myCmdSckTimeOut) = False Then Exit Function
Connect = True
End Function

Function Disconnect() As Boolean
Disconnect = False
SckCmd.Close
If WaitCommandSocketState(sckClosed, myCmdSckTimeOut) = False Then Exit Function
If CloseDataSocket = False Then Exit Function
Disconnect = True
End Function

Sub Execute(FTPCommand As String)
If CommandState = CS_CONNECTED Then SckCmd.SendData FTPCommand & vbCrLf
End Sub

Function OpenDataSocket() As Boolean
OpenDataSocket = False
ListFileFlag = 0
'If CloseDataSocket = False Then Exit Function
If myCmdState <> CS_CONNECTED Then Exit Function
If myPassive = True Then 'Passive mode
    SckCmd.SendData "PASV" & vbCrLf
    OpenDataSocket = True
Else 'Port mode
    On Error Resume Next
    With SckData
    Dim LocalIP As String, Port1&, Port2&, iCount As Integer
    LocalIP = .LocalIP
    LocalIP = Replace$(LocalIP, ".", ",")
    For iCount = 1 To 3
        Randomize Timer
        Port1& = Int(Rnd * (32768 - 1024)) + 1024
        .Close
        If WaitDataSocketState(sckClosed, myDatSckTimeOut) = True Then
            .LocalPort = Port1&
            .Listen
            Exit For
        End If
    Next
    If .State = sckListening Then
        SckCmd.SendData "PORT " & LocalIP & "," & Trim$(Str$(Port1& \ 256)) & "," & Trim$(Str$(Port1& Mod 256)) & vbCrLf
        If WaitReplyCode(COMMAND_OK_200, myDatSckTimeOut) = True Then
            DataState = DS_INITIATED
            OpenDataSocket = True
        End If
    End If
    End With
End If
End Function

Function CloseDataSocket() As Boolean
CloseDataSocket = False
SckData.Close
If WaitDataSocketState(sckClosed, myDatSckTimeOut) = False Then Exit Function
CloseDataSocket = True
End Function

Sub SendData(Data As String)
If SckData.State = sckConnected Then SckData.SendData Data
End Sub

Sub GetData(Data As String, Length As Long)
If SckData.State = sckConnected Then SckData.GetData Data, vbString, Length
End Sub

Function UploadFile(RemoteFile As String) As Boolean
UploadFile = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "STOR " & RemoteFile & vbCrLf
    If WaitReplyCode(TRANSFER_STARTING_125, myReplyTimeOut) = True Then
        UploadFile = True
    End If
End If
End Function

Function AppendFile(RemoteFile As String) As Boolean
AppendFile = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "APPE " & RemoteFile & vbCrLf
    If WaitReplyCode(TRANSFER_STARTING_125, myReplyTimeOut) = True Then
        AppendFile = True
    End If
End If
End Function

Function SetBytePosition(BytePosition As Long) As Boolean
SetBytePosition = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "REST " & Trim$(CStr(BytePosition)) & vbCrLf
    If WaitReplyCode(FILE_ACTION_PENDING_350, myReplyTimeOut) = True Then
        SetBytePosition = True
    End If
End If
End Function

Function DownloadFile(RemoteFile As String) As Boolean
DownloadFile = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "RETR " & RemoteFile & vbCrLf
    If WaitReplyCode(TRANSFER_STARTING_125, myReplyTimeOut) = True Then
        DownloadFile = True
    End If
End If
End Function

Function MakeDirectory(Directory As String) As Boolean
MakeDirectory = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "MKD " & Directory & vbCrLf
    If WaitReplyCode(PATH_CREATED_257, myReplyTimeOut) = True Then
        MakeDirectory = True
    End If
End If
End Function

Function ChangeDirectory(Directory As String) As Boolean
ChangeDirectory = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "CWD " & Directory & vbCrLf
    If WaitReplyCode(FILE_ACTION_COMPLETED_250, myReplyTimeOut) = True Then
        ChangeDirectory = True
    End If
End If
End Function

Function DeleteDirectory(Directory As String) As Boolean
DeleteDirectory = False
If myCmdState = CS_CONNECTED Then
    SckCmd.SendData "RMD " & Directory & vbCrLf
    If WaitReplyCode(FILE_ACTION_COMPLETED_250, myReplyTimeOut) = True Then
        DeleteDirectory = True
    End If
End If
End Function

Public Property Get FileCount() As Long
FileCount = FileEntryCount
End Property

Public Property Get FileInfo(Index As Long, InfoType As enmFileInfo) As String
If FileCount Then
    If InfoType < 6 Then
        FileInfo = FILEDATA(InfoType, Index)
    'ElseIf InfoType = FI_PATH Then
        'FileList = RemoteDirPath
    End If
End If
End Property

Sub ListFiles(File As String)
If myCmdState = CS_CONNECTED Then
    ListFileFlag = 1
    SckCmd.SendData "NLST " & File & vbCrLf
End If
End Sub

Sub ListFilesInfo(File As String)
If myCmdState = CS_CONNECTED Then
    ListFileFlag = 2
    SckCmd.SendData "LIST " & File & vbCrLf
End If
End Sub

Private Sub ParseDirListing(ParseInfo As Boolean)
Dim startPnt As Long, Count As Long, Separator As Long
Dim SkipParCount As Long, strRet As String, DirPoint As Long
Dim sFileInfo As String, DirLen As Long, I As Long
Erase FILEDATA
ReDim FILEDATA(5, 0)
FileEntryCount = 0
DirPoint = 1
DirLen = Len(sDirListing)
If ParseInfo = True Then
ReEvaluate:
Select Case myHostType
Case enmHostType.HT_UNIX_STANDARD
Do Until DirPoint > DirLen
    GoSub GetRawFileInfo
    FileEntryCount = FileEntryCount + 1
    ReDim Preserve FILEDATA(5, FileEntryCount)
    GoSub ParseUNIX
Loop
Case enmHostType.HT_WINDOWS_NT
Do Until DirPoint > DirLen
    GoSub GetRawFileInfo
    FileEntryCount = FileEntryCount + 1
    ReDim Preserve FILEDATA(5, FileEntryCount)
    GoSub ParseWinNT
Loop
Case enmHostType.HT_AUTODETECT
GoSub GetRawFileInfo
    FileEntryCount = FileEntryCount + 1
    ReDim Preserve FILEDATA(5, FileEntryCount)
    
    If IsNumeric(Left$(sFileInfo, 1)) Then
        myHostType = HT_WINDOWS_NT
        GoSub ParseWinNT
    Else
        myHostType = HT_UNIX_STANDARD
        GoSub ParseUNIX
    End If
GoTo ReEvaluate
End Select
Else
Do Until DirPoint > DirLen
    GoSub GetRawFileInfo
    FileEntryCount = FileEntryCount + 1
    ReDim Preserve FILEDATA(5, FileEntryCount)
    FILEDATA(enmFileInfo.FI_NAME, FileEntryCount) = Trim$(sFileInfo)
Loop
End If
Exit Sub

GetRawFileInfo:
Separator = InStr(DirPoint, sDirListing, vbCrLf)
If Separator Then
    sFileInfo = Mid$(sDirListing, DirPoint, Separator - DirPoint)
    DirPoint = Separator + 2
Else
    sFileInfo = Mid$(sDirListing, DirPoint)
    DirPoint = Len(sDirListing) + 2
End If
sFileInfo = Trim$(sFileInfo) & " "
Return

ParseUNIX:
    'First character is attribute
    FILEDATA(enmFileInfo.FI_ATTRIB, FileEntryCount) = IIf(LCase$(Left$(sFileInfo, 1)) = "d", "d", "f")
    'File size
    SkipParCount = 4
    GoSub GetSingleParam
    FILEDATA(enmFileInfo.FI_SIZE, FileEntryCount) = strRet
    'Created date
    SkipParCount = 5
    GoSub GetSingleParam
    FILEDATA(enmFileInfo.FI_CREATED_DATE, FileEntryCount) = strRet
    'Created Date
    SkipParCount = 6
    GoSub GetSingleParam
    FILEDATA(enmFileInfo.FI_CREATED_DATE, FileEntryCount) = strRet
    'Created time
    SkipParCount = 7
    GoSub GetSingleParam
    FILEDATA(enmFileInfo.FI_CREATED_DATE, FileEntryCount) = strRet
    'File name
    SkipParCount = 8
    GoSub GetSingleParam
    FILEDATA(enmFileInfo.FI_NAME, FileEntryCount) = strRet
Return

ParseWinNT:
'Date
FILEDATA(enmFileInfo.FI_CREATED_DATE, FileEntryCount) = Left$(sFileInfo, InStr(1, sFileInfo, " ") - 1)
'Time
SkipParCount = 1
GoSub GetSingleParam
FILEDATA(enmFileInfo.FI_CREATED_TIME, FileEntryCount) = strRet
SkipParCount = 2
GoSub GetSingleParam
If IsNumeric(Left$(strRet, 1)) Then
    'Size
    FILEDATA(enmFileInfo.FI_SIZE, FileEntryCount) = strRet
    FILEDATA(enmFileInfo.FI_ATTRIB, FileEntryCount) = "f"
Else
    'Directory
    FILEDATA(enmFileInfo.FI_ATTRIB, FileEntryCount) = "d"
End If
SkipParCount = 3
GoSub GetSingleParam
FILEDATA(enmFileInfo.FI_NAME, FileEntryCount) = strRet
Return
GetSingleParam:
'Variables:SkipParCount,strRET
    startPnt = 1
For I = 1 To SkipParCount
    startPnt = InStr(startPnt, sFileInfo, " ")
    Do Until Mid$(sFileInfo, startPnt, 1) <> " "
        startPnt = startPnt + 1
    Loop
Next
Separator = InStr(startPnt, sFileInfo, " ")
strRet = Mid$(sFileInfo, startPnt, Separator - startPnt)
Return
End Sub

Function DeleteFile(File As String) As Boolean
DeleteFile = False
If SckCmd.State = sckConnected Then
    SckCmd.SendData "DELE " & File & vbCrLf
    If WaitReplyCode(FILE_ACTION_COMPLETED_250, myReplyTimeOut) = True Then
        DeleteFile = True
    End If
End If
End Function

Function RenameFile(OldName As String, NewName As String) As Boolean
RenameFile = False
If SckCmd.State = sckConnected Then
    SckCmd.SendData "RNFR " & OldName & vbCrLf
    If WaitReplyCode(FILE_ACTION_PENDING_350, myReplyTimeOut) = True Then
        SckCmd.SendData "RNTO " & NewName & vbCrLf
        If WaitReplyCode(FILE_ACTION_COMPLETED_250, myReplyTimeOut) = True Then
            RenameFile = True
        End If
    End If
End If
End Function

Function WaitCommandState(State As enmCommandConnectionState, Optional ByVal msTimeOut As Long) As Boolean
WaitCommandState = False
Dim T As Long
T = GetTickCount + msTimeOut
If msTimeOut = 0 Then T = T + minmsTimeOut
Do Until GetTickCount >= T
    If CommandState = State Then
        WaitCommandState = True
        Exit Do
    End If
    DoEvents
Loop
End Function

Function WaitDataState(State As enmDataConnectionState, Optional ByVal msTimeOut As Long) As Boolean
WaitDataState = False
Dim T As Long
T = GetTickCount + msTimeOut
If msTimeOut = 0 Then T = T + minmsTimeOut
Do Until GetTickCount >= T
    If DataState = State Then
        WaitDataState = True
        Exit Do
    End If
    DoEvents
Loop
End Function

Private Function WaitCommandSocketState(State As StateConstants, Optional ByVal msTimeOut As Long) As Boolean
WaitCommandSocketState = False
Dim T As Long
T = GetTickCount + msTimeOut
If msTimeOut = 0 Then T = T + myCmdSckTimeOut
Do Until GetTickCount >= T
    If SckCmd.State = State Then
        WaitCommandSocketState = True
        Exit Do
    End If
    DoEvents
Loop
If WaitCommandSocketState = False Then RaiseEvent CommandSocketTimeOut
End Function

Private Function WaitDataSocketState(State As StateConstants, Optional ByVal msTimeOut As Long) As Boolean
WaitDataSocketState = False
Dim T As Long
T = GetTickCount + msTimeOut
If msTimeOut = 0 Then T = T + myDatSckTimeOut
Do Until GetTickCount >= T
    If SckData.State = State Then
        WaitDataSocketState = True
        Exit Do
    End If
    DoEvents
Loop
If WaitDataSocketState = False Then RaiseEvent DataSocketTimeOut
End Function

Function WaitReplyCode(ReplyCode As enmReplyCode, Optional ByVal msTimeOut As Long) As Boolean
WaitReplyCode = False
Dim T As Long
T = GetTickCount + msTimeOut
If msTimeOut = 0 Then T = T + myReplyTimeOut
Do Until GetTickCount >= T
    If myReplyCode = ReplyCode Then
        WaitReplyCode = True
        Exit Do
    End If
    DoEvents
Loop
End Function

Private Sub ProcessReplyCode(ByVal cmdData As String)
Dim ReplyCode As Long
ReplyCode = CLng(Val(Left$(cmdData, 3)))
myReplyCode = ReplyCode
Select Case ReplyCode
Case enmReplyCode.SERVICE_READY_220
    SckCmd.SendData "USER " & myUser & vbCrLf
    'CommandState = CS_LOGGING_IN
Case enmReplyCode.USER_OK_NEED_PASSWORD_331
    SckCmd.SendData "PASS " & myPass & vbCrLf
    'CommandState = CS_VERIFYING_CLIENT
Case enmReplyCode.ACCOUNT_REQUIRED_532
    SckCmd.SendData "ACCT " & myAcc & vbCrLf
    'CommandState = CS_VERIFYING_CLIENT
Case enmReplyCode.LOGIN_NEED_ACCOUNT_332
    SckCmd.SendData "ACCT " & myAcc & vbCrLf
    'CommandState = CS_VERIFYING_CLIENT
Case enmReplyCode.USER_LOGGED_IN_230
    CommandState = CS_CONNECTED
Case enmReplyCode.CLOSING_DATA_CONNECTION_226
    RaiseEvent TransferComplete
Case enmReplyCode.ENTERING_PASSIVE_MODE_227
    'Parse replied data to get remote IP and port
    Dim LeftSep%, RightSep%, Extracts$
    Dim IP$, PRT$, Port1&, Port2&, Port&
    LeftSep% = InStr(1, cmdData, "(")
    RightSep% = InStr(LeftSep%, cmdData, ")")
    Extracts$ = Mid$(cmdData, LeftSep% + 1, RightSep% - LeftSep% - 1)
    Extracts$ = Replace$(Extracts$, ",", ".", 1, 3)
    PRT$ = Mid$(Extracts$, InStr(1, Extracts$, ",") + 1)
    IP$ = Mid$(Extracts$, 1, InStr(1, Extracts$, ",") - 1)
    Port1& = CLng(Val(Trim$(Mid$(PRT$, 1, InStr(1, PRT$, ",") - 1))))
    Port2& = CLng(Val(Trim$(Mid$(PRT$, InStr(1, PRT$, ",") + 1))))
    Port& = (Port1& * 256) + Port2&
    'IP$ and Port&
    Call CloseDataSocket
    SckData.RemoteHost = IP$
    SckData.RemotePort = Port&
    DataState = DS_INITIATING
    SckData.Connect
'Case Is >= 400
'An error occured
'RaiseEvent LastActionError(RCode)
End Select
End Sub

Private Sub Class_Initialize()
Set frmSck = New frmSckWrapper
Load frmSck
frmSck.Visible = False
Set SckCmd = frmSck.SckCmd
Set SckData = frmSck.SckData
'Default value
myReplyTimeOut = minmsTimeOut
myCmdSckTimeOut = minmsTimeOut
myDatSckTimeOut = minmsTimeOut
myRemPort = 21
myHostType = HT_AUTODETECT
End Sub

Private Sub SckCmd_DataArrival(ByVal bytesTotal As Long)
Dim sData As String, ReplyCode As Long, cmdData As String
Dim SplitPoint As Long, StartPoint As Long
    SplitPoint = 1: StartPoint = 1
    SckCmd.GetData sData
    Do
        'vbCrLf separates each individual command line
        SplitPoint = InStr(StartPoint, sData, vbCrLf)
        If SplitPoint < 1 Then Exit Do
        cmdData = Mid$(sData, StartPoint, SplitPoint - StartPoint)
        'RCode = CLng(Val(Left$(cmdData, 3)))
        StartPoint = SplitPoint + 2
        RaiseEvent IncomingReply(cmdData)
        Call ProcessReplyCode(cmdData)
        DoEvents
    Loop
End Sub

Private Sub SckData_Close()
'Data socket is closed; ready to flush buffer
If DataState = DS_TRANSFERRING Then
    If ListFileFlag = 1 Then
        ParseDirListing False
    ElseIf ListFileFlag = 2 Then
        ParseDirListing True
    End If
    ListFileFlag = 0
    sDirListing = ""
    DataState = DS_TRANSFERRED_AND_DISCONNECTED
ElseIf DataState = DS_DISCONNECTED Then
'Do nothing
Else
    DataState = DS_DISCONNECTED
End If
End Sub

Private Sub SckData_Connect()
DataState = DS_INITIATED
End Sub

Private Sub SckData_ConnectionRequest(ByVal requestID As Long)
If CloseDataSocket = True Then
    SckData.Accept requestID
    DataState = DS_INITIATED
End If
End Sub

Private Sub SckData_DataArrival(ByVal bytesTotal As Long)
If SckData.State = sckConnected Then
    DataState = DS_TRANSFERRING
    If ListFileFlag = 0 Then
        RaiseEvent IncomingData(bytesTotal)
    Else
        Dim Data As String
        SckData.GetData Data
        sDirListing = sDirListing & Data
    End If
End If
End Sub

Private Sub SckData_SendComplete()
RaiseEvent NeedData
End Sub
