VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Memos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1
Public Name As String
Public Nick As String
Public ID As String
Public Email As String
Public Onchannels As New Collection
Public Idle As Long
Public Index As Long
Public NewUser As Boolean
Public BS As Long
Public BR As Long
Public Identified As Boolean
Public IRCOp As Boolean
Public Ident As String
Public SignOn As Long
Public DNS As String
Public NR As Boolean
Public OwnerOf As New Collection
Public SentQuit As Boolean
Public Modes As New Collection
Public Away As Boolean
Public AwayMsg As String
Public MsgsSent As Long
Public HasRegistered As Boolean
Public Killed As Boolean

Public Function GetOnChans() As String
Dim i As Long, Chan As clsChannel, NewChan As String
For i = 1 To Onchannels.Count
    Set Chan = ChanToObject(Onchannels(i))
    If Chan.IsOp(Nick) Then
        NewChan = "@" & Onchannels(i)
    ElseIf Chan.IsVoice(Nick) Then
        NewChan = "+" & Onchannels(i)
    Else
        NewChan = Onchannels(i)
    End If
    GetOnChans = GetOnChans & " " & NewChan
Next i
GetOnChans = Trim(GetOnChans)
End Function

Public Function GetWhois(NickName As String) As String
GetWhois = ":" & ServerName & " 311 " & NickName & " " & Nick & " " & Ident & " " & DNS & " * :" & Name & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 307 " & NickName & " " & Nick & " DNS " & DNS & vbCrLf
If Modes.Count <> 0 Then GetWhois = GetWhois & ":" & ServerName & " 309 " & NickName & " " & Nick & " :is using modes: +" & GetModes & vbCrLf
If Onchannels.Count <> 0 Then GetWhois = GetWhois & ":" & ServerName & " 319 " & NickName & " " & Nick & " :" & GetOnChans & vbCrLf
If Away Then GetWhois = GetWhois & ":" & ServerName & " 301 " & NickName & " " & Nick & " :" & AwayMsg & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 312 " & NickName & " " & Nick & " " & ServerName & " :" & ServerDesc & vbCrLf
If IRCOp Then
    GetWhois = GetWhois & ":" & ServerName & " 313 " & NickName & " " & Nick & " :is an IRC Operator - Server Administrator" & vbCrLf
    GetWhois = GetWhois & ":" & ServerName & " 310 " & NickName & " " & Nick & " :is a network helper" & vbCrLf
End If
GetWhois = GetWhois & ":" & ServerName & " 317 " & NickName & " " & Nick & " " & (UnixTime - Idle) & " " & SignOn & " :seconds idle, signon time" & vbCrLf
GetWhois = GetWhois & ":" & ServerName & " 318 " & NickName & " " & Nick & " :End of /WHOIS list." & vbCrLf
Debug.Print GetWhois
End Function

Public Function IsOwner(Channel As String) As Boolean
Dim i As Long
For i = 1 To OwnerOf.Count
    If OwnerOf(i) = Channel Then
        IsOwner = True
        Exit Function
    End If
Next i
End Function

Public Sub ClearOwnerShip()
Dim i As Long
For i = 1 To OwnerOf.Count
    OwnerOf.Remove i
Next i
End Sub

Public Property Get GetModes() As String
Dim i As Long
For i = 1 To Modes.Count
    GetModes = GetModes + Modes(i)
Next i
End Property

Public Function AddModes(NewModes As String) As String
    On Error GoTo ErrTrap
    Dim i As Long, sTemp As String
    For i = 1 To Len(NewModes)
        Modes.Add Mid(NewModes, i, 1), Mid(NewModes, i, 1)
        sTemp$ = sTemp$ & Mid$(NewModes$, i&, 1)
    Next i
    AddModes$ = sTemp$
Exit Function
ErrTrap:
    If Err.Number = 457 Then
        AddModes$ = sTemp$
    End If
End Function

Public Function IsMode(Mode As String) As Boolean
Dim i As Long
For i = 1 To Modes.Count
    If Modes(i) = Mode Then
        IsMode = True
        Exit Function
    End If
Next i
End Function

Public Function GetChanModes(ChanName As String) As String
Dim CurChan As clsChannel
Set CurChan = ChanToObject(ChanName)
If Not CurChan.IsOnChan(Nick) Then Exit Function
If CurChan.IsVoice(Nick) Then GetChanModes = GetChanModes & "+"
If CurChan.IsOp(Nick) Then GetChanModes = GetChanModes & "@"
End Function

Public Function GetMask() As String
If InStr(DNS, ".") > 0 Then
    GetMask = Mid(DNS, 1, InStr(1, DNS, ".") - 1)
    GetMask = Replace(DNS, GetMask, "*")
Else
    GetMask = "*"
End If
End Function

Public Function IsOnChan(Chan As String) As Boolean
Dim i As Long
For i = 1 To Onchannels.Count
    If UCase(Onchannels(i)) = UCase(Chan) Then
        IsOnChan = True
        Exit Function
    End If
Next i
End Function
