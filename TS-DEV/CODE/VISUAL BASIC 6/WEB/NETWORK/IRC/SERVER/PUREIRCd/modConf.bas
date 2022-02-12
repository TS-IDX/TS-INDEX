Attribute VB_Name = "modConf"
Option Explicit

'/*
' *  MODES: OoKkMAYIiwsRPpHLQ
' *     O - Global Operator
' *     o - Local Operator
' *     K - Can Set/Unset Global K-Line (O Required.)
' *     k - Can Set/Unset Local K-Line (O/o Required.)
' *     M - Can Set Local M-Line (O/o Required.)
' *     A - Can Set Local A-Line (O/o Required.)
' *     Y - Can Set Local Y-Line (O/o Required.)
' *     I - Can Set Local I-Line (O/o Required.)
' *     i - Invisible
' *     w - Receive WALLOPS
' *     s - Receive Server Messages
' *     R - Only Registered nicks may send PRIVMSG
' *     P - Can Set Local P-Line (O/o Required.)
' *     p - All Channels Unlisted on WHOIS (O/o Required.)
' *     H - Can Set Local H-Line (O Required.)
' *     L - Can Set Local L-Line (O Required.)
' *     Q - Can Set Local Q-Line (O/o Required.)
' */

  
Private Type STRUCT_M_LINE
    'M:HostName:BindAddr:TextName
    HostName As String
    BindAddr As String
    TextName As String
End Type

Private Type STRUCT_A_LINE
    'A:ServName:ServLoc:ServEmail
    ServName As String
    ServLoc As String
    ServEmail As String
End Type

Private Type STRUCT_Y_LINE
    'Y:ClassNum:PingFreq:ConFreq:MaxLinks:MaxSendQ
    ClassNum As Integer
    PingFreq As Integer
    ConFreq As Integer
    MaxLinks As Integer
    MaxSendQ As Long
End Type

Private Type STRUCT_I_LINE
    'I:HostMatch:Passwd:ClassNum
    HostMatch As String
    Passwd As String
    ClassNum As Integer
End Type
    
Private Type STRUCT_O_LINE
    'O/o:HostName:Passwd:OpNick:Flags:ClassNum
    HostName As String
    Passwd As String
    OpNick As String
    Flags As String
    ClassNum As Integer
    Global As Integer
End Type

Private Type STRUCT_H_LINE
    'H:ServAddr:Passwd:MaxLeaf
    ServAddr As String
    Passwd As String
    MaxLeaf As Integer
End Type

Private Type STRUCT_L_LINE
    'L:ServAddr:Passwd:MaxDepth
    ServAddr As String
    Passwd As String
    MaxDepth As Integer
End Type

Private Type STRUCT_P_LINE
    'P:HostName:Passwd:HostPort:MaxCon
    HostName As String
    Passwd As String
    HostPort As Integer
    MaxCon As Integer
End Type

Private Type STRUCT_K_LINE
    'K/k:HostName
    HostName As String
    Global As Integer
End Type

Private Type STRUCT_Q_LINE
    'Q:NickMask:Reason
    NickMask As String
    Reason As String
End Type

Public M_Line() As STRUCT_M_LINE
Public Y_Line() As STRUCT_Y_LINE
Public I_Line() As STRUCT_I_LINE
Public H_Line() As STRUCT_H_LINE
Public L_Line() As STRUCT_L_LINE
Public A_Line() As STRUCT_A_LINE
Public P_Line() As STRUCT_P_LINE
Public O_Line() As STRUCT_O_LINE
Public K_Line() As STRUCT_K_LINE
Public Q_Line() As STRUCT_Q_LINE

Public Sub LoadIRCdConf(Optional sDir As String)
    Dim nOpenFile As Integer, sString As String
    Dim sArray() As String, nK As Integer
    '/* load ircd.conf into memory for processing */
    ReDim M_Line(1)
    ReDim Y_Line(1)
    ReDim I_Line(1)
    ReDim H_Line(1)
    ReDim L_Line(1)
    ReDim A_Line(1)
    ReDim P_Line(1)
    ReDim O_Line(1)
    ReDim K_Line(1)
    ReDim Q_Line(1)
    If sDir$ = "" Then sDir$ = App.Path & "\ircd.conf"
    nOpenFile% = FreeFile
    Open sDir$ For Binary As #nOpenFile%
        sString$ = Space(LOF(nOpenFile%))
        Get #nOpenFile%, , sString$
    Close #nOpenFile%
    '/* parse ircd.conf for line processing */
    sString$ = Replace$(sString$, Chr$(13), "")
    sArray() = Split(sString$, Chr$(10))
    For nK% = 0 To UBound(sArray())
        If LCase(Left$(sArray(nK%), 9)) = ".include " Then
            Call LoadIRCdConf(Trim$(Mid$(sArray(nK%), 9)))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "k:" Then
            Call Process_K_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "m:" Then
            Call Process_M_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "p:" Then
            Call Process_P_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "q:" Then
            Call Process_Q_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "o:" Then
            Call Process_O_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "y:" Then
            Call Process_Y_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "i:" Then
            Call Process_I_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "h:" Then
            Call Process_H_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "l:" Then
            Call Process_L_Line(sArray(nK%))
        ElseIf LCase$(Left$(sArray(nK%), 2)) = "a:" Then
            Call Process_A_Line(sArray(nK%))
        End If
    Next nK%
End Sub

Private Sub Process_O_Line(sLine As String)
    'O/o:HostName:Passwd:OpNick:Flags:ClassNum
    Dim sArray() As String, nK As Integer
    sArray() = Split(sLine$, ":")
    For nK% = 0 To UBound(sArray())
        If nK% = 0 Then
            If sArray(nK%) = "O" Then
                ReDim Preserve O_Line(UBound(O_Line()) + 1)
                O_Line(UBound(O_Line())).Global = 1
            ElseIf sArray(nK%) = "o" Then
                ReDim Preserve O_Line(UBound(O_Line()) + 1)
                O_Line(UBound(O_Line())).Global = 0
            Else
                Exit Sub
            End If
        ElseIf nK% = 1 Then
            If sArray(nK%) <> "" Then
                O_Line(UBound(O_Line())).HostName = sArray(nK%)
            Else
                O_Line(UBound(O_Line())).HostName = "*"
            End If
        ElseIf nK% = 2 Then
            If sArray(nK%) <> "" Then
                O_Line(UBound(O_Line())).Passwd = sArray(nK%)
            Else
                O_Line(UBound(O_Line())).Passwd = "DF4vfet5G6vcfGT5t4v5Vrtv5Vv4Vth34cSgrn"
            End If
        ElseIf nK% = 3 Then
            If sArray(nK%) <> "" Then
                O_Line(UBound(O_Line())).OpNick = sArray(nK%)
            Else
                O_Line(UBound(O_Line())).OpNick = "Server"
            End If
        ElseIf nK% = 4 Then
            If sArray(nK%) <> "" Then
                O_Line(UBound(O_Line())).Flags = sArray(nK%)
            Else
                O_Line(UBound(O_Line())).Flags = "Ooipwk"
            End If
        ElseIf nK% = 5 Then
            If sArray(nK%) <> "" Then
                If IsNumeric(sArray(nK%)) Then
                    O_Line(UBound(O_Line())).ClassNum = sArray(nK%)
                Else
                    O_Line(UBound(O_Line())).ClassNum = 0
                End If
            Else
                O_Line(UBound(O_Line())).ClassNum = 0
            End If
        End If
    Next nK%
End Sub

Private Sub Process_M_Line(sLine As String)

End Sub

Private Sub Process_P_Line(sLine As String)

End Sub

Private Sub Process_Q_Line(sLine As String)

End Sub

Private Sub Process_K_Line(sLine As String)

End Sub

Private Sub Process_Y_Line(sLine As String)

End Sub

Private Sub Process_I_Line(sLine As String)

End Sub

Private Sub Process_H_Line(sLine As String)

End Sub

Private Sub Process_L_Line(sLine As String)

End Sub

Private Sub Process_A_Line(sLine As String)

End Sub
