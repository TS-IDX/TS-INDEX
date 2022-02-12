Attribute VB_Name = "modMyFunctions"

'   ************************************************************
'    ..........................................................
'     Application Name: TechFTP 1.00
'                       Complete FTP Client like CuteFTP.
'     Developer/Programmer: Muktak Joshi (muktak@techprotean.com) - http://www.techprotean.com/muktak
'    ..........................................................
'     Module Name: modMyFunctions
'     Module File: modMyFunctions.bas
'     Module Type: Form
'     Module Description:
'    ..........................................................
'     
'    ..........................................................
'   ************************************************************

Option Explicit
Public AsciiList As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'........................................
'Name: GetRightWord
'Description: 
'........................................
Public Function GetRightWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    Dim pos As Integer
    pos = InStrRev(Sentense, StopChar, -1, vbTextCompare)
    If pos = 0 Then
        GetRightWord = Sentense
    Else
        GetRightWord = Right(Sentense, Len(Sentense) - pos)
    End If

    If RemoveIt = True Then Sentense = Left$(Sentense, pos)

End Function



'........................................
'Name: GetLeftWord
'Description: 
'........................................
Public Function GetLeftWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    Dim pos As Integer
    pos = InStr(1, Sentense, StopChar, vbTextCompare)
    If pos = 0 Then
        GetLeftWord = Sentense
    Else
        GetLeftWord = Left$(Sentense, pos - 1)
    End If

    If RemoveIt = True Then Sentense = Right(Sentense, Len(Sentense) - pos)

End Function



'........................................
'Name: ChangeButState
'Description: 
'........................................
Public Sub ChangeButState(Butt As Object, State As Integer)
    Dim cnt As Control
    If TypeOf Butt Is Form Then
        For Each cnt In Butt.Controls
            If TypeOf cnt Is Label Then
                If cnt.BackColor = &H800080 Then
                    cnt.BackColor = &HC0C0C0
                    cnt.ForeColor = vbBlack
                End If
            End If
        Next
    Else


        If State = 1 Then
            Butt.BackColor = &H800080
            Butt.ForeColor = vbWhite
        Else
            Butt.BackColor = &HC0C0C0
            Butt.ForeColor = vbBlack
        End If

    End If
End Sub

