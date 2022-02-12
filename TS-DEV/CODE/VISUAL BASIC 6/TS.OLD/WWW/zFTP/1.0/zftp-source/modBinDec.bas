Attribute VB_Name = "modBinDec"

'   ************************************************************
'    ..........................................................
'     Application Name: TechFTP 1.00
'                       Complete FTP Client like CuteFTP.
'     Developer/Programmer: Muktak Joshi (muktak@techprotean.com) - http://www.techprotean.com/muktak
'    ..........................................................
'     Module Name: modBinDec
'     Module File: modBinDec.bas
'     Module Type: Form
'     Module Description:
'    ..........................................................
'     
'    ..........................................................
'   ************************************************************


Option Explicit



'........................................
'Name: Bin2Dec
'Description: 
'........................................
Public Function Bin2Dec(ByVal sBin As String) As Long
    Dim i As Integer


    For i = 1 To Len(sBin)
        Bin2Dec = Bin2Dec + CLng(CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1))
    Next i
End Function



'........................................
'Name: Bin2Hex
'Description: 
'........................................
Public Function Bin2Hex(ByVal sBin As String) As String
    Dim i As Integer
    Dim nDec As Long
    sBin = String(4 - Len(sBin) Mod 4, "0") & sBin 'Add zero To complete Byte


    For i = 1 To Len(sBin)
        nDec = nDec + CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1)
    Next i
    Bin2Hex = Hex(nDec)
    If Len(Bin2Hex) Mod 2 = 1 Then Bin2Hex = "0" & Bin2Hex
End Function



'........................................
'Name: Dec2Bin
'Description: 
'........................................
Public Function Dec2Bin(ByVal nDec As Integer) As String
    'This function is the same then Hex2Bin,
    '     but it has been copied to speed up proce
    '     ss
    Dim i As Integer
    Dim j As Integer
    Dim sHex As String
    Const HexChar As String = "0123456789ABCDEF"

    sHex = Hex(nDec) 'That the only part that is different


    For i = 1 To Len(sHex)
        nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1


        For j = 3 To 0 Step -1
            Dec2Bin = Dec2Bin & nDec \ 2 ^ j
            nDec = nDec Mod 2 ^ j
        Next j
    Next i
    'Remove the first unused 0
    i = InStr(1, Dec2Bin, "1")
    If i <> 0 Then Dec2Bin = Mid(Dec2Bin, i)
End Function



'........................................
'Name: Hex2Bin
'Description: 
'........................................
Public Function Hex2Bin(ByVal sHex As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim nDec As Long
    Const HexChar As String = "0123456789ABCDEF"



    For i = 1 To Len(sHex)
        nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1


        For j = 3 To 0 Step -1
            Hex2Bin = Hex2Bin & nDec \ 2 ^ j
            nDec = nDec Mod 2 ^ j
        Next j
    Next i
    'Remove the first unused 0
    i = InStr(1, Hex2Bin, "1")
    If i <> 0 Then Hex2Bin = Mid(Hex2Bin, i)
End Function



'........................................
'Name: Hex2Dec
'Description: 
'........................................
Public Function Hex2Dec(ByVal sHex As String) As Long
    Dim i As Integer
    Dim nDec As Long
    Const HexChar As String = "0123456789ABCDEF"



    For i = Len(sHex) To 1 Step -1
        nDec = nDec + (InStr(1, HexChar, Mid(sHex, i, 1)) - 1) * 16 ^ (Len(sHex) - i)
    Next i
    Hex2Dec = CStr(nDec)
End Function



'........................................
'Name: HiWord
'Description: 
'........................................
Public Function HiWord(ByVal DWord As Long) As Long
    HiWord = (DWord \ 65536) And &HFFFF
End Function



'........................................
'Name: LoWord
'Description: 
'........................................
Public Function LoWord(ByVal DWord As Long) As Long
    LoWord = DWord And &HFFFF
End Function



'........................................
'Name: DWord
'Description: 
'........................................
Public Function DWord(ByVal HiWord As Long, ByVal LoWord As Long) As Long
    DWord = ((LoWord And 65536) Or ((HiWord And 65536) * 65536))
End Function


