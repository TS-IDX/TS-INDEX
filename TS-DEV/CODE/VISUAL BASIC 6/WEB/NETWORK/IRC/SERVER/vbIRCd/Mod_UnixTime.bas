Attribute VB_Name = "Mod_UnixTime"
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


Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long




Public Type tm
    tm_sec As Long ' seconds (0 - 59)
    tm_min As Long ' minutes (0 - 59)
    tm_hour As Long ' hours (0 - 23)
    tm_mday As Long ' day of month (1 - 31)
    tm_mon As Long ' month of year (0 - 11)
    tm_year As Long ' year - 1900
    tm_wday As Long ' day of week (Sunday = 0), Not used
    tm_yday As Long ' day of year (0 - 365), Not used
    tm_isdst As Long ' Daylight Savings Time (0, 1), Not used
End Type

Public Function sUnixDate(ByVal lValue As Long) As String
    ' Now for the LocalTime function. Take
    '     the long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string YYYY/MM/DD H
    '     H:MM:SS
    Dim lSecPerYear
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Dim Hour As Long
    Dim Minute As Long
    Dim Second As Long
    Dim Temp As Long
    ' [0] = normal year, [1] = leap year
    lSecPerYear = Array(31536000, 31622400)
    lSecPerDay = 86400 ' 60*60*24
    lSecPerHour = 3600 ' 60 * 60
    Year = 70 ' starting point
    ' Calculate the year


    Do While (lValue > 0)
        Temp = isLeapYear(Year)


        If (lValue - lSecPerYear(Temp)) > 0 Then
            lValue = lValue - lSecPerYear(Temp)
            Year = Year + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Year = " & Year
    ' Calculate the month
    Month = 1


    Do While (lValue > 0)
        Temp = secsInMonth(Year, Month)


        If (lValue - Temp) > 0 Then
            lValue = lValue - Temp
            Month = Month + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Month = " & Month
    ' Now calculate day
    Day = 1


    Do While (lValue > 0)


        If (lValue - lSecPerDay) > 0 Then
            lValue = lValue - lSecPerDay
            Day = Day + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Day = " & Day
    ' Now calculate Hour
    Hour = 0


    Do While (lValue > 0)


        If (lValue - lSecPerHour) > 0 Then
            lValue = lValue - lSecPerHour
            Hour = Hour + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Hour = " & Hour
    Minute = Int(lValue / 60)
    'Debug.Print "Minute = " & Minute
    Second = lValue Mod 60
    'Debug.Print "Second = " & Second
    ' Standard date format is YYYY/MM/DD HH:
    '     MM:SS
    'If Year < 100 Then
    Year = Year + 1900
    sUnixDate = Month & "/" & Day & "/" & Year & ", " & Hour & ":" & Minute & ":" & Second
End Function

Private Function isLeapYear(Year As Long) As Integer
    ' Determine if given ANSI datetime struc
    '     t represents a leap year
    ' Private function: assumes valid parame
    '     ters
    Dim nYear As Integer
    Dim nIsLeap As Integer
    nYear = Year + 1900


    If ((nYear Mod 4 = 0 And Not (nYear Mod 100)) Or nYear Mod 400 = 0) Then
        nIsLeap = 1 ' its a leap year
    Else
        nIsLeap = 0 ' Not a leap year
    End If
    isLeapYear = nIsLeap
End Function

Private Function secsInMonth(Year As Long, Month As Long) As Long
    ' Return total number of seconds in the
    '     given month
    ' Private function: assumes valid parame
    '     ters
    Dim lResult As Long
    Dim lSecPerMonth
    lSecPerMonth = Array(2678400, 2419200, 2678400, 2592000, _
    2678400, 2592000, 2678400, 2678400, _
    2592000, 2678400, 2592000, 2678400)
    ' Compute result
    lResult = lSecPerMonth(Month - 1)


    If (isLeapYear(Year) And Month = 2) Then
        lResult = lResult + 86400 ' its February In a leap year
    End If
    secsInMonth = lResult
End Function


Private Function secsInYears(Year As Long) As Double
    ' Return sum of seconds for years since
    '     Jan 1, 1970 00:00
    ' up to but excluding the given year.
    ' Private function: assumes valid parame
    '     ters
    Dim lResult As Long
    Dim thisYear As Long
    Dim Temp As Long
    lResult = 0
    ' 0 = normal year, 1 = leap year
    Dim lSecPerYear
    lSecPerYear = Array(31536000, 31622400)


    If (Year > 97) Then
        ' shorten summation iterations for typic
        '     al cases
        lResult = 883612800 ' seconds To Jan 1,1998 00:00:00
        thisYear = 98
    Else
        ' sum all years since 1970
        thisYear = 70
    End If
    ' Sum total seconds since Jan 1, 1970 00
    '     :00


    While (thisYear < Year)
        'for ( ; thisYear < year; thisYear++
        '     )
        Temp = isLeapYear(thisYear)
        lResult = lResult + lSecPerYear(Temp)
        thisYear = thisYear + 1
    Wend
    secsInYears = lResult
End Function


Function GetLocalTZ(Optional ByRef strTZName As String) As Long
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    lngResult = GetTimeZoneInformation&(objTimeZone)


    Select Case lngResult
        Case 0&, 1& 'use standard time
        GetLocalTZ = -(objTimeZone.Bias + objTimeZone.StandardBias) * 60 'into minutes


        For i = 0 To 31
            If objTimeZone.StandardName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.StandardName(i))
        Next
        Case 2& 'use daylight savings time
        GetLocalTZ = -(objTimeZone.Bias + objTimeZone.DaylightBias) * 60 'into minutes


        For i = 0 To 31
            If objTimeZone.DaylightName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
        Next
    End Select
End Function

Function GetTime() As Double
Dim TheDate As Date 'target date
Dim iResult As Long
Dim SecondsToTarget As Long
    'set target date.
    TheDate = "01/01/1970"
    'compute # of seconds left to target date
    SecondsToTarget = DateDiff("s", Now, TheDate)
    'iResult = (GetLocalTZ / 30) * 30
    GetTime = Mid(SecondsToTarget, 2) - GetLocalTZ
    '6 * 6 = 36 | 60 * 60 = 3600
End Function

