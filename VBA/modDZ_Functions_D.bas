Attribute VB_Name = "modDZ_Functions"
Option Compare Database
Option Explicit

Public DZADDON_gCal

Function DaysToString(lDays As Long)
    Dim d, M, Y
10  DaysToString = Format(lDays, "yy") & "y "
End Function

Function DZ_CalcWorkTime(StartDate, EndDate, StartHour, EndHour)
10  On Error GoTo ERR_DZ_CalcWorkTime

    ' Using the StartHour and EndHour as the work time hours,
    ' calculate the total amount of work time used between
    ' StartDate and EndDate.

    Dim NumSecs As Double
    Dim NumDays As Double
    Dim WorkSecs As Double
    Dim NonWorkSecs As Double
    Dim NumWEs As Double

20  If IsNull(StartDate) Or IsNull(EndDate) Then
30      DZ_CalcWorkTime = 0
40  Else
50      If IsNull(StartHour) Or IsNull(EndHour) Then
60          DZ_CalcWorkTime = DateDiff("s", StartDate, EndDate)
70      Else
            ' Calculate amount of non working hours between two
            ' dates.
            ' We use the formula:
            '   NumSecs = Total number of minutes between two dates
            '   NumDays = Total number of days between two dates
            '   WorkSecs = Total working minutes in a day
            '   NonWorkSecs = Total non working minutes in a day
            '   NumWEs = Total number of weekends between two dates
            ' Therefore:
            ' TotalWorkTimeSecs = NumSecs - ((NumDays * NonWorkSecs) +
            '                                (NumWEs * WorkSecs))
            '
80          NumSecs = DateDiff("s", StartDate, EndDate)
90          NumDays = DateDiff("d", StartDate, EndDate)
100         WorkSecs = DateDiff("s", StartHour, EndHour)
110         NonWorkSecs = 86400 - WorkSecs
120         NumWEs = DateDiff("ww", StartDate, EndDate)
130         DZ_CalcWorkTime = NumSecs - ((NumDays * NonWorkSecs) + (NumWEs * WorkSecs * 2))
140     End If
150 End If

EXIT_DZ_CalcWorkTime:
160 Exit Function

ERR_DZ_CalcWorkTime:
170 If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_CalcWorkTime", err) Then Resume EXIT_DZ_CalcWorkTime
180 Resume Next

End Function

Function DZ_Calendar(gDefault)
10  On Error GoTo ErrorHandle

20  DZADDON_gCal = gDefault
30  DoCmd.OpenForm "Calendar", , , , , acDialog
40  DZ_Calendar = DZADDON_gCal

ErrorExit:
50  On Error GoTo 0
60  Exit Function

ErrorHandle:
70  If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_Calendar", err) Then Resume ErrorExit
80  Resume Next

End Function

Function DZ_CalendarClick(sCtrl As Access.Control)
10  On Error GoTo ErrorHandle

20  DZADDON_gCal = sCtrl.Value
30  DoCmd.OpenForm "Calendar", , , , , acDialog
40  sCtrl.Value = DZADDON_gCal

ErrorExit:
50  On Error GoTo 0
60  Exit Function

ErrorHandle:
70  If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_CalendarClick", err) Then Resume ErrorExit
80  Resume Next

End Function

Function DZ_ChangeAccessCaption(Caption As String)
    Dim dbs As DAO.Database
    Dim prp As DAO.Property

10  On Error Resume Next
20  Set dbs = CurrentDb
30  dbs.Properties!AppTitle = Caption$
40  If err = 3270 Then
50      Set prp = dbs.CreateProperty("AppTitle", dbText, Caption$)
60      dbs.Properties.Append prp
70  End If
80  Application.RefreshTitleBar
End Function

Function DZ_ConvSecsToTime(Secs As Double)
10  On Error GoTo ERR_DZ_ConvSecsToTime

    ' Return a string showing the time in seconds displayed as
    ' dd days hh hours mm minutes ss seconds
    Dim dDays As Double
    Dim dHours As Double
    Dim dMinutes As Double
    Dim dSeconds As Double
    Dim sTime As String

20  dDays = Int(Secs / 86400)
30  dHours = Int((Secs - dDays * 86400) / 3600)
40  dMinutes = Int((Secs - dDays * 86400 - dHours * 3600) / 60)
50  dSeconds = Int(Secs - dDays * 86400 - dHours * 3600 - dMinutes * 60)

60  sTime = ""
70  If dDays Then
80      sTime = dDays & " d"
90  End If

100 If dHours Or Len(sTime) Then
110     sTime = sTime & " " & dHours & " h"
120 End If

130 sTime = sTime & " " & dMinutes & " m " & dSeconds & " s"

140 DZ_ConvSecsToTime = sTime

EXIT_DZ_ConvSecsToTime:
150 Exit Function

ERR_DZ_ConvSecsToTime:
160 If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_ConvSecsToTime", err) Then Resume EXIT_DZ_ConvSecsToTime
170 Resume Next

End Function

Function DZ_Count(sStr, sToken As String)
10  On Error GoTo ErrorHandle

    'This counts the number of FIELDS in a string based on the token
    ' ie, ABC-DEF contains 2 FIELDS, so dz_count("ABC-DEF", "-") returns 2.

    Dim WC As Integer
    Dim pos As Integer
20  If Len(sStr) = 0 Then
30      DZ_Count = 0
40      Exit Function
50  End If
60  WC = 1
70  pos = InStr(sStr, sToken)
80  Do While pos > 0
90      WC = WC + 1
100     pos = InStr(pos + 1, sStr, sToken)
110 Loop
120 DZ_Count = WC

ErrorExit:
130 On Error GoTo 0
140 Exit Function

ErrorHandle:
150 If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_Count", err) Then Resume ErrorExit
160 Resume Next

End Function

Function DZ_GetPath(sFile) As String
10  On Error GoTo ErrorHandle

20  DZ_GetPath = Left(sFile, Len(sFile) - Len(DZ_Parse(sFile, "\", DZ_Count(sFile, "\"))))

ErrorExit:
30  On Error GoTo 0
40  Exit Function

ErrorHandle:
50  If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_GetPath", err) Then Resume ErrorExit
60  Resume Next

End Function

Function DZ_GetFile(sFile) As String
10  On Error GoTo ErrorHandle

20  DZ_GetFile = Mid(sFile, InStrRev(sFile, "\") + 1)

ErrorExit:
30  On Error GoTo 0
40  Exit Function

ErrorHandle:
50  If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_GetFile", err) Then Resume ErrorExit
60  Resume Next

End Function

Function DZ_Parse(sStr, sToken, iCnt As Integer)
10  On Error GoTo ErrorHandle

    Const C_QUOTE = """"

    Dim iOffset As Integer
    Dim i As Integer
    Dim iLen As Integer

20  If IsNull(sStr) Then Exit Function

30  iOffset = 0
40  For i = 1 To iCnt - 1
50      If Mid(sStr, iOffset + 1, 1) = C_QUOTE Then
60          iOffset = InStr(iOffset + 2, sStr, C_QUOTE)
70          If iOffset = 0 Then Exit For
80      End If
90      iOffset = InStr(iOffset + 1, sStr, sToken)
100     If iOffset = 0 Then Exit For
        ' Check for Quote Delimiter?
110 Next i

    ' Check for Quote Delimiter?
120 If Mid(sStr, iOffset + 1, 1) = """" Then
130     i = InStr(iOffset + 2, sStr, C_QUOTE)
140     iLen = InStr(i, sStr, sToken)
150 Else
160     iLen = InStr(iOffset + 1, sStr, sToken)
170 End If

180 If iLen > 0 Then
190     iLen = iLen - iOffset - 1
200 Else
210     iLen = Len(sStr)
220 End If
230 If iCnt <> 1 And iOffset = 0 Then
240     DZ_Parse = ""
250 Else
260     DZ_Parse = Mid(sStr, iOffset + 1, iLen)
270 End If

ErrorExit:
280 On Error GoTo 0
290 Exit Function

ErrorHandle:
300 If DZ_ErrorLog("DZADDON.modDZ_Functions.DZ_Parse", err) Then Resume ErrorExit
310 Resume Next

End Function

Function DZ_Range(sFrom, sTo)
    Const C_FMT = "dd/mm/yyyy"
10  On Error Resume Next
20  If Not (IsBlank(sFrom)) Then
30      If Not (IsBlank(sTo)) Then
40          DZ_Range = "Between " & Format(CVDate(sFrom), C_FMT) & " and " & Format(CVDate(sTo), C_FMT)
50      Else
60          DZ_Range = Format(CVDate(sFrom), C_FMT) & " or later"
70      End If
80  Else
90      If Not (IsBlank(sTo)) Then
100         DZ_Range = Format(CVDate(sTo), C_FMT) & " or earlier"
110     Else
120         DZ_Range = "All Dates"
130     End If
140 End If
End Function

Function DZ_RangeSql(sFrom, sTo)
    Const C_FMT = "dd/mm/yyyy"
10  On Error Resume Next
20  If Not (IsBlank(sFrom)) Then
30      If Not (IsBlank(sTo)) Then
40          DZ_RangeSql = " Between " & USDateH(CVDate(sFrom)) & " and " & USDateH(CVDate(sTo))
50      Else
60          DZ_RangeSql = " >= " & USDateH(CVDate(sFrom))
70      End If
80  Else
90      If Not (IsBlank(sTo)) Then
100         DZ_RangeSql = " <= " & USDateH(CVDate(sTo))
110     Else
120         DZ_RangeSql = " > #1/1/1901#"
130     End If
140 End If
End Function

