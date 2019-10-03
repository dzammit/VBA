Attribute VB_Name = "modDZ_ErrorLog"
Option Compare Database
Option Explicit
'Global gbLOG As Boolean
'Global gbLOGREAD As Boolean

Function DZ_ErrorLog(ProgramName, ErrorNum, Optional sWho As String = "", Optional sType As String = "ERROR", Optional sMsg As String = "")
    Dim sErrMsg
10  sErrMsg = Error
    Dim ErrLogNum As Long
    Dim errorline As Integer
    Dim sMsg2 As String

20  DoCmd.Echo True
30  DoCmd.SetWarnings True
40  errorline = Erl
50  If Not (IsNumeric(ErrorNum)) Then
        ' Custom Error Message
60      sErrMsg = ErrorNum
70      ErrorNum = -1
80  End If

90  On Error Resume Next
100 sMsg2 = GetConfigStr(Trim(str(ErrorNum)))
110 If sMsg2 = "IGNORE" Then
        ' This error has been flagged as being safe to ignore.
120     DZ_ErrorLog = False        ' Always continue on such errors
130     Exit Function
140 End If

150 If sMsg = "" Then sMsg = sMsg2

    Dim myDB As DAO.Database
    Dim myset As DAO.Recordset
160 err = 0
170 Set myDB = CurrentDb()
180 Set myset = myDB.OpenRecordset("ErrorLog", dbOpenDynaset)
190 myset.AddNew
200 myset!Program = ProgramName
210 myset!Program = myset!Program & "|"
220 myset!Program = myset!Program & CodeContextObject.Name
230 myset!Form = Screen.ActiveForm.Name
240 myset!Form = myset!Form & " (Report) " & Screen.ActiveReport.Name

250 myset!Control = Screen.ActiveControl.Name
260 myset!ControlValue = Screen.ActiveControl.ControlSource
270 myset!ControlValue = myset!ControlValue & "=" & Screen.ActiveControl.Value
280 myset!ControlValue = myset!ControlValue & "=" & Screen.ActiveControl.Caption
290 myset!Type = sType & " Line " & errorline
300 If sWho = "" Then sWho = fOSUserName & " on " & fOSMachineName
310 myset!Who = sWho
320 myset!ErrorCode = ErrorNum
330 myset!ErrorDesc = sErrMsg
340 myset!ErrorDesc = myset!ErrorDesc & "|" & sMsg
350 myset!Date = Now
360 ErrLogNum = myset!Counter
370 myset.Update
380 myset.Close

    '       Call UpdateErrorLogDetail(ErrLogNum)
    ' If the error that has occurred is a known error then display the message for it...
390 If Len(sMsg) Then
400     MsgBox sMsg, 48, "Error"
410 Else
420     If Len(Dir("C:\debug.dat")) Then 'Or err Then
430         If MsgBox("DEBUG: " & sType & " has occurred: " & vbCrLf & "'" & sErrMsg & "'" & vbCrLf & " in '" & ProgramName & "', Line " & errorline & vbCrLf & vbCrLf & "Continue executing current function?", 20, "Error") = vbNo Then Stop
440     End If
450     If 0 Then          'GetConfigBool("ERRDSP") Or DevEnv Then
460         If ProgramName = "DataDefinition.AddField" Then Exit Function
470         If sType <> "ERROR" Then Exit Function
480         DZ_ErrorLog = (MsgBox("An error has occurred: " & vbCrLf & "'" & sErrMsg & "'" & vbCrLf & " in '" & ProgramName & "', Line " & errorline & vbCrLf & vbCrLf & "Continue executing current function?", 20, "Error") <> 6)
            'If GetConfigBool("_DEBUG") Then
            '    Stop
            'End If
490         Exit Function
500     End If
510 End If

520 DZ_ErrorLog = False            'GetConfigBool("ERROR") ' True = Don't Continue, False = Continue anyway

EXIT_DZ_ErrorLog:
530 Exit Function

ERR_DZ_ErrorLog:
540 MsgBox "Error logging error! Call David Zammit. Module: DZ_Standard.DZ_ErrorLog - " & Error
550 Resume Next

End Function

Sub Log(sMsg As String, sRoutine As String, Optional sDetail As String = "", Optional iLvl As Integer = 1)
    Const cFRMTABLE = "frmLogTable"
    Const LOGTABLE = "_LOG"
    Dim rs As DAO.Recordset
    Dim void

    On Error Resume Next
    
    Static gbLOGREAD As Boolean
    Static gbLOG As Boolean

    ' If log table exists, save log, otherwise don't bother.
10  If Not gbLOGREAD Then
20      gbLOG = GetConfigBool("LOG")
30      gbLOGREAD = True
40  End If

    ' IF flag for LOG = FALSE then don't log.
50  If gbLOG = False Then Exit Sub

60  err = 0
70  On Error Resume Next
80  void = CurrentDb().TableDefs(LOGTABLE).Name
90  If err Then
100     gbLOG = False   ' Error occurred : save flag as no log wanted.
110     gbLOGREAD = True
        Exit Sub
120 End If

130 Set rs = CurrentDb().OpenRecordset(LOGTABLE)
140 With rs
150     .AddNew
160     !msg = sMsg
170     !Routine = sRoutine
180     !Detail = sDetail
190     !Lvl = iLvl
200     .Update
210     .Close
220 End With
230 Set rs = Nothing
240 'If iLvl >= Forms(cFRMTABLE)!frmFilter Then Forms(cFRMTABLE).Requery
    '    Forms(cFRMTABLE).Repaint
    '    DoEvents
End Sub
