Attribute VB_Name = "modDZ_Functions"
Option Compare Database
Option Explicit

Public Function CalcAge(vDate1, Optional vdate2 As String = "")
    Dim vMonths, vDays, vYears
    ' Comments  : calculates the age in Years, Months and Days
    ' Parameters:
    '    vDate1 - D.O.B.
    '    vDate2 - Date to calculate age based on
    '    vYears - will hold the Years difference
    '    vMonths - will hold the Months difference
    '    vDays - will hold the Days difference

10  If vdate2 = "" Then vdate2 = Date
20  If Nz(vDate1, "") = "" Then Exit Function

30  vMonths = DateDiff("m", vDate1, vdate2)
40  vDays = DateDiff("d", DateAdd("m", vMonths, vDate1), vdate2)
50  If vDays < 0 Then
        ' wierd way that DateDiff works, fix it here
60      vMonths = vMonths - 1
70      vDays = DateDiff("d", DateAdd("m", vMonths, vDate1), vdate2)
80  End If
90  vYears = vMonths \ 12            ' integer division
100 vMonths = vMonths Mod 12            ' only want leftover less than one year
110 CalcAge = vYears & "y " & vMonths & "m " & vDays & "d"
End Function

Function CheckSave()
10  On Error GoTo ERR_HANDLER
20  CheckSave = False
30  If Screen.ActiveForm.Dirty And Not DZADDON_gIGNOREDIRTY Then
40      Select Case MsgBox("Save Changes ?", vbYesNoCancel, Screen.ActiveForm.Caption)
            Case vbYes
50              CheckSave = True
60          Case vbNo
70              Screen.ActiveForm.Undo
80          Case Else
90              DoCmd.CancelEvent
100     End Select
110 End If
ERR_EXIT:
120 Exit Function

ERR_HANDLER:
130 If DZ_ErrorLog("modSys.CheckSave", err) Then Resume ERR_EXIT
140 Resume Next
End Function

Sub ContinuousUpDown(frm As Access.Form, KeyCode As Integer)
10  On Error GoTo Err_ContinuousUpDown
    'Purpose:   Respond to Up/Down in continuous form, by moving record.
20  Select Case KeyCode
        Case vbKeyUp
30          If ContinuousUpDownOk Then
                'Save any edits
40              If frm.Dirty Then
50                  RunCommand acCmdSaveRecord
60              End If
                'Go previous: error if already there.
70              RunCommand acCmdRecordsGoToPrevious
80              KeyCode = 0    'Destroy the keystroke
90          End If


100     Case vbKeyDown
110         If ContinuousUpDownOk Then
                'Save any edits
120             If frm.Dirty Then
130                 RunCommand acCmdSaveRecord
140             End If
                'Go to the next record, unless at a new record.
150             If Not frm.NewRecord Then
160                 RunCommand acCmdRecordsGoToNext
170             End If
180             KeyCode = 0    'Destroy the keystroke
190         End If
200 End Select

Exit_ContinuousUpDown:
210 Exit Sub


Err_ContinuousUpDown:
220 Select Case err.Number
        Case 2046, 2101    'Already at first record, or save failed.
230         KeyCode = 0
240     Case Else
250         MsgBox err.Description, vbExclamation, "Error " & err.Number
260 End Select
270 Resume Exit_ContinuousUpDown
End Sub

Private Function ContinuousUpDownOk() As Boolean
10  On Error GoTo Err_ContinuousUpDownOk
    'Purpose: Suppress moving up/down a record in a continuous form if:
    '           - control is not in the Detail section, or
    '           - multi-line text box (vertical scrollbar, or
    '                 EnterKeyBehavior is true).
    Dim bDontDoIt As Boolean
    Dim ctl As Access.Control


20  Set ctl = Screen.ActiveControl
30  If ctl.Section = acDetail Then
40      If TypeOf ctl Is TextBox Then
50          bDontDoIt = ((ctl.EnterKeyBehavior) Or (ctl.ScrollBars > 1))
60      End If
70  Else
80      bDontDoIt = True
90  End If

Exit_ContinuousUpDownOk:
100 ContinuousUpDownOk = Not bDontDoIt
110 Set ctl = Nothing
120 Exit Function

Err_ContinuousUpDownOk:
130 If err.Number <> 2474 Then          'There's no active control
140     MsgBox err.Description, vbExclamation, "Error " & err.Number
150 End If
160 Resume Exit_ContinuousUpDownOk
End Function

Function CountCSWords(ByVal pvS, ByVal pvT) As Integer
    'Counts the words in a string that are separated by commas.
    Dim iWC As Integer
    Dim iPos As Integer
10  On Error GoTo CountCSWords_Error

20  If VarType(pvS) <> 8 Or Len(pvS) = 0 Then
30      CountCSWords = 0
40      Exit Function
50  End If
60  iWC = 1
70  iPos = InStr(pvS, pvT)
80  Do While iPos > 0
90      iWC = iWC + 1
100     iPos = InStr(iPos + 1, pvS, pvT)
110 Loop
120 CountCSWords = iWC

CountCSWords_Exit:
130 On Error GoTo 0
140 Exit Function

CountCSWords_Error:

150 If DZ_ErrorLog("DZ_Functions.CountCSWords", err) Then Resume CountCSWords_Exit
160 Resume CountCSWords_Exit

End Function
