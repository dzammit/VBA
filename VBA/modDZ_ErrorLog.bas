Attribute VB_Name = "modDZ_ErrorLog"
Option Compare Database
Option Explicit
'Global gbLOG As Boolean
'Global gbLOGREAD As Boolean

'---------------------------------------------------------------------------------------
' Procedure : DZ_ErrorLog
' Author    : David Zammit
' Date      : Unknown
' Purpose   : Logs errors to the ErrorLog table and provides options for user notification.
' Parameters:
'   ProgramName (String): The name of the program or module where the error occurred.
'   ErrorNum (Long/String): The error number or a custom error message string.
'   sWho (Optional String): Identifier for the user or process that encountered the error. Defaults to OS username and machine name.
'   sType (Optional String): The type of error (e.g., "ERROR", "WARNING"). Defaults to "ERROR".
'   sMsg (Optional String): An additional custom message to be logged alongside the error.
' Returns   : Boolean - Typically False, indicating whether to continue execution (depends on configuration).
'---------------------------------------------------------------------------------------
Function DZ_ErrorLog(ProgramName, ErrorNum, Optional sWho As String = "", Optional sType As String = "ERROR", Optional sMsg As String = "")
    ' --- Variable Declarations ---
    Dim sErrMsg As String         ' Stores the description of the error.
10  sErrMsg = Error             ' Capture the built-in error message immediately.
    Dim ErrLogNum As Long       ' Stores the ID of the error log entry.
    Dim errorline As Integer    ' Stores the line number where the error occurred.
    Dim sMsg2 As String         ' Stores a configured message for the error number.

    ' --- Initial Setup and Error Handling ---
20  DoCmd.Echo True             ' Ensure screen updating is enabled.
30  DoCmd.SetWarnings True      ' Ensure system warnings are enabled.
40  errorline = Erl             ' Get the line number where the error occurred.

    ' Check if ErrorNum is a custom string message rather than a numeric error code.
50  If Not (IsNumeric(ErrorNum)) Then
        ' Custom Error Message provided
60      sErrMsg = ErrorNum      ' Use the provided string as the error message.
70      ErrorNum = -1           ' Set a generic error number for custom messages.
80  End If

    ' --- Check for Ignorable Errors ---
90  On Error Resume Next        ' Temporarily ignore errors during configuration string retrieval.
100 sMsg2 = GetConfigStr(Trim(str(ErrorNum))) ' Attempt to get a pre-configured message for this error number.
110 If sMsg2 = "IGNORE" Then
        ' This error has been flagged as being safe to ignore in the configuration.
120     DZ_ErrorLog = False        ' Always continue execution on such errors.
130     Exit Function           ' Stop further processing of this error.
140 End If
    ' --- Restore Normal Error Handling ---
    On Error GoTo ERR_DZ_ErrorLog ' Reinstate standard error handling for the rest of the function.


150 If sMsg = "" Then sMsg = sMsg2 ' If no specific message was passed in, use the configured message.

    ' --- Database Operations: Logging the Error ---
    Dim myDB As DAO.Database    ' DAO Database object.
    Dim myset As DAO.Recordset  ' DAO Recordset object.
160 err = 0                     ' Clear any previous error value.
170 Set myDB = CurrentDb()      ' Get a reference to the current database.
180 Set myset = myDB.OpenRecordset("ErrorLog", dbOpenDynaset) ' Open the ErrorLog table.
190 myset.AddNew                ' Create a new record to log the error.
    ' Populate the fields of the new error log record.
200 myset!Program = ProgramName
210 myset!Program = myset!Program & "|" ' Append a separator.
220 myset!Program = myset!Program & CodeContextObject.Name ' Append the name of the object (e.g., Form, Report, Module) where the error occurred.
230 myset!Form = Screen.ActiveForm.Name ' Log the name of the active form.
240 myset!Form = myset!Form & " (Report) " & Screen.ActiveReport.Name ' Append active report name if any.

250 myset!Control = Screen.ActiveControl.Name ' Log the name of the active control.
260 myset!ControlValue = Screen.ActiveControl.ControlSource ' Log the control's source.
270 myset!ControlValue = myset!ControlValue & "=" & Screen.ActiveControl.Value ' Append the control's value.
280 myset!ControlValue = myset!ControlValue & "=" & Screen.ActiveControl.Caption ' Append the control's caption.
290 myset!Type = sType & " Line " & errorline ' Log the error type and line number.
300 If sWho = "" Then sWho = fOSUserName & " on " & fOSMachineName ' If sWho is not provided, get OS username and machine name.
310 myset!Who = sWho
320 myset!ErrorCode = ErrorNum
330 myset!ErrorDesc = sErrMsg       ' Log the primary error description.
340 myset!ErrorDesc = myset!ErrorDesc & "|" & sMsg ' Append any additional message.
350 myset!Date = Now                ' Log the current date and time.
360 ErrLogNum = myset!Counter       ' Get the unique ID of this error log entry.
370 myset.Update                  ' Save the new record to the ErrorLog table.
380 myset.Close                   ' Close the recordset.

    ' --- User Notification and Execution Control ---
    ' The line below was commented out, possibly for future use or was part of a removed feature.
    '       Call UpdateErrorLogDetail(ErrLogNum)

    ' If a specific message (either passed in or from config) exists, display it.
390 If Len(sMsg) Then
400     MsgBox sMsg, 48, "Error" ' Display the message in a simple message box.
410 Else
        ' If no specific message, check for debug mode or other conditions.
420     If Len(Dir("C:\debug.dat")) Then ' Check if a "debug.dat" file exists (acts as a debug flag). 'Or err Then
            ' In debug mode, offer to stop execution.
430         If MsgBox("DEBUG: " & sType & " has occurred: " & vbCrLf & "'" & sErrMsg & "'" & vbCrLf & " in '" & ProgramName & "', Line " & errorline & vbCrLf & vbCrLf & "Continue executing current function?", 20, "Error") = vbNo Then Stop
440     End If

        ' This block is currently disabled (If 0 Then).
        ' It seems to be related to configuration settings "ERRDSP" or a "DevEnv" flag.
450     If 0 Then          'GetConfigBool("ERRDSP") Or DevEnv Then
460         If ProgramName = "DataDefinition.AddField" Then Exit Function ' Specific exit condition for a particular program name.
470         If sType <> "ERROR" Then Exit Function ' Only proceed for "ERROR" type if this block were active.
            ' Display a message box asking the user whether to continue.
            ' The result of the MsgBox (vbYes = 6, vbNo = 7) would determine the return value of DZ_ErrorLog.
480         DZ_ErrorLog = (MsgBox("An error has occurred: " & vbCrLf & "'" & sErrMsg & "'" & vbCrLf & " in '" & ProgramName & "', Line " & errorline & vbCrLf & vbCrLf & "Continue executing current function?", 20, "Error") <> 6)
            'If GetConfigBool("_DEBUG") Then ' Another commented out debug check.
            '    Stop
            'End If
490         Exit Function ' Exit after handling if this block were active.
500     End If
510 End If

    ' --- Determine Return Value (Continue Execution or Not) ---
    ' The return value is typically False, meaning "continue execution".
    ' This might have been intended to be configurable via "GetConfigBool("ERROR")".
    ' True would mean "Don't Continue", False means "Continue anyway".
520 DZ_ErrorLog = False            'GetConfigBool("ERROR")

EXIT_DZ_ErrorLog:
    ' Standard exit point for the function.
530 Exit Function

ERR_DZ_ErrorLog:
    ' Error handler for errors occurring *within* this error logging function itself.
540 MsgBox "Error logging error! Call David Zammit. Module: DZ_Standard.DZ_ErrorLog - " & Error
550 Resume Next ' Resume execution at the next line (can be problematic, might skip important cleanup).

End Function

'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Unknown (likely David Zammit based on DZ_ErrorLog)
' Date      : Unknown
' Purpose   : Writes a general log message to the _LOG table if logging is enabled
'             in the application's configuration.
' Parameters:
'   sMsg (String): The main message to be logged.
'   sRoutine (String): The name of the routine or procedure calling this log function.
'   sDetail (Optional String): Additional details for the log entry. Defaults to an empty string.
'   iLvl (Optional Integer): A level indicator for the log entry (e.g., for severity or verbosity). Defaults to 1.
'---------------------------------------------------------------------------------------
Sub Log(sMsg As String, sRoutine As String, Optional sDetail As String = "", Optional iLvl As Integer = 1)
    ' --- Constants ---
    Const cFRMTABLE = "frmLogTable" ' Constant for a form name, likely used to display logs (currently commented out).
    Const LOGTABLE = "_LOG"         ' The name of the database table where logs are stored.

    ' --- Variable Declarations ---
    Dim rs As DAO.Recordset         ' DAO Recordset object for interacting with the LOGTABLE.
    Dim void As Variant             ' Used to temporarily store the result of checking table existence.

    ' --- Error Handling Setup ---
    On Error Resume Next            ' Basic error handling: if an error occurs, execution continues on the next line.
                                    ' This is used throughout the sub, which can hide issues.
    
    ' --- Static Variables for Configuration Cache ---
    ' These static variables cache the logging configuration to avoid repeated calls to GetConfigBool.
    Static gbLOGREAD As Boolean     ' Flag indicating if the "LOG" configuration has been read.
    Static gbLOG As Boolean         ' Flag indicating if logging is enabled (True) or disabled (False).

    ' --- Read Logging Configuration (if not already read) ---
    ' This block executes only once per session (or until the module is reset)
    ' to determine if logging should be performed.
10  If Not gbLOGREAD Then
20      gbLOG = GetConfigBool("LOG") ' Read the "LOG" setting from the application's configuration.
30      gbLOGREAD = True             ' Mark the configuration as read.
40  End If

    ' --- Check if Logging is Enabled ---
    ' If the "LOG" configuration is False, exit the subroutine immediately.
50  If gbLOG = False Then Exit Sub

    ' --- Verify Log Table Existence ---
60  err = 0                         ' Clear any pending error code.
70  On Error Resume Next            ' Temporarily set error handling to resume next for the table check.
    ' Attempt to access the Name property of the LOGTABLE. If the table doesn't exist, this will raise an error.
80  void = CurrentDb().TableDefs(LOGTABLE).Name
90  If err Then                     ' Check if an error occurred (i.e., table not found).
        ' If the LOGTABLE doesn't exist, disable logging for the rest of the session to prevent further errors.
100     gbLOG = False               ' Set logging flag to False.
110     gbLOGREAD = True            ' Ensure gbLOGREAD is true so this check isn't re-attempted.
        Exit Sub                    ' Exit the subroutine.
120 End If
    ' It's generally better to restore specific error handling (e.g., On Error GoTo 0 or a specific handler)
    ' after a targeted On Error Resume Next, but this sub relies on it broadly.

    ' --- Write Log Entry to Table ---
130 Set rs = CurrentDb().OpenRecordset(LOGTABLE) ' Open the _LOG table.
140 With rs
150     .AddNew                     ' Create a new record.
        ' Populate the fields of the new log record.
160     !msg = sMsg                 ' Main log message.
170     !Routine = sRoutine         ' Routine where the log was called from.
180     !Detail = sDetail           ' Additional details.
190     !Lvl = iLvl                 ' Log level.
200     .Update                     ' Save the new record.
210     .Close                      ' Close the recordset.
220 End With
230 Set rs = Nothing                ' Release the recordset object.

    ' --- Commented-Out UI Update Logic ---
    ' The following lines are commented out. They likely were intended to refresh a
    ' log display form (frmLogTable) if the newly added log entry's level (iLvl)
    ' met a certain filter criterion on that form.
240 'If iLvl >= Forms(cFRMTABLE)!frmFilter Then Forms(cFRMTABLE).Requery
    '    Forms(cFRMTABLE).Repaint
    '    DoEvents
End Sub
