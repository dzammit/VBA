Attribute VB_Name = "modDZ_Functions"
Option Compare Database
Option Explicit


Function GetConfigBool(ConfigOpt As String)
10  On Error Resume Next

    ' If 'ConfigOpt' doesn't exist we default to FALSE

    Dim vStr As Variant

20  GetConfigBool = False

30  If DCount("*", "Configuration", "ID = '" & ConfigOpt & "'") <> 0 Then
40      vStr = ELookup("NonOleObject", "Configuration", "ID = '" & ConfigOpt & "'")
50  Else
60      If TableExists(C_CS) Then vStr = ELookup("NonOleObject", "Configuration_Software", "ID = '" & ConfigOpt & "'")
70  End If

80  GetConfigBool = Nz(Eval(vStr), False)

End Function

Function GetConfigStr(ConfigOpt As String)
    If Not CheckValid Then Exit Function
10  On Error Resume Next

    Dim vStr As Variant

20  If DCount("*", "Configuration", "ID = '" & ConfigOpt & "'") <> 0 Then
30      If err <> 0 Then
40          If TableExists(C_CS) Then vStr = ELookup("NonOleObject", C_CS, "ID = '" & ConfigOpt & "'")
50      Else
60          vStr = ELookup("NonOleObject", "Configuration", "ID = '" & ConfigOpt & "'")
70      End If
80  Else
90      If TableExists(C_CS) Then vStr = ELookup("NonOleObject", "Configuration_Software", "ID = '" & ConfigOpt & "'")
100 End If

110 If IsNull(vStr) Then
120     GetConfigStr = ""
130 Else
140     GetConfigStr = vStr
150 End If

End Function

Function GetCSWord(ByVal pvS, ByVal pvT, pIndx As Integer)
    'Returns the nth word in a specific field.
    Dim iWC As Integer
    Dim iCount As Integer
    Dim iSPos As Integer
    Dim iEPos As Integer
    Dim iTokenLen As Byte

10  On Error GoTo GetCSWord_Error

20  iWC = CountCSWords(pvS, pvT)
30  If pIndx < 1 Or pIndx > iWC Then
40      GetCSWord = Null
50      Exit Function
60  End If
70  iCount = 1
80  iSPos = 1
    iTokenLen = Len(pvT)
90  For iCount = 2 To pIndx
100     iSPos = InStr(iSPos, pvS, pvT) + iTokenLen
110 Next iCount
120 iEPos = InStr(iSPos, pvS & pvT, pvT) - 1
130 If iEPos <= 0 Then
140     If pIndx > 1 Then
150         iEPos = Len(pvS)
160     Else
170         GetCSWord = ""
180         Exit Function
190     End If
200 End If
210 GetCSWord = Trim(Mid(pvS, iSPos, iEPos - iSPos + 1))

GetCSWord_Exit:
220 On Error GoTo 0
230 Exit Function

GetCSWord_Error:

240 If DZ_ErrorLog("DZ_Functions.GetCSWord", err) Then Resume GetCSWord_Exit
250 Resume GetCSWord_Exit

End Function

Function GetElement(psStr As String, psElement As String, Optional psDelim1 As String = ";", Optional psDelim2 As String = "=")
    'USAGE:
    'GetElement("T=1;B=2", "B") returns 2
    Dim i As Integer
    Dim sSubStr As String

10  On Error GoTo GetPWD_Error

20  For i = 1 To CountCSWords(psStr, psDelim1)
30      sSubStr = GetCSWord(psStr, psDelim1, i)
40      If Left(sSubStr, Len(psElement) + 1) = psElement & psDelim2 Then
50          GetElement = GetCSWord(sSubStr, psDelim2, 2)
60          Exit For
70      End If
80  Next i

GetPWD_Exit:
90  On Error GoTo 0
100 Exit Function

GetPWD_Error:
110 If DZ_ErrorLog("DZ_Functions.GetElement", err) Then Resume GetPWD_Exit
120 Resume Next
End Function

Function GetFormValue(Optional sFrm As String, Optional sField As String)
    ' In case the form is NOT loaded, it will return NULL (without errors)
    On Error Resume Next
    If IsMissing(sFrm) Then sFrm = Screen.ActiveForm.Name
    If IsMissing(sField) Then sField = Screen.ActiveControl.Name
    If Nz(sFrm, "") = "" Then sFrm = Screen.ActiveForm.Name
    If Nz(sField, "") = "" Then sField = Screen.ActiveControl.Name
    
    GetFormValue = Forms(sFrm)(sField).Value

End Function

Public Function GetParentCtl(frm As Access.Form) As Control
    Dim ctl As Access.Control
10  On Error GoTo GetSubTag_Error
20  Set GetParentCtl = Null
30  For Each ctl In frm.Parent.Controls
40      With ctl
50          If .ControlType = acTextBox Or _
               .ControlType = acCommandButton Or _
               .ControlType = acCheckBox Or _
               .ControlType = acListBox Or _
               .ControlType = acComboBox Then
60              If .Enabled And .Visible Then
70                  Set GetParentCtl = ctl
80                  Exit For
90              End If
100         End If
110     End With
120 Next

GetSubTag_Exit:
130 On Error GoTo 0
140 Exit Function

GetSubTag_Error:
150 If DZ_ErrorLog("DZ_Functions.GetSubTag", err) Then Resume GetSubTag_Exit
160 Resume Next
End Function

Public Function GetSubCtl(frm As Access.Form) As Control
    Dim ctl As Access.Control
10  On Error GoTo GetSubTag_Error

20  For Each ctl In frm.Parent.Controls
30      With ctl
40          If .ControlType = acSubform Then
50              If .SourceObject = frm.Name Then
60                  Set GetSubCtl = ctl
70                  Exit For
80              End If
90          End If
100     End With
110 Next

GetSubTag_Exit:
120 On Error GoTo 0
130 Exit Function

GetSubTag_Error:
140 If DZ_ErrorLog("DZ_Functions.GetSubTag", err) Then Resume GetSubTag_Exit
150 Resume Next
End Function

Public Function GetSubTag(frm As Access.Form)
    Dim ctl As Access.Control
10  On Error GoTo GetSubTag_Error

20  For Each ctl In frm.Parent.Controls
30      With ctl
40          If .ControlType = acSubform Then
50              If .SourceObject = frm.Name Then
60                  GetSubTag = .Tag
70                  Exit For
80              End If
90          End If
100     End With
110 Next

GetSubTag_Exit:
120 On Error GoTo 0
130 Exit Function

GetSubTag_Error:
140 If DZ_ErrorLog("DZ_Functions.GetSubTag", err) Then Resume GetSubTag_Exit
150 Resume Next
End Function

Function GetSysName()
10  On Error Resume Next
20  If IsNull(DZADDON_gSysName) Then
30      DZADDON_gSysName = GetConfigStr("SYSNAM")
40  End If
50  If Len(DZADDON_gSysName) = 0 Then
60      DZADDON_gSysName = GetConfigStr("SYSNAM")
70  End If

80  GetSysName = DZADDON_gSysName
End Function

