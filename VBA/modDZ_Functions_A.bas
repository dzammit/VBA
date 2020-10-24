Attribute VB_Name = "modDZ_Functions"
Option Compare Database
Option Explicit

Private Declare Function GetKeyState& Lib "User32" (ByVal nKey As Long)

Function AllUCase()
10  On Error Resume Next
20  If Screen.ActiveForm.Dirty Then
30      Screen.ActiveControl.Text = UCase(Screen.ActiveControl.Text)
40  End If
End Function

Function AutoTab(iLen As Integer)
10  On Error Resume Next
20  If Len(Nz(Screen.ActiveControl.Text, "")) = iLen Then SendKeys "{tab}"
End Function

