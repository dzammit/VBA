Attribute VB_Name = "modDZ_Functions"
Option Compare Database
Option Explicit

Function fConcatChild(strChildTable As String, _
                      strIDName As String, _
                      strFldConcat As String, _
                      strIdType As String, _
                      varIDvalue As Variant, _
                      Optional sSeperator As String = ", ") _
                      As String
    'Returns a field from the Many table of a 1:M relationship
    'in a semi-colon separated format.
    '
    'Usage Examples:
    '   ?fConcatChild("Order Details", "OrderID", "Quantity", _
        "Long", 10255)
    'Where  Order Details = Many side table
    '       OrderID       = Primary Key of One side table
    '       Quantity      = Field name to concatenate
    '       Long          = Dataprivate type of Primary Key of One Side Table
    '       10255         = Value on which return concatenated Quantity
    '
    Dim db As DAO.Database          'Replaced by OfficeConverter 8.0.1 on line 130 ' original =           Dim db As Database
    Dim rs As DAO.Recordset          'Replaced by OfficeConverter 8.0.1 on line 131 ' original =           Dim rs As Recordset
    Dim varConcat As Variant
    Dim strCriteria As String
    Dim strSQL As String
10  On Error GoTo Err_fConcatChild

20  varConcat = Null
30  Set db = CurrentDb
40  strSQL = "Select [" & strFldConcat & "] From [" & strChildTable & "]"
50  strSQL = strSQL & " Where "

60  Select Case strIdType
        Case "String":
70          strSQL = strSQL & "[" & strIDName & "] = '" & varIDvalue & "'"
80      Case "Long", "Integer", "Double"
            'AutoNumber is private type Long
90          strSQL = strSQL & "[" & strIDName & "] = " & varIDvalue
100     Case Else
110         GoTo Err_fConcatChild
120 End Select

130 Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    'Are we sure that 'sub' records exist
140 With rs
150     If .RecordCount <> 0 Then
            'start concatenating records
160         Do While Not rs.EOF
170             varConcat = varConcat & rs(strFldConcat) & sSeperator
180             .MoveNext
190         Loop
200     End If
210 End With

    'That's it... you should have a concatenated string now
    'Just Trim the trailing ;
    If Len(sSeperator) > 0 Then
220     fConcatChild = Left(varConcat, Len(varConcat) - Len(sSeperator))
    Else
        fConcatChild = varConcat
    End If
Exit_fConcatChild:
230 Set rs = Nothing
240 Set db = Nothing
250 Exit Function
Err_fConcatChild:
260 Resume Exit_fConcatChild
End Function
