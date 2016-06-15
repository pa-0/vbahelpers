Attribute VB_Name = "wXlThreadMode"
Function XlThreadModeFromString(value As String) As XlThreadMode
    If IsNumeric(value) Then
        XlThreadModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlThreadModeAutomatic": XlThreadModeFromString = xlThreadModeAutomatic
        Case "xlThreadModeManual": XlThreadModeFromString = xlThreadModeManual
    End Select
End Function

Function XlThreadModeToString(value As XlThreadMode) As String
    Select Case value
        Case xlThreadModeAutomatic: XlThreadModeToString = "xlThreadModeAutomatic"
        Case xlThreadModeManual: XlThreadModeToString = "xlThreadModeManual"
    End Select
End Function
