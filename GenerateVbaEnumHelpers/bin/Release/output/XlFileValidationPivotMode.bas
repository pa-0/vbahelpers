Attribute VB_Name = "wXlFileValidationPivotMode"
Function XlFileValidationPivotModeFromString(value As String) As XlFileValidationPivotMode
    If IsNumeric(value) Then
        XlFileValidationPivotModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFileValidationPivotDefault": XlFileValidationPivotModeFromString = xlFileValidationPivotDefault
        Case "xlFileValidationPivotRun": XlFileValidationPivotModeFromString = xlFileValidationPivotRun
        Case "xlFileValidationPivotSkip": XlFileValidationPivotModeFromString = xlFileValidationPivotSkip
    End Select
End Function

Function XlFileValidationPivotModeToString(value As XlFileValidationPivotMode) As String
    Select Case value
        Case xlFileValidationPivotDefault: XlFileValidationPivotModeToString = "xlFileValidationPivotDefault"
        Case xlFileValidationPivotRun: XlFileValidationPivotModeToString = "xlFileValidationPivotRun"
        Case xlFileValidationPivotSkip: XlFileValidationPivotModeToString = "xlFileValidationPivotSkip"
    End Select
End Function
