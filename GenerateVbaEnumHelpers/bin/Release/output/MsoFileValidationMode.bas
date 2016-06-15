Attribute VB_Name = "wMsoFileValidationMode"
Function MsoFileValidationModeFromString(value As String) As MsoFileValidationMode
    If IsNumeric(value) Then
        MsoFileValidationModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFileValidationDefault": MsoFileValidationModeFromString = msoFileValidationDefault
        Case "msoFileValidationSkip": MsoFileValidationModeFromString = msoFileValidationSkip
    End Select
End Function

Function MsoFileValidationModeToString(value As MsoFileValidationMode) As String
    Select Case value
        Case msoFileValidationDefault: MsoFileValidationModeToString = "msoFileValidationDefault"
        Case msoFileValidationSkip: MsoFileValidationModeToString = "msoFileValidationSkip"
    End Select
End Function
