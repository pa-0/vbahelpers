Attribute VB_Name = "wOlRecurrenceType"
Function OlRecurrenceTypeFromString(value As String) As OlRecurrenceType
    If IsNumeric(value) Then
        OlRecurrenceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olRecursDaily": OlRecurrenceTypeFromString = olRecursDaily
        Case "olRecursWeekly": OlRecurrenceTypeFromString = olRecursWeekly
        Case "olRecursMonthly": OlRecurrenceTypeFromString = olRecursMonthly
        Case "olRecursMonthNth": OlRecurrenceTypeFromString = olRecursMonthNth
        Case "olRecursYearly": OlRecurrenceTypeFromString = olRecursYearly
        Case "olRecursYearNth": OlRecurrenceTypeFromString = olRecursYearNth
    End Select
End Function

Function OlRecurrenceTypeToString(value As OlRecurrenceType) As String
    Select Case value
        Case olRecursDaily: OlRecurrenceTypeToString = "olRecursDaily"
        Case olRecursWeekly: OlRecurrenceTypeToString = "olRecursWeekly"
        Case olRecursMonthly: OlRecurrenceTypeToString = "olRecursMonthly"
        Case olRecursMonthNth: OlRecurrenceTypeToString = "olRecursMonthNth"
        Case olRecursYearly: OlRecurrenceTypeToString = "olRecursYearly"
        Case olRecursYearNth: OlRecurrenceTypeToString = "olRecursYearNth"
    End Select
End Function
