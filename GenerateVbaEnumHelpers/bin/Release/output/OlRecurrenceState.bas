Attribute VB_Name = "wOlRecurrenceState"
Function OlRecurrenceStateFromString(value As String) As OlRecurrenceState
    If IsNumeric(value) Then
        OlRecurrenceStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olApptNotRecurring": OlRecurrenceStateFromString = olApptNotRecurring
        Case "olApptMaster": OlRecurrenceStateFromString = olApptMaster
        Case "olApptOccurrence": OlRecurrenceStateFromString = olApptOccurrence
        Case "olApptException": OlRecurrenceStateFromString = olApptException
    End Select
End Function

Function OlRecurrenceStateToString(value As OlRecurrenceState) As String
    Select Case value
        Case olApptNotRecurring: OlRecurrenceStateToString = "olApptNotRecurring"
        Case olApptMaster: OlRecurrenceStateToString = "olApptMaster"
        Case olApptOccurrence: OlRecurrenceStateToString = "olApptOccurrence"
        Case olApptException: OlRecurrenceStateToString = "olApptException"
    End Select
End Function
