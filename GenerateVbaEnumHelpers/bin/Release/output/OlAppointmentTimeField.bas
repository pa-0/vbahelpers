Attribute VB_Name = "wOlAppointmentTimeField"
Function OlAppointmentTimeFieldFromString(value As String) As OlAppointmentTimeField
    If IsNumeric(value) Then
        OlAppointmentTimeFieldFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olAppointmentTimeFieldNone": OlAppointmentTimeFieldFromString = olAppointmentTimeFieldNone
        Case "olAppointmentTimeFieldStart": OlAppointmentTimeFieldFromString = olAppointmentTimeFieldStart
        Case "olAppointmentTimeFieldEnd": OlAppointmentTimeFieldFromString = olAppointmentTimeFieldEnd
    End Select
End Function

Function OlAppointmentTimeFieldToString(value As OlAppointmentTimeField) As String
    Select Case value
        Case olAppointmentTimeFieldNone: OlAppointmentTimeFieldToString = "olAppointmentTimeFieldNone"
        Case olAppointmentTimeFieldStart: OlAppointmentTimeFieldToString = "olAppointmentTimeFieldStart"
        Case olAppointmentTimeFieldEnd: OlAppointmentTimeFieldToString = "olAppointmentTimeFieldEnd"
    End Select
End Function
