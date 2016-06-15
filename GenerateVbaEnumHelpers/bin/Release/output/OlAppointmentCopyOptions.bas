Attribute VB_Name = "wOlAppointmentCopyOptions"
Function OlAppointmentCopyOptionsFromString(value As String) As OlAppointmentCopyOptions
    If IsNumeric(value) Then
        OlAppointmentCopyOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olPromptUser": OlAppointmentCopyOptionsFromString = olPromptUser
        Case "olCreateAppointment": OlAppointmentCopyOptionsFromString = olCreateAppointment
        Case "olCopyAsAccept": OlAppointmentCopyOptionsFromString = olCopyAsAccept
    End Select
End Function

Function OlAppointmentCopyOptionsToString(value As OlAppointmentCopyOptions) As String
    Select Case value
        Case olPromptUser: OlAppointmentCopyOptionsToString = "olPromptUser"
        Case olCreateAppointment: OlAppointmentCopyOptionsToString = "olCreateAppointment"
        Case olCopyAsAccept: OlAppointmentCopyOptionsToString = "olCopyAsAccept"
    End Select
End Function
