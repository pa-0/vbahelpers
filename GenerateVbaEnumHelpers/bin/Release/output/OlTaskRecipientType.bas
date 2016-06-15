Attribute VB_Name = "wOlTaskRecipientType"
Function OlTaskRecipientTypeFromString(value As String) As OlTaskRecipientType
    If IsNumeric(value) Then
        OlTaskRecipientTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUpdate": OlTaskRecipientTypeFromString = olUpdate
        Case "olFinalStatus": OlTaskRecipientTypeFromString = olFinalStatus
    End Select
End Function

Function OlTaskRecipientTypeToString(value As OlTaskRecipientType) As String
    Select Case value
        Case olUpdate: OlTaskRecipientTypeToString = "olUpdate"
        Case olFinalStatus: OlTaskRecipientTypeToString = "olFinalStatus"
    End Select
End Function
