Attribute VB_Name = "wOlBusyStatus"
Function OlBusyStatusFromString(value As String) As OlBusyStatus
    If IsNumeric(value) Then
        OlBusyStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFree": OlBusyStatusFromString = olFree
        Case "olTentative": OlBusyStatusFromString = olTentative
        Case "olBusy": OlBusyStatusFromString = olBusy
        Case "olOutOfOffice": OlBusyStatusFromString = olOutOfOffice
    End Select
End Function

Function OlBusyStatusToString(value As OlBusyStatus) As String
    Select Case value
        Case olFree: OlBusyStatusToString = "olFree"
        Case olTentative: OlBusyStatusToString = "olTentative"
        Case olBusy: OlBusyStatusToString = "olBusy"
        Case olOutOfOffice: OlBusyStatusToString = "olOutOfOffice"
    End Select
End Function
