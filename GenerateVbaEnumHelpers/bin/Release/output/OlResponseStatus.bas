Attribute VB_Name = "wOlResponseStatus"
Function OlResponseStatusFromString(value As String) As OlResponseStatus
    If IsNumeric(value) Then
        OlResponseStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olResponseNone": OlResponseStatusFromString = olResponseNone
        Case "olResponseOrganized": OlResponseStatusFromString = olResponseOrganized
        Case "olResponseTentative": OlResponseStatusFromString = olResponseTentative
        Case "olResponseAccepted": OlResponseStatusFromString = olResponseAccepted
        Case "olResponseDeclined": OlResponseStatusFromString = olResponseDeclined
        Case "olResponseNotResponded": OlResponseStatusFromString = olResponseNotResponded
    End Select
End Function

Function OlResponseStatusToString(value As OlResponseStatus) As String
    Select Case value
        Case olResponseNone: OlResponseStatusToString = "olResponseNone"
        Case olResponseOrganized: OlResponseStatusToString = "olResponseOrganized"
        Case olResponseTentative: OlResponseStatusToString = "olResponseTentative"
        Case olResponseAccepted: OlResponseStatusToString = "olResponseAccepted"
        Case olResponseDeclined: OlResponseStatusToString = "olResponseDeclined"
        Case olResponseNotResponded: OlResponseStatusToString = "olResponseNotResponded"
    End Select
End Function
