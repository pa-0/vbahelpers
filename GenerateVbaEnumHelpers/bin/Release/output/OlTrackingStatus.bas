Attribute VB_Name = "wOlTrackingStatus"
Function OlTrackingStatusFromString(value As String) As OlTrackingStatus
    If IsNumeric(value) Then
        OlTrackingStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTrackingNone": OlTrackingStatusFromString = olTrackingNone
        Case "olTrackingDelivered": OlTrackingStatusFromString = olTrackingDelivered
        Case "olTrackingNotDelivered": OlTrackingStatusFromString = olTrackingNotDelivered
        Case "olTrackingNotRead": OlTrackingStatusFromString = olTrackingNotRead
        Case "olTrackingRecallFailure": OlTrackingStatusFromString = olTrackingRecallFailure
        Case "olTrackingRecallSuccess": OlTrackingStatusFromString = olTrackingRecallSuccess
        Case "olTrackingRead": OlTrackingStatusFromString = olTrackingRead
        Case "olTrackingReplied": OlTrackingStatusFromString = olTrackingReplied
    End Select
End Function

Function OlTrackingStatusToString(value As OlTrackingStatus) As String
    Select Case value
        Case olTrackingNone: OlTrackingStatusToString = "olTrackingNone"
        Case olTrackingDelivered: OlTrackingStatusToString = "olTrackingDelivered"
        Case olTrackingNotDelivered: OlTrackingStatusToString = "olTrackingNotDelivered"
        Case olTrackingNotRead: OlTrackingStatusToString = "olTrackingNotRead"
        Case olTrackingRecallFailure: OlTrackingStatusToString = "olTrackingRecallFailure"
        Case olTrackingRecallSuccess: OlTrackingStatusToString = "olTrackingRecallSuccess"
        Case olTrackingRead: OlTrackingStatusToString = "olTrackingRead"
        Case olTrackingReplied: OlTrackingStatusToString = "olTrackingReplied"
    End Select
End Function
