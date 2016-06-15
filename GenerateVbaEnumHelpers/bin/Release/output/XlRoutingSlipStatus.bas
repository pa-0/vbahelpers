Attribute VB_Name = "wXlRoutingSlipStatus"
Function XlRoutingSlipStatusFromString(value As String) As XlRoutingSlipStatus
    If IsNumeric(value) Then
        XlRoutingSlipStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNotYetRouted": XlRoutingSlipStatusFromString = xlNotYetRouted
        Case "xlRoutingInProgress": XlRoutingSlipStatusFromString = xlRoutingInProgress
        Case "xlRoutingComplete": XlRoutingSlipStatusFromString = xlRoutingComplete
    End Select
End Function

Function XlRoutingSlipStatusToString(value As XlRoutingSlipStatus) As String
    Select Case value
        Case xlNotYetRouted: XlRoutingSlipStatusToString = "xlNotYetRouted"
        Case xlRoutingInProgress: XlRoutingSlipStatusToString = "xlRoutingInProgress"
        Case xlRoutingComplete: XlRoutingSlipStatusToString = "xlRoutingComplete"
    End Select
End Function
