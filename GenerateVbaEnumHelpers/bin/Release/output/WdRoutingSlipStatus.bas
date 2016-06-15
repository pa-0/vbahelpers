Attribute VB_Name = "wWdRoutingSlipStatus"
Function WdRoutingSlipStatusFromString(value As String) As WdRoutingSlipStatus
    If IsNumeric(value) Then
        WdRoutingSlipStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNotYetRouted": WdRoutingSlipStatusFromString = wdNotYetRouted
        Case "wdRouteInProgress": WdRoutingSlipStatusFromString = wdRouteInProgress
        Case "wdRouteComplete": WdRoutingSlipStatusFromString = wdRouteComplete
    End Select
End Function

Function WdRoutingSlipStatusToString(value As WdRoutingSlipStatus) As String
    Select Case value
        Case wdNotYetRouted: WdRoutingSlipStatusToString = "wdNotYetRouted"
        Case wdRouteInProgress: WdRoutingSlipStatusToString = "wdRouteInProgress"
        Case wdRouteComplete: WdRoutingSlipStatusToString = "wdRouteComplete"
    End Select
End Function
