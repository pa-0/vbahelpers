Attribute VB_Name = "wWdRoutingSlipDelivery"
Function WdRoutingSlipDeliveryFromString(value As String) As WdRoutingSlipDelivery
    If IsNumeric(value) Then
        WdRoutingSlipDeliveryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOneAfterAnother": WdRoutingSlipDeliveryFromString = wdOneAfterAnother
        Case "wdAllAtOnce": WdRoutingSlipDeliveryFromString = wdAllAtOnce
    End Select
End Function

Function WdRoutingSlipDeliveryToString(value As WdRoutingSlipDelivery) As String
    Select Case value
        Case wdOneAfterAnother: WdRoutingSlipDeliveryToString = "wdOneAfterAnother"
        Case wdAllAtOnce: WdRoutingSlipDeliveryToString = "wdAllAtOnce"
    End Select
End Function
