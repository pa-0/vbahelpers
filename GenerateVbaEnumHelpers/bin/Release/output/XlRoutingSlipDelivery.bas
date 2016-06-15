Attribute VB_Name = "wXlRoutingSlipDelivery"
Function XlRoutingSlipDeliveryFromString(value As String) As XlRoutingSlipDelivery
    If IsNumeric(value) Then
        XlRoutingSlipDeliveryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOneAfterAnother": XlRoutingSlipDeliveryFromString = xlOneAfterAnother
        Case "xlAllAtOnce": XlRoutingSlipDeliveryFromString = xlAllAtOnce
    End Select
End Function

Function XlRoutingSlipDeliveryToString(value As XlRoutingSlipDelivery) As String
    Select Case value
        Case xlOneAfterAnother: XlRoutingSlipDeliveryToString = "xlOneAfterAnother"
        Case xlAllAtOnce: XlRoutingSlipDeliveryToString = "xlAllAtOnce"
    End Select
End Function
