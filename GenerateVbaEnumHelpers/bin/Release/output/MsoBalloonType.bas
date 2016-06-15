Attribute VB_Name = "wMsoBalloonType"
Function MsoBalloonTypeFromString(value As String) As MsoBalloonType
    If IsNumeric(value) Then
        MsoBalloonTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBalloonTypeButtons": MsoBalloonTypeFromString = msoBalloonTypeButtons
        Case "msoBalloonTypeBullets": MsoBalloonTypeFromString = msoBalloonTypeBullets
        Case "msoBalloonTypeNumbers": MsoBalloonTypeFromString = msoBalloonTypeNumbers
    End Select
End Function

Function MsoBalloonTypeToString(value As MsoBalloonType) As String
    Select Case value
        Case msoBalloonTypeButtons: MsoBalloonTypeToString = "msoBalloonTypeButtons"
        Case msoBalloonTypeBullets: MsoBalloonTypeToString = "msoBalloonTypeBullets"
        Case msoBalloonTypeNumbers: MsoBalloonTypeToString = "msoBalloonTypeNumbers"
    End Select
End Function
