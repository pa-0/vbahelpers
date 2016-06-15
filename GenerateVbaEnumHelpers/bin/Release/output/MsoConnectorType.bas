Attribute VB_Name = "wMsoConnectorType"
Function MsoConnectorTypeFromString(value As String) As MsoConnectorType
    If IsNumeric(value) Then
        MsoConnectorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoConnectorStraight": MsoConnectorTypeFromString = msoConnectorStraight
        Case "msoConnectorElbow": MsoConnectorTypeFromString = msoConnectorElbow
        Case "msoConnectorCurve": MsoConnectorTypeFromString = msoConnectorCurve
        Case "msoConnectorTypeMixed": MsoConnectorTypeFromString = msoConnectorTypeMixed
    End Select
End Function

Function MsoConnectorTypeToString(value As MsoConnectorType) As String
    Select Case value
        Case msoConnectorStraight: MsoConnectorTypeToString = "msoConnectorStraight"
        Case msoConnectorElbow: MsoConnectorTypeToString = "msoConnectorElbow"
        Case msoConnectorCurve: MsoConnectorTypeToString = "msoConnectorCurve"
        Case msoConnectorTypeMixed: MsoConnectorTypeToString = "msoConnectorTypeMixed"
    End Select
End Function
