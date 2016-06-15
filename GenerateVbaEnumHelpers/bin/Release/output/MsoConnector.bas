Attribute VB_Name = "wMsoConnector"
Function MsoConnectorFromString(value As String) As MsoConnector
    If IsNumeric(value) Then
        MsoConnectorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoConnectorAnd": MsoConnectorFromString = msoConnectorAnd
        Case "msoConnectorOr": MsoConnectorFromString = msoConnectorOr
    End Select
End Function

Function MsoConnectorToString(value As MsoConnector) As String
    Select Case value
        Case msoConnectorAnd: MsoConnectorToString = "msoConnectorAnd"
        Case msoConnectorOr: MsoConnectorToString = "msoConnectorOr"
    End Select
End Function
