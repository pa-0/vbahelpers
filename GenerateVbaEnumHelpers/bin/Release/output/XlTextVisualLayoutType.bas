Attribute VB_Name = "wXlTextVisualLayoutType"
Function XlTextVisualLayoutTypeFromString(value As String) As XlTextVisualLayoutType
    If IsNumeric(value) Then
        XlTextVisualLayoutTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTextVisualLTR": XlTextVisualLayoutTypeFromString = xlTextVisualLTR
        Case "xlTextVisualRTL": XlTextVisualLayoutTypeFromString = xlTextVisualRTL
    End Select
End Function

Function XlTextVisualLayoutTypeToString(value As XlTextVisualLayoutType) As String
    Select Case value
        Case xlTextVisualLTR: XlTextVisualLayoutTypeToString = "xlTextVisualLTR"
        Case xlTextVisualRTL: XlTextVisualLayoutTypeToString = "xlTextVisualRTL"
    End Select
End Function
