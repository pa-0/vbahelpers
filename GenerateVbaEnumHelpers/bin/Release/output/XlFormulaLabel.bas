Attribute VB_Name = "wXlFormulaLabel"
Function XlFormulaLabelFromString(value As String) As XlFormulaLabel
    If IsNumeric(value) Then
        XlFormulaLabelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRowLabels": XlFormulaLabelFromString = xlRowLabels
        Case "xlColumnLabels": XlFormulaLabelFromString = xlColumnLabels
        Case "xlMixedLabels": XlFormulaLabelFromString = xlMixedLabels
        Case "xlNoLabels": XlFormulaLabelFromString = xlNoLabels
    End Select
End Function

Function XlFormulaLabelToString(value As XlFormulaLabel) As String
    Select Case value
        Case xlRowLabels: XlFormulaLabelToString = "xlRowLabels"
        Case xlColumnLabels: XlFormulaLabelToString = "xlColumnLabels"
        Case xlMixedLabels: XlFormulaLabelToString = "xlMixedLabels"
        Case xlNoLabels: XlFormulaLabelToString = "xlNoLabels"
    End Select
End Function
