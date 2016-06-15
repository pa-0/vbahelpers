Attribute VB_Name = "wXlErrorChecks"
Function XlErrorChecksFromString(value As String) As XlErrorChecks
    If IsNumeric(value) Then
        XlErrorChecksFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlEvaluateToError": XlErrorChecksFromString = xlEvaluateToError
        Case "xlTextDate": XlErrorChecksFromString = xlTextDate
        Case "xlNumberAsText": XlErrorChecksFromString = xlNumberAsText
        Case "xlInconsistentFormula": XlErrorChecksFromString = xlInconsistentFormula
        Case "xlOmittedCells": XlErrorChecksFromString = xlOmittedCells
        Case "xlUnlockedFormulaCells": XlErrorChecksFromString = xlUnlockedFormulaCells
        Case "xlEmptyCellReferences": XlErrorChecksFromString = xlEmptyCellReferences
        Case "xlListDataValidation": XlErrorChecksFromString = xlListDataValidation
        Case "xlInconsistentListFormula": XlErrorChecksFromString = xlInconsistentListFormula
    End Select
End Function

Function XlErrorChecksToString(value As XlErrorChecks) As String
    Select Case value
        Case xlEvaluateToError: XlErrorChecksToString = "xlEvaluateToError"
        Case xlTextDate: XlErrorChecksToString = "xlTextDate"
        Case xlNumberAsText: XlErrorChecksToString = "xlNumberAsText"
        Case xlInconsistentFormula: XlErrorChecksToString = "xlInconsistentFormula"
        Case xlOmittedCells: XlErrorChecksToString = "xlOmittedCells"
        Case xlUnlockedFormulaCells: XlErrorChecksToString = "xlUnlockedFormulaCells"
        Case xlEmptyCellReferences: XlErrorChecksToString = "xlEmptyCellReferences"
        Case xlListDataValidation: XlErrorChecksToString = "xlListDataValidation"
        Case xlInconsistentListFormula: XlErrorChecksToString = "xlInconsistentListFormula"
    End Select
End Function
