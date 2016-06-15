Attribute VB_Name = "wXlPasteSpecialOperation"
Function XlPasteSpecialOperationFromString(value As String) As XlPasteSpecialOperation
    If IsNumeric(value) Then
        XlPasteSpecialOperationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPasteSpecialOperationAdd": XlPasteSpecialOperationFromString = xlPasteSpecialOperationAdd
        Case "xlPasteSpecialOperationSubtract": XlPasteSpecialOperationFromString = xlPasteSpecialOperationSubtract
        Case "xlPasteSpecialOperationMultiply": XlPasteSpecialOperationFromString = xlPasteSpecialOperationMultiply
        Case "xlPasteSpecialOperationDivide": XlPasteSpecialOperationFromString = xlPasteSpecialOperationDivide
        Case "xlPasteSpecialOperationNone": XlPasteSpecialOperationFromString = xlPasteSpecialOperationNone
    End Select
End Function

Function XlPasteSpecialOperationToString(value As XlPasteSpecialOperation) As String
    Select Case value
        Case xlPasteSpecialOperationAdd: XlPasteSpecialOperationToString = "xlPasteSpecialOperationAdd"
        Case xlPasteSpecialOperationSubtract: XlPasteSpecialOperationToString = "xlPasteSpecialOperationSubtract"
        Case xlPasteSpecialOperationMultiply: XlPasteSpecialOperationToString = "xlPasteSpecialOperationMultiply"
        Case xlPasteSpecialOperationDivide: XlPasteSpecialOperationToString = "xlPasteSpecialOperationDivide"
        Case xlPasteSpecialOperationNone: XlPasteSpecialOperationToString = "xlPasteSpecialOperationNone"
    End Select
End Function
