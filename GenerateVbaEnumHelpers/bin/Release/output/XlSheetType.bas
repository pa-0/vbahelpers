Attribute VB_Name = "wXlSheetType"
Function XlSheetTypeFromString(value As String) As XlSheetType
    If IsNumeric(value) Then
        XlSheetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlExcel4MacroSheet": XlSheetTypeFromString = xlExcel4MacroSheet
        Case "xlExcel4IntlMacroSheet": XlSheetTypeFromString = xlExcel4IntlMacroSheet
        Case "xlWorksheet": XlSheetTypeFromString = xlWorksheet
        Case "xlDialogSheet": XlSheetTypeFromString = xlDialogSheet
        Case "xlChart": XlSheetTypeFromString = xlChart
    End Select
End Function

Function XlSheetTypeToString(value As XlSheetType) As String
    Select Case value
        Case xlExcel4MacroSheet: XlSheetTypeToString = "xlExcel4MacroSheet"
        Case xlExcel4IntlMacroSheet: XlSheetTypeToString = "xlExcel4IntlMacroSheet"
        Case xlWorksheet: XlSheetTypeToString = "xlWorksheet"
        Case xlDialogSheet: XlSheetTypeToString = "xlDialogSheet"
        Case xlChart: XlSheetTypeToString = "xlChart"
    End Select
End Function
