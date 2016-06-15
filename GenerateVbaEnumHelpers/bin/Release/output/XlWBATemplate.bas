Attribute VB_Name = "wXlWBATemplate"
Function XlWBATemplateFromString(value As String) As XlWBATemplate
    If IsNumeric(value) Then
        XlWBATemplateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWBATExcel4MacroSheet": XlWBATemplateFromString = xlWBATExcel4MacroSheet
        Case "xlWBATExcel4IntlMacroSheet": XlWBATemplateFromString = xlWBATExcel4IntlMacroSheet
        Case "xlWBATWorksheet": XlWBATemplateFromString = xlWBATWorksheet
        Case "xlWBATChart": XlWBATemplateFromString = xlWBATChart
    End Select
End Function

Function XlWBATemplateToString(value As XlWBATemplate) As String
    Select Case value
        Case xlWBATExcel4MacroSheet: XlWBATemplateToString = "xlWBATExcel4MacroSheet"
        Case xlWBATExcel4IntlMacroSheet: XlWBATemplateToString = "xlWBATExcel4IntlMacroSheet"
        Case xlWBATWorksheet: XlWBATemplateToString = "xlWBATWorksheet"
        Case xlWBATChart: XlWBATemplateToString = "xlWBATChart"
    End Select
End Function
