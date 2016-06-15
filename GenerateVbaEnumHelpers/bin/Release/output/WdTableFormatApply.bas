Attribute VB_Name = "wWdTableFormatApply"
Function WdTableFormatApplyFromString(value As String) As WdTableFormatApply
    If IsNumeric(value) Then
        WdTableFormatApplyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTableFormatApplyBorders": WdTableFormatApplyFromString = wdTableFormatApplyBorders
        Case "wdTableFormatApplyShading": WdTableFormatApplyFromString = wdTableFormatApplyShading
        Case "wdTableFormatApplyFont": WdTableFormatApplyFromString = wdTableFormatApplyFont
        Case "wdTableFormatApplyColor": WdTableFormatApplyFromString = wdTableFormatApplyColor
        Case "wdTableFormatApplyAutoFit": WdTableFormatApplyFromString = wdTableFormatApplyAutoFit
        Case "wdTableFormatApplyHeadingRows": WdTableFormatApplyFromString = wdTableFormatApplyHeadingRows
        Case "wdTableFormatApplyLastRow": WdTableFormatApplyFromString = wdTableFormatApplyLastRow
        Case "wdTableFormatApplyFirstColumn": WdTableFormatApplyFromString = wdTableFormatApplyFirstColumn
        Case "wdTableFormatApplyLastColumn": WdTableFormatApplyFromString = wdTableFormatApplyLastColumn
    End Select
End Function

Function WdTableFormatApplyToString(value As WdTableFormatApply) As String
    Select Case value
        Case wdTableFormatApplyBorders: WdTableFormatApplyToString = "wdTableFormatApplyBorders"
        Case wdTableFormatApplyShading: WdTableFormatApplyToString = "wdTableFormatApplyShading"
        Case wdTableFormatApplyFont: WdTableFormatApplyToString = "wdTableFormatApplyFont"
        Case wdTableFormatApplyColor: WdTableFormatApplyToString = "wdTableFormatApplyColor"
        Case wdTableFormatApplyAutoFit: WdTableFormatApplyToString = "wdTableFormatApplyAutoFit"
        Case wdTableFormatApplyHeadingRows: WdTableFormatApplyToString = "wdTableFormatApplyHeadingRows"
        Case wdTableFormatApplyLastRow: WdTableFormatApplyToString = "wdTableFormatApplyLastRow"
        Case wdTableFormatApplyFirstColumn: WdTableFormatApplyToString = "wdTableFormatApplyFirstColumn"
        Case wdTableFormatApplyLastColumn: WdTableFormatApplyToString = "wdTableFormatApplyLastColumn"
    End Select
End Function
