Attribute VB_Name = "wWdRecoveryType"
Function WdRecoveryTypeFromString(value As String) As WdRecoveryType
    If IsNumeric(value) Then
        WdRecoveryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPasteDefault": WdRecoveryTypeFromString = wdPasteDefault
        Case "wdSingleCellText": WdRecoveryTypeFromString = wdSingleCellText
        Case "wdSingleCellTable": WdRecoveryTypeFromString = wdSingleCellTable
        Case "wdListContinueNumbering": WdRecoveryTypeFromString = wdListContinueNumbering
        Case "wdListRestartNumbering": WdRecoveryTypeFromString = wdListRestartNumbering
        Case "wdTableAppendTable": WdRecoveryTypeFromString = wdTableAppendTable
        Case "wdTableInsertAsRows": WdRecoveryTypeFromString = wdTableInsertAsRows
        Case "wdTableOriginalFormatting": WdRecoveryTypeFromString = wdTableOriginalFormatting
        Case "wdChartPicture": WdRecoveryTypeFromString = wdChartPicture
        Case "wdChart": WdRecoveryTypeFromString = wdChart
        Case "wdChartLinked": WdRecoveryTypeFromString = wdChartLinked
        Case "wdFormatOriginalFormatting": WdRecoveryTypeFromString = wdFormatOriginalFormatting
        Case "wdUseDestinationStylesRecovery": WdRecoveryTypeFromString = wdUseDestinationStylesRecovery
        Case "wdFormatSurroundingFormattingWithEmphasis": WdRecoveryTypeFromString = wdFormatSurroundingFormattingWithEmphasis
        Case "wdFormatPlainText": WdRecoveryTypeFromString = wdFormatPlainText
        Case "wdTableOverwriteCells": WdRecoveryTypeFromString = wdTableOverwriteCells
        Case "wdListCombineWithExistingList": WdRecoveryTypeFromString = wdListCombineWithExistingList
        Case "wdListDontMerge": WdRecoveryTypeFromString = wdListDontMerge
    End Select
End Function

Function WdRecoveryTypeToString(value As WdRecoveryType) As String
    Select Case value
        Case wdPasteDefault: WdRecoveryTypeToString = "wdPasteDefault"
        Case wdSingleCellText: WdRecoveryTypeToString = "wdSingleCellText"
        Case wdSingleCellTable: WdRecoveryTypeToString = "wdSingleCellTable"
        Case wdListContinueNumbering: WdRecoveryTypeToString = "wdListContinueNumbering"
        Case wdListRestartNumbering: WdRecoveryTypeToString = "wdListRestartNumbering"
        Case wdTableAppendTable: WdRecoveryTypeToString = "wdTableAppendTable"
        Case wdTableInsertAsRows: WdRecoveryTypeToString = "wdTableInsertAsRows"
        Case wdTableOriginalFormatting: WdRecoveryTypeToString = "wdTableOriginalFormatting"
        Case wdChartPicture: WdRecoveryTypeToString = "wdChartPicture"
        Case wdChart: WdRecoveryTypeToString = "wdChart"
        Case wdChartLinked: WdRecoveryTypeToString = "wdChartLinked"
        Case wdFormatOriginalFormatting: WdRecoveryTypeToString = "wdFormatOriginalFormatting"
        Case wdUseDestinationStylesRecovery: WdRecoveryTypeToString = "wdUseDestinationStylesRecovery"
        Case wdFormatSurroundingFormattingWithEmphasis: WdRecoveryTypeToString = "wdFormatSurroundingFormattingWithEmphasis"
        Case wdFormatPlainText: WdRecoveryTypeToString = "wdFormatPlainText"
        Case wdTableOverwriteCells: WdRecoveryTypeToString = "wdTableOverwriteCells"
        Case wdListCombineWithExistingList: WdRecoveryTypeToString = "wdListCombineWithExistingList"
        Case wdListDontMerge: WdRecoveryTypeToString = "wdListDontMerge"
    End Select
End Function
