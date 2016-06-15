Attribute VB_Name = "wWdUnits"
Function WdUnitsFromString(value As String) As WdUnits
    If IsNumeric(value) Then
        WdUnitsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCharacter": WdUnitsFromString = wdCharacter
        Case "wdWord": WdUnitsFromString = wdWord
        Case "wdSentence": WdUnitsFromString = wdSentence
        Case "wdParagraph": WdUnitsFromString = wdParagraph
        Case "wdLine": WdUnitsFromString = wdLine
        Case "wdStory": WdUnitsFromString = wdStory
        Case "wdScreen": WdUnitsFromString = wdScreen
        Case "wdSection": WdUnitsFromString = wdSection
        Case "wdColumn": WdUnitsFromString = wdColumn
        Case "wdRow": WdUnitsFromString = wdRow
        Case "wdWindow": WdUnitsFromString = wdWindow
        Case "wdCell": WdUnitsFromString = wdCell
        Case "wdCharacterFormatting": WdUnitsFromString = wdCharacterFormatting
        Case "wdParagraphFormatting": WdUnitsFromString = wdParagraphFormatting
        Case "wdTable": WdUnitsFromString = wdTable
        Case "wdItem": WdUnitsFromString = wdItem
    End Select
End Function

Function WdUnitsToString(value As WdUnits) As String
    Select Case value
        Case wdCharacter: WdUnitsToString = "wdCharacter"
        Case wdWord: WdUnitsToString = "wdWord"
        Case wdSentence: WdUnitsToString = "wdSentence"
        Case wdParagraph: WdUnitsToString = "wdParagraph"
        Case wdLine: WdUnitsToString = "wdLine"
        Case wdStory: WdUnitsToString = "wdStory"
        Case wdScreen: WdUnitsToString = "wdScreen"
        Case wdSection: WdUnitsToString = "wdSection"
        Case wdColumn: WdUnitsToString = "wdColumn"
        Case wdRow: WdUnitsToString = "wdRow"
        Case wdWindow: WdUnitsToString = "wdWindow"
        Case wdCell: WdUnitsToString = "wdCell"
        Case wdCharacterFormatting: WdUnitsToString = "wdCharacterFormatting"
        Case wdParagraphFormatting: WdUnitsToString = "wdParagraphFormatting"
        Case wdTable: WdUnitsToString = "wdTable"
        Case wdItem: WdUnitsToString = "wdItem"
    End Select
End Function
