Attribute VB_Name = "wWdGoToItem"
Function WdGoToItemFromString(value As String) As WdGoToItem
    If IsNumeric(value) Then
        WdGoToItemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGoToSection": WdGoToItemFromString = wdGoToSection
        Case "wdGoToPage": WdGoToItemFromString = wdGoToPage
        Case "wdGoToTable": WdGoToItemFromString = wdGoToTable
        Case "wdGoToLine": WdGoToItemFromString = wdGoToLine
        Case "wdGoToFootnote": WdGoToItemFromString = wdGoToFootnote
        Case "wdGoToEndnote": WdGoToItemFromString = wdGoToEndnote
        Case "wdGoToComment": WdGoToItemFromString = wdGoToComment
        Case "wdGoToField": WdGoToItemFromString = wdGoToField
        Case "wdGoToGraphic": WdGoToItemFromString = wdGoToGraphic
        Case "wdGoToObject": WdGoToItemFromString = wdGoToObject
        Case "wdGoToEquation": WdGoToItemFromString = wdGoToEquation
        Case "wdGoToHeading": WdGoToItemFromString = wdGoToHeading
        Case "wdGoToPercent": WdGoToItemFromString = wdGoToPercent
        Case "wdGoToSpellingError": WdGoToItemFromString = wdGoToSpellingError
        Case "wdGoToGrammaticalError": WdGoToItemFromString = wdGoToGrammaticalError
        Case "wdGoToProofreadingError": WdGoToItemFromString = wdGoToProofreadingError
        Case "wdGoToBookmark": WdGoToItemFromString = wdGoToBookmark
    End Select
End Function

Function WdGoToItemToString(value As WdGoToItem) As String
    Select Case value
        Case wdGoToSection: WdGoToItemToString = "wdGoToSection"
        Case wdGoToPage: WdGoToItemToString = "wdGoToPage"
        Case wdGoToTable: WdGoToItemToString = "wdGoToTable"
        Case wdGoToLine: WdGoToItemToString = "wdGoToLine"
        Case wdGoToFootnote: WdGoToItemToString = "wdGoToFootnote"
        Case wdGoToEndnote: WdGoToItemToString = "wdGoToEndnote"
        Case wdGoToComment: WdGoToItemToString = "wdGoToComment"
        Case wdGoToField: WdGoToItemToString = "wdGoToField"
        Case wdGoToGraphic: WdGoToItemToString = "wdGoToGraphic"
        Case wdGoToObject: WdGoToItemToString = "wdGoToObject"
        Case wdGoToEquation: WdGoToItemToString = "wdGoToEquation"
        Case wdGoToHeading: WdGoToItemToString = "wdGoToHeading"
        Case wdGoToPercent: WdGoToItemToString = "wdGoToPercent"
        Case wdGoToSpellingError: WdGoToItemToString = "wdGoToSpellingError"
        Case wdGoToGrammaticalError: WdGoToItemToString = "wdGoToGrammaticalError"
        Case wdGoToProofreadingError: WdGoToItemToString = "wdGoToProofreadingError"
        Case wdGoToBookmark: WdGoToItemToString = "wdGoToBookmark"
    End Select
End Function
