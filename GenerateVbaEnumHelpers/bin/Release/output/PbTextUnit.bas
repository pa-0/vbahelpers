Attribute VB_Name = "wPbTextUnit"
Function PbTextUnitFromString(value As String) As PbTextUnit
    If IsNumeric(value) Then
        PbTextUnitFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTextUnitCharacter": PbTextUnitFromString = pbTextUnitCharacter
        Case "pbTextUnitWord": PbTextUnitFromString = pbTextUnitWord
        Case "pbTextUnitSentence": PbTextUnitFromString = pbTextUnitSentence
        Case "pbTextUnitParagraph": PbTextUnitFromString = pbTextUnitParagraph
        Case "pbTextUnitLine": PbTextUnitFromString = pbTextUnitLine
        Case "pbTextUnitStory": PbTextUnitFromString = pbTextUnitStory
        Case "pbTextUnitScreen": PbTextUnitFromString = pbTextUnitScreen
        Case "pbTextUnitSection": PbTextUnitFromString = pbTextUnitSection
        Case "pbTextUnitColumn": PbTextUnitFromString = pbTextUnitColumn
        Case "pbTextUnitRow": PbTextUnitFromString = pbTextUnitRow
        Case "pbTextUnitWindow": PbTextUnitFromString = pbTextUnitWindow
        Case "pbTextUnitCell": PbTextUnitFromString = pbTextUnitCell
        Case "pbTextUnitCharFormat": PbTextUnitFromString = pbTextUnitCharFormat
        Case "pbTextUnitParaFormat": PbTextUnitFromString = pbTextUnitParaFormat
        Case "pbTextUnitTable": PbTextUnitFromString = pbTextUnitTable
        Case "pbTextUnitObject": PbTextUnitFromString = pbTextUnitObject
        Case "pbTextUnitCodePoint": PbTextUnitFromString = pbTextUnitCodePoint
    End Select
End Function

Function PbTextUnitToString(value As PbTextUnit) As String
    Select Case value
        Case pbTextUnitCharacter: PbTextUnitToString = "pbTextUnitCharacter"
        Case pbTextUnitWord: PbTextUnitToString = "pbTextUnitWord"
        Case pbTextUnitSentence: PbTextUnitToString = "pbTextUnitSentence"
        Case pbTextUnitParagraph: PbTextUnitToString = "pbTextUnitParagraph"
        Case pbTextUnitLine: PbTextUnitToString = "pbTextUnitLine"
        Case pbTextUnitStory: PbTextUnitToString = "pbTextUnitStory"
        Case pbTextUnitScreen: PbTextUnitToString = "pbTextUnitScreen"
        Case pbTextUnitSection: PbTextUnitToString = "pbTextUnitSection"
        Case pbTextUnitColumn: PbTextUnitToString = "pbTextUnitColumn"
        Case pbTextUnitRow: PbTextUnitToString = "pbTextUnitRow"
        Case pbTextUnitWindow: PbTextUnitToString = "pbTextUnitWindow"
        Case pbTextUnitCell: PbTextUnitToString = "pbTextUnitCell"
        Case pbTextUnitCharFormat: PbTextUnitToString = "pbTextUnitCharFormat"
        Case pbTextUnitParaFormat: PbTextUnitToString = "pbTextUnitParaFormat"
        Case pbTextUnitTable: PbTextUnitToString = "pbTextUnitTable"
        Case pbTextUnitObject: PbTextUnitToString = "pbTextUnitObject"
        Case pbTextUnitCodePoint: PbTextUnitToString = "pbTextUnitCodePoint"
    End Select
End Function
