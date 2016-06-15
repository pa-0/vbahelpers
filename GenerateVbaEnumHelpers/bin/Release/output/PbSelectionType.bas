Attribute VB_Name = "wPbSelectionType"
Function PbSelectionTypeFromString(value As String) As PbSelectionType
    If IsNumeric(value) Then
        PbSelectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbSelectionNone": PbSelectionTypeFromString = pbSelectionNone
        Case "pbSelectionShape": PbSelectionTypeFromString = pbSelectionShape
        Case "pbSelectionText": PbSelectionTypeFromString = pbSelectionText
        Case "pbSelectionTableCells": PbSelectionTypeFromString = pbSelectionTableCells
        Case "pbSelectionShapeSubSelection": PbSelectionTypeFromString = pbSelectionShapeSubSelection
    End Select
End Function

Function PbSelectionTypeToString(value As PbSelectionType) As String
    Select Case value
        Case pbSelectionNone: PbSelectionTypeToString = "pbSelectionNone"
        Case pbSelectionShape: PbSelectionTypeToString = "pbSelectionShape"
        Case pbSelectionText: PbSelectionTypeToString = "pbSelectionText"
        Case pbSelectionTableCells: PbSelectionTypeToString = "pbSelectionTableCells"
        Case pbSelectionShapeSubSelection: PbSelectionTypeToString = "pbSelectionShapeSubSelection"
    End Select
End Function
