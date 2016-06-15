Attribute VB_Name = "wWdSelectionType"
Function WdSelectionTypeFromString(value As String) As WdSelectionType
    If IsNumeric(value) Then
        WdSelectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNoSelection": WdSelectionTypeFromString = wdNoSelection
        Case "wdSelectionIP": WdSelectionTypeFromString = wdSelectionIP
        Case "wdSelectionNormal": WdSelectionTypeFromString = wdSelectionNormal
        Case "wdSelectionFrame": WdSelectionTypeFromString = wdSelectionFrame
        Case "wdSelectionColumn": WdSelectionTypeFromString = wdSelectionColumn
        Case "wdSelectionRow": WdSelectionTypeFromString = wdSelectionRow
        Case "wdSelectionBlock": WdSelectionTypeFromString = wdSelectionBlock
        Case "wdSelectionInlineShape": WdSelectionTypeFromString = wdSelectionInlineShape
        Case "wdSelectionShape": WdSelectionTypeFromString = wdSelectionShape
    End Select
End Function

Function WdSelectionTypeToString(value As WdSelectionType) As String
    Select Case value
        Case wdNoSelection: WdSelectionTypeToString = "wdNoSelection"
        Case wdSelectionIP: WdSelectionTypeToString = "wdSelectionIP"
        Case wdSelectionNormal: WdSelectionTypeToString = "wdSelectionNormal"
        Case wdSelectionFrame: WdSelectionTypeToString = "wdSelectionFrame"
        Case wdSelectionColumn: WdSelectionTypeToString = "wdSelectionColumn"
        Case wdSelectionRow: WdSelectionTypeToString = "wdSelectionRow"
        Case wdSelectionBlock: WdSelectionTypeToString = "wdSelectionBlock"
        Case wdSelectionInlineShape: WdSelectionTypeToString = "wdSelectionInlineShape"
        Case wdSelectionShape: WdSelectionTypeToString = "wdSelectionShape"
    End Select
End Function
