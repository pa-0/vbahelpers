Attribute VB_Name = "wXlPTSelectionMode"
Function XlPTSelectionModeFromString(value As String) As XlPTSelectionMode
    If IsNumeric(value) Then
        XlPTSelectionModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataAndLabel": XlPTSelectionModeFromString = xlDataAndLabel
        Case "xlLabelOnly": XlPTSelectionModeFromString = xlLabelOnly
        Case "xlDataOnly": XlPTSelectionModeFromString = xlDataOnly
        Case "xlOrigin": XlPTSelectionModeFromString = xlOrigin
        Case "xlBlanks": XlPTSelectionModeFromString = xlBlanks
        Case "xlButton": XlPTSelectionModeFromString = xlButton
        Case "xlFirstRow": XlPTSelectionModeFromString = xlFirstRow
    End Select
End Function

Function XlPTSelectionModeToString(value As XlPTSelectionMode) As String
    Select Case value
        Case xlDataAndLabel: XlPTSelectionModeToString = "xlDataAndLabel"
        Case xlLabelOnly: XlPTSelectionModeToString = "xlLabelOnly"
        Case xlDataOnly: XlPTSelectionModeToString = "xlDataOnly"
        Case xlOrigin: XlPTSelectionModeToString = "xlOrigin"
        Case xlBlanks: XlPTSelectionModeToString = "xlBlanks"
        Case xlButton: XlPTSelectionModeToString = "xlButton"
        Case xlFirstRow: XlPTSelectionModeToString = "xlFirstRow"
    End Select
End Function
