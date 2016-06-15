Attribute VB_Name = "wWdRectangleType"
Function WdRectangleTypeFromString(value As String) As WdRectangleType
    If IsNumeric(value) Then
        WdRectangleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTextRectangle": WdRectangleTypeFromString = wdTextRectangle
        Case "wdShapeRectangle": WdRectangleTypeFromString = wdShapeRectangle
        Case "wdMarkupRectangle": WdRectangleTypeFromString = wdMarkupRectangle
        Case "wdMarkupRectangleButton": WdRectangleTypeFromString = wdMarkupRectangleButton
        Case "wdPageBorderRectangle": WdRectangleTypeFromString = wdPageBorderRectangle
        Case "wdLineBetweenColumnRectangle": WdRectangleTypeFromString = wdLineBetweenColumnRectangle
        Case "wdSelection": WdRectangleTypeFromString = wdSelection
        Case "wdSystem": WdRectangleTypeFromString = wdSystem
        Case "wdMarkupRectangleArea": WdRectangleTypeFromString = wdMarkupRectangleArea
        Case "wdReadingModeNavigation": WdRectangleTypeFromString = wdReadingModeNavigation
        Case "wdMarkupRectangleMoveMatch": WdRectangleTypeFromString = wdMarkupRectangleMoveMatch
        Case "wdReadingModePanningArea": WdRectangleTypeFromString = wdReadingModePanningArea
        Case "wdMailNavArea": WdRectangleTypeFromString = wdMailNavArea
        Case "wdDocumentControlRectangle": WdRectangleTypeFromString = wdDocumentControlRectangle
    End Select
End Function

Function WdRectangleTypeToString(value As WdRectangleType) As String
    Select Case value
        Case wdTextRectangle: WdRectangleTypeToString = "wdTextRectangle"
        Case wdShapeRectangle: WdRectangleTypeToString = "wdShapeRectangle"
        Case wdMarkupRectangle: WdRectangleTypeToString = "wdMarkupRectangle"
        Case wdMarkupRectangleButton: WdRectangleTypeToString = "wdMarkupRectangleButton"
        Case wdPageBorderRectangle: WdRectangleTypeToString = "wdPageBorderRectangle"
        Case wdLineBetweenColumnRectangle: WdRectangleTypeToString = "wdLineBetweenColumnRectangle"
        Case wdSelection: WdRectangleTypeToString = "wdSelection"
        Case wdSystem: WdRectangleTypeToString = "wdSystem"
        Case wdMarkupRectangleArea: WdRectangleTypeToString = "wdMarkupRectangleArea"
        Case wdReadingModeNavigation: WdRectangleTypeToString = "wdReadingModeNavigation"
        Case wdMarkupRectangleMoveMatch: WdRectangleTypeToString = "wdMarkupRectangleMoveMatch"
        Case wdReadingModePanningArea: WdRectangleTypeToString = "wdReadingModePanningArea"
        Case wdMailNavArea: WdRectangleTypeToString = "wdMailNavArea"
        Case wdDocumentControlRectangle: WdRectangleTypeToString = "wdDocumentControlRectangle"
    End Select
End Function
