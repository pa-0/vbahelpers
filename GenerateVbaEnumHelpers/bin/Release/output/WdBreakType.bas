Attribute VB_Name = "wWdBreakType"
Function WdBreakTypeFromString(value As String) As WdBreakType
    If IsNumeric(value) Then
        WdBreakTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSectionBreakNextPage": WdBreakTypeFromString = wdSectionBreakNextPage
        Case "wdSectionBreakContinuous": WdBreakTypeFromString = wdSectionBreakContinuous
        Case "wdSectionBreakEvenPage": WdBreakTypeFromString = wdSectionBreakEvenPage
        Case "wdSectionBreakOddPage": WdBreakTypeFromString = wdSectionBreakOddPage
        Case "wdLineBreak": WdBreakTypeFromString = wdLineBreak
        Case "wdPageBreak": WdBreakTypeFromString = wdPageBreak
        Case "wdColumnBreak": WdBreakTypeFromString = wdColumnBreak
        Case "wdLineBreakClearLeft": WdBreakTypeFromString = wdLineBreakClearLeft
        Case "wdLineBreakClearRight": WdBreakTypeFromString = wdLineBreakClearRight
        Case "wdTextWrappingBreak": WdBreakTypeFromString = wdTextWrappingBreak
    End Select
End Function

Function WdBreakTypeToString(value As WdBreakType) As String
    Select Case value
        Case wdSectionBreakNextPage: WdBreakTypeToString = "wdSectionBreakNextPage"
        Case wdSectionBreakContinuous: WdBreakTypeToString = "wdSectionBreakContinuous"
        Case wdSectionBreakEvenPage: WdBreakTypeToString = "wdSectionBreakEvenPage"
        Case wdSectionBreakOddPage: WdBreakTypeToString = "wdSectionBreakOddPage"
        Case wdLineBreak: WdBreakTypeToString = "wdLineBreak"
        Case wdPageBreak: WdBreakTypeToString = "wdPageBreak"
        Case wdColumnBreak: WdBreakTypeToString = "wdColumnBreak"
        Case wdLineBreakClearLeft: WdBreakTypeToString = "wdLineBreakClearLeft"
        Case wdLineBreakClearRight: WdBreakTypeToString = "wdLineBreakClearRight"
        Case wdTextWrappingBreak: WdBreakTypeToString = "wdTextWrappingBreak"
    End Select
End Function
