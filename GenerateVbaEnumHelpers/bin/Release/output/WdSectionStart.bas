Attribute VB_Name = "wWdSectionStart"
Function WdSectionStartFromString(value As String) As WdSectionStart
    If IsNumeric(value) Then
        WdSectionStartFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSectionContinuous": WdSectionStartFromString = wdSectionContinuous
        Case "wdSectionNewColumn": WdSectionStartFromString = wdSectionNewColumn
        Case "wdSectionNewPage": WdSectionStartFromString = wdSectionNewPage
        Case "wdSectionEvenPage": WdSectionStartFromString = wdSectionEvenPage
        Case "wdSectionOddPage": WdSectionStartFromString = wdSectionOddPage
    End Select
End Function

Function WdSectionStartToString(value As WdSectionStart) As String
    Select Case value
        Case wdSectionContinuous: WdSectionStartToString = "wdSectionContinuous"
        Case wdSectionNewColumn: WdSectionStartToString = "wdSectionNewColumn"
        Case wdSectionNewPage: WdSectionStartToString = "wdSectionNewPage"
        Case wdSectionEvenPage: WdSectionStartToString = "wdSectionEvenPage"
        Case wdSectionOddPage: WdSectionStartToString = "wdSectionOddPage"
    End Select
End Function
