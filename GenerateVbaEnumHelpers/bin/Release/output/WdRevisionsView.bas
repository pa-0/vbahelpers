Attribute VB_Name = "wWdRevisionsView"
Function WdRevisionsViewFromString(value As String) As WdRevisionsView
    If IsNumeric(value) Then
        WdRevisionsViewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRevisionsViewFinal": WdRevisionsViewFromString = wdRevisionsViewFinal
        Case "wdRevisionsViewOriginal": WdRevisionsViewFromString = wdRevisionsViewOriginal
    End Select
End Function

Function WdRevisionsViewToString(value As WdRevisionsView) As String
    Select Case value
        Case wdRevisionsViewFinal: WdRevisionsViewToString = "wdRevisionsViewFinal"
        Case wdRevisionsViewOriginal: WdRevisionsViewToString = "wdRevisionsViewOriginal"
    End Select
End Function
