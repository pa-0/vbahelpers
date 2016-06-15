Attribute VB_Name = "wWdUseFormattingFrom"
Function WdUseFormattingFromFromString(value As String) As WdUseFormattingFrom
    If IsNumeric(value) Then
        WdUseFormattingFromFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFormattingFromCurrent": WdUseFormattingFromFromString = wdFormattingFromCurrent
        Case "wdFormattingFromSelected": WdUseFormattingFromFromString = wdFormattingFromSelected
        Case "wdFormattingFromPrompt": WdUseFormattingFromFromString = wdFormattingFromPrompt
    End Select
End Function

Function WdUseFormattingFromToString(value As WdUseFormattingFrom) As String
    Select Case value
        Case wdFormattingFromCurrent: WdUseFormattingFromToString = "wdFormattingFromCurrent"
        Case wdFormattingFromSelected: WdUseFormattingFromToString = "wdFormattingFromSelected"
        Case wdFormattingFromPrompt: WdUseFormattingFromToString = "wdFormattingFromPrompt"
    End Select
End Function
