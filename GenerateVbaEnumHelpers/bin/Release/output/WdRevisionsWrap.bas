Attribute VB_Name = "wWdRevisionsWrap"
Function WdRevisionsWrapFromString(value As String) As WdRevisionsWrap
    If IsNumeric(value) Then
        WdRevisionsWrapFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWrapNever": WdRevisionsWrapFromString = wdWrapNever
        Case "wdWrapAlways": WdRevisionsWrapFromString = wdWrapAlways
        Case "wdWrapAsk": WdRevisionsWrapFromString = wdWrapAsk
    End Select
End Function

Function WdRevisionsWrapToString(value As WdRevisionsWrap) As String
    Select Case value
        Case wdWrapNever: WdRevisionsWrapToString = "wdWrapNever"
        Case wdWrapAlways: WdRevisionsWrapToString = "wdWrapAlways"
        Case wdWrapAsk: WdRevisionsWrapToString = "wdWrapAsk"
    End Select
End Function
