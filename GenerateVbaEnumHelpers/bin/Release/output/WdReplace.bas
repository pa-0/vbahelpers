Attribute VB_Name = "wWdReplace"
Function WdReplaceFromString(value As String) As WdReplace
    If IsNumeric(value) Then
        WdReplaceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdReplaceNone": WdReplaceFromString = wdReplaceNone
        Case "wdReplaceOne": WdReplaceFromString = wdReplaceOne
        Case "wdReplaceAll": WdReplaceFromString = wdReplaceAll
    End Select
End Function

Function WdReplaceToString(value As WdReplace) As String
    Select Case value
        Case wdReplaceNone: WdReplaceToString = "wdReplaceNone"
        Case wdReplaceOne: WdReplaceToString = "wdReplaceOne"
        Case wdReplaceAll: WdReplaceToString = "wdReplaceAll"
    End Select
End Function
