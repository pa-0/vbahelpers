Attribute VB_Name = "wWdPortugueseReform"
Function WdPortugueseReformFromString(value As String) As WdPortugueseReform
    If IsNumeric(value) Then
        WdPortugueseReformFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPortuguesePreReform": WdPortugueseReformFromString = wdPortuguesePreReform
        Case "wdPortuguesePostReform": WdPortugueseReformFromString = wdPortuguesePostReform
        Case "wdPortugueseBoth": WdPortugueseReformFromString = wdPortugueseBoth
    End Select
End Function

Function WdPortugueseReformToString(value As WdPortugueseReform) As String
    Select Case value
        Case wdPortuguesePreReform: WdPortugueseReformToString = "wdPortuguesePreReform"
        Case wdPortuguesePostReform: WdPortugueseReformToString = "wdPortuguesePostReform"
        Case wdPortugueseBoth: WdPortugueseReformToString = "wdPortugueseBoth"
    End Select
End Function
