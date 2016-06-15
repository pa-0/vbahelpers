Attribute VB_Name = "wWdGranularity"
Function WdGranularityFromString(value As String) As WdGranularity
    If IsNumeric(value) Then
        WdGranularityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGranularityCharLevel": WdGranularityFromString = wdGranularityCharLevel
        Case "wdGranularityWordLevel": WdGranularityFromString = wdGranularityWordLevel
    End Select
End Function

Function WdGranularityToString(value As WdGranularity) As String
    Select Case value
        Case wdGranularityCharLevel: WdGranularityToString = "wdGranularityCharLevel"
        Case wdGranularityWordLevel: WdGranularityToString = "wdGranularityWordLevel"
    End Select
End Function
