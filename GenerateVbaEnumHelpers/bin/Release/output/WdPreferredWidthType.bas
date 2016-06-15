Attribute VB_Name = "wWdPreferredWidthType"
Function WdPreferredWidthTypeFromString(value As String) As WdPreferredWidthType
    If IsNumeric(value) Then
        WdPreferredWidthTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPreferredWidthAuto": WdPreferredWidthTypeFromString = wdPreferredWidthAuto
        Case "wdPreferredWidthPercent": WdPreferredWidthTypeFromString = wdPreferredWidthPercent
        Case "wdPreferredWidthPoints": WdPreferredWidthTypeFromString = wdPreferredWidthPoints
    End Select
End Function

Function WdPreferredWidthTypeToString(value As WdPreferredWidthType) As String
    Select Case value
        Case wdPreferredWidthAuto: WdPreferredWidthTypeToString = "wdPreferredWidthAuto"
        Case wdPreferredWidthPercent: WdPreferredWidthTypeToString = "wdPreferredWidthPercent"
        Case wdPreferredWidthPoints: WdPreferredWidthTypeToString = "wdPreferredWidthPoints"
    End Select
End Function
