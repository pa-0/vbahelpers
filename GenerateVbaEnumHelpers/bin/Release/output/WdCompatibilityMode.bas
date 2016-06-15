Attribute VB_Name = "wWdCompatibilityMode"
Function WdCompatibilityModeFromString(value As String) As WdCompatibilityMode
    If IsNumeric(value) Then
        WdCompatibilityModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWord2003": WdCompatibilityModeFromString = wdWord2003
        Case "wdWord2007": WdCompatibilityModeFromString = wdWord2007
        Case "wdWord2010": WdCompatibilityModeFromString = wdWord2010
        Case "wdCurrent": WdCompatibilityModeFromString = wdCurrent
    End Select
End Function

Function WdCompatibilityModeToString(value As WdCompatibilityMode) As String
    Select Case value
        Case wdWord2003: WdCompatibilityModeToString = "wdWord2003"
        Case wdWord2007: WdCompatibilityModeToString = "wdWord2007"
        Case wdWord2010: WdCompatibilityModeToString = "wdWord2010"
        Case wdCurrent: WdCompatibilityModeToString = "wdCurrent"
    End Select
End Function
