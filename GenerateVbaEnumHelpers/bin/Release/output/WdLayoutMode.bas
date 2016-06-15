Attribute VB_Name = "wWdLayoutMode"
Function WdLayoutModeFromString(value As String) As WdLayoutMode
    If IsNumeric(value) Then
        WdLayoutModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLayoutModeDefault": WdLayoutModeFromString = wdLayoutModeDefault
        Case "wdLayoutModeGrid": WdLayoutModeFromString = wdLayoutModeGrid
        Case "wdLayoutModeLineGrid": WdLayoutModeFromString = wdLayoutModeLineGrid
        Case "wdLayoutModeGenko": WdLayoutModeFromString = wdLayoutModeGenko
    End Select
End Function

Function WdLayoutModeToString(value As WdLayoutMode) As String
    Select Case value
        Case wdLayoutModeDefault: WdLayoutModeToString = "wdLayoutModeDefault"
        Case wdLayoutModeGrid: WdLayoutModeToString = "wdLayoutModeGrid"
        Case wdLayoutModeLineGrid: WdLayoutModeToString = "wdLayoutModeLineGrid"
        Case wdLayoutModeGenko: WdLayoutModeToString = "wdLayoutModeGenko"
    End Select
End Function
