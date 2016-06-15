Attribute VB_Name = "wWdWindowState"
Function WdWindowStateFromString(value As String) As WdWindowState
    If IsNumeric(value) Then
        WdWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWindowStateNormal": WdWindowStateFromString = wdWindowStateNormal
        Case "wdWindowStateMaximize": WdWindowStateFromString = wdWindowStateMaximize
        Case "wdWindowStateMinimize": WdWindowStateFromString = wdWindowStateMinimize
    End Select
End Function

Function WdWindowStateToString(value As WdWindowState) As String
    Select Case value
        Case wdWindowStateNormal: WdWindowStateToString = "wdWindowStateNormal"
        Case wdWindowStateMaximize: WdWindowStateToString = "wdWindowStateMaximize"
        Case wdWindowStateMinimize: WdWindowStateToString = "wdWindowStateMinimize"
    End Select
End Function
