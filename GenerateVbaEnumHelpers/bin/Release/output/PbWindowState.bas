Attribute VB_Name = "wPbWindowState"
Function PbWindowStateFromString(value As String) As PbWindowState
    If IsNumeric(value) Then
        PbWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWindowStateMaximize": PbWindowStateFromString = pbWindowStateMaximize
        Case "pbWindowStateMinimize": PbWindowStateFromString = pbWindowStateMinimize
        Case "pbWindowStateNormal": PbWindowStateFromString = pbWindowStateNormal
    End Select
End Function

Function PbWindowStateToString(value As PbWindowState) As String
    Select Case value
        Case pbWindowStateMaximize: PbWindowStateToString = "pbWindowStateMaximize"
        Case pbWindowStateMinimize: PbWindowStateToString = "pbWindowStateMinimize"
        Case pbWindowStateNormal: PbWindowStateToString = "pbWindowStateNormal"
    End Select
End Function
