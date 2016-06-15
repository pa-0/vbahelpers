Attribute VB_Name = "wPpWindowState"
Function PpWindowStateFromString(value As String) As PpWindowState
    If IsNumeric(value) Then
        PpWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppWindowNormal": PpWindowStateFromString = ppWindowNormal
        Case "ppWindowMinimized": PpWindowStateFromString = ppWindowMinimized
        Case "ppWindowMaximized": PpWindowStateFromString = ppWindowMaximized
    End Select
End Function

Function PpWindowStateToString(value As PpWindowState) As String
    Select Case value
        Case ppWindowNormal: PpWindowStateToString = "ppWindowNormal"
        Case ppWindowMinimized: PpWindowStateToString = "ppWindowMinimized"
        Case ppWindowMaximized: PpWindowStateToString = "ppWindowMaximized"
    End Select
End Function
