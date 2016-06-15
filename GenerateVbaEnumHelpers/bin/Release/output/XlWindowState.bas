Attribute VB_Name = "wXlWindowState"
Function XlWindowStateFromString(value As String) As XlWindowState
    If IsNumeric(value) Then
        XlWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNormal": XlWindowStateFromString = xlNormal
        Case "xlMinimized": XlWindowStateFromString = xlMinimized
        Case "xlMaximized": XlWindowStateFromString = xlMaximized
    End Select
End Function

Function XlWindowStateToString(value As XlWindowState) As String
    Select Case value
        Case xlNormal: XlWindowStateToString = "xlNormal"
        Case xlMinimized: XlWindowStateToString = "xlMinimized"
        Case xlMaximized: XlWindowStateToString = "xlMaximized"
    End Select
End Function
