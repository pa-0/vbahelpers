Attribute VB_Name = "wXlProtectedViewWindowState"
Function XlProtectedViewWindowStateFromString(value As String) As XlProtectedViewWindowState
    If IsNumeric(value) Then
        XlProtectedViewWindowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlProtectedViewWindowNormal": XlProtectedViewWindowStateFromString = xlProtectedViewWindowNormal
        Case "xlProtectedViewWindowMinimized": XlProtectedViewWindowStateFromString = xlProtectedViewWindowMinimized
        Case "xlProtectedViewWindowMaximized": XlProtectedViewWindowStateFromString = xlProtectedViewWindowMaximized
    End Select
End Function

Function XlProtectedViewWindowStateToString(value As XlProtectedViewWindowState) As String
    Select Case value
        Case xlProtectedViewWindowNormal: XlProtectedViewWindowStateToString = "xlProtectedViewWindowNormal"
        Case xlProtectedViewWindowMinimized: XlProtectedViewWindowStateToString = "xlProtectedViewWindowMinimized"
        Case xlProtectedViewWindowMaximized: XlProtectedViewWindowStateToString = "xlProtectedViewWindowMaximized"
    End Select
End Function
