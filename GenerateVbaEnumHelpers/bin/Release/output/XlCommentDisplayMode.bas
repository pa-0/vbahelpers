Attribute VB_Name = "wXlCommentDisplayMode"
Function XlCommentDisplayModeFromString(value As String) As XlCommentDisplayMode
    If IsNumeric(value) Then
        XlCommentDisplayModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoIndicator": XlCommentDisplayModeFromString = xlNoIndicator
        Case "xlCommentAndIndicator": XlCommentDisplayModeFromString = xlCommentAndIndicator
        Case "xlCommentIndicatorOnly": XlCommentDisplayModeFromString = xlCommentIndicatorOnly
    End Select
End Function

Function XlCommentDisplayModeToString(value As XlCommentDisplayMode) As String
    Select Case value
        Case xlNoIndicator: XlCommentDisplayModeToString = "xlNoIndicator"
        Case xlCommentAndIndicator: XlCommentDisplayModeToString = "xlCommentAndIndicator"
        Case xlCommentIndicatorOnly: XlCommentDisplayModeToString = "xlCommentIndicatorOnly"
    End Select
End Function
