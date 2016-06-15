Attribute VB_Name = "wXlSubtototalLocationType"
Function XlSubtototalLocationTypeFromString(value As String) As XlSubtototalLocationType
    If IsNumeric(value) Then
        XlSubtototalLocationTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAtTop": XlSubtototalLocationTypeFromString = xlAtTop
        Case "xlAtBottom": XlSubtototalLocationTypeFromString = xlAtBottom
    End Select
End Function

Function XlSubtototalLocationTypeToString(value As XlSubtototalLocationType) As String
    Select Case value
        Case xlAtTop: XlSubtototalLocationTypeToString = "xlAtTop"
        Case xlAtBottom: XlSubtototalLocationTypeToString = "xlAtBottom"
    End Select
End Function
