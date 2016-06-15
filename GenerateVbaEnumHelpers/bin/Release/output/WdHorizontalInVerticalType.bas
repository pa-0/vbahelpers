Attribute VB_Name = "wWdHorizontalInVerticalType"
Function WdHorizontalInVerticalTypeFromString(value As String) As WdHorizontalInVerticalType
    If IsNumeric(value) Then
        WdHorizontalInVerticalTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHorizontalInVerticalNone": WdHorizontalInVerticalTypeFromString = wdHorizontalInVerticalNone
        Case "wdHorizontalInVerticalFitInLine": WdHorizontalInVerticalTypeFromString = wdHorizontalInVerticalFitInLine
        Case "wdHorizontalInVerticalResizeLine": WdHorizontalInVerticalTypeFromString = wdHorizontalInVerticalResizeLine
    End Select
End Function

Function WdHorizontalInVerticalTypeToString(value As WdHorizontalInVerticalType) As String
    Select Case value
        Case wdHorizontalInVerticalNone: WdHorizontalInVerticalTypeToString = "wdHorizontalInVerticalNone"
        Case wdHorizontalInVerticalFitInLine: WdHorizontalInVerticalTypeToString = "wdHorizontalInVerticalFitInLine"
        Case wdHorizontalInVerticalResizeLine: WdHorizontalInVerticalTypeToString = "wdHorizontalInVerticalResizeLine"
    End Select
End Function
