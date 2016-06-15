Attribute VB_Name = "wWdHorizontalLineWidthType"
Function WdHorizontalLineWidthTypeFromString(value As String) As WdHorizontalLineWidthType
    If IsNumeric(value) Then
        WdHorizontalLineWidthTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHorizontalLineFixedWidth": WdHorizontalLineWidthTypeFromString = wdHorizontalLineFixedWidth
        Case "wdHorizontalLinePercentWidth": WdHorizontalLineWidthTypeFromString = wdHorizontalLinePercentWidth
    End Select
End Function

Function WdHorizontalLineWidthTypeToString(value As WdHorizontalLineWidthType) As String
    Select Case value
        Case wdHorizontalLineFixedWidth: WdHorizontalLineWidthTypeToString = "wdHorizontalLineFixedWidth"
        Case wdHorizontalLinePercentWidth: WdHorizontalLineWidthTypeToString = "wdHorizontalLinePercentWidth"
    End Select
End Function
