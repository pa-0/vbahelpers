Attribute VB_Name = "wXlFormatFilterTypes"
Function XlFormatFilterTypesFromString(value As String) As XlFormatFilterTypes
    If IsNumeric(value) Then
        XlFormatFilterTypesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFilterBottom": XlFormatFilterTypesFromString = xlFilterBottom
        Case "xlFilterTop": XlFormatFilterTypesFromString = xlFilterTop
        Case "xlFilterBottomPercent": XlFormatFilterTypesFromString = xlFilterBottomPercent
        Case "xlFilterTopPercent": XlFormatFilterTypesFromString = xlFilterTopPercent
    End Select
End Function

Function XlFormatFilterTypesToString(value As XlFormatFilterTypes) As String
    Select Case value
        Case xlFilterBottom: XlFormatFilterTypesToString = "xlFilterBottom"
        Case xlFilterTop: XlFormatFilterTypesToString = "xlFilterTop"
        Case xlFilterBottomPercent: XlFormatFilterTypesToString = "xlFilterBottomPercent"
        Case xlFilterTopPercent: XlFormatFilterTypesToString = "xlFilterTopPercent"
    End Select
End Function
