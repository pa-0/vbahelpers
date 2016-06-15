Attribute VB_Name = "wXlFixedFormatQuality"
Function XlFixedFormatQualityFromString(value As String) As XlFixedFormatQuality
    If IsNumeric(value) Then
        XlFixedFormatQualityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlQualityStandard": XlFixedFormatQualityFromString = xlQualityStandard
        Case "xlQualityMinimum": XlFixedFormatQualityFromString = xlQualityMinimum
    End Select
End Function

Function XlFixedFormatQualityToString(value As XlFixedFormatQuality) As String
    Select Case value
        Case xlQualityStandard: XlFixedFormatQualityToString = "xlQualityStandard"
        Case xlQualityMinimum: XlFixedFormatQualityToString = "xlQualityMinimum"
    End Select
End Function
