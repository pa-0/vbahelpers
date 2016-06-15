Attribute VB_Name = "wXlArrowHeadWidth"
Function XlArrowHeadWidthFromString(value As String) As XlArrowHeadWidth
    If IsNumeric(value) Then
        XlArrowHeadWidthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArrowHeadWidthNarrow": XlArrowHeadWidthFromString = xlArrowHeadWidthNarrow
        Case "xlArrowHeadWidthWide": XlArrowHeadWidthFromString = xlArrowHeadWidthWide
        Case "xlArrowHeadWidthMedium": XlArrowHeadWidthFromString = xlArrowHeadWidthMedium
    End Select
End Function

Function XlArrowHeadWidthToString(value As XlArrowHeadWidth) As String
    Select Case value
        Case xlArrowHeadWidthNarrow: XlArrowHeadWidthToString = "xlArrowHeadWidthNarrow"
        Case xlArrowHeadWidthWide: XlArrowHeadWidthToString = "xlArrowHeadWidthWide"
        Case xlArrowHeadWidthMedium: XlArrowHeadWidthToString = "xlArrowHeadWidthMedium"
    End Select
End Function
