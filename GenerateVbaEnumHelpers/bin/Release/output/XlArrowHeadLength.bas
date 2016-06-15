Attribute VB_Name = "wXlArrowHeadLength"
Function XlArrowHeadLengthFromString(value As String) As XlArrowHeadLength
    If IsNumeric(value) Then
        XlArrowHeadLengthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArrowHeadLengthShort": XlArrowHeadLengthFromString = xlArrowHeadLengthShort
        Case "xlArrowHeadLengthLong": XlArrowHeadLengthFromString = xlArrowHeadLengthLong
        Case "xlArrowHeadLengthMedium": XlArrowHeadLengthFromString = xlArrowHeadLengthMedium
    End Select
End Function

Function XlArrowHeadLengthToString(value As XlArrowHeadLength) As String
    Select Case value
        Case xlArrowHeadLengthShort: XlArrowHeadLengthToString = "xlArrowHeadLengthShort"
        Case xlArrowHeadLengthLong: XlArrowHeadLengthToString = "xlArrowHeadLengthLong"
        Case xlArrowHeadLengthMedium: XlArrowHeadLengthToString = "xlArrowHeadLengthMedium"
    End Select
End Function
