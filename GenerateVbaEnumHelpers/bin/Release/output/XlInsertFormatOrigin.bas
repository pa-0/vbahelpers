Attribute VB_Name = "wXlInsertFormatOrigin"
Function XlInsertFormatOriginFromString(value As String) As XlInsertFormatOrigin
    If IsNumeric(value) Then
        XlInsertFormatOriginFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFormatFromLeftOrAbove": XlInsertFormatOriginFromString = xlFormatFromLeftOrAbove
        Case "xlFormatFromRightOrBelow": XlInsertFormatOriginFromString = xlFormatFromRightOrBelow
    End Select
End Function

Function XlInsertFormatOriginToString(value As XlInsertFormatOrigin) As String
    Select Case value
        Case xlFormatFromLeftOrAbove: XlInsertFormatOriginToString = "xlFormatFromLeftOrAbove"
        Case xlFormatFromRightOrBelow: XlInsertFormatOriginToString = "xlFormatFromRightOrBelow"
    End Select
End Function
