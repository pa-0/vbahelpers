Attribute VB_Name = "wXlLayoutRowType"
Function XlLayoutRowTypeFromString(value As String) As XlLayoutRowType
    If IsNumeric(value) Then
        XlLayoutRowTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCompactRow": XlLayoutRowTypeFromString = xlCompactRow
        Case "xlTabularRow": XlLayoutRowTypeFromString = xlTabularRow
        Case "xlOutlineRow": XlLayoutRowTypeFromString = xlOutlineRow
    End Select
End Function

Function XlLayoutRowTypeToString(value As XlLayoutRowType) As String
    Select Case value
        Case xlCompactRow: XlLayoutRowTypeToString = "xlCompactRow"
        Case xlTabularRow: XlLayoutRowTypeToString = "xlTabularRow"
        Case xlOutlineRow: XlLayoutRowTypeToString = "xlOutlineRow"
    End Select
End Function
