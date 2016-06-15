Attribute VB_Name = "wXlSparklineRowCol"
Function XlSparklineRowColFromString(value As String) As XlSparklineRowCol
    If IsNumeric(value) Then
        XlSparklineRowColFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSparklineNonSquare": XlSparklineRowColFromString = xlSparklineNonSquare
        Case "xlSparklineRowsSquare": XlSparklineRowColFromString = xlSparklineRowsSquare
        Case "xlSparklineColumnsSquare": XlSparklineRowColFromString = xlSparklineColumnsSquare
    End Select
End Function

Function XlSparklineRowColToString(value As XlSparklineRowCol) As String
    Select Case value
        Case xlSparklineNonSquare: XlSparklineRowColToString = "xlSparklineNonSquare"
        Case xlSparklineRowsSquare: XlSparklineRowColToString = "xlSparklineRowsSquare"
        Case xlSparklineColumnsSquare: XlSparklineRowColToString = "xlSparklineColumnsSquare"
    End Select
End Function
