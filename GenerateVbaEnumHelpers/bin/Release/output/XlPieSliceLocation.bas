Attribute VB_Name = "wXlPieSliceLocation"
Function XlPieSliceLocationFromString(value As String) As XlPieSliceLocation
    If IsNumeric(value) Then
        XlPieSliceLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHorizontalCoordinate": XlPieSliceLocationFromString = xlHorizontalCoordinate
        Case "xlVerticalCoordinate": XlPieSliceLocationFromString = xlVerticalCoordinate
    End Select
End Function

Function XlPieSliceLocationToString(value As XlPieSliceLocation) As String
    Select Case value
        Case xlHorizontalCoordinate: XlPieSliceLocationToString = "xlHorizontalCoordinate"
        Case xlVerticalCoordinate: XlPieSliceLocationToString = "xlVerticalCoordinate"
    End Select
End Function
