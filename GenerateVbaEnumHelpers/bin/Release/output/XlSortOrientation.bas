Attribute VB_Name = "wXlSortOrientation"
Function XlSortOrientationFromString(value As String) As XlSortOrientation
    If IsNumeric(value) Then
        XlSortOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSortColumns": XlSortOrientationFromString = xlSortColumns
        Case "xlSortRows": XlSortOrientationFromString = xlSortRows
    End Select
End Function

Function XlSortOrientationToString(value As XlSortOrientation) As String
    Select Case value
        Case xlSortColumns: XlSortOrientationToString = "xlSortColumns"
        Case xlSortRows: XlSortOrientationToString = "xlSortRows"
    End Select
End Function
