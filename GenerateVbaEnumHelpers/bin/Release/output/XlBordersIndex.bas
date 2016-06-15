Attribute VB_Name = "wXlBordersIndex"
Function XlBordersIndexFromString(value As String) As XlBordersIndex
    If IsNumeric(value) Then
        XlBordersIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDiagonalDown": XlBordersIndexFromString = xlDiagonalDown
        Case "xlDiagonalUp": XlBordersIndexFromString = xlDiagonalUp
        Case "xlEdgeLeft": XlBordersIndexFromString = xlEdgeLeft
        Case "xlEdgeTop": XlBordersIndexFromString = xlEdgeTop
        Case "xlEdgeBottom": XlBordersIndexFromString = xlEdgeBottom
        Case "xlEdgeRight": XlBordersIndexFromString = xlEdgeRight
        Case "xlInsideVertical": XlBordersIndexFromString = xlInsideVertical
        Case "xlInsideHorizontal": XlBordersIndexFromString = xlInsideHorizontal
    End Select
End Function

Function XlBordersIndexToString(value As XlBordersIndex) As String
    Select Case value
        Case xlDiagonalDown: XlBordersIndexToString = "xlDiagonalDown"
        Case xlDiagonalUp: XlBordersIndexToString = "xlDiagonalUp"
        Case xlEdgeLeft: XlBordersIndexToString = "xlEdgeLeft"
        Case xlEdgeTop: XlBordersIndexToString = "xlEdgeTop"
        Case xlEdgeBottom: XlBordersIndexToString = "xlEdgeBottom"
        Case xlEdgeRight: XlBordersIndexToString = "xlEdgeRight"
        Case xlInsideVertical: XlBordersIndexToString = "xlInsideVertical"
        Case xlInsideHorizontal: XlBordersIndexToString = "xlInsideHorizontal"
    End Select
End Function
