Attribute VB_Name = "wWdBorderType"
Function WdBorderTypeFromString(value As String) As WdBorderType
    If IsNumeric(value) Then
        WdBorderTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBorderDiagonalUp": WdBorderTypeFromString = wdBorderDiagonalUp
        Case "wdBorderDiagonalDown": WdBorderTypeFromString = wdBorderDiagonalDown
        Case "wdBorderVertical": WdBorderTypeFromString = wdBorderVertical
        Case "wdBorderHorizontal": WdBorderTypeFromString = wdBorderHorizontal
        Case "wdBorderRight": WdBorderTypeFromString = wdBorderRight
        Case "wdBorderBottom": WdBorderTypeFromString = wdBorderBottom
        Case "wdBorderLeft": WdBorderTypeFromString = wdBorderLeft
        Case "wdBorderTop": WdBorderTypeFromString = wdBorderTop
    End Select
End Function

Function WdBorderTypeToString(value As WdBorderType) As String
    Select Case value
        Case wdBorderDiagonalUp: WdBorderTypeToString = "wdBorderDiagonalUp"
        Case wdBorderDiagonalDown: WdBorderTypeToString = "wdBorderDiagonalDown"
        Case wdBorderVertical: WdBorderTypeToString = "wdBorderVertical"
        Case wdBorderHorizontal: WdBorderTypeToString = "wdBorderHorizontal"
        Case wdBorderRight: WdBorderTypeToString = "wdBorderRight"
        Case wdBorderBottom: WdBorderTypeToString = "wdBorderBottom"
        Case wdBorderLeft: WdBorderTypeToString = "wdBorderLeft"
        Case wdBorderTop: WdBorderTypeToString = "wdBorderTop"
    End Select
End Function
