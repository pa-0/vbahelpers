Attribute VB_Name = "wXlTopBottom"
Function XlTopBottomFromString(value As String) As XlTopBottom
    If IsNumeric(value) Then
        XlTopBottomFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTop10Bottom": XlTopBottomFromString = xlTop10Bottom
        Case "xlTop10Top": XlTopBottomFromString = xlTop10Top
    End Select
End Function

Function XlTopBottomToString(value As XlTopBottom) As String
    Select Case value
        Case xlTop10Bottom: XlTopBottomToString = "xlTop10Bottom"
        Case xlTop10Top: XlTopBottomToString = "xlTop10Top"
    End Select
End Function
