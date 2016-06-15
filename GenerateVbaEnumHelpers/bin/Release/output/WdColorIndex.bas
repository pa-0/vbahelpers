Attribute VB_Name = "wWdColorIndex"
Function WdColorIndexFromString(value As String) As WdColorIndex
    If IsNumeric(value) Then
        WdColorIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAuto": WdColorIndexFromString = wdAuto
        Case "wdNoHighlight": WdColorIndexFromString = wdNoHighlight
        Case "wdBlack": WdColorIndexFromString = wdBlack
        Case "wdBlue": WdColorIndexFromString = wdBlue
        Case "wdTurquoise": WdColorIndexFromString = wdTurquoise
        Case "wdBrightGreen": WdColorIndexFromString = wdBrightGreen
        Case "wdPink": WdColorIndexFromString = wdPink
        Case "wdRed": WdColorIndexFromString = wdRed
        Case "wdYellow": WdColorIndexFromString = wdYellow
        Case "wdWhite": WdColorIndexFromString = wdWhite
        Case "wdDarkBlue": WdColorIndexFromString = wdDarkBlue
        Case "wdTeal": WdColorIndexFromString = wdTeal
        Case "wdGreen": WdColorIndexFromString = wdGreen
        Case "wdViolet": WdColorIndexFromString = wdViolet
        Case "wdDarkRed": WdColorIndexFromString = wdDarkRed
        Case "wdDarkYellow": WdColorIndexFromString = wdDarkYellow
        Case "wdGray50": WdColorIndexFromString = wdGray50
        Case "wdGray25": WdColorIndexFromString = wdGray25
        Case "wdByAuthor": WdColorIndexFromString = wdByAuthor
    End Select
End Function

Function WdColorIndexToString(value As WdColorIndex) As String
    Select Case value
        Case wdAuto: WdColorIndexToString = "wdAuto"
        Case wdNoHighlight: WdColorIndexToString = "wdNoHighlight"
        Case wdBlack: WdColorIndexToString = "wdBlack"
        Case wdBlue: WdColorIndexToString = "wdBlue"
        Case wdTurquoise: WdColorIndexToString = "wdTurquoise"
        Case wdBrightGreen: WdColorIndexToString = "wdBrightGreen"
        Case wdPink: WdColorIndexToString = "wdPink"
        Case wdRed: WdColorIndexToString = "wdRed"
        Case wdYellow: WdColorIndexToString = "wdYellow"
        Case wdWhite: WdColorIndexToString = "wdWhite"
        Case wdDarkBlue: WdColorIndexToString = "wdDarkBlue"
        Case wdTeal: WdColorIndexToString = "wdTeal"
        Case wdGreen: WdColorIndexToString = "wdGreen"
        Case wdViolet: WdColorIndexToString = "wdViolet"
        Case wdDarkRed: WdColorIndexToString = "wdDarkRed"
        Case wdDarkYellow: WdColorIndexToString = "wdDarkYellow"
        Case wdGray50: WdColorIndexToString = "wdGray50"
        Case wdGray25: WdColorIndexToString = "wdGray25"
        Case wdByAuthor: WdColorIndexToString = "wdByAuthor"
    End Select
End Function
