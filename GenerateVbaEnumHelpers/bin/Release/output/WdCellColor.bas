Attribute VB_Name = "wWdCellColor"
Function WdCellColorFromString(value As String) As WdCellColor
    If IsNumeric(value) Then
        WdCellColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCellColorNoHighlight": WdCellColorFromString = wdCellColorNoHighlight
        Case "wdCellColorPink": WdCellColorFromString = wdCellColorPink
        Case "wdCellColorLightBlue": WdCellColorFromString = wdCellColorLightBlue
        Case "wdCellColorLightYellow": WdCellColorFromString = wdCellColorLightYellow
        Case "wdCellColorLightPurple": WdCellColorFromString = wdCellColorLightPurple
        Case "wdCellColorLightOrange": WdCellColorFromString = wdCellColorLightOrange
        Case "wdCellColorLightGreen": WdCellColorFromString = wdCellColorLightGreen
        Case "wdCellColorLightGray": WdCellColorFromString = wdCellColorLightGray
        Case "wdCellColorByAuthor": WdCellColorFromString = wdCellColorByAuthor
    End Select
End Function

Function WdCellColorToString(value As WdCellColor) As String
    Select Case value
        Case wdCellColorNoHighlight: WdCellColorToString = "wdCellColorNoHighlight"
        Case wdCellColorPink: WdCellColorToString = "wdCellColorPink"
        Case wdCellColorLightBlue: WdCellColorToString = "wdCellColorLightBlue"
        Case wdCellColorLightYellow: WdCellColorToString = "wdCellColorLightYellow"
        Case wdCellColorLightPurple: WdCellColorToString = "wdCellColorLightPurple"
        Case wdCellColorLightOrange: WdCellColorToString = "wdCellColorLightOrange"
        Case wdCellColorLightGreen: WdCellColorToString = "wdCellColorLightGreen"
        Case wdCellColorLightGray: WdCellColorToString = "wdCellColorLightGray"
        Case wdCellColorByAuthor: WdCellColorToString = "wdCellColorByAuthor"
    End Select
End Function
