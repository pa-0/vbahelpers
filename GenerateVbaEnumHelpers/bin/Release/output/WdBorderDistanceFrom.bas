Attribute VB_Name = "wWdBorderDistanceFrom"
Function WdBorderDistanceFromFromString(value As String) As WdBorderDistanceFrom
    If IsNumeric(value) Then
        WdBorderDistanceFromFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBorderDistanceFromText": WdBorderDistanceFromFromString = wdBorderDistanceFromText
        Case "wdBorderDistanceFromPageEdge": WdBorderDistanceFromFromString = wdBorderDistanceFromPageEdge
    End Select
End Function

Function WdBorderDistanceFromToString(value As WdBorderDistanceFrom) As String
    Select Case value
        Case wdBorderDistanceFromText: WdBorderDistanceFromToString = "wdBorderDistanceFromText"
        Case wdBorderDistanceFromPageEdge: WdBorderDistanceFromToString = "wdBorderDistanceFromPageEdge"
    End Select
End Function
