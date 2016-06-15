Attribute VB_Name = "wWdShapePositionRelative"
Function WdShapePositionRelativeFromString(value As String) As WdShapePositionRelative
    If IsNumeric(value) Then
        WdShapePositionRelativeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdShapePositionRelativeNone": WdShapePositionRelativeFromString = wdShapePositionRelativeNone
    End Select
End Function

Function WdShapePositionRelativeToString(value As WdShapePositionRelative) As String
    Select Case value
        Case wdShapePositionRelativeNone: WdShapePositionRelativeToString = "wdShapePositionRelativeNone"
    End Select
End Function
