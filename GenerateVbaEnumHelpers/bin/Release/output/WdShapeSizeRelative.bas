Attribute VB_Name = "wWdShapeSizeRelative"
Function WdShapeSizeRelativeFromString(value As String) As WdShapeSizeRelative
    If IsNumeric(value) Then
        WdShapeSizeRelativeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdShapeSizeRelativeNone": WdShapeSizeRelativeFromString = wdShapeSizeRelativeNone
    End Select
End Function

Function WdShapeSizeRelativeToString(value As WdShapeSizeRelative) As String
    Select Case value
        Case wdShapeSizeRelativeNone: WdShapeSizeRelativeToString = "wdShapeSizeRelativeNone"
    End Select
End Function
