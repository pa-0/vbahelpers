Attribute VB_Name = "wWdOMathShapeType"
Function WdOMathShapeTypeFromString(value As String) As WdOMathShapeType
    If IsNumeric(value) Then
        WdOMathShapeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathShapeCentered": WdOMathShapeTypeFromString = wdOMathShapeCentered
        Case "wdOMathShapeMatch": WdOMathShapeTypeFromString = wdOMathShapeMatch
    End Select
End Function

Function WdOMathShapeTypeToString(value As WdOMathShapeType) As String
    Select Case value
        Case wdOMathShapeCentered: WdOMathShapeTypeToString = "wdOMathShapeCentered"
        Case wdOMathShapeMatch: WdOMathShapeTypeToString = "wdOMathShapeMatch"
    End Select
End Function
