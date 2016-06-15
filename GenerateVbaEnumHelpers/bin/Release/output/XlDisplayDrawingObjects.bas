Attribute VB_Name = "wXlDisplayDrawingObjects"
Function XlDisplayDrawingObjectsFromString(value As String) As XlDisplayDrawingObjects
    If IsNumeric(value) Then
        XlDisplayDrawingObjectsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPlaceholders": XlDisplayDrawingObjectsFromString = xlPlaceholders
        Case "xlHide": XlDisplayDrawingObjectsFromString = xlHide
        Case "xlDisplayShapes": XlDisplayDrawingObjectsFromString = xlDisplayShapes
    End Select
End Function

Function XlDisplayDrawingObjectsToString(value As XlDisplayDrawingObjects) As String
    Select Case value
        Case xlPlaceholders: XlDisplayDrawingObjectsToString = "xlPlaceholders"
        Case xlHide: XlDisplayDrawingObjectsToString = "xlHide"
        Case xlDisplayShapes: XlDisplayDrawingObjectsToString = "xlDisplayShapes"
    End Select
End Function
