Attribute VB_Name = "wXlBarShape"
Function XlBarShapeFromString(value As String) As XlBarShape
    If IsNumeric(value) Then
        XlBarShapeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBox": XlBarShapeFromString = xlBox
        Case "xlPyramidToPoint": XlBarShapeFromString = xlPyramidToPoint
        Case "xlPyramidToMax": XlBarShapeFromString = xlPyramidToMax
        Case "xlCylinder": XlBarShapeFromString = xlCylinder
        Case "xlConeToPoint": XlBarShapeFromString = xlConeToPoint
        Case "xlConeToMax": XlBarShapeFromString = xlConeToMax
    End Select
End Function

Function XlBarShapeToString(value As XlBarShape) As String
    Select Case value
        Case xlBox: XlBarShapeToString = "xlBox"
        Case xlPyramidToPoint: XlBarShapeToString = "xlPyramidToPoint"
        Case xlPyramidToMax: XlBarShapeToString = "xlPyramidToMax"
        Case xlCylinder: XlBarShapeToString = "xlCylinder"
        Case xlConeToPoint: XlBarShapeToString = "xlConeToPoint"
        Case xlConeToMax: XlBarShapeToString = "xlConeToMax"
    End Select
End Function
