Attribute VB_Name = "wXlPlacement"
Function XlPlacementFromString(value As String) As XlPlacement
    If IsNumeric(value) Then
        XlPlacementFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMoveAndSize": XlPlacementFromString = xlMoveAndSize
        Case "xlMove": XlPlacementFromString = xlMove
        Case "xlFreeFloating": XlPlacementFromString = xlFreeFloating
    End Select
End Function

Function XlPlacementToString(value As XlPlacement) As String
    Select Case value
        Case xlMoveAndSize: XlPlacementToString = "xlMoveAndSize"
        Case xlMove: XlPlacementToString = "xlMove"
        Case xlFreeFloating: XlPlacementToString = "xlFreeFloating"
    End Select
End Function
