Attribute VB_Name = "wWdOLEPlacement"
Function WdOLEPlacementFromString(value As String) As WdOLEPlacement
    If IsNumeric(value) Then
        WdOLEPlacementFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInLine": WdOLEPlacementFromString = wdInLine
        Case "wdFloatOverText": WdOLEPlacementFromString = wdFloatOverText
    End Select
End Function

Function WdOLEPlacementToString(value As WdOLEPlacement) As String
    Select Case value
        Case wdInLine: WdOLEPlacementToString = "wdInLine"
        Case wdFloatOverText: WdOLEPlacementToString = "wdFloatOverText"
    End Select
End Function
