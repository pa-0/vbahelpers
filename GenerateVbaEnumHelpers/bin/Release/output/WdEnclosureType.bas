Attribute VB_Name = "wWdEnclosureType"
Function WdEnclosureTypeFromString(value As String) As WdEnclosureType
    If IsNumeric(value) Then
        WdEnclosureTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEnclosureCircle": WdEnclosureTypeFromString = wdEnclosureCircle
        Case "wdEnclosureSquare": WdEnclosureTypeFromString = wdEnclosureSquare
        Case "wdEnclosureTriangle": WdEnclosureTypeFromString = wdEnclosureTriangle
        Case "wdEnclosureDiamond": WdEnclosureTypeFromString = wdEnclosureDiamond
    End Select
End Function

Function WdEnclosureTypeToString(value As WdEnclosureType) As String
    Select Case value
        Case wdEnclosureCircle: WdEnclosureTypeToString = "wdEnclosureCircle"
        Case wdEnclosureSquare: WdEnclosureTypeToString = "wdEnclosureSquare"
        Case wdEnclosureTriangle: WdEnclosureTypeToString = "wdEnclosureTriangle"
        Case wdEnclosureDiamond: WdEnclosureTypeToString = "wdEnclosureDiamond"
    End Select
End Function
