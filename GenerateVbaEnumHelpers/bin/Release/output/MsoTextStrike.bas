Attribute VB_Name = "wMsoTextStrike"
Function MsoTextStrikeFromString(value As String) As MsoTextStrike
    If IsNumeric(value) Then
        MsoTextStrikeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoNoStrike": MsoTextStrikeFromString = msoNoStrike
        Case "msoSingleStrike": MsoTextStrikeFromString = msoSingleStrike
        Case "msoDoubleStrike": MsoTextStrikeFromString = msoDoubleStrike
        Case "msoStrikeMixed": MsoTextStrikeFromString = msoStrikeMixed
    End Select
End Function

Function MsoTextStrikeToString(value As MsoTextStrike) As String
    Select Case value
        Case msoNoStrike: MsoTextStrikeToString = "msoNoStrike"
        Case msoSingleStrike: MsoTextStrikeToString = "msoSingleStrike"
        Case msoDoubleStrike: MsoTextStrikeToString = "msoDoubleStrike"
        Case msoStrikeMixed: MsoTextStrikeToString = "msoStrikeMixed"
    End Select
End Function
