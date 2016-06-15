Attribute VB_Name = "wWdStyleSheetPrecedence"
Function WdStyleSheetPrecedenceFromString(value As String) As WdStyleSheetPrecedence
    If IsNumeric(value) Then
        WdStyleSheetPrecedenceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStyleSheetPrecedenceLowest": WdStyleSheetPrecedenceFromString = wdStyleSheetPrecedenceLowest
        Case "wdStyleSheetPrecedenceHighest": WdStyleSheetPrecedenceFromString = wdStyleSheetPrecedenceHighest
        Case "wdStyleSheetPrecedenceLower": WdStyleSheetPrecedenceFromString = wdStyleSheetPrecedenceLower
        Case "wdStyleSheetPrecedenceHigher": WdStyleSheetPrecedenceFromString = wdStyleSheetPrecedenceHigher
    End Select
End Function

Function WdStyleSheetPrecedenceToString(value As WdStyleSheetPrecedence) As String
    Select Case value
        Case wdStyleSheetPrecedenceLowest: WdStyleSheetPrecedenceToString = "wdStyleSheetPrecedenceLowest"
        Case wdStyleSheetPrecedenceHighest: WdStyleSheetPrecedenceToString = "wdStyleSheetPrecedenceHighest"
        Case wdStyleSheetPrecedenceLower: WdStyleSheetPrecedenceToString = "wdStyleSheetPrecedenceLower"
        Case wdStyleSheetPrecedenceHigher: WdStyleSheetPrecedenceToString = "wdStyleSheetPrecedenceHigher"
    End Select
End Function
