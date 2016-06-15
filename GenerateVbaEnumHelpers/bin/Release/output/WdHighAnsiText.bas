Attribute VB_Name = "wWdHighAnsiText"
Function WdHighAnsiTextFromString(value As String) As WdHighAnsiText
    If IsNumeric(value) Then
        WdHighAnsiTextFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHighAnsiIsFarEast": WdHighAnsiTextFromString = wdHighAnsiIsFarEast
        Case "wdHighAnsiIsHighAnsi": WdHighAnsiTextFromString = wdHighAnsiIsHighAnsi
        Case "wdAutoDetectHighAnsiFarEast": WdHighAnsiTextFromString = wdAutoDetectHighAnsiFarEast
    End Select
End Function

Function WdHighAnsiTextToString(value As WdHighAnsiText) As String
    Select Case value
        Case wdHighAnsiIsFarEast: WdHighAnsiTextToString = "wdHighAnsiIsFarEast"
        Case wdHighAnsiIsHighAnsi: WdHighAnsiTextToString = "wdHighAnsiIsHighAnsi"
        Case wdAutoDetectHighAnsiFarEast: WdHighAnsiTextToString = "wdAutoDetectHighAnsiFarEast"
    End Select
End Function
