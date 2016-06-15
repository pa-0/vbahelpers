Attribute VB_Name = "wXlYesNoGuess"
Function XlYesNoGuessFromString(value As String) As XlYesNoGuess
    If IsNumeric(value) Then
        XlYesNoGuessFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlGuess": XlYesNoGuessFromString = xlGuess
        Case "xlYes": XlYesNoGuessFromString = xlYes
        Case "xlNo": XlYesNoGuessFromString = xlNo
    End Select
End Function

Function XlYesNoGuessToString(value As XlYesNoGuess) As String
    Select Case value
        Case xlGuess: XlYesNoGuessToString = "xlGuess"
        Case xlYes: XlYesNoGuessToString = "xlYes"
        Case xlNo: XlYesNoGuessToString = "xlNo"
    End Select
End Function
