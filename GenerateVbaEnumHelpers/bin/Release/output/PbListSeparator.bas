Attribute VB_Name = "wPbListSeparator"
Function PbListSeparatorFromString(value As String) As PbListSeparator
    If IsNumeric(value) Then
        PbListSeparatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbListSeparatorParenthesis": PbListSeparatorFromString = pbListSeparatorParenthesis
        Case "pbListSeparatorDoubleParen": PbListSeparatorFromString = pbListSeparatorDoubleParen
        Case "pbListSeparatorPeriod": PbListSeparatorFromString = pbListSeparatorPeriod
        Case "pbListSeparatorPlain": PbListSeparatorFromString = pbListSeparatorPlain
        Case "pbListSeparatorSquare": PbListSeparatorFromString = pbListSeparatorSquare
        Case "pbListSeparatorColon": PbListSeparatorFromString = pbListSeparatorColon
        Case "pbListSeparatorDoubleSquare": PbListSeparatorFromString = pbListSeparatorDoubleSquare
        Case "pbListSeparatorDoubleHyphen": PbListSeparatorFromString = pbListSeparatorDoubleHyphen
        Case "pbListSeparatorWideComma": PbListSeparatorFromString = pbListSeparatorWideComma
    End Select
End Function

Function PbListSeparatorToString(value As PbListSeparator) As String
    Select Case value
        Case pbListSeparatorParenthesis: PbListSeparatorToString = "pbListSeparatorParenthesis"
        Case pbListSeparatorDoubleParen: PbListSeparatorToString = "pbListSeparatorDoubleParen"
        Case pbListSeparatorPeriod: PbListSeparatorToString = "pbListSeparatorPeriod"
        Case pbListSeparatorPlain: PbListSeparatorToString = "pbListSeparatorPlain"
        Case pbListSeparatorSquare: PbListSeparatorToString = "pbListSeparatorSquare"
        Case pbListSeparatorColon: PbListSeparatorToString = "pbListSeparatorColon"
        Case pbListSeparatorDoubleSquare: PbListSeparatorToString = "pbListSeparatorDoubleSquare"
        Case pbListSeparatorDoubleHyphen: PbListSeparatorToString = "pbListSeparatorDoubleHyphen"
        Case pbListSeparatorWideComma: PbListSeparatorToString = "pbListSeparatorWideComma"
    End Select
End Function
