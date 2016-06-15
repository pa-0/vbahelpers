Attribute VB_Name = "wWdLetterStyle"
Function WdLetterStyleFromString(value As String) As WdLetterStyle
    If IsNumeric(value) Then
        WdLetterStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFullBlock": WdLetterStyleFromString = wdFullBlock
        Case "wdModifiedBlock": WdLetterStyleFromString = wdModifiedBlock
        Case "wdSemiBlock": WdLetterStyleFromString = wdSemiBlock
    End Select
End Function

Function WdLetterStyleToString(value As WdLetterStyle) As String
    Select Case value
        Case wdFullBlock: WdLetterStyleToString = "wdFullBlock"
        Case wdModifiedBlock: WdLetterStyleToString = "wdModifiedBlock"
        Case wdSemiBlock: WdLetterStyleToString = "wdSemiBlock"
    End Select
End Function
