Attribute VB_Name = "wWdStatistic"
Function WdStatisticFromString(value As String) As WdStatistic
    If IsNumeric(value) Then
        WdStatisticFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStatisticWords": WdStatisticFromString = wdStatisticWords
        Case "wdStatisticLines": WdStatisticFromString = wdStatisticLines
        Case "wdStatisticPages": WdStatisticFromString = wdStatisticPages
        Case "wdStatisticCharacters": WdStatisticFromString = wdStatisticCharacters
        Case "wdStatisticParagraphs": WdStatisticFromString = wdStatisticParagraphs
        Case "wdStatisticCharactersWithSpaces": WdStatisticFromString = wdStatisticCharactersWithSpaces
        Case "wdStatisticFarEastCharacters": WdStatisticFromString = wdStatisticFarEastCharacters
    End Select
End Function

Function WdStatisticToString(value As WdStatistic) As String
    Select Case value
        Case wdStatisticWords: WdStatisticToString = "wdStatisticWords"
        Case wdStatisticLines: WdStatisticToString = "wdStatisticLines"
        Case wdStatisticPages: WdStatisticToString = "wdStatisticPages"
        Case wdStatisticCharacters: WdStatisticToString = "wdStatisticCharacters"
        Case wdStatisticParagraphs: WdStatisticToString = "wdStatisticParagraphs"
        Case wdStatisticCharactersWithSpaces: WdStatisticToString = "wdStatisticCharactersWithSpaces"
        Case wdStatisticFarEastCharacters: WdStatisticToString = "wdStatisticFarEastCharacters"
    End Select
End Function
