Attribute VB_Name = "wMsoAnimTextUnitEffect"
Function MsoAnimTextUnitEffectFromString(value As String) As MsoAnimTextUnitEffect
    If IsNumeric(value) Then
        MsoAnimTextUnitEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimTextUnitEffectByParagraph": MsoAnimTextUnitEffectFromString = msoAnimTextUnitEffectByParagraph
        Case "msoAnimTextUnitEffectByCharacter": MsoAnimTextUnitEffectFromString = msoAnimTextUnitEffectByCharacter
        Case "msoAnimTextUnitEffectByWord": MsoAnimTextUnitEffectFromString = msoAnimTextUnitEffectByWord
        Case "msoAnimTextUnitEffectMixed": MsoAnimTextUnitEffectFromString = msoAnimTextUnitEffectMixed
    End Select
End Function

Function MsoAnimTextUnitEffectToString(value As MsoAnimTextUnitEffect) As String
    Select Case value
        Case msoAnimTextUnitEffectByParagraph: MsoAnimTextUnitEffectToString = "msoAnimTextUnitEffectByParagraph"
        Case msoAnimTextUnitEffectByCharacter: MsoAnimTextUnitEffectToString = "msoAnimTextUnitEffectByCharacter"
        Case msoAnimTextUnitEffectByWord: MsoAnimTextUnitEffectToString = "msoAnimTextUnitEffectByWord"
        Case msoAnimTextUnitEffectMixed: MsoAnimTextUnitEffectToString = "msoAnimTextUnitEffectMixed"
    End Select
End Function
