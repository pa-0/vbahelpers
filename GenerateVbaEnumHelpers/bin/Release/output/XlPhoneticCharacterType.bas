Attribute VB_Name = "wXlPhoneticCharacterType"
Function XlPhoneticCharacterTypeFromString(value As String) As XlPhoneticCharacterType
    If IsNumeric(value) Then
        XlPhoneticCharacterTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlKatakanaHalf": XlPhoneticCharacterTypeFromString = xlKatakanaHalf
        Case "xlKatakana": XlPhoneticCharacterTypeFromString = xlKatakana
        Case "xlHiragana": XlPhoneticCharacterTypeFromString = xlHiragana
        Case "xlNoConversion": XlPhoneticCharacterTypeFromString = xlNoConversion
    End Select
End Function

Function XlPhoneticCharacterTypeToString(value As XlPhoneticCharacterType) As String
    Select Case value
        Case xlKatakanaHalf: XlPhoneticCharacterTypeToString = "xlKatakanaHalf"
        Case xlKatakana: XlPhoneticCharacterTypeToString = "xlKatakana"
        Case xlHiragana: XlPhoneticCharacterTypeToString = "xlHiragana"
        Case xlNoConversion: XlPhoneticCharacterTypeToString = "xlNoConversion"
    End Select
End Function
