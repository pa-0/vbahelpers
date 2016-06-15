Attribute VB_Name = "wWdArabicNumeral"
Function WdArabicNumeralFromString(value As String) As WdArabicNumeral
    If IsNumeric(value) Then
        WdArabicNumeralFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNumeralArabic": WdArabicNumeralFromString = wdNumeralArabic
        Case "wdNumeralHindi": WdArabicNumeralFromString = wdNumeralHindi
        Case "wdNumeralContext": WdArabicNumeralFromString = wdNumeralContext
        Case "wdNumeralSystem": WdArabicNumeralFromString = wdNumeralSystem
    End Select
End Function

Function WdArabicNumeralToString(value As WdArabicNumeral) As String
    Select Case value
        Case wdNumeralArabic: WdArabicNumeralToString = "wdNumeralArabic"
        Case wdNumeralHindi: WdArabicNumeralToString = "wdNumeralHindi"
        Case wdNumeralContext: WdArabicNumeralToString = "wdNumeralContext"
        Case wdNumeralSystem: WdArabicNumeralToString = "wdNumeralSystem"
    End Select
End Function
