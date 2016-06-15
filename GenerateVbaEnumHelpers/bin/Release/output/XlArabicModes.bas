Attribute VB_Name = "wXlArabicModes"
Function XlArabicModesFromString(value As String) As XlArabicModes
    If IsNumeric(value) Then
        XlArabicModesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlArabicNone": XlArabicModesFromString = xlArabicNone
        Case "xlArabicStrictAlefHamza": XlArabicModesFromString = xlArabicStrictAlefHamza
        Case "xlArabicStrictFinalYaa": XlArabicModesFromString = xlArabicStrictFinalYaa
        Case "xlArabicBothStrict": XlArabicModesFromString = xlArabicBothStrict
    End Select
End Function

Function XlArabicModesToString(value As XlArabicModes) As String
    Select Case value
        Case xlArabicNone: XlArabicModesToString = "xlArabicNone"
        Case xlArabicStrictAlefHamza: XlArabicModesToString = "xlArabicStrictAlefHamza"
        Case xlArabicStrictFinalYaa: XlArabicModesToString = "xlArabicStrictFinalYaa"
        Case xlArabicBothStrict: XlArabicModesToString = "xlArabicBothStrict"
    End Select
End Function
