Attribute VB_Name = "wWdMonthNames"
Function WdMonthNamesFromString(value As String) As WdMonthNames
    If IsNumeric(value) Then
        WdMonthNamesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMonthNamesArabic": WdMonthNamesFromString = wdMonthNamesArabic
        Case "wdMonthNamesEnglish": WdMonthNamesFromString = wdMonthNamesEnglish
        Case "wdMonthNamesFrench": WdMonthNamesFromString = wdMonthNamesFrench
    End Select
End Function

Function WdMonthNamesToString(value As WdMonthNames) As String
    Select Case value
        Case wdMonthNamesArabic: WdMonthNamesToString = "wdMonthNamesArabic"
        Case wdMonthNamesEnglish: WdMonthNamesToString = "wdMonthNamesEnglish"
        Case wdMonthNamesFrench: WdMonthNamesToString = "wdMonthNamesFrench"
    End Select
End Function
