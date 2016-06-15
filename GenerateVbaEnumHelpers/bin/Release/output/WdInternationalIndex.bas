Attribute VB_Name = "wWdInternationalIndex"
Function WdInternationalIndexFromString(value As String) As WdInternationalIndex
    If IsNumeric(value) Then
        WdInternationalIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdListSeparator": WdInternationalIndexFromString = wdListSeparator
        Case "wdDecimalSeparator": WdInternationalIndexFromString = wdDecimalSeparator
        Case "wdThousandsSeparator": WdInternationalIndexFromString = wdThousandsSeparator
        Case "wdCurrencyCode": WdInternationalIndexFromString = wdCurrencyCode
        Case "wd24HourClock": WdInternationalIndexFromString = wd24HourClock
        Case "wdInternationalAM": WdInternationalIndexFromString = wdInternationalAM
        Case "wdInternationalPM": WdInternationalIndexFromString = wdInternationalPM
        Case "wdTimeSeparator": WdInternationalIndexFromString = wdTimeSeparator
        Case "wdDateSeparator": WdInternationalIndexFromString = wdDateSeparator
        Case "wdProductLanguageID": WdInternationalIndexFromString = wdProductLanguageID
    End Select
End Function

Function WdInternationalIndexToString(value As WdInternationalIndex) As String
    Select Case value
        Case wdListSeparator: WdInternationalIndexToString = "wdListSeparator"
        Case wdDecimalSeparator: WdInternationalIndexToString = "wdDecimalSeparator"
        Case wdThousandsSeparator: WdInternationalIndexToString = "wdThousandsSeparator"
        Case wdCurrencyCode: WdInternationalIndexToString = "wdCurrencyCode"
        Case wd24HourClock: WdInternationalIndexToString = "wd24HourClock"
        Case wdInternationalAM: WdInternationalIndexToString = "wdInternationalAM"
        Case wdInternationalPM: WdInternationalIndexToString = "wdInternationalPM"
        Case wdTimeSeparator: WdInternationalIndexToString = "wdTimeSeparator"
        Case wdDateSeparator: WdInternationalIndexToString = "wdDateSeparator"
        Case wdProductLanguageID: WdInternationalIndexToString = "wdProductLanguageID"
    End Select
End Function
