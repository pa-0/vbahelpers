Attribute VB_Name = "wWdIndexFilter"
Function WdIndexFilterFromString(value As String) As WdIndexFilter
    If IsNumeric(value) Then
        WdIndexFilterFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdIndexFilterNone": WdIndexFilterFromString = wdIndexFilterNone
        Case "wdIndexFilterAiueo": WdIndexFilterFromString = wdIndexFilterAiueo
        Case "wdIndexFilterAkasatana": WdIndexFilterFromString = wdIndexFilterAkasatana
        Case "wdIndexFilterChosung": WdIndexFilterFromString = wdIndexFilterChosung
        Case "wdIndexFilterLow": WdIndexFilterFromString = wdIndexFilterLow
        Case "wdIndexFilterMedium": WdIndexFilterFromString = wdIndexFilterMedium
        Case "wdIndexFilterFull": WdIndexFilterFromString = wdIndexFilterFull
    End Select
End Function

Function WdIndexFilterToString(value As WdIndexFilter) As String
    Select Case value
        Case wdIndexFilterNone: WdIndexFilterToString = "wdIndexFilterNone"
        Case wdIndexFilterAiueo: WdIndexFilterToString = "wdIndexFilterAiueo"
        Case wdIndexFilterAkasatana: WdIndexFilterToString = "wdIndexFilterAkasatana"
        Case wdIndexFilterChosung: WdIndexFilterToString = "wdIndexFilterChosung"
        Case wdIndexFilterLow: WdIndexFilterToString = "wdIndexFilterLow"
        Case wdIndexFilterMedium: WdIndexFilterToString = "wdIndexFilterMedium"
        Case wdIndexFilterFull: WdIndexFilterToString = "wdIndexFilterFull"
    End Select
End Function
