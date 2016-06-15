Attribute VB_Name = "wWdSortFieldType"
Function WdSortFieldTypeFromString(value As String) As WdSortFieldType
    If IsNumeric(value) Then
        WdSortFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSortFieldAlphanumeric": WdSortFieldTypeFromString = wdSortFieldAlphanumeric
        Case "wdSortFieldNumeric": WdSortFieldTypeFromString = wdSortFieldNumeric
        Case "wdSortFieldDate": WdSortFieldTypeFromString = wdSortFieldDate
        Case "wdSortFieldSyllable": WdSortFieldTypeFromString = wdSortFieldSyllable
        Case "wdSortFieldJapanJIS": WdSortFieldTypeFromString = wdSortFieldJapanJIS
        Case "wdSortFieldStroke": WdSortFieldTypeFromString = wdSortFieldStroke
        Case "wdSortFieldKoreaKS": WdSortFieldTypeFromString = wdSortFieldKoreaKS
    End Select
End Function

Function WdSortFieldTypeToString(value As WdSortFieldType) As String
    Select Case value
        Case wdSortFieldAlphanumeric: WdSortFieldTypeToString = "wdSortFieldAlphanumeric"
        Case wdSortFieldNumeric: WdSortFieldTypeToString = "wdSortFieldNumeric"
        Case wdSortFieldDate: WdSortFieldTypeToString = "wdSortFieldDate"
        Case wdSortFieldSyllable: WdSortFieldTypeToString = "wdSortFieldSyllable"
        Case wdSortFieldJapanJIS: WdSortFieldTypeToString = "wdSortFieldJapanJIS"
        Case wdSortFieldStroke: WdSortFieldTypeToString = "wdSortFieldStroke"
        Case wdSortFieldKoreaKS: WdSortFieldTypeToString = "wdSortFieldKoreaKS"
    End Select
End Function
