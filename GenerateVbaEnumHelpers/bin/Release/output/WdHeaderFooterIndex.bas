Attribute VB_Name = "wWdHeaderFooterIndex"
Function WdHeaderFooterIndexFromString(value As String) As WdHeaderFooterIndex
    If IsNumeric(value) Then
        WdHeaderFooterIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHeaderFooterPrimary": WdHeaderFooterIndexFromString = wdHeaderFooterPrimary
        Case "wdHeaderFooterFirstPage": WdHeaderFooterIndexFromString = wdHeaderFooterFirstPage
        Case "wdHeaderFooterEvenPages": WdHeaderFooterIndexFromString = wdHeaderFooterEvenPages
    End Select
End Function

Function WdHeaderFooterIndexToString(value As WdHeaderFooterIndex) As String
    Select Case value
        Case wdHeaderFooterPrimary: WdHeaderFooterIndexToString = "wdHeaderFooterPrimary"
        Case wdHeaderFooterFirstPage: WdHeaderFooterIndexToString = "wdHeaderFooterFirstPage"
        Case wdHeaderFooterEvenPages: WdHeaderFooterIndexToString = "wdHeaderFooterEvenPages"
    End Select
End Function
