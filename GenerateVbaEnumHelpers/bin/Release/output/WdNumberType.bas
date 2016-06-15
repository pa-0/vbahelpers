Attribute VB_Name = "wWdNumberType"
Function WdNumberTypeFromString(value As String) As WdNumberType
    If IsNumeric(value) Then
        WdNumberTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNumberParagraph": WdNumberTypeFromString = wdNumberParagraph
        Case "wdNumberListNum": WdNumberTypeFromString = wdNumberListNum
        Case "wdNumberAllNumbers": WdNumberTypeFromString = wdNumberAllNumbers
    End Select
End Function

Function WdNumberTypeToString(value As WdNumberType) As String
    Select Case value
        Case wdNumberParagraph: WdNumberTypeToString = "wdNumberParagraph"
        Case wdNumberListNum: WdNumberTypeToString = "wdNumberListNum"
        Case wdNumberAllNumbers: WdNumberTypeToString = "wdNumberAllNumbers"
    End Select
End Function
