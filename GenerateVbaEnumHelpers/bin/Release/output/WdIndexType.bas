Attribute VB_Name = "wWdIndexType"
Function WdIndexTypeFromString(value As String) As WdIndexType
    If IsNumeric(value) Then
        WdIndexTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdIndexIndent": WdIndexTypeFromString = wdIndexIndent
        Case "wdIndexRunin": WdIndexTypeFromString = wdIndexRunin
    End Select
End Function

Function WdIndexTypeToString(value As WdIndexType) As String
    Select Case value
        Case wdIndexIndent: WdIndexTypeToString = "wdIndexIndent"
        Case wdIndexRunin: WdIndexTypeToString = "wdIndexRunin"
    End Select
End Function
