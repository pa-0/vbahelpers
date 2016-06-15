Attribute VB_Name = "wWdDocPartInsertOptions"
Function WdDocPartInsertOptionsFromString(value As String) As WdDocPartInsertOptions
    If IsNumeric(value) Then
        WdDocPartInsertOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInsertContent": WdDocPartInsertOptionsFromString = wdInsertContent
        Case "wdInsertParagraph": WdDocPartInsertOptionsFromString = wdInsertParagraph
        Case "wdInsertPage": WdDocPartInsertOptionsFromString = wdInsertPage
    End Select
End Function

Function WdDocPartInsertOptionsToString(value As WdDocPartInsertOptions) As String
    Select Case value
        Case wdInsertContent: WdDocPartInsertOptionsToString = "wdInsertContent"
        Case wdInsertParagraph: WdDocPartInsertOptionsToString = "wdInsertParagraph"
        Case wdInsertPage: WdDocPartInsertOptionsToString = "wdInsertPage"
    End Select
End Function
