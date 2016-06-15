Attribute VB_Name = "wXlRunAutoMacro"
Function XlRunAutoMacroFromString(value As String) As XlRunAutoMacro
    If IsNumeric(value) Then
        XlRunAutoMacroFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAutoOpen": XlRunAutoMacroFromString = xlAutoOpen
        Case "xlAutoClose": XlRunAutoMacroFromString = xlAutoClose
        Case "xlAutoActivate": XlRunAutoMacroFromString = xlAutoActivate
        Case "xlAutoDeactivate": XlRunAutoMacroFromString = xlAutoDeactivate
    End Select
End Function

Function XlRunAutoMacroToString(value As XlRunAutoMacro) As String
    Select Case value
        Case xlAutoOpen: XlRunAutoMacroToString = "xlAutoOpen"
        Case xlAutoClose: XlRunAutoMacroToString = "xlAutoClose"
        Case xlAutoActivate: XlRunAutoMacroToString = "xlAutoActivate"
        Case xlAutoDeactivate: XlRunAutoMacroToString = "xlAutoDeactivate"
    End Select
End Function
