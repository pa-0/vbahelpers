Attribute VB_Name = "wWdNumberForm"
Function WdNumberFormFromString(value As String) As WdNumberForm
    If IsNumeric(value) Then
        WdNumberFormFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNumberFormDefault": WdNumberFormFromString = wdNumberFormDefault
        Case "wdNumberFormLining": WdNumberFormFromString = wdNumberFormLining
        Case "wdNumberFormOldStyle": WdNumberFormFromString = wdNumberFormOldStyle
    End Select
End Function

Function WdNumberFormToString(value As WdNumberForm) As String
    Select Case value
        Case wdNumberFormDefault: WdNumberFormToString = "wdNumberFormDefault"
        Case wdNumberFormLining: WdNumberFormToString = "wdNumberFormLining"
        Case wdNumberFormOldStyle: WdNumberFormToString = "wdNumberFormOldStyle"
    End Select
End Function
