Attribute VB_Name = "wWdNumberSpacing"
Function WdNumberSpacingFromString(value As String) As WdNumberSpacing
    If IsNumeric(value) Then
        WdNumberSpacingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNumberSpacingDefault": WdNumberSpacingFromString = wdNumberSpacingDefault
        Case "wdNumberSpacingProportional": WdNumberSpacingFromString = wdNumberSpacingProportional
        Case "wdNumberSpacingTabular": WdNumberSpacingFromString = wdNumberSpacingTabular
    End Select
End Function

Function WdNumberSpacingToString(value As WdNumberSpacing) As String
    Select Case value
        Case wdNumberSpacingDefault: WdNumberSpacingToString = "wdNumberSpacingDefault"
        Case wdNumberSpacingProportional: WdNumberSpacingToString = "wdNumberSpacingProportional"
        Case wdNumberSpacingTabular: WdNumberSpacingToString = "wdNumberSpacingTabular"
    End Select
End Function
