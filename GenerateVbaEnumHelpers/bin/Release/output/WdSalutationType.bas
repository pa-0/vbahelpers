Attribute VB_Name = "wWdSalutationType"
Function WdSalutationTypeFromString(value As String) As WdSalutationType
    If IsNumeric(value) Then
        WdSalutationTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSalutationInformal": WdSalutationTypeFromString = wdSalutationInformal
        Case "wdSalutationFormal": WdSalutationTypeFromString = wdSalutationFormal
        Case "wdSalutationBusiness": WdSalutationTypeFromString = wdSalutationBusiness
        Case "wdSalutationOther": WdSalutationTypeFromString = wdSalutationOther
    End Select
End Function

Function WdSalutationTypeToString(value As WdSalutationType) As String
    Select Case value
        Case wdSalutationInformal: WdSalutationTypeToString = "wdSalutationInformal"
        Case wdSalutationFormal: WdSalutationTypeToString = "wdSalutationFormal"
        Case wdSalutationBusiness: WdSalutationTypeToString = "wdSalutationBusiness"
        Case wdSalutationOther: WdSalutationTypeToString = "wdSalutationOther"
    End Select
End Function
