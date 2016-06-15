Attribute VB_Name = "wMsoMixedType"
Function MsoMixedTypeFromString(value As String) As MsoMixedType
    If IsNumeric(value) Then
        MsoMixedTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoIntegerMixed": MsoMixedTypeFromString = msoIntegerMixed
        Case "msoSingleMixed": MsoMixedTypeFromString = msoSingleMixed
    End Select
End Function

Function MsoMixedTypeToString(value As MsoMixedType) As String
    Select Case value
        Case msoIntegerMixed: MsoMixedTypeToString = "msoIntegerMixed"
        Case msoSingleMixed: MsoMixedTypeToString = "msoSingleMixed"
    End Select
End Function
