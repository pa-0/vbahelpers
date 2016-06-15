Attribute VB_Name = "wPbFilterConjunction"
Function PbFilterConjunctionFromString(value As String) As PbFilterConjunction
    If IsNumeric(value) Then
        PbFilterConjunctionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbConjunctionAnd": PbFilterConjunctionFromString = pbConjunctionAnd
        Case "pbConjunctionOr": PbFilterConjunctionFromString = pbConjunctionOr
    End Select
End Function

Function PbFilterConjunctionToString(value As PbFilterConjunction) As String
    Select Case value
        Case pbConjunctionAnd: PbFilterConjunctionToString = "pbConjunctionAnd"
        Case pbConjunctionOr: PbFilterConjunctionToString = "pbConjunctionOr"
    End Select
End Function
