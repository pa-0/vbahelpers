Attribute VB_Name = "wXlPortugueseReform"
Function XlPortugueseReformFromString(value As String) As XlPortugueseReform
    If IsNumeric(value) Then
        XlPortugueseReformFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPortuguesePreReform": XlPortugueseReformFromString = xlPortuguesePreReform
        Case "xlPortuguesePostReform": XlPortugueseReformFromString = xlPortuguesePostReform
        Case "xlPortugueseBoth": XlPortugueseReformFromString = xlPortugueseBoth
    End Select
End Function

Function XlPortugueseReformToString(value As XlPortugueseReform) As String
    Select Case value
        Case xlPortuguesePreReform: XlPortugueseReformToString = "xlPortuguesePreReform"
        Case xlPortuguesePostReform: XlPortugueseReformToString = "xlPortuguesePostReform"
        Case xlPortugueseBoth: XlPortugueseReformToString = "xlPortugueseBoth"
    End Select
End Function
