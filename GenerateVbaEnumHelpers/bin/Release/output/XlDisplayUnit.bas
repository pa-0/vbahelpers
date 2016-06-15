Attribute VB_Name = "wXlDisplayUnit"
Function XlDisplayUnitFromString(value As String) As XlDisplayUnit
    If IsNumeric(value) Then
        XlDisplayUnitFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMillionMillions": XlDisplayUnitFromString = xlMillionMillions
        Case "xlThousandMillions": XlDisplayUnitFromString = xlThousandMillions
        Case "xlHundredMillions": XlDisplayUnitFromString = xlHundredMillions
        Case "xlTenMillions": XlDisplayUnitFromString = xlTenMillions
        Case "xlMillions": XlDisplayUnitFromString = xlMillions
        Case "xlHundredThousands": XlDisplayUnitFromString = xlHundredThousands
        Case "xlTenThousands": XlDisplayUnitFromString = xlTenThousands
        Case "xlThousands": XlDisplayUnitFromString = xlThousands
        Case "xlHundreds": XlDisplayUnitFromString = xlHundreds
    End Select
End Function

Function XlDisplayUnitToString(value As XlDisplayUnit) As String
    Select Case value
        Case xlMillionMillions: XlDisplayUnitToString = "xlMillionMillions"
        Case xlThousandMillions: XlDisplayUnitToString = "xlThousandMillions"
        Case xlHundredMillions: XlDisplayUnitToString = "xlHundredMillions"
        Case xlTenMillions: XlDisplayUnitToString = "xlTenMillions"
        Case xlMillions: XlDisplayUnitToString = "xlMillions"
        Case xlHundredThousands: XlDisplayUnitToString = "xlHundredThousands"
        Case xlTenThousands: XlDisplayUnitToString = "xlTenThousands"
        Case xlThousands: XlDisplayUnitToString = "xlThousands"
        Case xlHundreds: XlDisplayUnitToString = "xlHundreds"
    End Select
End Function
