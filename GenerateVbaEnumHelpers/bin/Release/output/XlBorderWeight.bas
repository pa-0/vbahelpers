Attribute VB_Name = "wXlBorderWeight"
Function XlBorderWeightFromString(value As String) As XlBorderWeight
    If IsNumeric(value) Then
        XlBorderWeightFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHairline": XlBorderWeightFromString = xlHairline
        Case "xlThin": XlBorderWeightFromString = xlThin
        Case "xlThick": XlBorderWeightFromString = xlThick
        Case "xlMedium": XlBorderWeightFromString = xlMedium
    End Select
End Function

Function XlBorderWeightToString(value As XlBorderWeight) As String
    Select Case value
        Case xlHairline: XlBorderWeightToString = "xlHairline"
        Case xlThin: XlBorderWeightToString = "xlThin"
        Case xlThick: XlBorderWeightToString = "xlThick"
        Case xlMedium: XlBorderWeightToString = "xlMedium"
    End Select
End Function
