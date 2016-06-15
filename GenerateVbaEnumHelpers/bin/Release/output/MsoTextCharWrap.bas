Attribute VB_Name = "wMsoTextCharWrap"
Function MsoTextCharWrapFromString(value As String) As MsoTextCharWrap
    If IsNumeric(value) Then
        MsoTextCharWrapFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoNoCharWrap": MsoTextCharWrapFromString = msoNoCharWrap
        Case "msoStandardCharWrap": MsoTextCharWrapFromString = msoStandardCharWrap
        Case "msoStrictCharWrap": MsoTextCharWrapFromString = msoStrictCharWrap
        Case "msoCustomCharWrap": MsoTextCharWrapFromString = msoCustomCharWrap
        Case "msoCharWrapMixed": MsoTextCharWrapFromString = msoCharWrapMixed
    End Select
End Function

Function MsoTextCharWrapToString(value As MsoTextCharWrap) As String
    Select Case value
        Case msoNoCharWrap: MsoTextCharWrapToString = "msoNoCharWrap"
        Case msoStandardCharWrap: MsoTextCharWrapToString = "msoStandardCharWrap"
        Case msoStrictCharWrap: MsoTextCharWrapToString = "msoStrictCharWrap"
        Case msoCustomCharWrap: MsoTextCharWrapToString = "msoCustomCharWrap"
        Case msoCharWrapMixed: MsoTextCharWrapToString = "msoCharWrapMixed"
    End Select
End Function
