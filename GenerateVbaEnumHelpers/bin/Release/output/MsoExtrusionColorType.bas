Attribute VB_Name = "wMsoExtrusionColorType"
Function MsoExtrusionColorTypeFromString(value As String) As MsoExtrusionColorType
    If IsNumeric(value) Then
        MsoExtrusionColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoExtrusionColorAutomatic": MsoExtrusionColorTypeFromString = msoExtrusionColorAutomatic
        Case "msoExtrusionColorCustom": MsoExtrusionColorTypeFromString = msoExtrusionColorCustom
        Case "msoExtrusionColorTypeMixed": MsoExtrusionColorTypeFromString = msoExtrusionColorTypeMixed
    End Select
End Function

Function MsoExtrusionColorTypeToString(value As MsoExtrusionColorType) As String
    Select Case value
        Case msoExtrusionColorAutomatic: MsoExtrusionColorTypeToString = "msoExtrusionColorAutomatic"
        Case msoExtrusionColorCustom: MsoExtrusionColorTypeToString = "msoExtrusionColorCustom"
        Case msoExtrusionColorTypeMixed: MsoExtrusionColorTypeToString = "msoExtrusionColorTypeMixed"
    End Select
End Function
