Attribute VB_Name = "wMsoColorType"
Function MsoColorTypeFromString(value As String) As MsoColorType
    If IsNumeric(value) Then
        MsoColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoColorTypeRGB": MsoColorTypeFromString = msoColorTypeRGB
        Case "msoColorTypeScheme": MsoColorTypeFromString = msoColorTypeScheme
        Case "msoColorTypeCMYK": MsoColorTypeFromString = msoColorTypeCMYK
        Case "msoColorTypeCMS": MsoColorTypeFromString = msoColorTypeCMS
        Case "msoColorTypeInk": MsoColorTypeFromString = msoColorTypeInk
        Case "msoColorTypeMixed": MsoColorTypeFromString = msoColorTypeMixed
    End Select
End Function

Function MsoColorTypeToString(value As MsoColorType) As String
    Select Case value
        Case msoColorTypeRGB: MsoColorTypeToString = "msoColorTypeRGB"
        Case msoColorTypeScheme: MsoColorTypeToString = "msoColorTypeScheme"
        Case msoColorTypeCMYK: MsoColorTypeToString = "msoColorTypeCMYK"
        Case msoColorTypeCMS: MsoColorTypeToString = "msoColorTypeCMS"
        Case msoColorTypeInk: MsoColorTypeToString = "msoColorTypeInk"
        Case msoColorTypeMixed: MsoColorTypeToString = "msoColorTypeMixed"
    End Select
End Function
