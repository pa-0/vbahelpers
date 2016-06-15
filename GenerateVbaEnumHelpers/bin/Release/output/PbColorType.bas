Attribute VB_Name = "wPbColorType"
Function PbColorTypeFromString(value As String) As PbColorType
    If IsNumeric(value) Then
        PbColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbColorTypeRGB": PbColorTypeFromString = pbColorTypeRGB
        Case "pbColorTypeScheme": PbColorTypeFromString = pbColorTypeScheme
        Case "pbColorTypeCMYK": PbColorTypeFromString = pbColorTypeCMYK
        Case "pbColorTypeCMS": PbColorTypeFromString = pbColorTypeCMS
        Case "pbColorTypeInk": PbColorTypeFromString = pbColorTypeInk
        Case "pbColorTypeMixed": PbColorTypeFromString = pbColorTypeMixed
    End Select
End Function

Function PbColorTypeToString(value As PbColorType) As String
    Select Case value
        Case pbColorTypeRGB: PbColorTypeToString = "pbColorTypeRGB"
        Case pbColorTypeScheme: PbColorTypeToString = "pbColorTypeScheme"
        Case pbColorTypeCMYK: PbColorTypeToString = "pbColorTypeCMYK"
        Case pbColorTypeCMS: PbColorTypeToString = "pbColorTypeCMS"
        Case pbColorTypeInk: PbColorTypeToString = "pbColorTypeInk"
        Case pbColorTypeMixed: PbColorTypeToString = "pbColorTypeMixed"
    End Select
End Function
