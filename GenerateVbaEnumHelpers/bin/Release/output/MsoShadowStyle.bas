Attribute VB_Name = "wMsoShadowStyle"
Function MsoShadowStyleFromString(value As String) As MsoShadowStyle
    If IsNumeric(value) Then
        MsoShadowStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoShadowStyleInnerShadow": MsoShadowStyleFromString = msoShadowStyleInnerShadow
        Case "msoShadowStyleOuterShadow": MsoShadowStyleFromString = msoShadowStyleOuterShadow
        Case "msoShadowStyleMixed": MsoShadowStyleFromString = msoShadowStyleMixed
    End Select
End Function

Function MsoShadowStyleToString(value As MsoShadowStyle) As String
    Select Case value
        Case msoShadowStyleInnerShadow: MsoShadowStyleToString = "msoShadowStyleInnerShadow"
        Case msoShadowStyleOuterShadow: MsoShadowStyleToString = "msoShadowStyleOuterShadow"
        Case msoShadowStyleMixed: MsoShadowStyleToString = "msoShadowStyleMixed"
    End Select
End Function
