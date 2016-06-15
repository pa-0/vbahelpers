Attribute VB_Name = "wPpTextStyleType"
Function PpTextStyleTypeFromString(value As String) As PpTextStyleType
    If IsNumeric(value) Then
        PpTextStyleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppDefaultStyle": PpTextStyleTypeFromString = ppDefaultStyle
        Case "ppTitleStyle": PpTextStyleTypeFromString = ppTitleStyle
        Case "ppBodyStyle": PpTextStyleTypeFromString = ppBodyStyle
    End Select
End Function

Function PpTextStyleTypeToString(value As PpTextStyleType) As String
    Select Case value
        Case ppDefaultStyle: PpTextStyleTypeToString = "ppDefaultStyle"
        Case ppTitleStyle: PpTextStyleTypeToString = "ppTitleStyle"
        Case ppBodyStyle: PpTextStyleTypeToString = "ppBodyStyle"
    End Select
End Function
