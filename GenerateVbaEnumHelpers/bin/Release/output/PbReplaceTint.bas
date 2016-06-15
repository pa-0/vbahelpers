Attribute VB_Name = "wPbReplaceTint"
Function PbReplaceTintFromString(value As String) As PbReplaceTint
    If IsNumeric(value) Then
        PbReplaceTintFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbReplaceTintUseDefault": PbReplaceTintFromString = pbReplaceTintUseDefault
        Case "pbReplaceTintKeepTints": PbReplaceTintFromString = pbReplaceTintKeepTints
        Case "pbReplaceTintMaintainLuminosity": PbReplaceTintFromString = pbReplaceTintMaintainLuminosity
    End Select
End Function

Function PbReplaceTintToString(value As PbReplaceTint) As String
    Select Case value
        Case pbReplaceTintUseDefault: PbReplaceTintToString = "pbReplaceTintUseDefault"
        Case pbReplaceTintKeepTints: PbReplaceTintToString = "pbReplaceTintKeepTints"
        Case pbReplaceTintMaintainLuminosity: PbReplaceTintToString = "pbReplaceTintMaintainLuminosity"
    End Select
End Function
