Attribute VB_Name = "wPbFontSource"
Function PbFontSourceFromString(value As String) As PbFontSource
    If IsNumeric(value) Then
        PbFontSourceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFontSystem": PbFontSourceFromString = pbFontSystem
        Case "pbFontDocument": PbFontSourceFromString = pbFontDocument
        Case "pbFontMissing": PbFontSourceFromString = pbFontMissing
    End Select
End Function

Function PbFontSourceToString(value As PbFontSource) As String
    Select Case value
        Case pbFontSystem: PbFontSourceToString = "pbFontSystem"
        Case pbFontDocument: PbFontSourceToString = "pbFontDocument"
        Case pbFontMissing: PbFontSourceToString = "pbFontMissing"
    End Select
End Function
