Attribute VB_Name = "wPpHTMLVersion"
Function PpHTMLVersionFromString(value As String) As PpHTMLVersion
    If IsNumeric(value) Then
        PpHTMLVersionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppHTMLv3": PpHTMLVersionFromString = ppHTMLv3
        Case "ppHTMLv4": PpHTMLVersionFromString = ppHTMLv4
        Case "ppHTMLDual": PpHTMLVersionFromString = ppHTMLDual
        Case "ppHTMLAutodetect": PpHTMLVersionFromString = ppHTMLAutodetect
    End Select
End Function

Function PpHTMLVersionToString(value As PpHTMLVersion) As String
    Select Case value
        Case ppHTMLv3: PpHTMLVersionToString = "ppHTMLv3"
        Case ppHTMLv4: PpHTMLVersionToString = "ppHTMLv4"
        Case ppHTMLDual: PpHTMLVersionToString = "ppHTMLDual"
        Case ppHTMLAutodetect: PpHTMLVersionToString = "ppHTMLAutodetect"
    End Select
End Function
