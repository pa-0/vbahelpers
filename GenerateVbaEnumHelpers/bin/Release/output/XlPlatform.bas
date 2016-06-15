Attribute VB_Name = "wXlPlatform"
Function XlPlatformFromString(value As String) As XlPlatform
    If IsNumeric(value) Then
        XlPlatformFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMacintosh": XlPlatformFromString = xlMacintosh
        Case "xlWindows": XlPlatformFromString = xlWindows
        Case "xlMSDOS": XlPlatformFromString = xlMSDOS
    End Select
End Function

Function XlPlatformToString(value As XlPlatform) As String
    Select Case value
        Case xlMacintosh: XlPlatformToString = "xlMacintosh"
        Case xlWindows: XlPlatformToString = "xlWindows"
        Case xlMSDOS: XlPlatformToString = "xlMSDOS"
    End Select
End Function
