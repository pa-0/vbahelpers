Attribute VB_Name = "wXlPictureAppearance"
Function XlPictureAppearanceFromString(value As String) As XlPictureAppearance
    If IsNumeric(value) Then
        XlPictureAppearanceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlScreen": XlPictureAppearanceFromString = xlScreen
        Case "xlPrinter": XlPictureAppearanceFromString = xlPrinter
    End Select
End Function

Function XlPictureAppearanceToString(value As XlPictureAppearance) As String
    Select Case value
        Case xlScreen: XlPictureAppearanceToString = "xlScreen"
        Case xlPrinter: XlPictureAppearanceToString = "xlPrinter"
    End Select
End Function
