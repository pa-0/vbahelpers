Attribute VB_Name = "wXlObjectSize"
Function XlObjectSizeFromString(value As String) As XlObjectSize
    If IsNumeric(value) Then
        XlObjectSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlScreenSize": XlObjectSizeFromString = xlScreenSize
        Case "xlFitToPage": XlObjectSizeFromString = xlFitToPage
        Case "xlFullPage": XlObjectSizeFromString = xlFullPage
    End Select
End Function

Function XlObjectSizeToString(value As XlObjectSize) As String
    Select Case value
        Case xlScreenSize: XlObjectSizeToString = "xlScreenSize"
        Case xlFitToPage: XlObjectSizeToString = "xlFitToPage"
        Case xlFullPage: XlObjectSizeToString = "xlFullPage"
    End Select
End Function
