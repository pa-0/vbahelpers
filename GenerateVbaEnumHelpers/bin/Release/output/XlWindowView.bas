Attribute VB_Name = "wXlWindowView"
Function XlWindowViewFromString(value As String) As XlWindowView
    If IsNumeric(value) Then
        XlWindowViewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNormalView": XlWindowViewFromString = xlNormalView
        Case "xlPageBreakPreview": XlWindowViewFromString = xlPageBreakPreview
        Case "xlPageLayoutView": XlWindowViewFromString = xlPageLayoutView
    End Select
End Function

Function XlWindowViewToString(value As XlWindowView) As String
    Select Case value
        Case xlNormalView: XlWindowViewToString = "xlNormalView"
        Case xlPageBreakPreview: XlWindowViewToString = "xlPageBreakPreview"
        Case xlPageLayoutView: XlWindowViewToString = "xlPageLayoutView"
    End Select
End Function
