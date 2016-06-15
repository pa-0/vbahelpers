Attribute VB_Name = "wMsoFileFindView"
Function MsoFileFindViewFromString(value As String) As MsoFileFindView
    If IsNumeric(value) Then
        MsoFileFindViewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoViewFileInfo": MsoFileFindViewFromString = msoViewFileInfo
        Case "msoViewPreview": MsoFileFindViewFromString = msoViewPreview
        Case "msoViewSummaryInfo": MsoFileFindViewFromString = msoViewSummaryInfo
    End Select
End Function

Function MsoFileFindViewToString(value As MsoFileFindView) As String
    Select Case value
        Case msoViewFileInfo: MsoFileFindViewToString = "msoViewFileInfo"
        Case msoViewPreview: MsoFileFindViewToString = "msoViewPreview"
        Case msoViewSummaryInfo: MsoFileFindViewToString = "msoViewSummaryInfo"
    End Select
End Function
