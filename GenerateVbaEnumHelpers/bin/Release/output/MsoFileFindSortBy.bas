Attribute VB_Name = "wMsoFileFindSortBy"
Function MsoFileFindSortByFromString(value As String) As MsoFileFindSortBy
    If IsNumeric(value) Then
        MsoFileFindSortByFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFileFindSortbyAuthor": MsoFileFindSortByFromString = msoFileFindSortbyAuthor
        Case "msoFileFindSortbyDateCreated": MsoFileFindSortByFromString = msoFileFindSortbyDateCreated
        Case "msoFileFindSortbyLastSavedBy": MsoFileFindSortByFromString = msoFileFindSortbyLastSavedBy
        Case "msoFileFindSortbyDateSaved": MsoFileFindSortByFromString = msoFileFindSortbyDateSaved
        Case "msoFileFindSortbyFileName": MsoFileFindSortByFromString = msoFileFindSortbyFileName
        Case "msoFileFindSortbySize": MsoFileFindSortByFromString = msoFileFindSortbySize
        Case "msoFileFindSortbyTitle": MsoFileFindSortByFromString = msoFileFindSortbyTitle
    End Select
End Function

Function MsoFileFindSortByToString(value As MsoFileFindSortBy) As String
    Select Case value
        Case msoFileFindSortbyAuthor: MsoFileFindSortByToString = "msoFileFindSortbyAuthor"
        Case msoFileFindSortbyDateCreated: MsoFileFindSortByToString = "msoFileFindSortbyDateCreated"
        Case msoFileFindSortbyLastSavedBy: MsoFileFindSortByToString = "msoFileFindSortbyLastSavedBy"
        Case msoFileFindSortbyDateSaved: MsoFileFindSortByToString = "msoFileFindSortbyDateSaved"
        Case msoFileFindSortbyFileName: MsoFileFindSortByToString = "msoFileFindSortbyFileName"
        Case msoFileFindSortbySize: MsoFileFindSortByToString = "msoFileFindSortbySize"
        Case msoFileFindSortbyTitle: MsoFileFindSortByToString = "msoFileFindSortbyTitle"
    End Select
End Function
