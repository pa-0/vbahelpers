Attribute VB_Name = "wMsoSortBy"
Function MsoSortByFromString(value As String) As MsoSortBy
    If IsNumeric(value) Then
        MsoSortByFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSortByFileName": MsoSortByFromString = msoSortByFileName
        Case "msoSortBySize": MsoSortByFromString = msoSortBySize
        Case "msoSortByFileType": MsoSortByFromString = msoSortByFileType
        Case "msoSortByLastModified": MsoSortByFromString = msoSortByLastModified
        Case "msoSortByNone": MsoSortByFromString = msoSortByNone
    End Select
End Function

Function MsoSortByToString(value As MsoSortBy) As String
    Select Case value
        Case msoSortByFileName: MsoSortByToString = "msoSortByFileName"
        Case msoSortBySize: MsoSortByToString = "msoSortBySize"
        Case msoSortByFileType: MsoSortByToString = "msoSortByFileType"
        Case msoSortByLastModified: MsoSortByToString = "msoSortByLastModified"
        Case msoSortByNone: MsoSortByToString = "msoSortByNone"
    End Select
End Function
