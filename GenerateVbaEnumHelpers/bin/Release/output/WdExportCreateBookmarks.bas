Attribute VB_Name = "wWdExportCreateBookmarks"
Function WdExportCreateBookmarksFromString(value As String) As WdExportCreateBookmarks
    If IsNumeric(value) Then
        WdExportCreateBookmarksFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdExportCreateNoBookmarks": WdExportCreateBookmarksFromString = wdExportCreateNoBookmarks
        Case "wdExportCreateHeadingBookmarks": WdExportCreateBookmarksFromString = wdExportCreateHeadingBookmarks
        Case "wdExportCreateWordBookmarks": WdExportCreateBookmarksFromString = wdExportCreateWordBookmarks
    End Select
End Function

Function WdExportCreateBookmarksToString(value As WdExportCreateBookmarks) As String
    Select Case value
        Case wdExportCreateNoBookmarks: WdExportCreateBookmarksToString = "wdExportCreateNoBookmarks"
        Case wdExportCreateHeadingBookmarks: WdExportCreateBookmarksToString = "wdExportCreateHeadingBookmarks"
        Case wdExportCreateWordBookmarks: WdExportCreateBookmarksToString = "wdExportCreateWordBookmarks"
    End Select
End Function
