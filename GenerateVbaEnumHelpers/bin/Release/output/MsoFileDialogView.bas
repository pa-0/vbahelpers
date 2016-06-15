Attribute VB_Name = "wMsoFileDialogView"
Function MsoFileDialogViewFromString(value As String) As MsoFileDialogView
    If IsNumeric(value) Then
        MsoFileDialogViewFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFileDialogViewList": MsoFileDialogViewFromString = msoFileDialogViewList
        Case "msoFileDialogViewDetails": MsoFileDialogViewFromString = msoFileDialogViewDetails
        Case "msoFileDialogViewProperties": MsoFileDialogViewFromString = msoFileDialogViewProperties
        Case "msoFileDialogViewPreview": MsoFileDialogViewFromString = msoFileDialogViewPreview
        Case "msoFileDialogViewThumbnail": MsoFileDialogViewFromString = msoFileDialogViewThumbnail
        Case "msoFileDialogViewLargeIcons": MsoFileDialogViewFromString = msoFileDialogViewLargeIcons
        Case "msoFileDialogViewSmallIcons": MsoFileDialogViewFromString = msoFileDialogViewSmallIcons
        Case "msoFileDialogViewWebView": MsoFileDialogViewFromString = msoFileDialogViewWebView
        Case "msoFileDialogViewTiles": MsoFileDialogViewFromString = msoFileDialogViewTiles
    End Select
End Function

Function MsoFileDialogViewToString(value As MsoFileDialogView) As String
    Select Case value
        Case msoFileDialogViewList: MsoFileDialogViewToString = "msoFileDialogViewList"
        Case msoFileDialogViewDetails: MsoFileDialogViewToString = "msoFileDialogViewDetails"
        Case msoFileDialogViewProperties: MsoFileDialogViewToString = "msoFileDialogViewProperties"
        Case msoFileDialogViewPreview: MsoFileDialogViewToString = "msoFileDialogViewPreview"
        Case msoFileDialogViewThumbnail: MsoFileDialogViewToString = "msoFileDialogViewThumbnail"
        Case msoFileDialogViewLargeIcons: MsoFileDialogViewToString = "msoFileDialogViewLargeIcons"
        Case msoFileDialogViewSmallIcons: MsoFileDialogViewToString = "msoFileDialogViewSmallIcons"
        Case msoFileDialogViewWebView: MsoFileDialogViewToString = "msoFileDialogViewWebView"
        Case msoFileDialogViewTiles: MsoFileDialogViewToString = "msoFileDialogViewTiles"
    End Select
End Function
