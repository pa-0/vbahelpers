Attribute VB_Name = "wMsoFileDialogType"
Function MsoFileDialogTypeFromString(value As String) As MsoFileDialogType
    If IsNumeric(value) Then
        MsoFileDialogTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFileDialogOpen": MsoFileDialogTypeFromString = msoFileDialogOpen
        Case "msoFileDialogSaveAs": MsoFileDialogTypeFromString = msoFileDialogSaveAs
        Case "msoFileDialogFilePicker": MsoFileDialogTypeFromString = msoFileDialogFilePicker
        Case "msoFileDialogFolderPicker": MsoFileDialogTypeFromString = msoFileDialogFolderPicker
    End Select
End Function

Function MsoFileDialogTypeToString(value As MsoFileDialogType) As String
    Select Case value
        Case msoFileDialogOpen: MsoFileDialogTypeToString = "msoFileDialogOpen"
        Case msoFileDialogSaveAs: MsoFileDialogTypeToString = "msoFileDialogSaveAs"
        Case msoFileDialogFilePicker: MsoFileDialogTypeToString = "msoFileDialogFilePicker"
        Case msoFileDialogFolderPicker: MsoFileDialogTypeToString = "msoFileDialogFolderPicker"
    End Select
End Function
