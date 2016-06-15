Attribute VB_Name = "wPpFileDialogType"
Function PpFileDialogTypeFromString(value As String) As PpFileDialogType
    If IsNumeric(value) Then
        PpFileDialogTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFileDialogOpen": PpFileDialogTypeFromString = ppFileDialogOpen
        Case "ppFileDialogSave": PpFileDialogTypeFromString = ppFileDialogSave
    End Select
End Function

Function PpFileDialogTypeToString(value As PpFileDialogType) As String
    Select Case value
        Case ppFileDialogOpen: PpFileDialogTypeToString = "ppFileDialogOpen"
        Case ppFileDialogSave: PpFileDialogTypeToString = "ppFileDialogSave"
    End Select
End Function
