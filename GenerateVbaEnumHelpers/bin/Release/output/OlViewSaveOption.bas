Attribute VB_Name = "wOlViewSaveOption"
Function OlViewSaveOptionFromString(value As String) As OlViewSaveOption
    If IsNumeric(value) Then
        OlViewSaveOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olViewSaveOptionThisFolderEveryone": OlViewSaveOptionFromString = olViewSaveOptionThisFolderEveryone
        Case "olViewSaveOptionThisFolderOnlyMe": OlViewSaveOptionFromString = olViewSaveOptionThisFolderOnlyMe
        Case "olViewSaveOptionAllFoldersOfType": OlViewSaveOptionFromString = olViewSaveOptionAllFoldersOfType
    End Select
End Function

Function OlViewSaveOptionToString(value As OlViewSaveOption) As String
    Select Case value
        Case olViewSaveOptionThisFolderEveryone: OlViewSaveOptionToString = "olViewSaveOptionThisFolderEveryone"
        Case olViewSaveOptionThisFolderOnlyMe: OlViewSaveOptionToString = "olViewSaveOptionThisFolderOnlyMe"
        Case olViewSaveOptionAllFoldersOfType: OlViewSaveOptionToString = "olViewSaveOptionAllFoldersOfType"
    End Select
End Function
