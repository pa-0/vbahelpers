Attribute VB_Name = "wOlFolderDisplayMode"
Function OlFolderDisplayModeFromString(value As String) As OlFolderDisplayMode
    If IsNumeric(value) Then
        OlFolderDisplayModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFolderDisplayNormal": OlFolderDisplayModeFromString = olFolderDisplayNormal
        Case "olFolderDisplayFolderOnly": OlFolderDisplayModeFromString = olFolderDisplayFolderOnly
        Case "olFolderDisplayNoNavigation": OlFolderDisplayModeFromString = olFolderDisplayNoNavigation
    End Select
End Function

Function OlFolderDisplayModeToString(value As OlFolderDisplayMode) As String
    Select Case value
        Case olFolderDisplayNormal: OlFolderDisplayModeToString = "olFolderDisplayNormal"
        Case olFolderDisplayFolderOnly: OlFolderDisplayModeToString = "olFolderDisplayFolderOnly"
        Case olFolderDisplayNoNavigation: OlFolderDisplayModeToString = "olFolderDisplayNoNavigation"
    End Select
End Function
