Attribute VB_Name = "wOlFormRegionMode"
Function OlFormRegionModeFromString(value As String) As OlFormRegionMode
    If IsNumeric(value) Then
        OlFormRegionModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormRegionRead": OlFormRegionModeFromString = olFormRegionRead
        Case "olFormRegionCompose": OlFormRegionModeFromString = olFormRegionCompose
        Case "olFormRegionPreview": OlFormRegionModeFromString = olFormRegionPreview
    End Select
End Function

Function OlFormRegionModeToString(value As OlFormRegionMode) As String
    Select Case value
        Case olFormRegionRead: OlFormRegionModeToString = "olFormRegionRead"
        Case olFormRegionCompose: OlFormRegionModeToString = "olFormRegionCompose"
        Case olFormRegionPreview: OlFormRegionModeToString = "olFormRegionPreview"
    End Select
End Function
