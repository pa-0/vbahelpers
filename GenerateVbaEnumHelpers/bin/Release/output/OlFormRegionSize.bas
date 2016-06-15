Attribute VB_Name = "wOlFormRegionSize"
Function OlFormRegionSizeFromString(value As String) As OlFormRegionSize
    If IsNumeric(value) Then
        OlFormRegionSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormRegionTypeSeparate": OlFormRegionSizeFromString = olFormRegionTypeSeparate
        Case "olFormRegionTypeAdjoining": OlFormRegionSizeFromString = olFormRegionTypeAdjoining
    End Select
End Function

Function OlFormRegionSizeToString(value As OlFormRegionSize) As String
    Select Case value
        Case olFormRegionTypeSeparate: OlFormRegionSizeToString = "olFormRegionTypeSeparate"
        Case olFormRegionTypeAdjoining: OlFormRegionSizeToString = "olFormRegionTypeAdjoining"
    End Select
End Function
