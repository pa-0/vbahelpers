Attribute VB_Name = "wOlSolutionScope"
Function OlSolutionScopeFromString(value As String) As OlSolutionScope
    If IsNumeric(value) Then
        OlSolutionScopeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olHideInDefaultModules": OlSolutionScopeFromString = olHideInDefaultModules
        Case "olShowInDefaultModules": OlSolutionScopeFromString = olShowInDefaultModules
    End Select
End Function

Function OlSolutionScopeToString(value As OlSolutionScope) As String
    Select Case value
        Case olHideInDefaultModules: OlSolutionScopeToString = "olHideInDefaultModules"
        Case olShowInDefaultModules: OlSolutionScopeToString = "olShowInDefaultModules"
    End Select
End Function
