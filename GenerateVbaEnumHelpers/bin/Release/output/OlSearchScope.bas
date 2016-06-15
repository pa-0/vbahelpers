Attribute VB_Name = "wOlSearchScope"
Function OlSearchScopeFromString(value As String) As OlSearchScope
    If IsNumeric(value) Then
        OlSearchScopeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olSearchScopeCurrentFolder": OlSearchScopeFromString = olSearchScopeCurrentFolder
        Case "olSearchScopeAllFolders": OlSearchScopeFromString = olSearchScopeAllFolders
        Case "olSearchScopeAllOutlookItems": OlSearchScopeFromString = olSearchScopeAllOutlookItems
        Case "olSearchScopeSubfolders": OlSearchScopeFromString = olSearchScopeSubfolders
    End Select
End Function

Function OlSearchScopeToString(value As OlSearchScope) As String
    Select Case value
        Case olSearchScopeCurrentFolder: OlSearchScopeToString = "olSearchScopeCurrentFolder"
        Case olSearchScopeAllFolders: OlSearchScopeToString = "olSearchScopeAllFolders"
        Case olSearchScopeAllOutlookItems: OlSearchScopeToString = "olSearchScopeAllOutlookItems"
        Case olSearchScopeSubfolders: OlSearchScopeToString = "olSearchScopeSubfolders"
    End Select
End Function
