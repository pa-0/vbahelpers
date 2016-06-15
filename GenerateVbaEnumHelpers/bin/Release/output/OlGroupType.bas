Attribute VB_Name = "wOlGroupType"
Function OlGroupTypeFromString(value As String) As OlGroupType
    If IsNumeric(value) Then
        OlGroupTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olCustomFoldersGroup": OlGroupTypeFromString = olCustomFoldersGroup
        Case "olMyFoldersGroup": OlGroupTypeFromString = olMyFoldersGroup
        Case "olPeopleFoldersGroup": OlGroupTypeFromString = olPeopleFoldersGroup
        Case "olOtherFoldersGroup": OlGroupTypeFromString = olOtherFoldersGroup
        Case "olFavoriteFoldersGroup": OlGroupTypeFromString = olFavoriteFoldersGroup
        Case "olRoomsGroup": OlGroupTypeFromString = olRoomsGroup
        Case "olReadOnlyGroup": OlGroupTypeFromString = olReadOnlyGroup
    End Select
End Function

Function OlGroupTypeToString(value As OlGroupType) As String
    Select Case value
        Case olCustomFoldersGroup: OlGroupTypeToString = "olCustomFoldersGroup"
        Case olMyFoldersGroup: OlGroupTypeToString = "olMyFoldersGroup"
        Case olPeopleFoldersGroup: OlGroupTypeToString = "olPeopleFoldersGroup"
        Case olOtherFoldersGroup: OlGroupTypeToString = "olOtherFoldersGroup"
        Case olFavoriteFoldersGroup: OlGroupTypeToString = "olFavoriteFoldersGroup"
        Case olRoomsGroup: OlGroupTypeToString = "olRoomsGroup"
        Case olReadOnlyGroup: OlGroupTypeToString = "olReadOnlyGroup"
    End Select
End Function
