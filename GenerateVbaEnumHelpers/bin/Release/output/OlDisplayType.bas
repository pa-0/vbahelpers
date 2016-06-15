Attribute VB_Name = "wOlDisplayType"
Function OlDisplayTypeFromString(value As String) As OlDisplayType
    If IsNumeric(value) Then
        OlDisplayTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUser": OlDisplayTypeFromString = olUser
        Case "olDistList": OlDisplayTypeFromString = olDistList
        Case "olForum": OlDisplayTypeFromString = olForum
        Case "olAgent": OlDisplayTypeFromString = olAgent
        Case "olOrganization": OlDisplayTypeFromString = olOrganization
        Case "olPrivateDistList": OlDisplayTypeFromString = olPrivateDistList
        Case "olRemoteUser": OlDisplayTypeFromString = olRemoteUser
    End Select
End Function

Function OlDisplayTypeToString(value As OlDisplayType) As String
    Select Case value
        Case olUser: OlDisplayTypeToString = "olUser"
        Case olDistList: OlDisplayTypeToString = "olDistList"
        Case olForum: OlDisplayTypeToString = "olForum"
        Case olAgent: OlDisplayTypeToString = "olAgent"
        Case olOrganization: OlDisplayTypeToString = "olOrganization"
        Case olPrivateDistList: OlDisplayTypeToString = "olPrivateDistList"
        Case olRemoteUser: OlDisplayTypeToString = "olRemoteUser"
    End Select
End Function
