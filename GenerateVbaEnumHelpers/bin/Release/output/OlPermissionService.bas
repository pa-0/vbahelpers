Attribute VB_Name = "wOlPermissionService"
Function OlPermissionServiceFromString(value As String) As OlPermissionService
    If IsNumeric(value) Then
        OlPermissionServiceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUnknown": OlPermissionServiceFromString = olUnknown
        Case "olWindows": OlPermissionServiceFromString = olWindows
        Case "olPassport": OlPermissionServiceFromString = olPassport
    End Select
End Function

Function OlPermissionServiceToString(value As OlPermissionService) As String
    Select Case value
        Case olUnknown: OlPermissionServiceToString = "olUnknown"
        Case olWindows: OlPermissionServiceToString = "olWindows"
        Case olPassport: OlPermissionServiceToString = "olPassport"
    End Select
End Function
