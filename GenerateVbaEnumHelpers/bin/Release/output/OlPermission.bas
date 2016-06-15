Attribute VB_Name = "wOlPermission"
Function OlPermissionFromString(value As String) As OlPermission
    If IsNumeric(value) Then
        OlPermissionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olUnrestricted": OlPermissionFromString = olUnrestricted
        Case "olDoNotForward": OlPermissionFromString = olDoNotForward
        Case "olPermissionTemplate": OlPermissionFromString = olPermissionTemplate
    End Select
End Function

Function OlPermissionToString(value As OlPermission) As String
    Select Case value
        Case olUnrestricted: OlPermissionToString = "olUnrestricted"
        Case olDoNotForward: OlPermissionToString = "olDoNotForward"
        Case olPermissionTemplate: OlPermissionToString = "olPermissionTemplate"
    End Select
End Function
