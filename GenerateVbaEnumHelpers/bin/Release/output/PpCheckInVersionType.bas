Attribute VB_Name = "wPpCheckInVersionType"
Function PpCheckInVersionTypeFromString(value As String) As PpCheckInVersionType
    If IsNumeric(value) Then
        PpCheckInVersionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppCheckInMinorVersion": PpCheckInVersionTypeFromString = ppCheckInMinorVersion
        Case "ppCheckInMajorVersion": PpCheckInVersionTypeFromString = ppCheckInMajorVersion
        Case "ppCheckInOverwriteVersion": PpCheckInVersionTypeFromString = ppCheckInOverwriteVersion
    End Select
End Function

Function PpCheckInVersionTypeToString(value As PpCheckInVersionType) As String
    Select Case value
        Case ppCheckInMinorVersion: PpCheckInVersionTypeToString = "ppCheckInMinorVersion"
        Case ppCheckInMajorVersion: PpCheckInVersionTypeToString = "ppCheckInMajorVersion"
        Case ppCheckInOverwriteVersion: PpCheckInVersionTypeToString = "ppCheckInOverwriteVersion"
    End Select
End Function
