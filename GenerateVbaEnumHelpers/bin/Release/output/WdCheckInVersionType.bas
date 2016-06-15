Attribute VB_Name = "wWdCheckInVersionType"
Function WdCheckInVersionTypeFromString(value As String) As WdCheckInVersionType
    If IsNumeric(value) Then
        WdCheckInVersionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCheckInMinorVersion": WdCheckInVersionTypeFromString = wdCheckInMinorVersion
        Case "wdCheckInMajorVersion": WdCheckInVersionTypeFromString = wdCheckInMajorVersion
        Case "wdCheckInOverwriteVersion": WdCheckInVersionTypeFromString = wdCheckInOverwriteVersion
    End Select
End Function

Function WdCheckInVersionTypeToString(value As WdCheckInVersionType) As String
    Select Case value
        Case wdCheckInMinorVersion: WdCheckInVersionTypeToString = "wdCheckInMinorVersion"
        Case wdCheckInMajorVersion: WdCheckInVersionTypeToString = "wdCheckInMajorVersion"
        Case wdCheckInOverwriteVersion: WdCheckInVersionTypeToString = "wdCheckInOverwriteVersion"
    End Select
End Function
