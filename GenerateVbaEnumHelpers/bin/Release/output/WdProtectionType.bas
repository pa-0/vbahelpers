Attribute VB_Name = "wWdProtectionType"
Function WdProtectionTypeFromString(value As String) As WdProtectionType
    If IsNumeric(value) Then
        WdProtectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAllowOnlyRevisions": WdProtectionTypeFromString = wdAllowOnlyRevisions
        Case "wdAllowOnlyComments": WdProtectionTypeFromString = wdAllowOnlyComments
        Case "wdAllowOnlyFormFields": WdProtectionTypeFromString = wdAllowOnlyFormFields
        Case "wdAllowOnlyReading": WdProtectionTypeFromString = wdAllowOnlyReading
        Case "wdNoProtection": WdProtectionTypeFromString = wdNoProtection
    End Select
End Function

Function WdProtectionTypeToString(value As WdProtectionType) As String
    Select Case value
        Case wdAllowOnlyRevisions: WdProtectionTypeToString = "wdAllowOnlyRevisions"
        Case wdAllowOnlyComments: WdProtectionTypeToString = "wdAllowOnlyComments"
        Case wdAllowOnlyFormFields: WdProtectionTypeToString = "wdAllowOnlyFormFields"
        Case wdAllowOnlyReading: WdProtectionTypeToString = "wdAllowOnlyReading"
        Case wdNoProtection: WdProtectionTypeToString = "wdNoProtection"
    End Select
End Function
