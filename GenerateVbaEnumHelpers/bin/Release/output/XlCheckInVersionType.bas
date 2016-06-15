Attribute VB_Name = "wXlCheckInVersionType"
Function XlCheckInVersionTypeFromString(value As String) As XlCheckInVersionType
    If IsNumeric(value) Then
        XlCheckInVersionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCheckInMinorVersion": XlCheckInVersionTypeFromString = xlCheckInMinorVersion
        Case "xlCheckInMajorVersion": XlCheckInVersionTypeFromString = xlCheckInMajorVersion
        Case "xlCheckInOverwriteVersion": XlCheckInVersionTypeFromString = xlCheckInOverwriteVersion
    End Select
End Function

Function XlCheckInVersionTypeToString(value As XlCheckInVersionType) As String
    Select Case value
        Case xlCheckInMinorVersion: XlCheckInVersionTypeToString = "xlCheckInMinorVersion"
        Case xlCheckInMajorVersion: XlCheckInVersionTypeToString = "xlCheckInMajorVersion"
        Case xlCheckInOverwriteVersion: XlCheckInVersionTypeToString = "xlCheckInOverwriteVersion"
    End Select
End Function
