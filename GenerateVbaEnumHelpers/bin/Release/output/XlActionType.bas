Attribute VB_Name = "wXlActionType"
Function XlActionTypeFromString(value As String) As XlActionType
    If IsNumeric(value) Then
        XlActionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlActionTypeUrl": XlActionTypeFromString = xlActionTypeUrl
        Case "xlActionTypeRowset": XlActionTypeFromString = xlActionTypeRowset
        Case "xlActionTypeReport": XlActionTypeFromString = xlActionTypeReport
        Case "xlActionTypeDrillthrough": XlActionTypeFromString = xlActionTypeDrillthrough
    End Select
End Function

Function XlActionTypeToString(value As XlActionType) As String
    Select Case value
        Case xlActionTypeUrl: XlActionTypeToString = "xlActionTypeUrl"
        Case xlActionTypeRowset: XlActionTypeToString = "xlActionTypeRowset"
        Case xlActionTypeReport: XlActionTypeToString = "xlActionTypeReport"
        Case xlActionTypeDrillthrough: XlActionTypeToString = "xlActionTypeDrillthrough"
    End Select
End Function
