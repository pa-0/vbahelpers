Attribute VB_Name = "wXlListObjectSourceType"
Function XlListObjectSourceTypeFromString(value As String) As XlListObjectSourceType
    If IsNumeric(value) Then
        XlListObjectSourceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSrcExternal": XlListObjectSourceTypeFromString = xlSrcExternal
        Case "xlSrcRange": XlListObjectSourceTypeFromString = xlSrcRange
        Case "xlSrcXml": XlListObjectSourceTypeFromString = xlSrcXml
        Case "xlSrcQuery": XlListObjectSourceTypeFromString = xlSrcQuery
    End Select
End Function

Function XlListObjectSourceTypeToString(value As XlListObjectSourceType) As String
    Select Case value
        Case xlSrcExternal: XlListObjectSourceTypeToString = "xlSrcExternal"
        Case xlSrcRange: XlListObjectSourceTypeToString = "xlSrcRange"
        Case xlSrcXml: XlListObjectSourceTypeToString = "xlSrcXml"
        Case xlSrcQuery: XlListObjectSourceTypeToString = "xlSrcQuery"
    End Select
End Function
