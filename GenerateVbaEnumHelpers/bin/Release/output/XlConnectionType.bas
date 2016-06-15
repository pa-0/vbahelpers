Attribute VB_Name = "wXlConnectionType"
Function XlConnectionTypeFromString(value As String) As XlConnectionType
    If IsNumeric(value) Then
        XlConnectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlConnectionTypeOLEDB": XlConnectionTypeFromString = xlConnectionTypeOLEDB
        Case "xlConnectionTypeODBC": XlConnectionTypeFromString = xlConnectionTypeODBC
        Case "xlConnectionTypeXMLMAP": XlConnectionTypeFromString = xlConnectionTypeXMLMAP
        Case "xlConnectionTypeTEXT": XlConnectionTypeFromString = xlConnectionTypeTEXT
        Case "xlConnectionTypeWEB": XlConnectionTypeFromString = xlConnectionTypeWEB
    End Select
End Function

Function XlConnectionTypeToString(value As XlConnectionType) As String
    Select Case value
        Case xlConnectionTypeOLEDB: XlConnectionTypeToString = "xlConnectionTypeOLEDB"
        Case xlConnectionTypeODBC: XlConnectionTypeToString = "xlConnectionTypeODBC"
        Case xlConnectionTypeXMLMAP: XlConnectionTypeToString = "xlConnectionTypeXMLMAP"
        Case xlConnectionTypeTEXT: XlConnectionTypeToString = "xlConnectionTypeTEXT"
        Case xlConnectionTypeWEB: XlConnectionTypeToString = "xlConnectionTypeWEB"
    End Select
End Function
