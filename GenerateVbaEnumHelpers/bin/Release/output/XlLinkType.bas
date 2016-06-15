Attribute VB_Name = "wXlLinkType"
Function XlLinkTypeFromString(value As String) As XlLinkType
    If IsNumeric(value) Then
        XlLinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLinkTypeExcelLinks": XlLinkTypeFromString = xlLinkTypeExcelLinks
        Case "xlLinkTypeOLELinks": XlLinkTypeFromString = xlLinkTypeOLELinks
    End Select
End Function

Function XlLinkTypeToString(value As XlLinkType) As String
    Select Case value
        Case xlLinkTypeExcelLinks: XlLinkTypeToString = "xlLinkTypeExcelLinks"
        Case xlLinkTypeOLELinks: XlLinkTypeToString = "xlLinkTypeOLELinks"
    End Select
End Function
