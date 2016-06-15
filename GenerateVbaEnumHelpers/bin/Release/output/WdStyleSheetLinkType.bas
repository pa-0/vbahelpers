Attribute VB_Name = "wWdStyleSheetLinkType"
Function WdStyleSheetLinkTypeFromString(value As String) As WdStyleSheetLinkType
    If IsNumeric(value) Then
        WdStyleSheetLinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStyleSheetLinkTypeLinked": WdStyleSheetLinkTypeFromString = wdStyleSheetLinkTypeLinked
        Case "wdStyleSheetLinkTypeImported": WdStyleSheetLinkTypeFromString = wdStyleSheetLinkTypeImported
    End Select
End Function

Function WdStyleSheetLinkTypeToString(value As WdStyleSheetLinkType) As String
    Select Case value
        Case wdStyleSheetLinkTypeLinked: WdStyleSheetLinkTypeToString = "wdStyleSheetLinkTypeLinked"
        Case wdStyleSheetLinkTypeImported: WdStyleSheetLinkTypeToString = "wdStyleSheetLinkTypeImported"
    End Select
End Function
