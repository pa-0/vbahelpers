Attribute VB_Name = "wXlLink"
Function XlLinkFromString(value As String) As XlLink
    If IsNumeric(value) Then
        XlLinkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlExcelLinks": XlLinkFromString = xlExcelLinks
        Case "xlOLELinks": XlLinkFromString = xlOLELinks
        Case "xlPublishers": XlLinkFromString = xlPublishers
        Case "xlSubscribers": XlLinkFromString = xlSubscribers
    End Select
End Function

Function XlLinkToString(value As XlLink) As String
    Select Case value
        Case xlExcelLinks: XlLinkToString = "xlExcelLinks"
        Case xlOLELinks: XlLinkToString = "xlOLELinks"
        Case xlPublishers: XlLinkToString = "xlPublishers"
        Case xlSubscribers: XlLinkToString = "xlSubscribers"
    End Select
End Function
