Attribute VB_Name = "wXlLinkInfoType"
Function XlLinkInfoTypeFromString(value As String) As XlLinkInfoType
    If IsNumeric(value) Then
        XlLinkInfoTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLinkInfoOLELinks": XlLinkInfoTypeFromString = xlLinkInfoOLELinks
        Case "xlLinkInfoPublishers": XlLinkInfoTypeFromString = xlLinkInfoPublishers
        Case "xlLinkInfoSubscribers": XlLinkInfoTypeFromString = xlLinkInfoSubscribers
    End Select
End Function

Function XlLinkInfoTypeToString(value As XlLinkInfoType) As String
    Select Case value
        Case xlLinkInfoOLELinks: XlLinkInfoTypeToString = "xlLinkInfoOLELinks"
        Case xlLinkInfoPublishers: XlLinkInfoTypeToString = "xlLinkInfoPublishers"
        Case xlLinkInfoSubscribers: XlLinkInfoTypeToString = "xlLinkInfoSubscribers"
    End Select
End Function
