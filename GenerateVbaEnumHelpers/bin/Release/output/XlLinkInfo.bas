Attribute VB_Name = "wXlLinkInfo"
Function XlLinkInfoFromString(value As String) As XlLinkInfo
    If IsNumeric(value) Then
        XlLinkInfoFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUpdateState": XlLinkInfoFromString = xlUpdateState
        Case "xlEditionDate": XlLinkInfoFromString = xlEditionDate
        Case "xlLinkInfoStatus": XlLinkInfoFromString = xlLinkInfoStatus
    End Select
End Function

Function XlLinkInfoToString(value As XlLinkInfo) As String
    Select Case value
        Case xlUpdateState: XlLinkInfoToString = "xlUpdateState"
        Case xlEditionDate: XlLinkInfoToString = "xlEditionDate"
        Case xlLinkInfoStatus: XlLinkInfoToString = "xlLinkInfoStatus"
    End Select
End Function
