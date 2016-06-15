Attribute VB_Name = "wXlUpdateLinks"
Function XlUpdateLinksFromString(value As String) As XlUpdateLinks
    If IsNumeric(value) Then
        XlUpdateLinksFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUpdateLinksUserSetting": XlUpdateLinksFromString = xlUpdateLinksUserSetting
        Case "xlUpdateLinksNever": XlUpdateLinksFromString = xlUpdateLinksNever
        Case "xlUpdateLinksAlways": XlUpdateLinksFromString = xlUpdateLinksAlways
    End Select
End Function

Function XlUpdateLinksToString(value As XlUpdateLinks) As String
    Select Case value
        Case xlUpdateLinksUserSetting: XlUpdateLinksToString = "xlUpdateLinksUserSetting"
        Case xlUpdateLinksNever: XlUpdateLinksToString = "xlUpdateLinksNever"
        Case xlUpdateLinksAlways: XlUpdateLinksToString = "xlUpdateLinksAlways"
    End Select
End Function
