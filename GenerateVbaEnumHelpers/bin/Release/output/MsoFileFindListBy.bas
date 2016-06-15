Attribute VB_Name = "wMsoFileFindListBy"
Function MsoFileFindListByFromString(value As String) As MsoFileFindListBy
    If IsNumeric(value) Then
        MsoFileFindListByFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoListbyName": MsoFileFindListByFromString = msoListbyName
        Case "msoListbyTitle": MsoFileFindListByFromString = msoListbyTitle
    End Select
End Function

Function MsoFileFindListByToString(value As MsoFileFindListBy) As String
    Select Case value
        Case msoListbyName: MsoFileFindListByToString = "msoListbyName"
        Case msoListbyTitle: MsoFileFindListByToString = "msoListbyTitle"
    End Select
End Function
