Attribute VB_Name = "wXlCreator"
Function XlCreatorFromString(value As String) As XlCreator
    If IsNumeric(value) Then
        XlCreatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCreatorCode": XlCreatorFromString = xlCreatorCode
    End Select
End Function

Function XlCreatorToString(value As XlCreator) As String
    Select Case value
        Case xlCreatorCode: XlCreatorToString = "xlCreatorCode"
    End Select
End Function
