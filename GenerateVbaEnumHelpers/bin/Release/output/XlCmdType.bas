Attribute VB_Name = "wXlCmdType"
Function XlCmdTypeFromString(value As String) As XlCmdType
    If IsNumeric(value) Then
        XlCmdTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCmdCube": XlCmdTypeFromString = xlCmdCube
        Case "xlCmdSql": XlCmdTypeFromString = xlCmdSql
        Case "xlCmdTable": XlCmdTypeFromString = xlCmdTable
        Case "xlCmdDefault": XlCmdTypeFromString = xlCmdDefault
        Case "xlCmdList": XlCmdTypeFromString = xlCmdList
    End Select
End Function

Function XlCmdTypeToString(value As XlCmdType) As String
    Select Case value
        Case xlCmdCube: XlCmdTypeToString = "xlCmdCube"
        Case xlCmdSql: XlCmdTypeToString = "xlCmdSql"
        Case xlCmdTable: XlCmdTypeToString = "xlCmdTable"
        Case xlCmdDefault: XlCmdTypeToString = "xlCmdDefault"
        Case xlCmdList: XlCmdTypeToString = "xlCmdList"
    End Select
End Function
