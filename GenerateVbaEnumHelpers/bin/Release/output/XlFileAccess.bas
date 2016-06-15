Attribute VB_Name = "wXlFileAccess"
Function XlFileAccessFromString(value As String) As XlFileAccess
    If IsNumeric(value) Then
        XlFileAccessFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlReadWrite": XlFileAccessFromString = xlReadWrite
        Case "xlReadOnly": XlFileAccessFromString = xlReadOnly
    End Select
End Function

Function XlFileAccessToString(value As XlFileAccess) As String
    Select Case value
        Case xlReadWrite: XlFileAccessToString = "xlReadWrite"
        Case xlReadOnly: XlFileAccessToString = "xlReadOnly"
    End Select
End Function
