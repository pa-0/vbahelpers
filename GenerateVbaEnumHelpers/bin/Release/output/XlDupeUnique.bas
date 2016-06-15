Attribute VB_Name = "wXlDupeUnique"
Function XlDupeUniqueFromString(value As String) As XlDupeUnique
    If IsNumeric(value) Then
        XlDupeUniqueFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUnique": XlDupeUniqueFromString = xlUnique
        Case "xlDuplicate": XlDupeUniqueFromString = xlDuplicate
    End Select
End Function

Function XlDupeUniqueToString(value As XlDupeUnique) As String
    Select Case value
        Case xlUnique: XlDupeUniqueToString = "xlUnique"
        Case xlDuplicate: XlDupeUniqueToString = "xlDuplicate"
    End Select
End Function
