Attribute VB_Name = "wXlSearchDirection"
Function XlSearchDirectionFromString(value As String) As XlSearchDirection
    If IsNumeric(value) Then
        XlSearchDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNext": XlSearchDirectionFromString = xlNext
        Case "xlPrevious": XlSearchDirectionFromString = xlPrevious
    End Select
End Function

Function XlSearchDirectionToString(value As XlSearchDirection) As String
    Select Case value
        Case xlNext: XlSearchDirectionToString = "xlNext"
        Case xlPrevious: XlSearchDirectionToString = "xlPrevious"
    End Select
End Function
