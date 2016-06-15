Attribute VB_Name = "wXlDirection"
Function XlDirectionFromString(value As String) As XlDirection
    If IsNumeric(value) Then
        XlDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUp": XlDirectionFromString = xlUp
        Case "xlToRight": XlDirectionFromString = xlToRight
        Case "xlToLeft": XlDirectionFromString = xlToLeft
        Case "xlDown": XlDirectionFromString = xlDown
    End Select
End Function

Function XlDirectionToString(value As XlDirection) As String
    Select Case value
        Case xlUp: XlDirectionToString = "xlUp"
        Case xlToRight: XlDirectionToString = "xlToRight"
        Case xlToLeft: XlDirectionToString = "xlToLeft"
        Case xlDown: XlDirectionToString = "xlDown"
    End Select
End Function
