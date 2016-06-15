Attribute VB_Name = "wXlDataBarAxisPosition"
Function XlDataBarAxisPositionFromString(value As String) As XlDataBarAxisPosition
    If IsNumeric(value) Then
        XlDataBarAxisPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataBarAxisAutomatic": XlDataBarAxisPositionFromString = xlDataBarAxisAutomatic
        Case "xlDataBarAxisMidpoint": XlDataBarAxisPositionFromString = xlDataBarAxisMidpoint
        Case "xlDataBarAxisNone": XlDataBarAxisPositionFromString = xlDataBarAxisNone
    End Select
End Function

Function XlDataBarAxisPositionToString(value As XlDataBarAxisPosition) As String
    Select Case value
        Case xlDataBarAxisAutomatic: XlDataBarAxisPositionToString = "xlDataBarAxisAutomatic"
        Case xlDataBarAxisMidpoint: XlDataBarAxisPositionToString = "xlDataBarAxisMidpoint"
        Case xlDataBarAxisNone: XlDataBarAxisPositionToString = "xlDataBarAxisNone"
    End Select
End Function
