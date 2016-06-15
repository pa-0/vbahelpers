Attribute VB_Name = "wXlTabPosition"
Function XlTabPositionFromString(value As String) As XlTabPosition
    If IsNumeric(value) Then
        XlTabPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTabPositionFirst": XlTabPositionFromString = xlTabPositionFirst
        Case "xlTabPositionLast": XlTabPositionFromString = xlTabPositionLast
    End Select
End Function

Function XlTabPositionToString(value As XlTabPosition) As String
    Select Case value
        Case xlTabPositionFirst: XlTabPositionToString = "xlTabPositionFirst"
        Case xlTabPositionLast: XlTabPositionToString = "xlTabPositionLast"
    End Select
End Function
