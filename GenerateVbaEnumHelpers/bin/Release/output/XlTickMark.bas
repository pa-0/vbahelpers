Attribute VB_Name = "wXlTickMark"
Function XlTickMarkFromString(value As String) As XlTickMark
    If IsNumeric(value) Then
        XlTickMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTickMarkInside": XlTickMarkFromString = xlTickMarkInside
        Case "xlTickMarkOutside": XlTickMarkFromString = xlTickMarkOutside
        Case "xlTickMarkCross": XlTickMarkFromString = xlTickMarkCross
        Case "xlTickMarkNone": XlTickMarkFromString = xlTickMarkNone
    End Select
End Function

Function XlTickMarkToString(value As XlTickMark) As String
    Select Case value
        Case xlTickMarkInside: XlTickMarkToString = "xlTickMarkInside"
        Case xlTickMarkOutside: XlTickMarkToString = "xlTickMarkOutside"
        Case xlTickMarkCross: XlTickMarkToString = "xlTickMarkCross"
        Case xlTickMarkNone: XlTickMarkToString = "xlTickMarkNone"
    End Select
End Function
