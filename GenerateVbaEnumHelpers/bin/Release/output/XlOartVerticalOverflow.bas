Attribute VB_Name = "wXlOartVerticalOverflow"
Function XlOartVerticalOverflowFromString(value As String) As XlOartVerticalOverflow
    If IsNumeric(value) Then
        XlOartVerticalOverflowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOartVerticalOverflowOverflow": XlOartVerticalOverflowFromString = xlOartVerticalOverflowOverflow
        Case "xlOartVerticalOverflowClip": XlOartVerticalOverflowFromString = xlOartVerticalOverflowClip
        Case "xlOartVerticalOverflowEllipsis": XlOartVerticalOverflowFromString = xlOartVerticalOverflowEllipsis
    End Select
End Function

Function XlOartVerticalOverflowToString(value As XlOartVerticalOverflow) As String
    Select Case value
        Case xlOartVerticalOverflowOverflow: XlOartVerticalOverflowToString = "xlOartVerticalOverflowOverflow"
        Case xlOartVerticalOverflowClip: XlOartVerticalOverflowToString = "xlOartVerticalOverflowClip"
        Case xlOartVerticalOverflowEllipsis: XlOartVerticalOverflowToString = "xlOartVerticalOverflowEllipsis"
    End Select
End Function
