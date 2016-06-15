Attribute VB_Name = "wOlGridLineStyle"
Function OlGridLineStyleFromString(value As String) As OlGridLineStyle
    If IsNumeric(value) Then
        OlGridLineStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olGridLineNone": OlGridLineStyleFromString = olGridLineNone
        Case "olGridLineSmallDots": OlGridLineStyleFromString = olGridLineSmallDots
        Case "olGridLineLargeDots": OlGridLineStyleFromString = olGridLineLargeDots
        Case "olGridLineDashes": OlGridLineStyleFromString = olGridLineDashes
        Case "olGridLineSolid": OlGridLineStyleFromString = olGridLineSolid
    End Select
End Function

Function OlGridLineStyleToString(value As OlGridLineStyle) As String
    Select Case value
        Case olGridLineNone: OlGridLineStyleToString = "olGridLineNone"
        Case olGridLineSmallDots: OlGridLineStyleToString = "olGridLineSmallDots"
        Case olGridLineLargeDots: OlGridLineStyleToString = "olGridLineLargeDots"
        Case olGridLineDashes: OlGridLineStyleToString = "olGridLineDashes"
        Case olGridLineSolid: OlGridLineStyleToString = "olGridLineSolid"
    End Select
End Function
