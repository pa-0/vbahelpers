Attribute VB_Name = "wXlOartHorizontalOverflow"
Function XlOartHorizontalOverflowFromString(value As String) As XlOartHorizontalOverflow
    If IsNumeric(value) Then
        XlOartHorizontalOverflowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOartHorizontalOverflowOverflow": XlOartHorizontalOverflowFromString = xlOartHorizontalOverflowOverflow
        Case "xlOartHorizontalOverflowClip": XlOartHorizontalOverflowFromString = xlOartHorizontalOverflowClip
    End Select
End Function

Function XlOartHorizontalOverflowToString(value As XlOartHorizontalOverflow) As String
    Select Case value
        Case xlOartHorizontalOverflowOverflow: XlOartHorizontalOverflowToString = "xlOartHorizontalOverflowOverflow"
        Case xlOartHorizontalOverflowClip: XlOartHorizontalOverflowToString = "xlOartHorizontalOverflowClip"
    End Select
End Function
