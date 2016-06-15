Attribute VB_Name = "wMsoContactCardStyle"
Function MsoContactCardStyleFromString(value As String) As MsoContactCardStyle
    If IsNumeric(value) Then
        MsoContactCardStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoContactCardHover": MsoContactCardStyleFromString = msoContactCardHover
        Case "msoContactCardFull": MsoContactCardStyleFromString = msoContactCardFull
    End Select
End Function

Function MsoContactCardStyleToString(value As MsoContactCardStyle) As String
    Select Case value
        Case msoContactCardHover: MsoContactCardStyleToString = "msoContactCardHover"
        Case msoContactCardFull: MsoContactCardStyleToString = "msoContactCardFull"
    End Select
End Function
