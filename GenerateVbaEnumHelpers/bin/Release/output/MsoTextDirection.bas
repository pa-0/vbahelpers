Attribute VB_Name = "wMsoTextDirection"
Function MsoTextDirectionFromString(value As String) As MsoTextDirection
    If IsNumeric(value) Then
        MsoTextDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextDirectionLeftToRight": MsoTextDirectionFromString = msoTextDirectionLeftToRight
        Case "msoTextDirectionRightToLeft": MsoTextDirectionFromString = msoTextDirectionRightToLeft
        Case "msoTextDirectionMixed": MsoTextDirectionFromString = msoTextDirectionMixed
    End Select
End Function

Function MsoTextDirectionToString(value As MsoTextDirection) As String
    Select Case value
        Case msoTextDirectionLeftToRight: MsoTextDirectionToString = "msoTextDirectionLeftToRight"
        Case msoTextDirectionRightToLeft: MsoTextDirectionToString = "msoTextDirectionRightToLeft"
        Case msoTextDirectionMixed: MsoTextDirectionToString = "msoTextDirectionMixed"
    End Select
End Function
