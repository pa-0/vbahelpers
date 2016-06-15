Attribute VB_Name = "wPbTextDirection"
Function PbTextDirectionFromString(value As String) As PbTextDirection
    If IsNumeric(value) Then
        PbTextDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTextDirectionLeftToRight": PbTextDirectionFromString = pbTextDirectionLeftToRight
        Case "pbTextDirectionRightToLeft": PbTextDirectionFromString = pbTextDirectionRightToLeft
        Case "pbTextDirectionMixed": PbTextDirectionFromString = pbTextDirectionMixed
    End Select
End Function

Function PbTextDirectionToString(value As PbTextDirection) As String
    Select Case value
        Case pbTextDirectionLeftToRight: PbTextDirectionToString = "pbTextDirectionLeftToRight"
        Case pbTextDirectionRightToLeft: PbTextDirectionToString = "pbTextDirectionRightToLeft"
        Case pbTextDirectionMixed: PbTextDirectionToString = "pbTextDirectionMixed"
    End Select
End Function
