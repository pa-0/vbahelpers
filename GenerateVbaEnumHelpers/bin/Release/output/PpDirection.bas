Attribute VB_Name = "wPpDirection"
Function PpDirectionFromString(value As String) As PpDirection
    If IsNumeric(value) Then
        PpDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppDirectionLeftToRight": PpDirectionFromString = ppDirectionLeftToRight
        Case "ppDirectionRightToLeft": PpDirectionFromString = ppDirectionRightToLeft
        Case "ppDirectionMixed": PpDirectionFromString = ppDirectionMixed
    End Select
End Function

Function PpDirectionToString(value As PpDirection) As String
    Select Case value
        Case ppDirectionLeftToRight: PpDirectionToString = "ppDirectionLeftToRight"
        Case ppDirectionRightToLeft: PpDirectionToString = "ppDirectionRightToLeft"
        Case ppDirectionMixed: PpDirectionToString = "ppDirectionMixed"
    End Select
End Function
