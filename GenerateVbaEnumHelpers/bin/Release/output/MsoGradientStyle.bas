Attribute VB_Name = "wMsoGradientStyle"
Function MsoGradientStyleFromString(value As String) As MsoGradientStyle
    If IsNumeric(value) Then
        MsoGradientStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoGradientHorizontal": MsoGradientStyleFromString = msoGradientHorizontal
        Case "msoGradientVertical": MsoGradientStyleFromString = msoGradientVertical
        Case "msoGradientDiagonalUp": MsoGradientStyleFromString = msoGradientDiagonalUp
        Case "msoGradientDiagonalDown": MsoGradientStyleFromString = msoGradientDiagonalDown
        Case "msoGradientFromCorner": MsoGradientStyleFromString = msoGradientFromCorner
        Case "msoGradientFromTitle": MsoGradientStyleFromString = msoGradientFromTitle
        Case "msoGradientFromCenter": MsoGradientStyleFromString = msoGradientFromCenter
        Case "msoGradientMixed": MsoGradientStyleFromString = msoGradientMixed
    End Select
End Function

Function MsoGradientStyleToString(value As MsoGradientStyle) As String
    Select Case value
        Case msoGradientHorizontal: MsoGradientStyleToString = "msoGradientHorizontal"
        Case msoGradientVertical: MsoGradientStyleToString = "msoGradientVertical"
        Case msoGradientDiagonalUp: MsoGradientStyleToString = "msoGradientDiagonalUp"
        Case msoGradientDiagonalDown: MsoGradientStyleToString = "msoGradientDiagonalDown"
        Case msoGradientFromCorner: MsoGradientStyleToString = "msoGradientFromCorner"
        Case msoGradientFromTitle: MsoGradientStyleToString = "msoGradientFromTitle"
        Case msoGradientFromCenter: MsoGradientStyleToString = "msoGradientFromCenter"
        Case msoGradientMixed: MsoGradientStyleToString = "msoGradientMixed"
    End Select
End Function
