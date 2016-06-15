Attribute VB_Name = "wPpBorderType"
Function PpBorderTypeFromString(value As String) As PpBorderType
    If IsNumeric(value) Then
        PpBorderTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppBorderTop": PpBorderTypeFromString = ppBorderTop
        Case "ppBorderLeft": PpBorderTypeFromString = ppBorderLeft
        Case "ppBorderBottom": PpBorderTypeFromString = ppBorderBottom
        Case "ppBorderRight": PpBorderTypeFromString = ppBorderRight
        Case "ppBorderDiagonalDown": PpBorderTypeFromString = ppBorderDiagonalDown
        Case "ppBorderDiagonalUp": PpBorderTypeFromString = ppBorderDiagonalUp
    End Select
End Function

Function PpBorderTypeToString(value As PpBorderType) As String
    Select Case value
        Case ppBorderTop: PpBorderTypeToString = "ppBorderTop"
        Case ppBorderLeft: PpBorderTypeToString = "ppBorderLeft"
        Case ppBorderBottom: PpBorderTypeToString = "ppBorderBottom"
        Case ppBorderRight: PpBorderTypeToString = "ppBorderRight"
        Case ppBorderDiagonalDown: PpBorderTypeToString = "ppBorderDiagonalDown"
        Case ppBorderDiagonalUp: PpBorderTypeToString = "ppBorderDiagonalUp"
    End Select
End Function
