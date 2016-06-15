Attribute VB_Name = "wXlStdColorScale"
Function XlStdColorScaleFromString(value As String) As XlStdColorScale
    If IsNumeric(value) Then
        XlStdColorScaleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlColorScaleRYG": XlStdColorScaleFromString = xlColorScaleRYG
        Case "xlColorScaleGYR": XlStdColorScaleFromString = xlColorScaleGYR
        Case "xlColorScaleBlackWhite": XlStdColorScaleFromString = xlColorScaleBlackWhite
        Case "xlColorScaleWhiteBlack": XlStdColorScaleFromString = xlColorScaleWhiteBlack
    End Select
End Function

Function XlStdColorScaleToString(value As XlStdColorScale) As String
    Select Case value
        Case xlColorScaleRYG: XlStdColorScaleToString = "xlColorScaleRYG"
        Case xlColorScaleGYR: XlStdColorScaleToString = "xlColorScaleGYR"
        Case xlColorScaleBlackWhite: XlStdColorScaleToString = "xlColorScaleBlackWhite"
        Case xlColorScaleWhiteBlack: XlStdColorScaleToString = "xlColorScaleWhiteBlack"
    End Select
End Function
