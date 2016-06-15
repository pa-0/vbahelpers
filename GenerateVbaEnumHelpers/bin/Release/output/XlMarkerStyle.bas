Attribute VB_Name = "wXlMarkerStyle"
Function XlMarkerStyleFromString(value As String) As XlMarkerStyle
    If IsNumeric(value) Then
        XlMarkerStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlMarkerStyleSquare": XlMarkerStyleFromString = xlMarkerStyleSquare
        Case "xlMarkerStyleDiamond": XlMarkerStyleFromString = xlMarkerStyleDiamond
        Case "xlMarkerStyleTriangle": XlMarkerStyleFromString = xlMarkerStyleTriangle
        Case "xlMarkerStyleStar": XlMarkerStyleFromString = xlMarkerStyleStar
        Case "xlMarkerStyleCircle": XlMarkerStyleFromString = xlMarkerStyleCircle
        Case "xlMarkerStylePlus": XlMarkerStyleFromString = xlMarkerStylePlus
        Case "xlMarkerStyleX": XlMarkerStyleFromString = xlMarkerStyleX
        Case "xlMarkerStylePicture": XlMarkerStyleFromString = xlMarkerStylePicture
        Case "xlMarkerStyleNone": XlMarkerStyleFromString = xlMarkerStyleNone
        Case "xlMarkerStyleDot": XlMarkerStyleFromString = xlMarkerStyleDot
        Case "xlMarkerStyleDash": XlMarkerStyleFromString = xlMarkerStyleDash
        Case "xlMarkerStyleAutomatic": XlMarkerStyleFromString = xlMarkerStyleAutomatic
    End Select
End Function

Function XlMarkerStyleToString(value As XlMarkerStyle) As String
    Select Case value
        Case xlMarkerStyleSquare: XlMarkerStyleToString = "xlMarkerStyleSquare"
        Case xlMarkerStyleDiamond: XlMarkerStyleToString = "xlMarkerStyleDiamond"
        Case xlMarkerStyleTriangle: XlMarkerStyleToString = "xlMarkerStyleTriangle"
        Case xlMarkerStyleStar: XlMarkerStyleToString = "xlMarkerStyleStar"
        Case xlMarkerStyleCircle: XlMarkerStyleToString = "xlMarkerStyleCircle"
        Case xlMarkerStylePlus: XlMarkerStyleToString = "xlMarkerStylePlus"
        Case xlMarkerStyleX: XlMarkerStyleToString = "xlMarkerStyleX"
        Case xlMarkerStylePicture: XlMarkerStyleToString = "xlMarkerStylePicture"
        Case xlMarkerStyleNone: XlMarkerStyleToString = "xlMarkerStyleNone"
        Case xlMarkerStyleDot: XlMarkerStyleToString = "xlMarkerStyleDot"
        Case xlMarkerStyleDash: XlMarkerStyleToString = "xlMarkerStyleDash"
        Case xlMarkerStyleAutomatic: XlMarkerStyleToString = "xlMarkerStyleAutomatic"
    End Select
End Function
