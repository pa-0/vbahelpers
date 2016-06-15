Attribute VB_Name = "wXlPattern"
Function XlPatternFromString(value As String) As XlPattern
    If IsNumeric(value) Then
        XlPatternFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPatternSolid": XlPatternFromString = xlPatternSolid
        Case "xlPatternChecker": XlPatternFromString = xlPatternChecker
        Case "xlPatternSemiGray75": XlPatternFromString = xlPatternSemiGray75
        Case "xlPatternLightHorizontal": XlPatternFromString = xlPatternLightHorizontal
        Case "xlPatternLightVertical": XlPatternFromString = xlPatternLightVertical
        Case "xlPatternLightDown": XlPatternFromString = xlPatternLightDown
        Case "xlPatternLightUp": XlPatternFromString = xlPatternLightUp
        Case "xlPatternGrid": XlPatternFromString = xlPatternGrid
        Case "xlPatternCrissCross": XlPatternFromString = xlPatternCrissCross
        Case "xlPatternGray16": XlPatternFromString = xlPatternGray16
        Case "xlPatternGray8": XlPatternFromString = xlPatternGray8
        Case "xlPatternLinearGradient": XlPatternFromString = xlPatternLinearGradient
        Case "xlPatternRectangularGradient": XlPatternFromString = xlPatternRectangularGradient
        Case "xlPatternVertical": XlPatternFromString = xlPatternVertical
        Case "xlPatternUp": XlPatternFromString = xlPatternUp
        Case "xlPatternNone": XlPatternFromString = xlPatternNone
        Case "xlPatternHorizontal": XlPatternFromString = xlPatternHorizontal
        Case "xlPatternGray75": XlPatternFromString = xlPatternGray75
        Case "xlPatternGray50": XlPatternFromString = xlPatternGray50
        Case "xlPatternGray25": XlPatternFromString = xlPatternGray25
        Case "xlPatternDown": XlPatternFromString = xlPatternDown
        Case "xlPatternAutomatic": XlPatternFromString = xlPatternAutomatic
    End Select
End Function

Function XlPatternToString(value As XlPattern) As String
    Select Case value
        Case xlPatternSolid: XlPatternToString = "xlPatternSolid"
        Case xlPatternChecker: XlPatternToString = "xlPatternChecker"
        Case xlPatternSemiGray75: XlPatternToString = "xlPatternSemiGray75"
        Case xlPatternLightHorizontal: XlPatternToString = "xlPatternLightHorizontal"
        Case xlPatternLightVertical: XlPatternToString = "xlPatternLightVertical"
        Case xlPatternLightDown: XlPatternToString = "xlPatternLightDown"
        Case xlPatternLightUp: XlPatternToString = "xlPatternLightUp"
        Case xlPatternGrid: XlPatternToString = "xlPatternGrid"
        Case xlPatternCrissCross: XlPatternToString = "xlPatternCrissCross"
        Case xlPatternGray16: XlPatternToString = "xlPatternGray16"
        Case xlPatternGray8: XlPatternToString = "xlPatternGray8"
        Case xlPatternLinearGradient: XlPatternToString = "xlPatternLinearGradient"
        Case xlPatternRectangularGradient: XlPatternToString = "xlPatternRectangularGradient"
        Case xlPatternVertical: XlPatternToString = "xlPatternVertical"
        Case xlPatternUp: XlPatternToString = "xlPatternUp"
        Case xlPatternNone: XlPatternToString = "xlPatternNone"
        Case xlPatternHorizontal: XlPatternToString = "xlPatternHorizontal"
        Case xlPatternGray75: XlPatternToString = "xlPatternGray75"
        Case xlPatternGray50: XlPatternToString = "xlPatternGray50"
        Case xlPatternGray25: XlPatternToString = "xlPatternGray25"
        Case xlPatternDown: XlPatternToString = "xlPatternDown"
        Case xlPatternAutomatic: XlPatternToString = "xlPatternAutomatic"
    End Select
End Function
