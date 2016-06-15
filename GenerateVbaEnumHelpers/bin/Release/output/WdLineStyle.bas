Attribute VB_Name = "wWdLineStyle"
Function WdLineStyleFromString(value As String) As WdLineStyle
    If IsNumeric(value) Then
        WdLineStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLineStyleNone": WdLineStyleFromString = wdLineStyleNone
        Case "wdLineStyleSingle": WdLineStyleFromString = wdLineStyleSingle
        Case "wdLineStyleDot": WdLineStyleFromString = wdLineStyleDot
        Case "wdLineStyleDashSmallGap": WdLineStyleFromString = wdLineStyleDashSmallGap
        Case "wdLineStyleDashLargeGap": WdLineStyleFromString = wdLineStyleDashLargeGap
        Case "wdLineStyleDashDot": WdLineStyleFromString = wdLineStyleDashDot
        Case "wdLineStyleDashDotDot": WdLineStyleFromString = wdLineStyleDashDotDot
        Case "wdLineStyleDouble": WdLineStyleFromString = wdLineStyleDouble
        Case "wdLineStyleTriple": WdLineStyleFromString = wdLineStyleTriple
        Case "wdLineStyleThinThickSmallGap": WdLineStyleFromString = wdLineStyleThinThickSmallGap
        Case "wdLineStyleThickThinSmallGap": WdLineStyleFromString = wdLineStyleThickThinSmallGap
        Case "wdLineStyleThinThickThinSmallGap": WdLineStyleFromString = wdLineStyleThinThickThinSmallGap
        Case "wdLineStyleThinThickMedGap": WdLineStyleFromString = wdLineStyleThinThickMedGap
        Case "wdLineStyleThickThinMedGap": WdLineStyleFromString = wdLineStyleThickThinMedGap
        Case "wdLineStyleThinThickThinMedGap": WdLineStyleFromString = wdLineStyleThinThickThinMedGap
        Case "wdLineStyleThinThickLargeGap": WdLineStyleFromString = wdLineStyleThinThickLargeGap
        Case "wdLineStyleThickThinLargeGap": WdLineStyleFromString = wdLineStyleThickThinLargeGap
        Case "wdLineStyleThinThickThinLargeGap": WdLineStyleFromString = wdLineStyleThinThickThinLargeGap
        Case "wdLineStyleSingleWavy": WdLineStyleFromString = wdLineStyleSingleWavy
        Case "wdLineStyleDoubleWavy": WdLineStyleFromString = wdLineStyleDoubleWavy
        Case "wdLineStyleDashDotStroked": WdLineStyleFromString = wdLineStyleDashDotStroked
        Case "wdLineStyleEmboss3D": WdLineStyleFromString = wdLineStyleEmboss3D
        Case "wdLineStyleEngrave3D": WdLineStyleFromString = wdLineStyleEngrave3D
        Case "wdLineStyleOutset": WdLineStyleFromString = wdLineStyleOutset
        Case "wdLineStyleInset": WdLineStyleFromString = wdLineStyleInset
    End Select
End Function

Function WdLineStyleToString(value As WdLineStyle) As String
    Select Case value
        Case wdLineStyleNone: WdLineStyleToString = "wdLineStyleNone"
        Case wdLineStyleSingle: WdLineStyleToString = "wdLineStyleSingle"
        Case wdLineStyleDot: WdLineStyleToString = "wdLineStyleDot"
        Case wdLineStyleDashSmallGap: WdLineStyleToString = "wdLineStyleDashSmallGap"
        Case wdLineStyleDashLargeGap: WdLineStyleToString = "wdLineStyleDashLargeGap"
        Case wdLineStyleDashDot: WdLineStyleToString = "wdLineStyleDashDot"
        Case wdLineStyleDashDotDot: WdLineStyleToString = "wdLineStyleDashDotDot"
        Case wdLineStyleDouble: WdLineStyleToString = "wdLineStyleDouble"
        Case wdLineStyleTriple: WdLineStyleToString = "wdLineStyleTriple"
        Case wdLineStyleThinThickSmallGap: WdLineStyleToString = "wdLineStyleThinThickSmallGap"
        Case wdLineStyleThickThinSmallGap: WdLineStyleToString = "wdLineStyleThickThinSmallGap"
        Case wdLineStyleThinThickThinSmallGap: WdLineStyleToString = "wdLineStyleThinThickThinSmallGap"
        Case wdLineStyleThinThickMedGap: WdLineStyleToString = "wdLineStyleThinThickMedGap"
        Case wdLineStyleThickThinMedGap: WdLineStyleToString = "wdLineStyleThickThinMedGap"
        Case wdLineStyleThinThickThinMedGap: WdLineStyleToString = "wdLineStyleThinThickThinMedGap"
        Case wdLineStyleThinThickLargeGap: WdLineStyleToString = "wdLineStyleThinThickLargeGap"
        Case wdLineStyleThickThinLargeGap: WdLineStyleToString = "wdLineStyleThickThinLargeGap"
        Case wdLineStyleThinThickThinLargeGap: WdLineStyleToString = "wdLineStyleThinThickThinLargeGap"
        Case wdLineStyleSingleWavy: WdLineStyleToString = "wdLineStyleSingleWavy"
        Case wdLineStyleDoubleWavy: WdLineStyleToString = "wdLineStyleDoubleWavy"
        Case wdLineStyleDashDotStroked: WdLineStyleToString = "wdLineStyleDashDotStroked"
        Case wdLineStyleEmboss3D: WdLineStyleToString = "wdLineStyleEmboss3D"
        Case wdLineStyleEngrave3D: WdLineStyleToString = "wdLineStyleEngrave3D"
        Case wdLineStyleOutset: WdLineStyleToString = "wdLineStyleOutset"
        Case wdLineStyleInset: WdLineStyleToString = "wdLineStyleInset"
    End Select
End Function
