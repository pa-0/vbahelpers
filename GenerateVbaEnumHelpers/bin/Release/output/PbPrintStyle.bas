Attribute VB_Name = "wPbPrintStyle"
Function PbPrintStyleFromString(value As String) As PbPrintStyle
    If IsNumeric(value) Then
        PbPrintStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPrintStyleDefault": PbPrintStyleFromString = pbPrintStyleDefault
        Case "pbPrintStyleOnePagePerSheet": PbPrintStyleFromString = pbPrintStyleOnePagePerSheet
        Case "pbPrintStyleTiled": PbPrintStyleFromString = pbPrintStyleTiled
        Case "pbPrintStyleMultipleCopiesPerSheet": PbPrintStyleFromString = pbPrintStyleMultipleCopiesPerSheet
        Case "pbPrintStyleMultiplePagesPerSheet": PbPrintStyleFromString = pbPrintStyleMultiplePagesPerSheet
        Case "pbPrintStyleBookletSideFold": PbPrintStyleFromString = pbPrintStyleBookletSideFold
        Case "pbPrintStyleBookletTopFold": PbPrintStyleFromString = pbPrintStyleBookletTopFold
        Case "pbPrintStyleHalfFoldSide": PbPrintStyleFromString = pbPrintStyleHalfFoldSide
        Case "pbPrintStyleHalfFoldTop": PbPrintStyleFromString = pbPrintStyleHalfFoldTop
        Case "pbPrintStyleQuarterFoldTop": PbPrintStyleFromString = pbPrintStyleQuarterFoldTop
        Case "pbPrintStyleQuarterFoldSide": PbPrintStyleFromString = pbPrintStyleQuarterFoldSide
        Case "pbPrintStyleEnvelope": PbPrintStyleFromString = pbPrintStyleEnvelope
    End Select
End Function

Function PbPrintStyleToString(value As PbPrintStyle) As String
    Select Case value
        Case pbPrintStyleDefault: PbPrintStyleToString = "pbPrintStyleDefault"
        Case pbPrintStyleOnePagePerSheet: PbPrintStyleToString = "pbPrintStyleOnePagePerSheet"
        Case pbPrintStyleTiled: PbPrintStyleToString = "pbPrintStyleTiled"
        Case pbPrintStyleMultipleCopiesPerSheet: PbPrintStyleToString = "pbPrintStyleMultipleCopiesPerSheet"
        Case pbPrintStyleMultiplePagesPerSheet: PbPrintStyleToString = "pbPrintStyleMultiplePagesPerSheet"
        Case pbPrintStyleBookletSideFold: PbPrintStyleToString = "pbPrintStyleBookletSideFold"
        Case pbPrintStyleBookletTopFold: PbPrintStyleToString = "pbPrintStyleBookletTopFold"
        Case pbPrintStyleHalfFoldSide: PbPrintStyleToString = "pbPrintStyleHalfFoldSide"
        Case pbPrintStyleHalfFoldTop: PbPrintStyleToString = "pbPrintStyleHalfFoldTop"
        Case pbPrintStyleQuarterFoldTop: PbPrintStyleToString = "pbPrintStyleQuarterFoldTop"
        Case pbPrintStyleQuarterFoldSide: PbPrintStyleToString = "pbPrintStyleQuarterFoldSide"
        Case pbPrintStyleEnvelope: PbPrintStyleToString = "pbPrintStyleEnvelope"
    End Select
End Function
