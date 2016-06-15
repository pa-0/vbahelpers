Attribute VB_Name = "wXlIconSet"
Function XlIconSetFromString(value As String) As XlIconSet
    If IsNumeric(value) Then
        XlIconSetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xl3Arrows": XlIconSetFromString = xl3Arrows
        Case "xl3ArrowsGray": XlIconSetFromString = xl3ArrowsGray
        Case "xl3Flags": XlIconSetFromString = xl3Flags
        Case "xl3TrafficLights1": XlIconSetFromString = xl3TrafficLights1
        Case "xl3TrafficLights2": XlIconSetFromString = xl3TrafficLights2
        Case "xl3Signs": XlIconSetFromString = xl3Signs
        Case "xl3Symbols": XlIconSetFromString = xl3Symbols
        Case "xl3Symbols2": XlIconSetFromString = xl3Symbols2
        Case "xl4Arrows": XlIconSetFromString = xl4Arrows
        Case "xl4ArrowsGray": XlIconSetFromString = xl4ArrowsGray
        Case "xl4RedToBlack": XlIconSetFromString = xl4RedToBlack
        Case "xl4CRV": XlIconSetFromString = xl4CRV
        Case "xl4TrafficLights": XlIconSetFromString = xl4TrafficLights
        Case "xl5Arrows": XlIconSetFromString = xl5Arrows
        Case "xl5ArrowsGray": XlIconSetFromString = xl5ArrowsGray
        Case "xl5CRV": XlIconSetFromString = xl5CRV
        Case "xl5Quarters": XlIconSetFromString = xl5Quarters
        Case "xl3Stars": XlIconSetFromString = xl3Stars
        Case "xl3Triangles": XlIconSetFromString = xl3Triangles
        Case "xl5Boxes": XlIconSetFromString = xl5Boxes
        Case "xlCustomSet": XlIconSetFromString = xlCustomSet
    End Select
End Function

Function XlIconSetToString(value As XlIconSet) As String
    Select Case value
        Case xl3Arrows: XlIconSetToString = "xl3Arrows"
        Case xl3ArrowsGray: XlIconSetToString = "xl3ArrowsGray"
        Case xl3Flags: XlIconSetToString = "xl3Flags"
        Case xl3TrafficLights1: XlIconSetToString = "xl3TrafficLights1"
        Case xl3TrafficLights2: XlIconSetToString = "xl3TrafficLights2"
        Case xl3Signs: XlIconSetToString = "xl3Signs"
        Case xl3Symbols: XlIconSetToString = "xl3Symbols"
        Case xl3Symbols2: XlIconSetToString = "xl3Symbols2"
        Case xl4Arrows: XlIconSetToString = "xl4Arrows"
        Case xl4ArrowsGray: XlIconSetToString = "xl4ArrowsGray"
        Case xl4RedToBlack: XlIconSetToString = "xl4RedToBlack"
        Case xl4CRV: XlIconSetToString = "xl4CRV"
        Case xl4TrafficLights: XlIconSetToString = "xl4TrafficLights"
        Case xl5Arrows: XlIconSetToString = "xl5Arrows"
        Case xl5ArrowsGray: XlIconSetToString = "xl5ArrowsGray"
        Case xl5CRV: XlIconSetToString = "xl5CRV"
        Case xl5Quarters: XlIconSetToString = "xl5Quarters"
        Case xl3Stars: XlIconSetToString = "xl3Stars"
        Case xl3Triangles: XlIconSetToString = "xl3Triangles"
        Case xl5Boxes: XlIconSetToString = "xl5Boxes"
        Case xlCustomSet: XlIconSetToString = "xlCustomSet"
    End Select
End Function
