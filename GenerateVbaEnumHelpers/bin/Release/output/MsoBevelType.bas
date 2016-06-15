Attribute VB_Name = "wMsoBevelType"
Function MsoBevelTypeFromString(value As String) As MsoBevelType
    If IsNumeric(value) Then
        MsoBevelTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBevelNone": MsoBevelTypeFromString = msoBevelNone
        Case "msoBevelRelaxedInset": MsoBevelTypeFromString = msoBevelRelaxedInset
        Case "msoBevelCircle": MsoBevelTypeFromString = msoBevelCircle
        Case "msoBevelSlope": MsoBevelTypeFromString = msoBevelSlope
        Case "msoBevelCross": MsoBevelTypeFromString = msoBevelCross
        Case "msoBevelAngle": MsoBevelTypeFromString = msoBevelAngle
        Case "msoBevelSoftRound": MsoBevelTypeFromString = msoBevelSoftRound
        Case "msoBevelConvex": MsoBevelTypeFromString = msoBevelConvex
        Case "msoBevelCoolSlant": MsoBevelTypeFromString = msoBevelCoolSlant
        Case "msoBevelDivot": MsoBevelTypeFromString = msoBevelDivot
        Case "msoBevelRiblet": MsoBevelTypeFromString = msoBevelRiblet
        Case "msoBevelHardEdge": MsoBevelTypeFromString = msoBevelHardEdge
        Case "msoBevelArtDeco": MsoBevelTypeFromString = msoBevelArtDeco
        Case "msoBevelTypeMixed": MsoBevelTypeFromString = msoBevelTypeMixed
    End Select
End Function

Function MsoBevelTypeToString(value As MsoBevelType) As String
    Select Case value
        Case msoBevelNone: MsoBevelTypeToString = "msoBevelNone"
        Case msoBevelRelaxedInset: MsoBevelTypeToString = "msoBevelRelaxedInset"
        Case msoBevelCircle: MsoBevelTypeToString = "msoBevelCircle"
        Case msoBevelSlope: MsoBevelTypeToString = "msoBevelSlope"
        Case msoBevelCross: MsoBevelTypeToString = "msoBevelCross"
        Case msoBevelAngle: MsoBevelTypeToString = "msoBevelAngle"
        Case msoBevelSoftRound: MsoBevelTypeToString = "msoBevelSoftRound"
        Case msoBevelConvex: MsoBevelTypeToString = "msoBevelConvex"
        Case msoBevelCoolSlant: MsoBevelTypeToString = "msoBevelCoolSlant"
        Case msoBevelDivot: MsoBevelTypeToString = "msoBevelDivot"
        Case msoBevelRiblet: MsoBevelTypeToString = "msoBevelRiblet"
        Case msoBevelHardEdge: MsoBevelTypeToString = "msoBevelHardEdge"
        Case msoBevelArtDeco: MsoBevelTypeToString = "msoBevelArtDeco"
        Case msoBevelTypeMixed: MsoBevelTypeToString = "msoBevelTypeMixed"
    End Select
End Function
