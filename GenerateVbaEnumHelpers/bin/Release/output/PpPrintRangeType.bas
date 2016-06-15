Attribute VB_Name = "wPpPrintRangeType"
Function PpPrintRangeTypeFromString(value As String) As PpPrintRangeType
    If IsNumeric(value) Then
        PpPrintRangeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPrintAll": PpPrintRangeTypeFromString = ppPrintAll
        Case "ppPrintSelection": PpPrintRangeTypeFromString = ppPrintSelection
        Case "ppPrintCurrent": PpPrintRangeTypeFromString = ppPrintCurrent
        Case "ppPrintSlideRange": PpPrintRangeTypeFromString = ppPrintSlideRange
        Case "ppPrintNamedSlideShow": PpPrintRangeTypeFromString = ppPrintNamedSlideShow
        Case "ppPrintSection": PpPrintRangeTypeFromString = ppPrintSection
    End Select
End Function

Function PpPrintRangeTypeToString(value As PpPrintRangeType) As String
    Select Case value
        Case ppPrintAll: PpPrintRangeTypeToString = "ppPrintAll"
        Case ppPrintSelection: PpPrintRangeTypeToString = "ppPrintSelection"
        Case ppPrintCurrent: PpPrintRangeTypeToString = "ppPrintCurrent"
        Case ppPrintSlideRange: PpPrintRangeTypeToString = "ppPrintSlideRange"
        Case ppPrintNamedSlideShow: PpPrintRangeTypeToString = "ppPrintNamedSlideShow"
        Case ppPrintSection: PpPrintRangeTypeToString = "ppPrintSection"
    End Select
End Function
