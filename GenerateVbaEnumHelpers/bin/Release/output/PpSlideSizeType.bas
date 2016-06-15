Attribute VB_Name = "wPpSlideSizeType"
Function PpSlideSizeTypeFromString(value As String) As PpSlideSizeType
    If IsNumeric(value) Then
        PpSlideSizeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSlideSizeOnScreen": PpSlideSizeTypeFromString = ppSlideSizeOnScreen
        Case "ppSlideSizeLetterPaper": PpSlideSizeTypeFromString = ppSlideSizeLetterPaper
        Case "ppSlideSizeA4Paper": PpSlideSizeTypeFromString = ppSlideSizeA4Paper
        Case "ppSlideSize35MM": PpSlideSizeTypeFromString = ppSlideSize35MM
        Case "ppSlideSizeOverhead": PpSlideSizeTypeFromString = ppSlideSizeOverhead
        Case "ppSlideSizeBanner": PpSlideSizeTypeFromString = ppSlideSizeBanner
        Case "ppSlideSizeCustom": PpSlideSizeTypeFromString = ppSlideSizeCustom
        Case "ppSlideSizeLedgerPaper": PpSlideSizeTypeFromString = ppSlideSizeLedgerPaper
        Case "ppSlideSizeA3Paper": PpSlideSizeTypeFromString = ppSlideSizeA3Paper
        Case "ppSlideSizeB4ISOPaper": PpSlideSizeTypeFromString = ppSlideSizeB4ISOPaper
        Case "ppSlideSizeB5ISOPaper": PpSlideSizeTypeFromString = ppSlideSizeB5ISOPaper
        Case "ppSlideSizeB4JISPaper": PpSlideSizeTypeFromString = ppSlideSizeB4JISPaper
        Case "ppSlideSizeB5JISPaper": PpSlideSizeTypeFromString = ppSlideSizeB5JISPaper
        Case "ppSlideSizeHagakiCard": PpSlideSizeTypeFromString = ppSlideSizeHagakiCard
        Case "ppSlideSizeOnScreen16x9": PpSlideSizeTypeFromString = ppSlideSizeOnScreen16x9
        Case "ppSlideSizeOnScreen16x10": PpSlideSizeTypeFromString = ppSlideSizeOnScreen16x10
    End Select
End Function

Function PpSlideSizeTypeToString(value As PpSlideSizeType) As String
    Select Case value
        Case ppSlideSizeOnScreen: PpSlideSizeTypeToString = "ppSlideSizeOnScreen"
        Case ppSlideSizeLetterPaper: PpSlideSizeTypeToString = "ppSlideSizeLetterPaper"
        Case ppSlideSizeA4Paper: PpSlideSizeTypeToString = "ppSlideSizeA4Paper"
        Case ppSlideSize35MM: PpSlideSizeTypeToString = "ppSlideSize35MM"
        Case ppSlideSizeOverhead: PpSlideSizeTypeToString = "ppSlideSizeOverhead"
        Case ppSlideSizeBanner: PpSlideSizeTypeToString = "ppSlideSizeBanner"
        Case ppSlideSizeCustom: PpSlideSizeTypeToString = "ppSlideSizeCustom"
        Case ppSlideSizeLedgerPaper: PpSlideSizeTypeToString = "ppSlideSizeLedgerPaper"
        Case ppSlideSizeA3Paper: PpSlideSizeTypeToString = "ppSlideSizeA3Paper"
        Case ppSlideSizeB4ISOPaper: PpSlideSizeTypeToString = "ppSlideSizeB4ISOPaper"
        Case ppSlideSizeB5ISOPaper: PpSlideSizeTypeToString = "ppSlideSizeB5ISOPaper"
        Case ppSlideSizeB4JISPaper: PpSlideSizeTypeToString = "ppSlideSizeB4JISPaper"
        Case ppSlideSizeB5JISPaper: PpSlideSizeTypeToString = "ppSlideSizeB5JISPaper"
        Case ppSlideSizeHagakiCard: PpSlideSizeTypeToString = "ppSlideSizeHagakiCard"
        Case ppSlideSizeOnScreen16x9: PpSlideSizeTypeToString = "ppSlideSizeOnScreen16x9"
        Case ppSlideSizeOnScreen16x10: PpSlideSizeTypeToString = "ppSlideSizeOnScreen16x10"
    End Select
End Function
