Attribute VB_Name = "wPpViewType"
Function PpViewTypeFromString(value As String) As PpViewType
    If IsNumeric(value) Then
        PpViewTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppViewSlide": PpViewTypeFromString = ppViewSlide
        Case "ppViewSlideMaster": PpViewTypeFromString = ppViewSlideMaster
        Case "ppViewNotesPage": PpViewTypeFromString = ppViewNotesPage
        Case "ppViewHandoutMaster": PpViewTypeFromString = ppViewHandoutMaster
        Case "ppViewNotesMaster": PpViewTypeFromString = ppViewNotesMaster
        Case "ppViewOutline": PpViewTypeFromString = ppViewOutline
        Case "ppViewSlideSorter": PpViewTypeFromString = ppViewSlideSorter
        Case "ppViewTitleMaster": PpViewTypeFromString = ppViewTitleMaster
        Case "ppViewNormal": PpViewTypeFromString = ppViewNormal
        Case "ppViewPrintPreview": PpViewTypeFromString = ppViewPrintPreview
        Case "ppViewThumbnails": PpViewTypeFromString = ppViewThumbnails
        Case "ppViewMasterThumbnails": PpViewTypeFromString = ppViewMasterThumbnails
    End Select
End Function

Function PpViewTypeToString(value As PpViewType) As String
    Select Case value
        Case ppViewSlide: PpViewTypeToString = "ppViewSlide"
        Case ppViewSlideMaster: PpViewTypeToString = "ppViewSlideMaster"
        Case ppViewNotesPage: PpViewTypeToString = "ppViewNotesPage"
        Case ppViewHandoutMaster: PpViewTypeToString = "ppViewHandoutMaster"
        Case ppViewNotesMaster: PpViewTypeToString = "ppViewNotesMaster"
        Case ppViewOutline: PpViewTypeToString = "ppViewOutline"
        Case ppViewSlideSorter: PpViewTypeToString = "ppViewSlideSorter"
        Case ppViewTitleMaster: PpViewTypeToString = "ppViewTitleMaster"
        Case ppViewNormal: PpViewTypeToString = "ppViewNormal"
        Case ppViewPrintPreview: PpViewTypeToString = "ppViewPrintPreview"
        Case ppViewThumbnails: PpViewTypeToString = "ppViewThumbnails"
        Case ppViewMasterThumbnails: PpViewTypeToString = "ppViewMasterThumbnails"
    End Select
End Function
