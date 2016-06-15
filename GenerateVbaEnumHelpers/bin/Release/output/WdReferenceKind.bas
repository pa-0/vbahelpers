Attribute VB_Name = "wWdReferenceKind"
Function WdReferenceKindFromString(value As String) As WdReferenceKind
    If IsNumeric(value) Then
        WdReferenceKindFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEntireCaption": WdReferenceKindFromString = wdEntireCaption
        Case "wdOnlyLabelAndNumber": WdReferenceKindFromString = wdOnlyLabelAndNumber
        Case "wdOnlyCaptionText": WdReferenceKindFromString = wdOnlyCaptionText
        Case "wdFootnoteNumber": WdReferenceKindFromString = wdFootnoteNumber
        Case "wdEndnoteNumber": WdReferenceKindFromString = wdEndnoteNumber
        Case "wdPageNumber": WdReferenceKindFromString = wdPageNumber
        Case "wdPosition": WdReferenceKindFromString = wdPosition
        Case "wdFootnoteNumberFormatted": WdReferenceKindFromString = wdFootnoteNumberFormatted
        Case "wdEndnoteNumberFormatted": WdReferenceKindFromString = wdEndnoteNumberFormatted
        Case "wdNumberFullContext": WdReferenceKindFromString = wdNumberFullContext
        Case "wdNumberNoContext": WdReferenceKindFromString = wdNumberNoContext
        Case "wdNumberRelativeContext": WdReferenceKindFromString = wdNumberRelativeContext
        Case "wdContentText": WdReferenceKindFromString = wdContentText
    End Select
End Function

Function WdReferenceKindToString(value As WdReferenceKind) As String
    Select Case value
        Case wdEntireCaption: WdReferenceKindToString = "wdEntireCaption"
        Case wdOnlyLabelAndNumber: WdReferenceKindToString = "wdOnlyLabelAndNumber"
        Case wdOnlyCaptionText: WdReferenceKindToString = "wdOnlyCaptionText"
        Case wdFootnoteNumber: WdReferenceKindToString = "wdFootnoteNumber"
        Case wdEndnoteNumber: WdReferenceKindToString = "wdEndnoteNumber"
        Case wdPageNumber: WdReferenceKindToString = "wdPageNumber"
        Case wdPosition: WdReferenceKindToString = "wdPosition"
        Case wdFootnoteNumberFormatted: WdReferenceKindToString = "wdFootnoteNumberFormatted"
        Case wdEndnoteNumberFormatted: WdReferenceKindToString = "wdEndnoteNumberFormatted"
        Case wdNumberFullContext: WdReferenceKindToString = "wdNumberFullContext"
        Case wdNumberNoContext: WdReferenceKindToString = "wdNumberNoContext"
        Case wdNumberRelativeContext: WdReferenceKindToString = "wdNumberRelativeContext"
        Case wdContentText: WdReferenceKindToString = "wdContentText"
    End Select
End Function
