Attribute VB_Name = "wWdSpecialPane"
Function WdSpecialPaneFromString(value As String) As WdSpecialPane
    If IsNumeric(value) Then
        WdSpecialPaneFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPaneNone": WdSpecialPaneFromString = wdPaneNone
        Case "wdPanePrimaryHeader": WdSpecialPaneFromString = wdPanePrimaryHeader
        Case "wdPaneFirstPageHeader": WdSpecialPaneFromString = wdPaneFirstPageHeader
        Case "wdPaneEvenPagesHeader": WdSpecialPaneFromString = wdPaneEvenPagesHeader
        Case "wdPanePrimaryFooter": WdSpecialPaneFromString = wdPanePrimaryFooter
        Case "wdPaneFirstPageFooter": WdSpecialPaneFromString = wdPaneFirstPageFooter
        Case "wdPaneEvenPagesFooter": WdSpecialPaneFromString = wdPaneEvenPagesFooter
        Case "wdPaneFootnotes": WdSpecialPaneFromString = wdPaneFootnotes
        Case "wdPaneEndnotes": WdSpecialPaneFromString = wdPaneEndnotes
        Case "wdPaneFootnoteContinuationNotice": WdSpecialPaneFromString = wdPaneFootnoteContinuationNotice
        Case "wdPaneFootnoteContinuationSeparator": WdSpecialPaneFromString = wdPaneFootnoteContinuationSeparator
        Case "wdPaneFootnoteSeparator": WdSpecialPaneFromString = wdPaneFootnoteSeparator
        Case "wdPaneEndnoteContinuationNotice": WdSpecialPaneFromString = wdPaneEndnoteContinuationNotice
        Case "wdPaneEndnoteContinuationSeparator": WdSpecialPaneFromString = wdPaneEndnoteContinuationSeparator
        Case "wdPaneEndnoteSeparator": WdSpecialPaneFromString = wdPaneEndnoteSeparator
        Case "wdPaneComments": WdSpecialPaneFromString = wdPaneComments
        Case "wdPaneCurrentPageHeader": WdSpecialPaneFromString = wdPaneCurrentPageHeader
        Case "wdPaneCurrentPageFooter": WdSpecialPaneFromString = wdPaneCurrentPageFooter
        Case "wdPaneRevisions": WdSpecialPaneFromString = wdPaneRevisions
        Case "wdPaneRevisionsHoriz": WdSpecialPaneFromString = wdPaneRevisionsHoriz
        Case "wdPaneRevisionsVert": WdSpecialPaneFromString = wdPaneRevisionsVert
    End Select
End Function

Function WdSpecialPaneToString(value As WdSpecialPane) As String
    Select Case value
        Case wdPaneNone: WdSpecialPaneToString = "wdPaneNone"
        Case wdPanePrimaryHeader: WdSpecialPaneToString = "wdPanePrimaryHeader"
        Case wdPaneFirstPageHeader: WdSpecialPaneToString = "wdPaneFirstPageHeader"
        Case wdPaneEvenPagesHeader: WdSpecialPaneToString = "wdPaneEvenPagesHeader"
        Case wdPanePrimaryFooter: WdSpecialPaneToString = "wdPanePrimaryFooter"
        Case wdPaneFirstPageFooter: WdSpecialPaneToString = "wdPaneFirstPageFooter"
        Case wdPaneEvenPagesFooter: WdSpecialPaneToString = "wdPaneEvenPagesFooter"
        Case wdPaneFootnotes: WdSpecialPaneToString = "wdPaneFootnotes"
        Case wdPaneEndnotes: WdSpecialPaneToString = "wdPaneEndnotes"
        Case wdPaneFootnoteContinuationNotice: WdSpecialPaneToString = "wdPaneFootnoteContinuationNotice"
        Case wdPaneFootnoteContinuationSeparator: WdSpecialPaneToString = "wdPaneFootnoteContinuationSeparator"
        Case wdPaneFootnoteSeparator: WdSpecialPaneToString = "wdPaneFootnoteSeparator"
        Case wdPaneEndnoteContinuationNotice: WdSpecialPaneToString = "wdPaneEndnoteContinuationNotice"
        Case wdPaneEndnoteContinuationSeparator: WdSpecialPaneToString = "wdPaneEndnoteContinuationSeparator"
        Case wdPaneEndnoteSeparator: WdSpecialPaneToString = "wdPaneEndnoteSeparator"
        Case wdPaneComments: WdSpecialPaneToString = "wdPaneComments"
        Case wdPaneCurrentPageHeader: WdSpecialPaneToString = "wdPaneCurrentPageHeader"
        Case wdPaneCurrentPageFooter: WdSpecialPaneToString = "wdPaneCurrentPageFooter"
        Case wdPaneRevisions: WdSpecialPaneToString = "wdPaneRevisions"
        Case wdPaneRevisionsHoriz: WdSpecialPaneToString = "wdPaneRevisionsHoriz"
        Case wdPaneRevisionsVert: WdSpecialPaneToString = "wdPaneRevisionsVert"
    End Select
End Function
