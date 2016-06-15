Attribute VB_Name = "wWdInformation"
Function WdInformationFromString(value As String) As WdInformation
    If IsNumeric(value) Then
        WdInformationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdActiveEndAdjustedPageNumber": WdInformationFromString = wdActiveEndAdjustedPageNumber
        Case "wdActiveEndSectionNumber": WdInformationFromString = wdActiveEndSectionNumber
        Case "wdActiveEndPageNumber": WdInformationFromString = wdActiveEndPageNumber
        Case "wdNumberOfPagesInDocument": WdInformationFromString = wdNumberOfPagesInDocument
        Case "wdHorizontalPositionRelativeToPage": WdInformationFromString = wdHorizontalPositionRelativeToPage
        Case "wdVerticalPositionRelativeToPage": WdInformationFromString = wdVerticalPositionRelativeToPage
        Case "wdHorizontalPositionRelativeToTextBoundary": WdInformationFromString = wdHorizontalPositionRelativeToTextBoundary
        Case "wdVerticalPositionRelativeToTextBoundary": WdInformationFromString = wdVerticalPositionRelativeToTextBoundary
        Case "wdFirstCharacterColumnNumber": WdInformationFromString = wdFirstCharacterColumnNumber
        Case "wdFirstCharacterLineNumber": WdInformationFromString = wdFirstCharacterLineNumber
        Case "wdFrameIsSelected": WdInformationFromString = wdFrameIsSelected
        Case "wdWithInTable": WdInformationFromString = wdWithInTable
        Case "wdStartOfRangeRowNumber": WdInformationFromString = wdStartOfRangeRowNumber
        Case "wdEndOfRangeRowNumber": WdInformationFromString = wdEndOfRangeRowNumber
        Case "wdMaximumNumberOfRows": WdInformationFromString = wdMaximumNumberOfRows
        Case "wdStartOfRangeColumnNumber": WdInformationFromString = wdStartOfRangeColumnNumber
        Case "wdEndOfRangeColumnNumber": WdInformationFromString = wdEndOfRangeColumnNumber
        Case "wdMaximumNumberOfColumns": WdInformationFromString = wdMaximumNumberOfColumns
        Case "wdZoomPercentage": WdInformationFromString = wdZoomPercentage
        Case "wdSelectionMode": WdInformationFromString = wdSelectionMode
        Case "wdCapsLock": WdInformationFromString = wdCapsLock
        Case "wdNumLock": WdInformationFromString = wdNumLock
        Case "wdOverType": WdInformationFromString = wdOverType
        Case "wdRevisionMarking": WdInformationFromString = wdRevisionMarking
        Case "wdInFootnoteEndnotePane": WdInformationFromString = wdInFootnoteEndnotePane
        Case "wdInCommentPane": WdInformationFromString = wdInCommentPane
        Case "wdInHeaderFooter": WdInformationFromString = wdInHeaderFooter
        Case "wdAtEndOfRowMarker": WdInformationFromString = wdAtEndOfRowMarker
        Case "wdReferenceOfType": WdInformationFromString = wdReferenceOfType
        Case "wdHeaderFooterType": WdInformationFromString = wdHeaderFooterType
        Case "wdInMasterDocument": WdInformationFromString = wdInMasterDocument
        Case "wdInFootnote": WdInformationFromString = wdInFootnote
        Case "wdInEndnote": WdInformationFromString = wdInEndnote
        Case "wdInWordMail": WdInformationFromString = wdInWordMail
        Case "wdInClipboard": WdInformationFromString = wdInClipboard
    End Select
End Function

Function WdInformationToString(value As WdInformation) As String
    Select Case value
        Case wdActiveEndAdjustedPageNumber: WdInformationToString = "wdActiveEndAdjustedPageNumber"
        Case wdActiveEndSectionNumber: WdInformationToString = "wdActiveEndSectionNumber"
        Case wdActiveEndPageNumber: WdInformationToString = "wdActiveEndPageNumber"
        Case wdNumberOfPagesInDocument: WdInformationToString = "wdNumberOfPagesInDocument"
        Case wdHorizontalPositionRelativeToPage: WdInformationToString = "wdHorizontalPositionRelativeToPage"
        Case wdVerticalPositionRelativeToPage: WdInformationToString = "wdVerticalPositionRelativeToPage"
        Case wdHorizontalPositionRelativeToTextBoundary: WdInformationToString = "wdHorizontalPositionRelativeToTextBoundary"
        Case wdVerticalPositionRelativeToTextBoundary: WdInformationToString = "wdVerticalPositionRelativeToTextBoundary"
        Case wdFirstCharacterColumnNumber: WdInformationToString = "wdFirstCharacterColumnNumber"
        Case wdFirstCharacterLineNumber: WdInformationToString = "wdFirstCharacterLineNumber"
        Case wdFrameIsSelected: WdInformationToString = "wdFrameIsSelected"
        Case wdWithInTable: WdInformationToString = "wdWithInTable"
        Case wdStartOfRangeRowNumber: WdInformationToString = "wdStartOfRangeRowNumber"
        Case wdEndOfRangeRowNumber: WdInformationToString = "wdEndOfRangeRowNumber"
        Case wdMaximumNumberOfRows: WdInformationToString = "wdMaximumNumberOfRows"
        Case wdStartOfRangeColumnNumber: WdInformationToString = "wdStartOfRangeColumnNumber"
        Case wdEndOfRangeColumnNumber: WdInformationToString = "wdEndOfRangeColumnNumber"
        Case wdMaximumNumberOfColumns: WdInformationToString = "wdMaximumNumberOfColumns"
        Case wdZoomPercentage: WdInformationToString = "wdZoomPercentage"
        Case wdSelectionMode: WdInformationToString = "wdSelectionMode"
        Case wdCapsLock: WdInformationToString = "wdCapsLock"
        Case wdNumLock: WdInformationToString = "wdNumLock"
        Case wdOverType: WdInformationToString = "wdOverType"
        Case wdRevisionMarking: WdInformationToString = "wdRevisionMarking"
        Case wdInFootnoteEndnotePane: WdInformationToString = "wdInFootnoteEndnotePane"
        Case wdInCommentPane: WdInformationToString = "wdInCommentPane"
        Case wdInHeaderFooter: WdInformationToString = "wdInHeaderFooter"
        Case wdAtEndOfRowMarker: WdInformationToString = "wdAtEndOfRowMarker"
        Case wdReferenceOfType: WdInformationToString = "wdReferenceOfType"
        Case wdHeaderFooterType: WdInformationToString = "wdHeaderFooterType"
        Case wdInMasterDocument: WdInformationToString = "wdInMasterDocument"
        Case wdInFootnote: WdInformationToString = "wdInFootnote"
        Case wdInEndnote: WdInformationToString = "wdInEndnote"
        Case wdInWordMail: WdInformationToString = "wdInWordMail"
        Case wdInClipboard: WdInformationToString = "wdInClipboard"
    End Select
End Function
