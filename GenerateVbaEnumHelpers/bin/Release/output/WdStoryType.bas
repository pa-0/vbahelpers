Attribute VB_Name = "wWdStoryType"
Function WdStoryTypeFromString(value As String) As WdStoryType
    If IsNumeric(value) Then
        WdStoryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMainTextStory": WdStoryTypeFromString = wdMainTextStory
        Case "wdFootnotesStory": WdStoryTypeFromString = wdFootnotesStory
        Case "wdEndnotesStory": WdStoryTypeFromString = wdEndnotesStory
        Case "wdCommentsStory": WdStoryTypeFromString = wdCommentsStory
        Case "wdTextFrameStory": WdStoryTypeFromString = wdTextFrameStory
        Case "wdEvenPagesHeaderStory": WdStoryTypeFromString = wdEvenPagesHeaderStory
        Case "wdPrimaryHeaderStory": WdStoryTypeFromString = wdPrimaryHeaderStory
        Case "wdEvenPagesFooterStory": WdStoryTypeFromString = wdEvenPagesFooterStory
        Case "wdPrimaryFooterStory": WdStoryTypeFromString = wdPrimaryFooterStory
        Case "wdFirstPageHeaderStory": WdStoryTypeFromString = wdFirstPageHeaderStory
        Case "wdFirstPageFooterStory": WdStoryTypeFromString = wdFirstPageFooterStory
        Case "wdFootnoteSeparatorStory": WdStoryTypeFromString = wdFootnoteSeparatorStory
        Case "wdFootnoteContinuationSeparatorStory": WdStoryTypeFromString = wdFootnoteContinuationSeparatorStory
        Case "wdFootnoteContinuationNoticeStory": WdStoryTypeFromString = wdFootnoteContinuationNoticeStory
        Case "wdEndnoteSeparatorStory": WdStoryTypeFromString = wdEndnoteSeparatorStory
        Case "wdEndnoteContinuationSeparatorStory": WdStoryTypeFromString = wdEndnoteContinuationSeparatorStory
        Case "wdEndnoteContinuationNoticeStory": WdStoryTypeFromString = wdEndnoteContinuationNoticeStory
    End Select
End Function

Function WdStoryTypeToString(value As WdStoryType) As String
    Select Case value
        Case wdMainTextStory: WdStoryTypeToString = "wdMainTextStory"
        Case wdFootnotesStory: WdStoryTypeToString = "wdFootnotesStory"
        Case wdEndnotesStory: WdStoryTypeToString = "wdEndnotesStory"
        Case wdCommentsStory: WdStoryTypeToString = "wdCommentsStory"
        Case wdTextFrameStory: WdStoryTypeToString = "wdTextFrameStory"
        Case wdEvenPagesHeaderStory: WdStoryTypeToString = "wdEvenPagesHeaderStory"
        Case wdPrimaryHeaderStory: WdStoryTypeToString = "wdPrimaryHeaderStory"
        Case wdEvenPagesFooterStory: WdStoryTypeToString = "wdEvenPagesFooterStory"
        Case wdPrimaryFooterStory: WdStoryTypeToString = "wdPrimaryFooterStory"
        Case wdFirstPageHeaderStory: WdStoryTypeToString = "wdFirstPageHeaderStory"
        Case wdFirstPageFooterStory: WdStoryTypeToString = "wdFirstPageFooterStory"
        Case wdFootnoteSeparatorStory: WdStoryTypeToString = "wdFootnoteSeparatorStory"
        Case wdFootnoteContinuationSeparatorStory: WdStoryTypeToString = "wdFootnoteContinuationSeparatorStory"
        Case wdFootnoteContinuationNoticeStory: WdStoryTypeToString = "wdFootnoteContinuationNoticeStory"
        Case wdEndnoteSeparatorStory: WdStoryTypeToString = "wdEndnoteSeparatorStory"
        Case wdEndnoteContinuationSeparatorStory: WdStoryTypeToString = "wdEndnoteContinuationSeparatorStory"
        Case wdEndnoteContinuationNoticeStory: WdStoryTypeToString = "wdEndnoteContinuationNoticeStory"
    End Select
End Function
