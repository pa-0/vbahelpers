Attribute VB_Name = "wWdRelativeVerticalPosition"
Function WdRelativeVerticalPositionFromString(value As String) As WdRelativeVerticalPosition
    If IsNumeric(value) Then
        WdRelativeVerticalPositionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRelativeVerticalPositionMargin": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionMargin
        Case "wdRelativeVerticalPositionPage": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionPage
        Case "wdRelativeVerticalPositionParagraph": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionParagraph
        Case "wdRelativeVerticalPositionLine": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionLine
        Case "wdRelativeVerticalPositionTopMarginArea": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionTopMarginArea
        Case "wdRelativeVerticalPositionBottomMarginArea": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionBottomMarginArea
        Case "wdRelativeVerticalPositionInnerMarginArea": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionInnerMarginArea
        Case "wdRelativeVerticalPositionOuterMarginArea": WdRelativeVerticalPositionFromString = wdRelativeVerticalPositionOuterMarginArea
    End Select
End Function

Function WdRelativeVerticalPositionToString(value As WdRelativeVerticalPosition) As String
    Select Case value
        Case wdRelativeVerticalPositionMargin: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionMargin"
        Case wdRelativeVerticalPositionPage: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionPage"
        Case wdRelativeVerticalPositionParagraph: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionParagraph"
        Case wdRelativeVerticalPositionLine: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionLine"
        Case wdRelativeVerticalPositionTopMarginArea: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionTopMarginArea"
        Case wdRelativeVerticalPositionBottomMarginArea: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionBottomMarginArea"
        Case wdRelativeVerticalPositionInnerMarginArea: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionInnerMarginArea"
        Case wdRelativeVerticalPositionOuterMarginArea: WdRelativeVerticalPositionToString = "wdRelativeVerticalPositionOuterMarginArea"
    End Select
End Function
