Attribute VB_Name = "wWdRelativeVerticalSize"
Function WdRelativeVerticalSizeFromString(value As String) As WdRelativeVerticalSize
    If IsNumeric(value) Then
        WdRelativeVerticalSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRelativeVerticalSizeMargin": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizeMargin
        Case "wdRelativeVerticalSizePage": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizePage
        Case "wdRelativeVerticalSizeTopMarginArea": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizeTopMarginArea
        Case "wdRelativeVerticalSizeBottomMarginArea": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizeBottomMarginArea
        Case "wdRelativeVerticalSizeInnerMarginArea": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizeInnerMarginArea
        Case "wdRelativeVerticalSizeOuterMarginArea": WdRelativeVerticalSizeFromString = wdRelativeVerticalSizeOuterMarginArea
    End Select
End Function

Function WdRelativeVerticalSizeToString(value As WdRelativeVerticalSize) As String
    Select Case value
        Case wdRelativeVerticalSizeMargin: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizeMargin"
        Case wdRelativeVerticalSizePage: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizePage"
        Case wdRelativeVerticalSizeTopMarginArea: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizeTopMarginArea"
        Case wdRelativeVerticalSizeBottomMarginArea: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizeBottomMarginArea"
        Case wdRelativeVerticalSizeInnerMarginArea: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizeInnerMarginArea"
        Case wdRelativeVerticalSizeOuterMarginArea: WdRelativeVerticalSizeToString = "wdRelativeVerticalSizeOuterMarginArea"
    End Select
End Function
