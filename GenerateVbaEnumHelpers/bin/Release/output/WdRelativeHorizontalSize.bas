Attribute VB_Name = "wWdRelativeHorizontalSize"
Function WdRelativeHorizontalSizeFromString(value As String) As WdRelativeHorizontalSize
    If IsNumeric(value) Then
        WdRelativeHorizontalSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRelativeHorizontalSizeMargin": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizeMargin
        Case "wdRelativeHorizontalSizePage": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizePage
        Case "wdRelativeHorizontalSizeLeftMarginArea": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizeLeftMarginArea
        Case "wdRelativeHorizontalSizeRightMarginArea": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizeRightMarginArea
        Case "wdRelativeHorizontalSizeInnerMarginArea": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizeInnerMarginArea
        Case "wdRelativeHorizontalSizeOuterMarginArea": WdRelativeHorizontalSizeFromString = wdRelativeHorizontalSizeOuterMarginArea
    End Select
End Function

Function WdRelativeHorizontalSizeToString(value As WdRelativeHorizontalSize) As String
    Select Case value
        Case wdRelativeHorizontalSizeMargin: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizeMargin"
        Case wdRelativeHorizontalSizePage: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizePage"
        Case wdRelativeHorizontalSizeLeftMarginArea: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizeLeftMarginArea"
        Case wdRelativeHorizontalSizeRightMarginArea: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizeRightMarginArea"
        Case wdRelativeHorizontalSizeInnerMarginArea: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizeInnerMarginArea"
        Case wdRelativeHorizontalSizeOuterMarginArea: WdRelativeHorizontalSizeToString = "wdRelativeHorizontalSizeOuterMarginArea"
    End Select
End Function
