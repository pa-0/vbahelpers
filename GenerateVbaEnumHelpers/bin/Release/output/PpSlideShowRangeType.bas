Attribute VB_Name = "wPpSlideShowRangeType"
Function PpSlideShowRangeTypeFromString(value As String) As PpSlideShowRangeType
    If IsNumeric(value) Then
        PpSlideShowRangeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppShowAll": PpSlideShowRangeTypeFromString = ppShowAll
        Case "ppShowSlideRange": PpSlideShowRangeTypeFromString = ppShowSlideRange
        Case "ppShowNamedSlideShow": PpSlideShowRangeTypeFromString = ppShowNamedSlideShow
    End Select
End Function

Function PpSlideShowRangeTypeToString(value As PpSlideShowRangeType) As String
    Select Case value
        Case ppShowAll: PpSlideShowRangeTypeToString = "ppShowAll"
        Case ppShowSlideRange: PpSlideShowRangeTypeToString = "ppShowSlideRange"
        Case ppShowNamedSlideShow: PpSlideShowRangeTypeToString = "ppShowNamedSlideShow"
    End Select
End Function
