Attribute VB_Name = "wPpPublishSourceType"
Function PpPublishSourceTypeFromString(value As String) As PpPublishSourceType
    If IsNumeric(value) Then
        PpPublishSourceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPublishAll": PpPublishSourceTypeFromString = ppPublishAll
        Case "ppPublishSlideRange": PpPublishSourceTypeFromString = ppPublishSlideRange
        Case "ppPublishNamedSlideShow": PpPublishSourceTypeFromString = ppPublishNamedSlideShow
    End Select
End Function

Function PpPublishSourceTypeToString(value As PpPublishSourceType) As String
    Select Case value
        Case ppPublishAll: PpPublishSourceTypeToString = "ppPublishAll"
        Case ppPublishSlideRange: PpPublishSourceTypeToString = "ppPublishSlideRange"
        Case ppPublishNamedSlideShow: PpPublishSourceTypeToString = "ppPublishNamedSlideShow"
    End Select
End Function
