Attribute VB_Name = "wPbPageNumberType"
Function PbPageNumberTypeFromString(value As String) As PbPageNumberType
    If IsNumeric(value) Then
        PbPageNumberTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPageNumberCurrent": PbPageNumberTypeFromString = pbPageNumberCurrent
        Case "pbPageNumberNextInStory": PbPageNumberTypeFromString = pbPageNumberNextInStory
        Case "pbPageNumberPreviousInStory": PbPageNumberTypeFromString = pbPageNumberPreviousInStory
    End Select
End Function

Function PbPageNumberTypeToString(value As PbPageNumberType) As String
    Select Case value
        Case pbPageNumberCurrent: PbPageNumberTypeToString = "pbPageNumberCurrent"
        Case pbPageNumberNextInStory: PbPageNumberTypeToString = "pbPageNumberNextInStory"
        Case pbPageNumberPreviousInStory: PbPageNumberTypeToString = "pbPageNumberPreviousInStory"
    End Select
End Function
