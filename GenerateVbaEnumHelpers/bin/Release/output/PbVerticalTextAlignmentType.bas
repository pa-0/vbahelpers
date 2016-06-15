Attribute VB_Name = "wPbVerticalTextAlignmentType"
Function PbVerticalTextAlignmentTypeFromString(value As String) As PbVerticalTextAlignmentType
    If IsNumeric(value) Then
        PbVerticalTextAlignmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbVerticalTextAlignmentTop": PbVerticalTextAlignmentTypeFromString = pbVerticalTextAlignmentTop
        Case "pbVerticalTextAlignmentCenter": PbVerticalTextAlignmentTypeFromString = pbVerticalTextAlignmentCenter
        Case "pbVerticalTextAlignmentBottom": PbVerticalTextAlignmentTypeFromString = pbVerticalTextAlignmentBottom
    End Select
End Function

Function PbVerticalTextAlignmentTypeToString(value As PbVerticalTextAlignmentType) As String
    Select Case value
        Case pbVerticalTextAlignmentTop: PbVerticalTextAlignmentTypeToString = "pbVerticalTextAlignmentTop"
        Case pbVerticalTextAlignmentCenter: PbVerticalTextAlignmentTypeToString = "pbVerticalTextAlignmentCenter"
        Case pbVerticalTextAlignmentBottom: PbVerticalTextAlignmentTypeToString = "pbVerticalTextAlignmentBottom"
    End Select
End Function
