Attribute VB_Name = "wPbPhoneticGuideAlignmentType"
Function PbPhoneticGuideAlignmentTypeFromString(value As String) As PbPhoneticGuideAlignmentType
    If IsNumeric(value) Then
        PbPhoneticGuideAlignmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbPhoneticGuideAlignmentDefault": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentDefault
        Case "pbPhoneticGuideAlignmentZeroOneZero": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentZeroOneZero
        Case "pbPhoneticGuideAlignmentOneTwoOne": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentOneTwoOne
        Case "pbPhoneticGuideAlignmentCenter": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentCenter
        Case "pbPhoneticGuideAlignmentLeft": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentLeft
        Case "pbPhoneticGuideAlignmentRight": PbPhoneticGuideAlignmentTypeFromString = pbPhoneticGuideAlignmentRight
    End Select
End Function

Function PbPhoneticGuideAlignmentTypeToString(value As PbPhoneticGuideAlignmentType) As String
    Select Case value
        Case pbPhoneticGuideAlignmentDefault: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentDefault"
        Case pbPhoneticGuideAlignmentZeroOneZero: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentZeroOneZero"
        Case pbPhoneticGuideAlignmentOneTwoOne: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentOneTwoOne"
        Case pbPhoneticGuideAlignmentCenter: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentCenter"
        Case pbPhoneticGuideAlignmentLeft: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentLeft"
        Case pbPhoneticGuideAlignmentRight: PbPhoneticGuideAlignmentTypeToString = "pbPhoneticGuideAlignmentRight"
    End Select
End Function
