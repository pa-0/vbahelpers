Attribute VB_Name = "wWdPhoneticGuideAlignmentType"
Function WdPhoneticGuideAlignmentTypeFromString(value As String) As WdPhoneticGuideAlignmentType
    If IsNumeric(value) Then
        WdPhoneticGuideAlignmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPhoneticGuideAlignmentCenter": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentCenter
        Case "wdPhoneticGuideAlignmentZeroOneZero": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentZeroOneZero
        Case "wdPhoneticGuideAlignmentOneTwoOne": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentOneTwoOne
        Case "wdPhoneticGuideAlignmentLeft": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentLeft
        Case "wdPhoneticGuideAlignmentRight": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentRight
        Case "wdPhoneticGuideAlignmentRightVertical": WdPhoneticGuideAlignmentTypeFromString = wdPhoneticGuideAlignmentRightVertical
    End Select
End Function

Function WdPhoneticGuideAlignmentTypeToString(value As WdPhoneticGuideAlignmentType) As String
    Select Case value
        Case wdPhoneticGuideAlignmentCenter: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentCenter"
        Case wdPhoneticGuideAlignmentZeroOneZero: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentZeroOneZero"
        Case wdPhoneticGuideAlignmentOneTwoOne: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentOneTwoOne"
        Case wdPhoneticGuideAlignmentLeft: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentLeft"
        Case wdPhoneticGuideAlignmentRight: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentRight"
        Case wdPhoneticGuideAlignmentRightVertical: WdPhoneticGuideAlignmentTypeToString = "wdPhoneticGuideAlignmentRightVertical"
    End Select
End Function
