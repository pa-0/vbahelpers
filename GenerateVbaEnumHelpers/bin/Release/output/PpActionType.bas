Attribute VB_Name = "wPpActionType"
Function PpActionTypeFromString(value As String) As PpActionType
    If IsNumeric(value) Then
        PpActionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppActionNone": PpActionTypeFromString = ppActionNone
        Case "ppActionNextSlide": PpActionTypeFromString = ppActionNextSlide
        Case "ppActionPreviousSlide": PpActionTypeFromString = ppActionPreviousSlide
        Case "ppActionFirstSlide": PpActionTypeFromString = ppActionFirstSlide
        Case "ppActionLastSlide": PpActionTypeFromString = ppActionLastSlide
        Case "ppActionLastSlideViewed": PpActionTypeFromString = ppActionLastSlideViewed
        Case "ppActionEndShow": PpActionTypeFromString = ppActionEndShow
        Case "ppActionHyperlink": PpActionTypeFromString = ppActionHyperlink
        Case "ppActionRunMacro": PpActionTypeFromString = ppActionRunMacro
        Case "ppActionRunProgram": PpActionTypeFromString = ppActionRunProgram
        Case "ppActionNamedSlideShow": PpActionTypeFromString = ppActionNamedSlideShow
        Case "ppActionOLEVerb": PpActionTypeFromString = ppActionOLEVerb
        Case "ppActionPlay": PpActionTypeFromString = ppActionPlay
        Case "ppActionMixed": PpActionTypeFromString = ppActionMixed
    End Select
End Function

Function PpActionTypeToString(value As PpActionType) As String
    Select Case value
        Case ppActionNone: PpActionTypeToString = "ppActionNone"
        Case ppActionNextSlide: PpActionTypeToString = "ppActionNextSlide"
        Case ppActionPreviousSlide: PpActionTypeToString = "ppActionPreviousSlide"
        Case ppActionFirstSlide: PpActionTypeToString = "ppActionFirstSlide"
        Case ppActionLastSlide: PpActionTypeToString = "ppActionLastSlide"
        Case ppActionLastSlideViewed: PpActionTypeToString = "ppActionLastSlideViewed"
        Case ppActionEndShow: PpActionTypeToString = "ppActionEndShow"
        Case ppActionHyperlink: PpActionTypeToString = "ppActionHyperlink"
        Case ppActionRunMacro: PpActionTypeToString = "ppActionRunMacro"
        Case ppActionRunProgram: PpActionTypeToString = "ppActionRunProgram"
        Case ppActionNamedSlideShow: PpActionTypeToString = "ppActionNamedSlideShow"
        Case ppActionOLEVerb: PpActionTypeToString = "ppActionOLEVerb"
        Case ppActionPlay: PpActionTypeToString = "ppActionPlay"
        Case ppActionMixed: PpActionTypeToString = "ppActionMixed"
    End Select
End Function
