Attribute VB_Name = "wMsoAnimationType"
Function MsoAnimationTypeFromString(value As String) As MsoAnimationType
    If IsNumeric(value) Then
        MsoAnimationTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimationIdle": MsoAnimationTypeFromString = msoAnimationIdle
        Case "msoAnimationGreeting": MsoAnimationTypeFromString = msoAnimationGreeting
        Case "msoAnimationGoodbye": MsoAnimationTypeFromString = msoAnimationGoodbye
        Case "msoAnimationBeginSpeaking": MsoAnimationTypeFromString = msoAnimationBeginSpeaking
        Case "msoAnimationRestPose": MsoAnimationTypeFromString = msoAnimationRestPose
        Case "msoAnimationCharacterSuccessMajor": MsoAnimationTypeFromString = msoAnimationCharacterSuccessMajor
        Case "msoAnimationGetAttentionMajor": MsoAnimationTypeFromString = msoAnimationGetAttentionMajor
        Case "msoAnimationGetAttentionMinor": MsoAnimationTypeFromString = msoAnimationGetAttentionMinor
        Case "msoAnimationSearching": MsoAnimationTypeFromString = msoAnimationSearching
        Case "msoAnimationPrinting": MsoAnimationTypeFromString = msoAnimationPrinting
        Case "msoAnimationGestureRight": MsoAnimationTypeFromString = msoAnimationGestureRight
        Case "msoAnimationWritingNotingSomething": MsoAnimationTypeFromString = msoAnimationWritingNotingSomething
        Case "msoAnimationWorkingAtSomething": MsoAnimationTypeFromString = msoAnimationWorkingAtSomething
        Case "msoAnimationThinking": MsoAnimationTypeFromString = msoAnimationThinking
        Case "msoAnimationSendingMail": MsoAnimationTypeFromString = msoAnimationSendingMail
        Case "msoAnimationListensToComputer": MsoAnimationTypeFromString = msoAnimationListensToComputer
        Case "msoAnimationDisappear": MsoAnimationTypeFromString = msoAnimationDisappear
        Case "msoAnimationAppear": MsoAnimationTypeFromString = msoAnimationAppear
        Case "msoAnimationGetArtsy": MsoAnimationTypeFromString = msoAnimationGetArtsy
        Case "msoAnimationGetTechy": MsoAnimationTypeFromString = msoAnimationGetTechy
        Case "msoAnimationGetWizardy": MsoAnimationTypeFromString = msoAnimationGetWizardy
        Case "msoAnimationCheckingSomething": MsoAnimationTypeFromString = msoAnimationCheckingSomething
        Case "msoAnimationLookDown": MsoAnimationTypeFromString = msoAnimationLookDown
        Case "msoAnimationLookDownLeft": MsoAnimationTypeFromString = msoAnimationLookDownLeft
        Case "msoAnimationLookDownRight": MsoAnimationTypeFromString = msoAnimationLookDownRight
        Case "msoAnimationLookLeft": MsoAnimationTypeFromString = msoAnimationLookLeft
        Case "msoAnimationLookRight": MsoAnimationTypeFromString = msoAnimationLookRight
        Case "msoAnimationLookUp": MsoAnimationTypeFromString = msoAnimationLookUp
        Case "msoAnimationLookUpLeft": MsoAnimationTypeFromString = msoAnimationLookUpLeft
        Case "msoAnimationLookUpRight": MsoAnimationTypeFromString = msoAnimationLookUpRight
        Case "msoAnimationSaving": MsoAnimationTypeFromString = msoAnimationSaving
        Case "msoAnimationGestureDown": MsoAnimationTypeFromString = msoAnimationGestureDown
        Case "msoAnimationGestureLeft": MsoAnimationTypeFromString = msoAnimationGestureLeft
        Case "msoAnimationGestureUp": MsoAnimationTypeFromString = msoAnimationGestureUp
        Case "msoAnimationEmptyTrash": MsoAnimationTypeFromString = msoAnimationEmptyTrash
    End Select
End Function

Function MsoAnimationTypeToString(value As MsoAnimationType) As String
    Select Case value
        Case msoAnimationIdle: MsoAnimationTypeToString = "msoAnimationIdle"
        Case msoAnimationGreeting: MsoAnimationTypeToString = "msoAnimationGreeting"
        Case msoAnimationGoodbye: MsoAnimationTypeToString = "msoAnimationGoodbye"
        Case msoAnimationBeginSpeaking: MsoAnimationTypeToString = "msoAnimationBeginSpeaking"
        Case msoAnimationRestPose: MsoAnimationTypeToString = "msoAnimationRestPose"
        Case msoAnimationCharacterSuccessMajor: MsoAnimationTypeToString = "msoAnimationCharacterSuccessMajor"
        Case msoAnimationGetAttentionMajor: MsoAnimationTypeToString = "msoAnimationGetAttentionMajor"
        Case msoAnimationGetAttentionMinor: MsoAnimationTypeToString = "msoAnimationGetAttentionMinor"
        Case msoAnimationSearching: MsoAnimationTypeToString = "msoAnimationSearching"
        Case msoAnimationPrinting: MsoAnimationTypeToString = "msoAnimationPrinting"
        Case msoAnimationGestureRight: MsoAnimationTypeToString = "msoAnimationGestureRight"
        Case msoAnimationWritingNotingSomething: MsoAnimationTypeToString = "msoAnimationWritingNotingSomething"
        Case msoAnimationWorkingAtSomething: MsoAnimationTypeToString = "msoAnimationWorkingAtSomething"
        Case msoAnimationThinking: MsoAnimationTypeToString = "msoAnimationThinking"
        Case msoAnimationSendingMail: MsoAnimationTypeToString = "msoAnimationSendingMail"
        Case msoAnimationListensToComputer: MsoAnimationTypeToString = "msoAnimationListensToComputer"
        Case msoAnimationDisappear: MsoAnimationTypeToString = "msoAnimationDisappear"
        Case msoAnimationAppear: MsoAnimationTypeToString = "msoAnimationAppear"
        Case msoAnimationGetArtsy: MsoAnimationTypeToString = "msoAnimationGetArtsy"
        Case msoAnimationGetTechy: MsoAnimationTypeToString = "msoAnimationGetTechy"
        Case msoAnimationGetWizardy: MsoAnimationTypeToString = "msoAnimationGetWizardy"
        Case msoAnimationCheckingSomething: MsoAnimationTypeToString = "msoAnimationCheckingSomething"
        Case msoAnimationLookDown: MsoAnimationTypeToString = "msoAnimationLookDown"
        Case msoAnimationLookDownLeft: MsoAnimationTypeToString = "msoAnimationLookDownLeft"
        Case msoAnimationLookDownRight: MsoAnimationTypeToString = "msoAnimationLookDownRight"
        Case msoAnimationLookLeft: MsoAnimationTypeToString = "msoAnimationLookLeft"
        Case msoAnimationLookRight: MsoAnimationTypeToString = "msoAnimationLookRight"
        Case msoAnimationLookUp: MsoAnimationTypeToString = "msoAnimationLookUp"
        Case msoAnimationLookUpLeft: MsoAnimationTypeToString = "msoAnimationLookUpLeft"
        Case msoAnimationLookUpRight: MsoAnimationTypeToString = "msoAnimationLookUpRight"
        Case msoAnimationSaving: MsoAnimationTypeToString = "msoAnimationSaving"
        Case msoAnimationGestureDown: MsoAnimationTypeToString = "msoAnimationGestureDown"
        Case msoAnimationGestureLeft: MsoAnimationTypeToString = "msoAnimationGestureLeft"
        Case msoAnimationGestureUp: MsoAnimationTypeToString = "msoAnimationGestureUp"
        Case msoAnimationEmptyTrash: MsoAnimationTypeToString = "msoAnimationEmptyTrash"
    End Select
End Function
