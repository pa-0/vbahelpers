Attribute VB_Name = "wWdAnimation"
Function WdAnimationFromString(value As String) As WdAnimation
    If IsNumeric(value) Then
        WdAnimationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAnimationNone": WdAnimationFromString = wdAnimationNone
        Case "wdAnimationLasVegasLights": WdAnimationFromString = wdAnimationLasVegasLights
        Case "wdAnimationBlinkingBackground": WdAnimationFromString = wdAnimationBlinkingBackground
        Case "wdAnimationSparkleText": WdAnimationFromString = wdAnimationSparkleText
        Case "wdAnimationMarchingBlackAnts": WdAnimationFromString = wdAnimationMarchingBlackAnts
        Case "wdAnimationMarchingRedAnts": WdAnimationFromString = wdAnimationMarchingRedAnts
        Case "wdAnimationShimmer": WdAnimationFromString = wdAnimationShimmer
    End Select
End Function

Function WdAnimationToString(value As WdAnimation) As String
    Select Case value
        Case wdAnimationNone: WdAnimationToString = "wdAnimationNone"
        Case wdAnimationLasVegasLights: WdAnimationToString = "wdAnimationLasVegasLights"
        Case wdAnimationBlinkingBackground: WdAnimationToString = "wdAnimationBlinkingBackground"
        Case wdAnimationSparkleText: WdAnimationToString = "wdAnimationSparkleText"
        Case wdAnimationMarchingBlackAnts: WdAnimationToString = "wdAnimationMarchingBlackAnts"
        Case wdAnimationMarchingRedAnts: WdAnimationToString = "wdAnimationMarchingRedAnts"
        Case wdAnimationShimmer: WdAnimationToString = "wdAnimationShimmer"
    End Select
End Function
