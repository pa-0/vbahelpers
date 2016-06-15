Attribute VB_Name = "wPpPrintOutputType"
Function PpPrintOutputTypeFromString(value As String) As PpPrintOutputType
    If IsNumeric(value) Then
        PpPrintOutputTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPrintOutputSlides": PpPrintOutputTypeFromString = ppPrintOutputSlides
        Case "ppPrintOutputTwoSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputTwoSlideHandouts
        Case "ppPrintOutputThreeSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputThreeSlideHandouts
        Case "ppPrintOutputSixSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputSixSlideHandouts
        Case "ppPrintOutputNotesPages": PpPrintOutputTypeFromString = ppPrintOutputNotesPages
        Case "ppPrintOutputOutline": PpPrintOutputTypeFromString = ppPrintOutputOutline
        Case "ppPrintOutputBuildSlides": PpPrintOutputTypeFromString = ppPrintOutputBuildSlides
        Case "ppPrintOutputFourSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputFourSlideHandouts
        Case "ppPrintOutputNineSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputNineSlideHandouts
        Case "ppPrintOutputOneSlideHandouts": PpPrintOutputTypeFromString = ppPrintOutputOneSlideHandouts
    End Select
End Function

Function PpPrintOutputTypeToString(value As PpPrintOutputType) As String
    Select Case value
        Case ppPrintOutputSlides: PpPrintOutputTypeToString = "ppPrintOutputSlides"
        Case ppPrintOutputTwoSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputTwoSlideHandouts"
        Case ppPrintOutputThreeSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputThreeSlideHandouts"
        Case ppPrintOutputSixSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputSixSlideHandouts"
        Case ppPrintOutputNotesPages: PpPrintOutputTypeToString = "ppPrintOutputNotesPages"
        Case ppPrintOutputOutline: PpPrintOutputTypeToString = "ppPrintOutputOutline"
        Case ppPrintOutputBuildSlides: PpPrintOutputTypeToString = "ppPrintOutputBuildSlides"
        Case ppPrintOutputFourSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputFourSlideHandouts"
        Case ppPrintOutputNineSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputNineSlideHandouts"
        Case ppPrintOutputOneSlideHandouts: PpPrintOutputTypeToString = "ppPrintOutputOneSlideHandouts"
    End Select
End Function
