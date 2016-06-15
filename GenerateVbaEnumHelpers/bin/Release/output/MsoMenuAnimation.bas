Attribute VB_Name = "wMsoMenuAnimation"
Function MsoMenuAnimationFromString(value As String) As MsoMenuAnimation
    If IsNumeric(value) Then
        MsoMenuAnimationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoMenuAnimationNone": MsoMenuAnimationFromString = msoMenuAnimationNone
        Case "msoMenuAnimationRandom": MsoMenuAnimationFromString = msoMenuAnimationRandom
        Case "msoMenuAnimationUnfold": MsoMenuAnimationFromString = msoMenuAnimationUnfold
        Case "msoMenuAnimationSlide": MsoMenuAnimationFromString = msoMenuAnimationSlide
    End Select
End Function

Function MsoMenuAnimationToString(value As MsoMenuAnimation) As String
    Select Case value
        Case msoMenuAnimationNone: MsoMenuAnimationToString = "msoMenuAnimationNone"
        Case msoMenuAnimationRandom: MsoMenuAnimationToString = "msoMenuAnimationRandom"
        Case msoMenuAnimationUnfold: MsoMenuAnimationToString = "msoMenuAnimationUnfold"
        Case msoMenuAnimationSlide: MsoMenuAnimationToString = "msoMenuAnimationSlide"
    End Select
End Function
