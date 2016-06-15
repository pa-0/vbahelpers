Attribute VB_Name = "wPpResampleMediaProfile"
Function PpResampleMediaProfileFromString(value As String) As PpResampleMediaProfile
    If IsNumeric(value) Then
        PpResampleMediaProfileFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppResampleMediaProfileCustom": PpResampleMediaProfileFromString = ppResampleMediaProfileCustom
        Case "ppResampleMediaProfileSmall": PpResampleMediaProfileFromString = ppResampleMediaProfileSmall
        Case "ppResampleMediaProfileSmaller": PpResampleMediaProfileFromString = ppResampleMediaProfileSmaller
        Case "ppResampleMediaProfileSmallest": PpResampleMediaProfileFromString = ppResampleMediaProfileSmallest
    End Select
End Function

Function PpResampleMediaProfileToString(value As PpResampleMediaProfile) As String
    Select Case value
        Case ppResampleMediaProfileCustom: PpResampleMediaProfileToString = "ppResampleMediaProfileCustom"
        Case ppResampleMediaProfileSmall: PpResampleMediaProfileToString = "ppResampleMediaProfileSmall"
        Case ppResampleMediaProfileSmaller: PpResampleMediaProfileToString = "ppResampleMediaProfileSmaller"
        Case ppResampleMediaProfileSmallest: PpResampleMediaProfileToString = "ppResampleMediaProfileSmallest"
    End Select
End Function
