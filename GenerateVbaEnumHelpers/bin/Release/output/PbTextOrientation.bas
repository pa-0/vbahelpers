Attribute VB_Name = "wPbTextOrientation"
Function PbTextOrientationFromString(value As String) As PbTextOrientation
    If IsNumeric(value) Then
        PbTextOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTextOrientationHorizontal": PbTextOrientationFromString = pbTextOrientationHorizontal
        Case "pbTextOrientationVerticalEastAsia": PbTextOrientationFromString = pbTextOrientationVerticalEastAsia
        Case "pbTextOrientationRightToLeft": PbTextOrientationFromString = pbTextOrientationRightToLeft
        Case "pbTextOrientationMixed": PbTextOrientationFromString = pbTextOrientationMixed
    End Select
End Function

Function PbTextOrientationToString(value As PbTextOrientation) As String
    Select Case value
        Case pbTextOrientationHorizontal: PbTextOrientationToString = "pbTextOrientationHorizontal"
        Case pbTextOrientationVerticalEastAsia: PbTextOrientationToString = "pbTextOrientationVerticalEastAsia"
        Case pbTextOrientationRightToLeft: PbTextOrientationToString = "pbTextOrientationRightToLeft"
        Case pbTextOrientationMixed: PbTextOrientationToString = "pbTextOrientationMixed"
    End Select
End Function
