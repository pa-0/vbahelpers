Attribute VB_Name = "wPpBaselineAlignment"
Function PpBaselineAlignmentFromString(value As String) As PpBaselineAlignment
    If IsNumeric(value) Then
        PpBaselineAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppBaselineAlignBaseline": PpBaselineAlignmentFromString = ppBaselineAlignBaseline
        Case "ppBaselineAlignTop": PpBaselineAlignmentFromString = ppBaselineAlignTop
        Case "ppBaselineAlignCenter": PpBaselineAlignmentFromString = ppBaselineAlignCenter
        Case "ppBaselineAlignFarEast50": PpBaselineAlignmentFromString = ppBaselineAlignFarEast50
        Case "ppBaselineAlignAuto": PpBaselineAlignmentFromString = ppBaselineAlignAuto
        Case "ppBaselineAlignMixed": PpBaselineAlignmentFromString = ppBaselineAlignMixed
    End Select
End Function

Function PpBaselineAlignmentToString(value As PpBaselineAlignment) As String
    Select Case value
        Case ppBaselineAlignBaseline: PpBaselineAlignmentToString = "ppBaselineAlignBaseline"
        Case ppBaselineAlignTop: PpBaselineAlignmentToString = "ppBaselineAlignTop"
        Case ppBaselineAlignCenter: PpBaselineAlignmentToString = "ppBaselineAlignCenter"
        Case ppBaselineAlignFarEast50: PpBaselineAlignmentToString = "ppBaselineAlignFarEast50"
        Case ppBaselineAlignAuto: PpBaselineAlignmentToString = "ppBaselineAlignAuto"
        Case ppBaselineAlignMixed: PpBaselineAlignmentToString = "ppBaselineAlignMixed"
    End Select
End Function
