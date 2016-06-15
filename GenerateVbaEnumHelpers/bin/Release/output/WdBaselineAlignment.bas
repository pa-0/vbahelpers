Attribute VB_Name = "wWdBaselineAlignment"
Function WdBaselineAlignmentFromString(value As String) As WdBaselineAlignment
    If IsNumeric(value) Then
        WdBaselineAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBaselineAlignTop": WdBaselineAlignmentFromString = wdBaselineAlignTop
        Case "wdBaselineAlignCenter": WdBaselineAlignmentFromString = wdBaselineAlignCenter
        Case "wdBaselineAlignBaseline": WdBaselineAlignmentFromString = wdBaselineAlignBaseline
        Case "wdBaselineAlignFarEast50": WdBaselineAlignmentFromString = wdBaselineAlignFarEast50
        Case "wdBaselineAlignAuto": WdBaselineAlignmentFromString = wdBaselineAlignAuto
    End Select
End Function

Function WdBaselineAlignmentToString(value As WdBaselineAlignment) As String
    Select Case value
        Case wdBaselineAlignTop: WdBaselineAlignmentToString = "wdBaselineAlignTop"
        Case wdBaselineAlignCenter: WdBaselineAlignmentToString = "wdBaselineAlignCenter"
        Case wdBaselineAlignBaseline: WdBaselineAlignmentToString = "wdBaselineAlignBaseline"
        Case wdBaselineAlignFarEast50: WdBaselineAlignmentToString = "wdBaselineAlignFarEast50"
        Case wdBaselineAlignAuto: WdBaselineAlignmentToString = "wdBaselineAlignAuto"
    End Select
End Function
