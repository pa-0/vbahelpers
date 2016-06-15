Attribute VB_Name = "wMsoBaselineAlignment"
Function MsoBaselineAlignmentFromString(value As String) As MsoBaselineAlignment
    If IsNumeric(value) Then
        MsoBaselineAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBaselineAlignBaseline": MsoBaselineAlignmentFromString = msoBaselineAlignBaseline
        Case "msoBaselineAlignTop": MsoBaselineAlignmentFromString = msoBaselineAlignTop
        Case "msoBaselineAlignCenter": MsoBaselineAlignmentFromString = msoBaselineAlignCenter
        Case "msoBaselineAlignFarEast50": MsoBaselineAlignmentFromString = msoBaselineAlignFarEast50
        Case "msoBaselineAlignAuto": MsoBaselineAlignmentFromString = msoBaselineAlignAuto
        Case "msoBaselineAlignMixed": MsoBaselineAlignmentFromString = msoBaselineAlignMixed
    End Select
End Function

Function MsoBaselineAlignmentToString(value As MsoBaselineAlignment) As String
    Select Case value
        Case msoBaselineAlignBaseline: MsoBaselineAlignmentToString = "msoBaselineAlignBaseline"
        Case msoBaselineAlignTop: MsoBaselineAlignmentToString = "msoBaselineAlignTop"
        Case msoBaselineAlignCenter: MsoBaselineAlignmentToString = "msoBaselineAlignCenter"
        Case msoBaselineAlignFarEast50: MsoBaselineAlignmentToString = "msoBaselineAlignFarEast50"
        Case msoBaselineAlignAuto: MsoBaselineAlignmentToString = "msoBaselineAlignAuto"
        Case msoBaselineAlignMixed: MsoBaselineAlignmentToString = "msoBaselineAlignMixed"
    End Select
End Function
