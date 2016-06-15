Attribute VB_Name = "wXlPivotFieldRepeatLabels"
Function XlPivotFieldRepeatLabelsFromString(value As String) As XlPivotFieldRepeatLabels
    If IsNumeric(value) Then
        XlPivotFieldRepeatLabelsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDoNotRepeatLabels": XlPivotFieldRepeatLabelsFromString = xlDoNotRepeatLabels
        Case "xlRepeatLabels": XlPivotFieldRepeatLabelsFromString = xlRepeatLabels
    End Select
End Function

Function XlPivotFieldRepeatLabelsToString(value As XlPivotFieldRepeatLabels) As String
    Select Case value
        Case xlDoNotRepeatLabels: XlPivotFieldRepeatLabelsToString = "xlDoNotRepeatLabels"
        Case xlRepeatLabels: XlPivotFieldRepeatLabelsToString = "xlRepeatLabels"
    End Select
End Function
