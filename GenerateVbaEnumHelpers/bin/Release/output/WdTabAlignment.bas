Attribute VB_Name = "wWdTabAlignment"
Function WdTabAlignmentFromString(value As String) As WdTabAlignment
    If IsNumeric(value) Then
        WdTabAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlignTabLeft": WdTabAlignmentFromString = wdAlignTabLeft
        Case "wdAlignTabCenter": WdTabAlignmentFromString = wdAlignTabCenter
        Case "wdAlignTabRight": WdTabAlignmentFromString = wdAlignTabRight
        Case "wdAlignTabDecimal": WdTabAlignmentFromString = wdAlignTabDecimal
        Case "wdAlignTabBar": WdTabAlignmentFromString = wdAlignTabBar
        Case "wdAlignTabList": WdTabAlignmentFromString = wdAlignTabList
    End Select
End Function

Function WdTabAlignmentToString(value As WdTabAlignment) As String
    Select Case value
        Case wdAlignTabLeft: WdTabAlignmentToString = "wdAlignTabLeft"
        Case wdAlignTabCenter: WdTabAlignmentToString = "wdAlignTabCenter"
        Case wdAlignTabRight: WdTabAlignmentToString = "wdAlignTabRight"
        Case wdAlignTabDecimal: WdTabAlignmentToString = "wdAlignTabDecimal"
        Case wdAlignTabBar: WdTabAlignmentToString = "wdAlignTabBar"
        Case wdAlignTabList: WdTabAlignmentToString = "wdAlignTabList"
    End Select
End Function
