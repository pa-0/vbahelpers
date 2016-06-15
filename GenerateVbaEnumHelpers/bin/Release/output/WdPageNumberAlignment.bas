Attribute VB_Name = "wWdPageNumberAlignment"
Function WdPageNumberAlignmentFromString(value As String) As WdPageNumberAlignment
    If IsNumeric(value) Then
        WdPageNumberAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlignPageNumberLeft": WdPageNumberAlignmentFromString = wdAlignPageNumberLeft
        Case "wdAlignPageNumberCenter": WdPageNumberAlignmentFromString = wdAlignPageNumberCenter
        Case "wdAlignPageNumberRight": WdPageNumberAlignmentFromString = wdAlignPageNumberRight
        Case "wdAlignPageNumberInside": WdPageNumberAlignmentFromString = wdAlignPageNumberInside
        Case "wdAlignPageNumberOutside": WdPageNumberAlignmentFromString = wdAlignPageNumberOutside
    End Select
End Function

Function WdPageNumberAlignmentToString(value As WdPageNumberAlignment) As String
    Select Case value
        Case wdAlignPageNumberLeft: WdPageNumberAlignmentToString = "wdAlignPageNumberLeft"
        Case wdAlignPageNumberCenter: WdPageNumberAlignmentToString = "wdAlignPageNumberCenter"
        Case wdAlignPageNumberRight: WdPageNumberAlignmentToString = "wdAlignPageNumberRight"
        Case wdAlignPageNumberInside: WdPageNumberAlignmentToString = "wdAlignPageNumberInside"
        Case wdAlignPageNumberOutside: WdPageNumberAlignmentToString = "wdAlignPageNumberOutside"
    End Select
End Function
