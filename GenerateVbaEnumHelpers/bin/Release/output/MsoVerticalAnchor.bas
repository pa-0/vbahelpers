Attribute VB_Name = "wMsoVerticalAnchor"
Function MsoVerticalAnchorFromString(value As String) As MsoVerticalAnchor
    If IsNumeric(value) Then
        MsoVerticalAnchorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnchorTop": MsoVerticalAnchorFromString = msoAnchorTop
        Case "msoAnchorTopBaseline": MsoVerticalAnchorFromString = msoAnchorTopBaseline
        Case "msoAnchorMiddle": MsoVerticalAnchorFromString = msoAnchorMiddle
        Case "msoAnchorBottom": MsoVerticalAnchorFromString = msoAnchorBottom
        Case "msoAnchorBottomBaseLine": MsoVerticalAnchorFromString = msoAnchorBottomBaseLine
        Case "msoVerticalAnchorMixed": MsoVerticalAnchorFromString = msoVerticalAnchorMixed
    End Select
End Function

Function MsoVerticalAnchorToString(value As MsoVerticalAnchor) As String
    Select Case value
        Case msoAnchorTop: MsoVerticalAnchorToString = "msoAnchorTop"
        Case msoAnchorTopBaseline: MsoVerticalAnchorToString = "msoAnchorTopBaseline"
        Case msoAnchorMiddle: MsoVerticalAnchorToString = "msoAnchorMiddle"
        Case msoAnchorBottom: MsoVerticalAnchorToString = "msoAnchorBottom"
        Case msoAnchorBottomBaseLine: MsoVerticalAnchorToString = "msoAnchorBottomBaseLine"
        Case msoVerticalAnchorMixed: MsoVerticalAnchorToString = "msoVerticalAnchorMixed"
    End Select
End Function
