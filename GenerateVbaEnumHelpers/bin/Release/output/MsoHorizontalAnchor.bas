Attribute VB_Name = "wMsoHorizontalAnchor"
Function MsoHorizontalAnchorFromString(value As String) As MsoHorizontalAnchor
    If IsNumeric(value) Then
        MsoHorizontalAnchorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnchorNone": MsoHorizontalAnchorFromString = msoAnchorNone
        Case "msoAnchorCenter": MsoHorizontalAnchorFromString = msoAnchorCenter
        Case "msoHorizontalAnchorMixed": MsoHorizontalAnchorFromString = msoHorizontalAnchorMixed
    End Select
End Function

Function MsoHorizontalAnchorToString(value As MsoHorizontalAnchor) As String
    Select Case value
        Case msoAnchorNone: MsoHorizontalAnchorToString = "msoAnchorNone"
        Case msoAnchorCenter: MsoHorizontalAnchorToString = "msoAnchorCenter"
        Case msoHorizontalAnchorMixed: MsoHorizontalAnchorToString = "msoHorizontalAnchorMixed"
    End Select
End Function
