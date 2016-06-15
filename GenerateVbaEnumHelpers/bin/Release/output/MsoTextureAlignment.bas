Attribute VB_Name = "wMsoTextureAlignment"
Function MsoTextureAlignmentFromString(value As String) As MsoTextureAlignment
    If IsNumeric(value) Then
        MsoTextureAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextureTopLeft": MsoTextureAlignmentFromString = msoTextureTopLeft
        Case "msoTextureTop": MsoTextureAlignmentFromString = msoTextureTop
        Case "msoTextureTopRight": MsoTextureAlignmentFromString = msoTextureTopRight
        Case "msoTextureLeft": MsoTextureAlignmentFromString = msoTextureLeft
        Case "msoTextureCenter": MsoTextureAlignmentFromString = msoTextureCenter
        Case "msoTextureRight": MsoTextureAlignmentFromString = msoTextureRight
        Case "msoTextureBottomLeft": MsoTextureAlignmentFromString = msoTextureBottomLeft
        Case "msoTextureBottom": MsoTextureAlignmentFromString = msoTextureBottom
        Case "msoTextureBottomRight": MsoTextureAlignmentFromString = msoTextureBottomRight
        Case "msoTextureAlignmentMixed": MsoTextureAlignmentFromString = msoTextureAlignmentMixed
    End Select
End Function

Function MsoTextureAlignmentToString(value As MsoTextureAlignment) As String
    Select Case value
        Case msoTextureTopLeft: MsoTextureAlignmentToString = "msoTextureTopLeft"
        Case msoTextureTop: MsoTextureAlignmentToString = "msoTextureTop"
        Case msoTextureTopRight: MsoTextureAlignmentToString = "msoTextureTopRight"
        Case msoTextureLeft: MsoTextureAlignmentToString = "msoTextureLeft"
        Case msoTextureCenter: MsoTextureAlignmentToString = "msoTextureCenter"
        Case msoTextureRight: MsoTextureAlignmentToString = "msoTextureRight"
        Case msoTextureBottomLeft: MsoTextureAlignmentToString = "msoTextureBottomLeft"
        Case msoTextureBottom: MsoTextureAlignmentToString = "msoTextureBottom"
        Case msoTextureBottomRight: MsoTextureAlignmentToString = "msoTextureBottomRight"
        Case msoTextureAlignmentMixed: MsoTextureAlignmentToString = "msoTextureAlignmentMixed"
    End Select
End Function
