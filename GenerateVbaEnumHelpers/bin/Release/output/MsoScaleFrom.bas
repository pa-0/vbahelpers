Attribute VB_Name = "wMsoScaleFrom"
Function MsoScaleFromFromString(value As String) As MsoScaleFrom
    If IsNumeric(value) Then
        MsoScaleFromFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoScaleFromTopLeft": MsoScaleFromFromString = msoScaleFromTopLeft
        Case "msoScaleFromMiddle": MsoScaleFromFromString = msoScaleFromMiddle
        Case "msoScaleFromBottomRight": MsoScaleFromFromString = msoScaleFromBottomRight
    End Select
End Function

Function MsoScaleFromToString(value As MsoScaleFrom) As String
    Select Case value
        Case msoScaleFromTopLeft: MsoScaleFromToString = "msoScaleFromTopLeft"
        Case msoScaleFromMiddle: MsoScaleFromToString = "msoScaleFromMiddle"
        Case msoScaleFromBottomRight: MsoScaleFromToString = "msoScaleFromBottomRight"
    End Select
End Function
