Attribute VB_Name = "wMsoArrowheadStyle"
Function MsoArrowheadStyleFromString(value As String) As MsoArrowheadStyle
    If IsNumeric(value) Then
        MsoArrowheadStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoArrowheadNone": MsoArrowheadStyleFromString = msoArrowheadNone
        Case "msoArrowheadTriangle": MsoArrowheadStyleFromString = msoArrowheadTriangle
        Case "msoArrowheadOpen": MsoArrowheadStyleFromString = msoArrowheadOpen
        Case "msoArrowheadStealth": MsoArrowheadStyleFromString = msoArrowheadStealth
        Case "msoArrowheadDiamond": MsoArrowheadStyleFromString = msoArrowheadDiamond
        Case "msoArrowheadOval": MsoArrowheadStyleFromString = msoArrowheadOval
        Case "msoArrowheadStyleMixed": MsoArrowheadStyleFromString = msoArrowheadStyleMixed
    End Select
End Function

Function MsoArrowheadStyleToString(value As MsoArrowheadStyle) As String
    Select Case value
        Case msoArrowheadNone: MsoArrowheadStyleToString = "msoArrowheadNone"
        Case msoArrowheadTriangle: MsoArrowheadStyleToString = "msoArrowheadTriangle"
        Case msoArrowheadOpen: MsoArrowheadStyleToString = "msoArrowheadOpen"
        Case msoArrowheadStealth: MsoArrowheadStyleToString = "msoArrowheadStealth"
        Case msoArrowheadDiamond: MsoArrowheadStyleToString = "msoArrowheadDiamond"
        Case msoArrowheadOval: MsoArrowheadStyleToString = "msoArrowheadOval"
        Case msoArrowheadStyleMixed: MsoArrowheadStyleToString = "msoArrowheadStyleMixed"
    End Select
End Function
