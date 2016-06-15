Attribute VB_Name = "wMsoTextFontAlign"
Function MsoTextFontAlignFromString(value As String) As MsoTextFontAlign
    If IsNumeric(value) Then
        MsoTextFontAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFontAlignAuto": MsoTextFontAlignFromString = msoFontAlignAuto
        Case "msoFontAlignTop": MsoTextFontAlignFromString = msoFontAlignTop
        Case "msoFontAlignCenter": MsoTextFontAlignFromString = msoFontAlignCenter
        Case "msoFontAlignBaseline": MsoTextFontAlignFromString = msoFontAlignBaseline
        Case "msoFontAlignBottom": MsoTextFontAlignFromString = msoFontAlignBottom
        Case "msoFontAlignMixed": MsoTextFontAlignFromString = msoFontAlignMixed
    End Select
End Function

Function MsoTextFontAlignToString(value As MsoTextFontAlign) As String
    Select Case value
        Case msoFontAlignAuto: MsoTextFontAlignToString = "msoFontAlignAuto"
        Case msoFontAlignTop: MsoTextFontAlignToString = "msoFontAlignTop"
        Case msoFontAlignCenter: MsoTextFontAlignToString = "msoFontAlignCenter"
        Case msoFontAlignBaseline: MsoTextFontAlignToString = "msoFontAlignBaseline"
        Case msoFontAlignBottom: MsoTextFontAlignToString = "msoFontAlignBottom"
        Case msoFontAlignMixed: MsoTextFontAlignToString = "msoFontAlignMixed"
    End Select
End Function
