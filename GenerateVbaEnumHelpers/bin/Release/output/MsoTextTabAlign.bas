Attribute VB_Name = "wMsoTextTabAlign"
Function MsoTextTabAlignFromString(value As String) As MsoTextTabAlign
    If IsNumeric(value) Then
        MsoTextTabAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTabAlignLeft": MsoTextTabAlignFromString = msoTabAlignLeft
        Case "msoTabAlignCenter": MsoTextTabAlignFromString = msoTabAlignCenter
        Case "msoTabAlignRight": MsoTextTabAlignFromString = msoTabAlignRight
        Case "msoTabAlignDecimal": MsoTextTabAlignFromString = msoTabAlignDecimal
        Case "msoTabAlignMixed": MsoTextTabAlignFromString = msoTabAlignMixed
    End Select
End Function

Function MsoTextTabAlignToString(value As MsoTextTabAlign) As String
    Select Case value
        Case msoTabAlignLeft: MsoTextTabAlignToString = "msoTabAlignLeft"
        Case msoTabAlignCenter: MsoTextTabAlignToString = "msoTabAlignCenter"
        Case msoTabAlignRight: MsoTextTabAlignToString = "msoTabAlignRight"
        Case msoTabAlignDecimal: MsoTextTabAlignToString = "msoTabAlignDecimal"
        Case msoTabAlignMixed: MsoTextTabAlignToString = "msoTabAlignMixed"
    End Select
End Function
