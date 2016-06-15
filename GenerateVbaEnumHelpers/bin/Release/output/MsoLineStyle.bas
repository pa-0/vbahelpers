Attribute VB_Name = "wMsoLineStyle"
Function MsoLineStyleFromString(value As String) As MsoLineStyle
    If IsNumeric(value) Then
        MsoLineStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLineSingle": MsoLineStyleFromString = msoLineSingle
        Case "msoLineThinThin": MsoLineStyleFromString = msoLineThinThin
        Case "msoLineThinThick": MsoLineStyleFromString = msoLineThinThick
        Case "msoLineThickThin": MsoLineStyleFromString = msoLineThickThin
        Case "msoLineThickBetweenThin": MsoLineStyleFromString = msoLineThickBetweenThin
        Case "msoLineStyleMixed": MsoLineStyleFromString = msoLineStyleMixed
    End Select
End Function

Function MsoLineStyleToString(value As MsoLineStyle) As String
    Select Case value
        Case msoLineSingle: MsoLineStyleToString = "msoLineSingle"
        Case msoLineThinThin: MsoLineStyleToString = "msoLineThinThin"
        Case msoLineThinThick: MsoLineStyleToString = "msoLineThinThick"
        Case msoLineThickThin: MsoLineStyleToString = "msoLineThickThin"
        Case msoLineThickBetweenThin: MsoLineStyleToString = "msoLineThickBetweenThin"
        Case msoLineStyleMixed: MsoLineStyleToString = "msoLineStyleMixed"
    End Select
End Function
