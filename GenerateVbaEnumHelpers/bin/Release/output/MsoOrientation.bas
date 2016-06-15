Attribute VB_Name = "wMsoOrientation"
Function MsoOrientationFromString(value As String) As MsoOrientation
    If IsNumeric(value) Then
        MsoOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOrientationHorizontal": MsoOrientationFromString = msoOrientationHorizontal
        Case "msoOrientationVertical": MsoOrientationFromString = msoOrientationVertical
        Case "msoOrientationMixed": MsoOrientationFromString = msoOrientationMixed
    End Select
End Function

Function MsoOrientationToString(value As MsoOrientation) As String
    Select Case value
        Case msoOrientationHorizontal: MsoOrientationToString = "msoOrientationHorizontal"
        Case msoOrientationVertical: MsoOrientationToString = "msoOrientationVertical"
        Case msoOrientationMixed: MsoOrientationToString = "msoOrientationMixed"
    End Select
End Function
