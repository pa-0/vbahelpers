Attribute VB_Name = "wXlPivotFieldOrientation"
Function XlPivotFieldOrientationFromString(value As String) As XlPivotFieldOrientation
    If IsNumeric(value) Then
        XlPivotFieldOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHidden": XlPivotFieldOrientationFromString = xlHidden
        Case "xlRowField": XlPivotFieldOrientationFromString = xlRowField
        Case "xlColumnField": XlPivotFieldOrientationFromString = xlColumnField
        Case "xlPageField": XlPivotFieldOrientationFromString = xlPageField
        Case "xlDataField": XlPivotFieldOrientationFromString = xlDataField
    End Select
End Function

Function XlPivotFieldOrientationToString(value As XlPivotFieldOrientation) As String
    Select Case value
        Case xlHidden: XlPivotFieldOrientationToString = "xlHidden"
        Case xlRowField: XlPivotFieldOrientationToString = "xlRowField"
        Case xlColumnField: XlPivotFieldOrientationToString = "xlColumnField"
        Case xlPageField: XlPivotFieldOrientationToString = "xlPageField"
        Case xlDataField: XlPivotFieldOrientationToString = "xlDataField"
    End Select
End Function
