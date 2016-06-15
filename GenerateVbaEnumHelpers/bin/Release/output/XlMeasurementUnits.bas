Attribute VB_Name = "wXlMeasurementUnits"
Function XlMeasurementUnitsFromString(value As String) As XlMeasurementUnits
    If IsNumeric(value) Then
        XlMeasurementUnitsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlInches": XlMeasurementUnitsFromString = xlInches
        Case "xlCentimeters": XlMeasurementUnitsFromString = xlCentimeters
        Case "xlMillimeters": XlMeasurementUnitsFromString = xlMillimeters
    End Select
End Function

Function XlMeasurementUnitsToString(value As XlMeasurementUnits) As String
    Select Case value
        Case xlInches: XlMeasurementUnitsToString = "xlInches"
        Case xlCentimeters: XlMeasurementUnitsToString = "xlCentimeters"
        Case xlMillimeters: XlMeasurementUnitsToString = "xlMillimeters"
    End Select
End Function
