Attribute VB_Name = "wWdMeasurementUnits"
Function WdMeasurementUnitsFromString(value As String) As WdMeasurementUnits
    If IsNumeric(value) Then
        WdMeasurementUnitsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInches": WdMeasurementUnitsFromString = wdInches
        Case "wdCentimeters": WdMeasurementUnitsFromString = wdCentimeters
        Case "wdMillimeters": WdMeasurementUnitsFromString = wdMillimeters
        Case "wdPoints": WdMeasurementUnitsFromString = wdPoints
        Case "wdPicas": WdMeasurementUnitsFromString = wdPicas
    End Select
End Function

Function WdMeasurementUnitsToString(value As WdMeasurementUnits) As String
    Select Case value
        Case wdInches: WdMeasurementUnitsToString = "wdInches"
        Case wdCentimeters: WdMeasurementUnitsToString = "wdCentimeters"
        Case wdMillimeters: WdMeasurementUnitsToString = "wdMillimeters"
        Case wdPoints: WdMeasurementUnitsToString = "wdPoints"
        Case wdPicas: WdMeasurementUnitsToString = "wdPicas"
    End Select
End Function
