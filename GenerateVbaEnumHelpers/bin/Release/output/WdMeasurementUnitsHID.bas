Attribute VB_Name = "wWdMeasurementUnitsHID"
Function WdMeasurementUnitsHIDFromString(value As String) As WdMeasurementUnitsHID
    If IsNumeric(value) Then
        WdMeasurementUnitsHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdMeasurementUnitsHIDFromString = emptyenum
    End Select
End Function

Function WdMeasurementUnitsHIDToString(value As WdMeasurementUnitsHID) As String
    Select Case value
        Case emptyenum: WdMeasurementUnitsHIDToString = "emptyenum"
    End Select
End Function
