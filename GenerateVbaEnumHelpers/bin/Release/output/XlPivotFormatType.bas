Attribute VB_Name = "wXlPivotFormatType"
Function XlPivotFormatTypeFromString(value As String) As XlPivotFormatType
    If IsNumeric(value) Then
        XlPivotFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlReport1": XlPivotFormatTypeFromString = xlReport1
        Case "xlReport2": XlPivotFormatTypeFromString = xlReport2
        Case "xlReport3": XlPivotFormatTypeFromString = xlReport3
        Case "xlReport4": XlPivotFormatTypeFromString = xlReport4
        Case "xlReport5": XlPivotFormatTypeFromString = xlReport5
        Case "xlReport6": XlPivotFormatTypeFromString = xlReport6
        Case "xlReport7": XlPivotFormatTypeFromString = xlReport7
        Case "xlReport8": XlPivotFormatTypeFromString = xlReport8
        Case "xlReport9": XlPivotFormatTypeFromString = xlReport9
        Case "xlReport10": XlPivotFormatTypeFromString = xlReport10
        Case "xlTable1": XlPivotFormatTypeFromString = xlTable1
        Case "xlTable2": XlPivotFormatTypeFromString = xlTable2
        Case "xlTable3": XlPivotFormatTypeFromString = xlTable3
        Case "xlTable4": XlPivotFormatTypeFromString = xlTable4
        Case "xlTable5": XlPivotFormatTypeFromString = xlTable5
        Case "xlTable6": XlPivotFormatTypeFromString = xlTable6
        Case "xlTable7": XlPivotFormatTypeFromString = xlTable7
        Case "xlTable8": XlPivotFormatTypeFromString = xlTable8
        Case "xlTable9": XlPivotFormatTypeFromString = xlTable9
        Case "xlTable10": XlPivotFormatTypeFromString = xlTable10
        Case "xlPTClassic": XlPivotFormatTypeFromString = xlPTClassic
        Case "xlPTNone": XlPivotFormatTypeFromString = xlPTNone
    End Select
End Function

Function XlPivotFormatTypeToString(value As XlPivotFormatType) As String
    Select Case value
        Case xlReport1: XlPivotFormatTypeToString = "xlReport1"
        Case xlReport2: XlPivotFormatTypeToString = "xlReport2"
        Case xlReport3: XlPivotFormatTypeToString = "xlReport3"
        Case xlReport4: XlPivotFormatTypeToString = "xlReport4"
        Case xlReport5: XlPivotFormatTypeToString = "xlReport5"
        Case xlReport6: XlPivotFormatTypeToString = "xlReport6"
        Case xlReport7: XlPivotFormatTypeToString = "xlReport7"
        Case xlReport8: XlPivotFormatTypeToString = "xlReport8"
        Case xlReport9: XlPivotFormatTypeToString = "xlReport9"
        Case xlReport10: XlPivotFormatTypeToString = "xlReport10"
        Case xlTable1: XlPivotFormatTypeToString = "xlTable1"
        Case xlTable2: XlPivotFormatTypeToString = "xlTable2"
        Case xlTable3: XlPivotFormatTypeToString = "xlTable3"
        Case xlTable4: XlPivotFormatTypeToString = "xlTable4"
        Case xlTable5: XlPivotFormatTypeToString = "xlTable5"
        Case xlTable6: XlPivotFormatTypeToString = "xlTable6"
        Case xlTable7: XlPivotFormatTypeToString = "xlTable7"
        Case xlTable8: XlPivotFormatTypeToString = "xlTable8"
        Case xlTable9: XlPivotFormatTypeToString = "xlTable9"
        Case xlTable10: XlPivotFormatTypeToString = "xlTable10"
        Case xlPTClassic: XlPivotFormatTypeToString = "xlPTClassic"
        Case xlPTNone: XlPivotFormatTypeToString = "xlPTNone"
    End Select
End Function
