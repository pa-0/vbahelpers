Attribute VB_Name = "wXlCubeFieldSubType"
Function XlCubeFieldSubTypeFromString(value As String) As XlCubeFieldSubType
    If IsNumeric(value) Then
        XlCubeFieldSubTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCubeHierarchy": XlCubeFieldSubTypeFromString = xlCubeHierarchy
        Case "xlCubeMeasure": XlCubeFieldSubTypeFromString = xlCubeMeasure
        Case "xlCubeSet": XlCubeFieldSubTypeFromString = xlCubeSet
        Case "xlCubeAttribute": XlCubeFieldSubTypeFromString = xlCubeAttribute
        Case "xlCubeCalculatedMeasure": XlCubeFieldSubTypeFromString = xlCubeCalculatedMeasure
        Case "xlCubeKPIValue": XlCubeFieldSubTypeFromString = xlCubeKPIValue
        Case "xlCubeKPIGoal": XlCubeFieldSubTypeFromString = xlCubeKPIGoal
        Case "xlCubeKPIStatus": XlCubeFieldSubTypeFromString = xlCubeKPIStatus
        Case "xlCubeKPITrend": XlCubeFieldSubTypeFromString = xlCubeKPITrend
        Case "xlCubeKPIWeight": XlCubeFieldSubTypeFromString = xlCubeKPIWeight
    End Select
End Function

Function XlCubeFieldSubTypeToString(value As XlCubeFieldSubType) As String
    Select Case value
        Case xlCubeHierarchy: XlCubeFieldSubTypeToString = "xlCubeHierarchy"
        Case xlCubeMeasure: XlCubeFieldSubTypeToString = "xlCubeMeasure"
        Case xlCubeSet: XlCubeFieldSubTypeToString = "xlCubeSet"
        Case xlCubeAttribute: XlCubeFieldSubTypeToString = "xlCubeAttribute"
        Case xlCubeCalculatedMeasure: XlCubeFieldSubTypeToString = "xlCubeCalculatedMeasure"
        Case xlCubeKPIValue: XlCubeFieldSubTypeToString = "xlCubeKPIValue"
        Case xlCubeKPIGoal: XlCubeFieldSubTypeToString = "xlCubeKPIGoal"
        Case xlCubeKPIStatus: XlCubeFieldSubTypeToString = "xlCubeKPIStatus"
        Case xlCubeKPITrend: XlCubeFieldSubTypeToString = "xlCubeKPITrend"
        Case xlCubeKPIWeight: XlCubeFieldSubTypeToString = "xlCubeKPIWeight"
    End Select
End Function
