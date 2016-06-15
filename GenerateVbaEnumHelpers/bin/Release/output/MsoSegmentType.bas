Attribute VB_Name = "wMsoSegmentType"
Function MsoSegmentTypeFromString(value As String) As MsoSegmentType
    If IsNumeric(value) Then
        MsoSegmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSegmentLine": MsoSegmentTypeFromString = msoSegmentLine
        Case "msoSegmentCurve": MsoSegmentTypeFromString = msoSegmentCurve
    End Select
End Function

Function MsoSegmentTypeToString(value As MsoSegmentType) As String
    Select Case value
        Case msoSegmentLine: MsoSegmentTypeToString = "msoSegmentLine"
        Case msoSegmentCurve: MsoSegmentTypeToString = "msoSegmentCurve"
    End Select
End Function
