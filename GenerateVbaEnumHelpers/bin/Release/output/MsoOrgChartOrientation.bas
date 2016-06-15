Attribute VB_Name = "wMsoOrgChartOrientation"
Function MsoOrgChartOrientationFromString(value As String) As MsoOrgChartOrientation
    If IsNumeric(value) Then
        MsoOrgChartOrientationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOrgChartOrientationVertical": MsoOrgChartOrientationFromString = msoOrgChartOrientationVertical
        Case "msoOrgChartOrientationMixed": MsoOrgChartOrientationFromString = msoOrgChartOrientationMixed
    End Select
End Function

Function MsoOrgChartOrientationToString(value As MsoOrgChartOrientation) As String
    Select Case value
        Case msoOrgChartOrientationVertical: MsoOrgChartOrientationToString = "msoOrgChartOrientationVertical"
        Case msoOrgChartOrientationMixed: MsoOrgChartOrientationToString = "msoOrgChartOrientationMixed"
    End Select
End Function
