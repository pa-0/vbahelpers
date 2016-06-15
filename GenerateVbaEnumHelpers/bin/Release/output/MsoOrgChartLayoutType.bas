Attribute VB_Name = "wMsoOrgChartLayoutType"
Function MsoOrgChartLayoutTypeFromString(value As String) As MsoOrgChartLayoutType
    If IsNumeric(value) Then
        MsoOrgChartLayoutTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoOrgChartLayoutStandard": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutStandard
        Case "msoOrgChartLayoutBothHanging": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutBothHanging
        Case "msoOrgChartLayoutLeftHanging": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutLeftHanging
        Case "msoOrgChartLayoutRightHanging": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutRightHanging
        Case "msoOrgChartLayoutDefault": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutDefault
        Case "msoOrgChartLayoutMixed": MsoOrgChartLayoutTypeFromString = msoOrgChartLayoutMixed
    End Select
End Function

Function MsoOrgChartLayoutTypeToString(value As MsoOrgChartLayoutType) As String
    Select Case value
        Case msoOrgChartLayoutStandard: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutStandard"
        Case msoOrgChartLayoutBothHanging: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutBothHanging"
        Case msoOrgChartLayoutLeftHanging: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutLeftHanging"
        Case msoOrgChartLayoutRightHanging: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutRightHanging"
        Case msoOrgChartLayoutDefault: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutDefault"
        Case msoOrgChartLayoutMixed: MsoOrgChartLayoutTypeToString = "msoOrgChartLayoutMixed"
    End Select
End Function
