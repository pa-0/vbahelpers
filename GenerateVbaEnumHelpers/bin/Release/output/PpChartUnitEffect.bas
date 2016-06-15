Attribute VB_Name = "wPpChartUnitEffect"
Function PpChartUnitEffectFromString(value As String) As PpChartUnitEffect
    If IsNumeric(value) Then
        PpChartUnitEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAnimateBySeries": PpChartUnitEffectFromString = ppAnimateBySeries
        Case "ppAnimateByCategory": PpChartUnitEffectFromString = ppAnimateByCategory
        Case "ppAnimateBySeriesElements": PpChartUnitEffectFromString = ppAnimateBySeriesElements
        Case "ppAnimateByCategoryElements": PpChartUnitEffectFromString = ppAnimateByCategoryElements
        Case "ppAnimateChartAllAtOnce": PpChartUnitEffectFromString = ppAnimateChartAllAtOnce
        Case "ppAnimateChartMixed": PpChartUnitEffectFromString = ppAnimateChartMixed
    End Select
End Function

Function PpChartUnitEffectToString(value As PpChartUnitEffect) As String
    Select Case value
        Case ppAnimateBySeries: PpChartUnitEffectToString = "ppAnimateBySeries"
        Case ppAnimateByCategory: PpChartUnitEffectToString = "ppAnimateByCategory"
        Case ppAnimateBySeriesElements: PpChartUnitEffectToString = "ppAnimateBySeriesElements"
        Case ppAnimateByCategoryElements: PpChartUnitEffectToString = "ppAnimateByCategoryElements"
        Case ppAnimateChartAllAtOnce: PpChartUnitEffectToString = "ppAnimateChartAllAtOnce"
        Case ppAnimateChartMixed: PpChartUnitEffectToString = "ppAnimateChartMixed"
    End Select
End Function
