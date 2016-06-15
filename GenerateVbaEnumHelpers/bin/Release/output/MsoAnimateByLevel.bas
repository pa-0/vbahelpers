Attribute VB_Name = "wMsoAnimateByLevel"
Function MsoAnimateByLevelFromString(value As String) As MsoAnimateByLevel
    If IsNumeric(value) Then
        MsoAnimateByLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimateLevelNone": MsoAnimateByLevelFromString = msoAnimateLevelNone
        Case "msoAnimateTextByAllLevels": MsoAnimateByLevelFromString = msoAnimateTextByAllLevels
        Case "msoAnimateTextByFirstLevel": MsoAnimateByLevelFromString = msoAnimateTextByFirstLevel
        Case "msoAnimateTextBySecondLevel": MsoAnimateByLevelFromString = msoAnimateTextBySecondLevel
        Case "msoAnimateTextByThirdLevel": MsoAnimateByLevelFromString = msoAnimateTextByThirdLevel
        Case "msoAnimateTextByFourthLevel": MsoAnimateByLevelFromString = msoAnimateTextByFourthLevel
        Case "msoAnimateTextByFifthLevel": MsoAnimateByLevelFromString = msoAnimateTextByFifthLevel
        Case "msoAnimateChartAllAtOnce": MsoAnimateByLevelFromString = msoAnimateChartAllAtOnce
        Case "msoAnimateChartByCategory": MsoAnimateByLevelFromString = msoAnimateChartByCategory
        Case "msoAnimateChartByCategoryElements": MsoAnimateByLevelFromString = msoAnimateChartByCategoryElements
        Case "msoAnimateChartBySeries": MsoAnimateByLevelFromString = msoAnimateChartBySeries
        Case "msoAnimateChartBySeriesElements": MsoAnimateByLevelFromString = msoAnimateChartBySeriesElements
        Case "msoAnimateDiagramAllAtOnce": MsoAnimateByLevelFromString = msoAnimateDiagramAllAtOnce
        Case "msoAnimateDiagramDepthByNode": MsoAnimateByLevelFromString = msoAnimateDiagramDepthByNode
        Case "msoAnimateDiagramDepthByBranch": MsoAnimateByLevelFromString = msoAnimateDiagramDepthByBranch
        Case "msoAnimateDiagramBreadthByNode": MsoAnimateByLevelFromString = msoAnimateDiagramBreadthByNode
        Case "msoAnimateDiagramBreadthByLevel": MsoAnimateByLevelFromString = msoAnimateDiagramBreadthByLevel
        Case "msoAnimateDiagramClockwise": MsoAnimateByLevelFromString = msoAnimateDiagramClockwise
        Case "msoAnimateDiagramClockwiseIn": MsoAnimateByLevelFromString = msoAnimateDiagramClockwiseIn
        Case "msoAnimateDiagramClockwiseOut": MsoAnimateByLevelFromString = msoAnimateDiagramClockwiseOut
        Case "msoAnimateDiagramCounterClockwise": MsoAnimateByLevelFromString = msoAnimateDiagramCounterClockwise
        Case "msoAnimateDiagramCounterClockwiseIn": MsoAnimateByLevelFromString = msoAnimateDiagramCounterClockwiseIn
        Case "msoAnimateDiagramCounterClockwiseOut": MsoAnimateByLevelFromString = msoAnimateDiagramCounterClockwiseOut
        Case "msoAnimateDiagramInByRing": MsoAnimateByLevelFromString = msoAnimateDiagramInByRing
        Case "msoAnimateDiagramOutByRing": MsoAnimateByLevelFromString = msoAnimateDiagramOutByRing
        Case "msoAnimateDiagramUp": MsoAnimateByLevelFromString = msoAnimateDiagramUp
        Case "msoAnimateDiagramDown": MsoAnimateByLevelFromString = msoAnimateDiagramDown
        Case "msoAnimateLevelMixed": MsoAnimateByLevelFromString = msoAnimateLevelMixed
    End Select
End Function

Function MsoAnimateByLevelToString(value As MsoAnimateByLevel) As String
    Select Case value
        Case msoAnimateLevelNone: MsoAnimateByLevelToString = "msoAnimateLevelNone"
        Case msoAnimateTextByAllLevels: MsoAnimateByLevelToString = "msoAnimateTextByAllLevels"
        Case msoAnimateTextByFirstLevel: MsoAnimateByLevelToString = "msoAnimateTextByFirstLevel"
        Case msoAnimateTextBySecondLevel: MsoAnimateByLevelToString = "msoAnimateTextBySecondLevel"
        Case msoAnimateTextByThirdLevel: MsoAnimateByLevelToString = "msoAnimateTextByThirdLevel"
        Case msoAnimateTextByFourthLevel: MsoAnimateByLevelToString = "msoAnimateTextByFourthLevel"
        Case msoAnimateTextByFifthLevel: MsoAnimateByLevelToString = "msoAnimateTextByFifthLevel"
        Case msoAnimateChartAllAtOnce: MsoAnimateByLevelToString = "msoAnimateChartAllAtOnce"
        Case msoAnimateChartByCategory: MsoAnimateByLevelToString = "msoAnimateChartByCategory"
        Case msoAnimateChartByCategoryElements: MsoAnimateByLevelToString = "msoAnimateChartByCategoryElements"
        Case msoAnimateChartBySeries: MsoAnimateByLevelToString = "msoAnimateChartBySeries"
        Case msoAnimateChartBySeriesElements: MsoAnimateByLevelToString = "msoAnimateChartBySeriesElements"
        Case msoAnimateDiagramAllAtOnce: MsoAnimateByLevelToString = "msoAnimateDiagramAllAtOnce"
        Case msoAnimateDiagramDepthByNode: MsoAnimateByLevelToString = "msoAnimateDiagramDepthByNode"
        Case msoAnimateDiagramDepthByBranch: MsoAnimateByLevelToString = "msoAnimateDiagramDepthByBranch"
        Case msoAnimateDiagramBreadthByNode: MsoAnimateByLevelToString = "msoAnimateDiagramBreadthByNode"
        Case msoAnimateDiagramBreadthByLevel: MsoAnimateByLevelToString = "msoAnimateDiagramBreadthByLevel"
        Case msoAnimateDiagramClockwise: MsoAnimateByLevelToString = "msoAnimateDiagramClockwise"
        Case msoAnimateDiagramClockwiseIn: MsoAnimateByLevelToString = "msoAnimateDiagramClockwiseIn"
        Case msoAnimateDiagramClockwiseOut: MsoAnimateByLevelToString = "msoAnimateDiagramClockwiseOut"
        Case msoAnimateDiagramCounterClockwise: MsoAnimateByLevelToString = "msoAnimateDiagramCounterClockwise"
        Case msoAnimateDiagramCounterClockwiseIn: MsoAnimateByLevelToString = "msoAnimateDiagramCounterClockwiseIn"
        Case msoAnimateDiagramCounterClockwiseOut: MsoAnimateByLevelToString = "msoAnimateDiagramCounterClockwiseOut"
        Case msoAnimateDiagramInByRing: MsoAnimateByLevelToString = "msoAnimateDiagramInByRing"
        Case msoAnimateDiagramOutByRing: MsoAnimateByLevelToString = "msoAnimateDiagramOutByRing"
        Case msoAnimateDiagramUp: MsoAnimateByLevelToString = "msoAnimateDiagramUp"
        Case msoAnimateDiagramDown: MsoAnimateByLevelToString = "msoAnimateDiagramDown"
        Case msoAnimateLevelMixed: MsoAnimateByLevelToString = "msoAnimateLevelMixed"
    End Select
End Function
