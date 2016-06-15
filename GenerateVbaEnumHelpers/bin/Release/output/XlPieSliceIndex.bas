Attribute VB_Name = "wXlPieSliceIndex"
Function XlPieSliceIndexFromString(value As String) As XlPieSliceIndex
    If IsNumeric(value) Then
        XlPieSliceIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOuterCounterClockwisePoint": XlPieSliceIndexFromString = xlOuterCounterClockwisePoint
        Case "xlOuterCenterPoint": XlPieSliceIndexFromString = xlOuterCenterPoint
        Case "xlOuterClockwisePoint": XlPieSliceIndexFromString = xlOuterClockwisePoint
        Case "xlMidClockwiseRadiusPoint": XlPieSliceIndexFromString = xlMidClockwiseRadiusPoint
        Case "xlCenterPoint": XlPieSliceIndexFromString = xlCenterPoint
        Case "xlMidCounterClockwiseRadiusPoint": XlPieSliceIndexFromString = xlMidCounterClockwiseRadiusPoint
        Case "xlInnerClockwisePoint": XlPieSliceIndexFromString = xlInnerClockwisePoint
        Case "xlInnerCenterPoint": XlPieSliceIndexFromString = xlInnerCenterPoint
        Case "xlInnerCounterClockwisePoint": XlPieSliceIndexFromString = xlInnerCounterClockwisePoint
    End Select
End Function

Function XlPieSliceIndexToString(value As XlPieSliceIndex) As String
    Select Case value
        Case xlOuterCounterClockwisePoint: XlPieSliceIndexToString = "xlOuterCounterClockwisePoint"
        Case xlOuterCenterPoint: XlPieSliceIndexToString = "xlOuterCenterPoint"
        Case xlOuterClockwisePoint: XlPieSliceIndexToString = "xlOuterClockwisePoint"
        Case xlMidClockwiseRadiusPoint: XlPieSliceIndexToString = "xlMidClockwiseRadiusPoint"
        Case xlCenterPoint: XlPieSliceIndexToString = "xlCenterPoint"
        Case xlMidCounterClockwiseRadiusPoint: XlPieSliceIndexToString = "xlMidCounterClockwiseRadiusPoint"
        Case xlInnerClockwisePoint: XlPieSliceIndexToString = "xlInnerClockwisePoint"
        Case xlInnerCenterPoint: XlPieSliceIndexToString = "xlInnerCenterPoint"
        Case xlInnerCounterClockwisePoint: XlPieSliceIndexToString = "xlInnerCounterClockwisePoint"
    End Select
End Function
