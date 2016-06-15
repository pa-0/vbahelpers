Attribute VB_Name = "wXlSlicerCrossFilterType"
Function XlSlicerCrossFilterTypeFromString(value As String) As XlSlicerCrossFilterType
    If IsNumeric(value) Then
        XlSlicerCrossFilterTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSlicerNoCrossFilter": XlSlicerCrossFilterTypeFromString = xlSlicerNoCrossFilter
        Case "xlSlicerCrossFilterShowItemsWithDataAtTop": XlSlicerCrossFilterTypeFromString = xlSlicerCrossFilterShowItemsWithDataAtTop
        Case "xlSlicerCrossFilterShowItemsWithNoData": XlSlicerCrossFilterTypeFromString = xlSlicerCrossFilterShowItemsWithNoData
    End Select
End Function

Function XlSlicerCrossFilterTypeToString(value As XlSlicerCrossFilterType) As String
    Select Case value
        Case xlSlicerNoCrossFilter: XlSlicerCrossFilterTypeToString = "xlSlicerNoCrossFilter"
        Case xlSlicerCrossFilterShowItemsWithDataAtTop: XlSlicerCrossFilterTypeToString = "xlSlicerCrossFilterShowItemsWithDataAtTop"
        Case xlSlicerCrossFilterShowItemsWithNoData: XlSlicerCrossFilterTypeToString = "xlSlicerCrossFilterShowItemsWithNoData"
    End Select
End Function
