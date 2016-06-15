Attribute VB_Name = "wXlAutoFilterOperator"
Function XlAutoFilterOperatorFromString(value As String) As XlAutoFilterOperator
    If IsNumeric(value) Then
        XlAutoFilterOperatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAnd": XlAutoFilterOperatorFromString = xlAnd
        Case "xlOr": XlAutoFilterOperatorFromString = xlOr
        Case "xlTop10Items": XlAutoFilterOperatorFromString = xlTop10Items
        Case "xlBottom10Items": XlAutoFilterOperatorFromString = xlBottom10Items
        Case "xlTop10Percent": XlAutoFilterOperatorFromString = xlTop10Percent
        Case "xlBottom10Percent": XlAutoFilterOperatorFromString = xlBottom10Percent
        Case "xlFilterValues": XlAutoFilterOperatorFromString = xlFilterValues
        Case "xlFilterCellColor": XlAutoFilterOperatorFromString = xlFilterCellColor
        Case "xlFilterFontColor": XlAutoFilterOperatorFromString = xlFilterFontColor
        Case "xlFilterIcon": XlAutoFilterOperatorFromString = xlFilterIcon
        Case "xlFilterDynamic": XlAutoFilterOperatorFromString = xlFilterDynamic
        Case "xlFilterNoFill": XlAutoFilterOperatorFromString = xlFilterNoFill
        Case "xlFilterAutomaticFontColor": XlAutoFilterOperatorFromString = xlFilterAutomaticFontColor
        Case "xlFilterNoIcon": XlAutoFilterOperatorFromString = xlFilterNoIcon
    End Select
End Function

Function XlAutoFilterOperatorToString(value As XlAutoFilterOperator) As String
    Select Case value
        Case xlAnd: XlAutoFilterOperatorToString = "xlAnd"
        Case xlOr: XlAutoFilterOperatorToString = "xlOr"
        Case xlTop10Items: XlAutoFilterOperatorToString = "xlTop10Items"
        Case xlBottom10Items: XlAutoFilterOperatorToString = "xlBottom10Items"
        Case xlTop10Percent: XlAutoFilterOperatorToString = "xlTop10Percent"
        Case xlBottom10Percent: XlAutoFilterOperatorToString = "xlBottom10Percent"
        Case xlFilterValues: XlAutoFilterOperatorToString = "xlFilterValues"
        Case xlFilterCellColor: XlAutoFilterOperatorToString = "xlFilterCellColor"
        Case xlFilterFontColor: XlAutoFilterOperatorToString = "xlFilterFontColor"
        Case xlFilterIcon: XlAutoFilterOperatorToString = "xlFilterIcon"
        Case xlFilterDynamic: XlAutoFilterOperatorToString = "xlFilterDynamic"
        Case xlFilterNoFill: XlAutoFilterOperatorToString = "xlFilterNoFill"
        Case xlFilterAutomaticFontColor: XlAutoFilterOperatorToString = "xlFilterAutomaticFontColor"
        Case xlFilterNoIcon: XlAutoFilterOperatorToString = "xlFilterNoIcon"
    End Select
End Function
