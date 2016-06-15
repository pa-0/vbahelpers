Attribute VB_Name = "wXlTableStyleElementType"
Function XlTableStyleElementTypeFromString(value As String) As XlTableStyleElementType
    If IsNumeric(value) Then
        XlTableStyleElementTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWholeTable": XlTableStyleElementTypeFromString = xlWholeTable
        Case "xlHeaderRow": XlTableStyleElementTypeFromString = xlHeaderRow
        Case "xlTotalRow": XlTableStyleElementTypeFromString = xlTotalRow
        Case "xlGrandTotalRow": XlTableStyleElementTypeFromString = xlGrandTotalRow
        Case "xlFirstColumn": XlTableStyleElementTypeFromString = xlFirstColumn
        Case "xlLastColumn": XlTableStyleElementTypeFromString = xlLastColumn
        Case "xlGrandTotalColumn": XlTableStyleElementTypeFromString = xlGrandTotalColumn
        Case "xlRowStripe1": XlTableStyleElementTypeFromString = xlRowStripe1
        Case "xlRowStripe2": XlTableStyleElementTypeFromString = xlRowStripe2
        Case "xlColumnStripe1": XlTableStyleElementTypeFromString = xlColumnStripe1
        Case "xlColumnStripe2": XlTableStyleElementTypeFromString = xlColumnStripe2
        Case "xlFirstHeaderCell": XlTableStyleElementTypeFromString = xlFirstHeaderCell
        Case "xlLastHeaderCell": XlTableStyleElementTypeFromString = xlLastHeaderCell
        Case "xlFirstTotalCell": XlTableStyleElementTypeFromString = xlFirstTotalCell
        Case "xlLastTotalCell": XlTableStyleElementTypeFromString = xlLastTotalCell
        Case "xlSubtotalColumn1": XlTableStyleElementTypeFromString = xlSubtotalColumn1
        Case "xlSubtotalColumn2": XlTableStyleElementTypeFromString = xlSubtotalColumn2
        Case "xlSubtotalColumn3": XlTableStyleElementTypeFromString = xlSubtotalColumn3
        Case "xlSubtotalRow1": XlTableStyleElementTypeFromString = xlSubtotalRow1
        Case "xlSubtotalRow2": XlTableStyleElementTypeFromString = xlSubtotalRow2
        Case "xlSubtotalRow3": XlTableStyleElementTypeFromString = xlSubtotalRow3
        Case "xlBlankRow": XlTableStyleElementTypeFromString = xlBlankRow
        Case "xlColumnSubheading1": XlTableStyleElementTypeFromString = xlColumnSubheading1
        Case "xlColumnSubheading2": XlTableStyleElementTypeFromString = xlColumnSubheading2
        Case "xlColumnSubheading3": XlTableStyleElementTypeFromString = xlColumnSubheading3
        Case "xlRowSubheading1": XlTableStyleElementTypeFromString = xlRowSubheading1
        Case "xlRowSubheading2": XlTableStyleElementTypeFromString = xlRowSubheading2
        Case "xlRowSubheading3": XlTableStyleElementTypeFromString = xlRowSubheading3
        Case "xlPageFieldLabels": XlTableStyleElementTypeFromString = xlPageFieldLabels
        Case "xlPageFieldValues": XlTableStyleElementTypeFromString = xlPageFieldValues
        Case "xlSlicerUnselectedItemWithData": XlTableStyleElementTypeFromString = xlSlicerUnselectedItemWithData
        Case "xlSlicerUnselectedItemWithNoData": XlTableStyleElementTypeFromString = xlSlicerUnselectedItemWithNoData
        Case "xlSlicerSelectedItemWithData": XlTableStyleElementTypeFromString = xlSlicerSelectedItemWithData
        Case "xlSlicerSelectedItemWithNoData": XlTableStyleElementTypeFromString = xlSlicerSelectedItemWithNoData
        Case "xlSlicerHoveredUnselectedItemWithData": XlTableStyleElementTypeFromString = xlSlicerHoveredUnselectedItemWithData
        Case "xlSlicerHoveredSelectedItemWithData": XlTableStyleElementTypeFromString = xlSlicerHoveredSelectedItemWithData
        Case "xlSlicerHoveredUnselectedItemWithNoData": XlTableStyleElementTypeFromString = xlSlicerHoveredUnselectedItemWithNoData
        Case "xlSlicerHoveredSelectedItemWithNoData": XlTableStyleElementTypeFromString = xlSlicerHoveredSelectedItemWithNoData
    End Select
End Function

Function XlTableStyleElementTypeToString(value As XlTableStyleElementType) As String
    Select Case value
        Case xlWholeTable: XlTableStyleElementTypeToString = "xlWholeTable"
        Case xlHeaderRow: XlTableStyleElementTypeToString = "xlHeaderRow"
        Case xlTotalRow: XlTableStyleElementTypeToString = "xlTotalRow"
        Case xlGrandTotalRow: XlTableStyleElementTypeToString = "xlGrandTotalRow"
        Case xlFirstColumn: XlTableStyleElementTypeToString = "xlFirstColumn"
        Case xlLastColumn: XlTableStyleElementTypeToString = "xlLastColumn"
        Case xlGrandTotalColumn: XlTableStyleElementTypeToString = "xlGrandTotalColumn"
        Case xlRowStripe1: XlTableStyleElementTypeToString = "xlRowStripe1"
        Case xlRowStripe2: XlTableStyleElementTypeToString = "xlRowStripe2"
        Case xlColumnStripe1: XlTableStyleElementTypeToString = "xlColumnStripe1"
        Case xlColumnStripe2: XlTableStyleElementTypeToString = "xlColumnStripe2"
        Case xlFirstHeaderCell: XlTableStyleElementTypeToString = "xlFirstHeaderCell"
        Case xlLastHeaderCell: XlTableStyleElementTypeToString = "xlLastHeaderCell"
        Case xlFirstTotalCell: XlTableStyleElementTypeToString = "xlFirstTotalCell"
        Case xlLastTotalCell: XlTableStyleElementTypeToString = "xlLastTotalCell"
        Case xlSubtotalColumn1: XlTableStyleElementTypeToString = "xlSubtotalColumn1"
        Case xlSubtotalColumn2: XlTableStyleElementTypeToString = "xlSubtotalColumn2"
        Case xlSubtotalColumn3: XlTableStyleElementTypeToString = "xlSubtotalColumn3"
        Case xlSubtotalRow1: XlTableStyleElementTypeToString = "xlSubtotalRow1"
        Case xlSubtotalRow2: XlTableStyleElementTypeToString = "xlSubtotalRow2"
        Case xlSubtotalRow3: XlTableStyleElementTypeToString = "xlSubtotalRow3"
        Case xlBlankRow: XlTableStyleElementTypeToString = "xlBlankRow"
        Case xlColumnSubheading1: XlTableStyleElementTypeToString = "xlColumnSubheading1"
        Case xlColumnSubheading2: XlTableStyleElementTypeToString = "xlColumnSubheading2"
        Case xlColumnSubheading3: XlTableStyleElementTypeToString = "xlColumnSubheading3"
        Case xlRowSubheading1: XlTableStyleElementTypeToString = "xlRowSubheading1"
        Case xlRowSubheading2: XlTableStyleElementTypeToString = "xlRowSubheading2"
        Case xlRowSubheading3: XlTableStyleElementTypeToString = "xlRowSubheading3"
        Case xlPageFieldLabels: XlTableStyleElementTypeToString = "xlPageFieldLabels"
        Case xlPageFieldValues: XlTableStyleElementTypeToString = "xlPageFieldValues"
        Case xlSlicerUnselectedItemWithData: XlTableStyleElementTypeToString = "xlSlicerUnselectedItemWithData"
        Case xlSlicerUnselectedItemWithNoData: XlTableStyleElementTypeToString = "xlSlicerUnselectedItemWithNoData"
        Case xlSlicerSelectedItemWithData: XlTableStyleElementTypeToString = "xlSlicerSelectedItemWithData"
        Case xlSlicerSelectedItemWithNoData: XlTableStyleElementTypeToString = "xlSlicerSelectedItemWithNoData"
        Case xlSlicerHoveredUnselectedItemWithData: XlTableStyleElementTypeToString = "xlSlicerHoveredUnselectedItemWithData"
        Case xlSlicerHoveredSelectedItemWithData: XlTableStyleElementTypeToString = "xlSlicerHoveredSelectedItemWithData"
        Case xlSlicerHoveredUnselectedItemWithNoData: XlTableStyleElementTypeToString = "xlSlicerHoveredUnselectedItemWithNoData"
        Case xlSlicerHoveredSelectedItemWithNoData: XlTableStyleElementTypeToString = "xlSlicerHoveredSelectedItemWithNoData"
    End Select
End Function
