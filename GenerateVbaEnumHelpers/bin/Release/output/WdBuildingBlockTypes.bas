Attribute VB_Name = "wWdBuildingBlockTypes"
Function WdBuildingBlockTypesFromString(value As String) As WdBuildingBlockTypes
    If IsNumeric(value) Then
        WdBuildingBlockTypesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTypeQuickParts": WdBuildingBlockTypesFromString = wdTypeQuickParts
        Case "wdTypeCoverPage": WdBuildingBlockTypesFromString = wdTypeCoverPage
        Case "wdTypeEquations": WdBuildingBlockTypesFromString = wdTypeEquations
        Case "wdTypeFooters": WdBuildingBlockTypesFromString = wdTypeFooters
        Case "wdTypeHeaders": WdBuildingBlockTypesFromString = wdTypeHeaders
        Case "wdTypePageNumber": WdBuildingBlockTypesFromString = wdTypePageNumber
        Case "wdTypeTables": WdBuildingBlockTypesFromString = wdTypeTables
        Case "wdTypeWatermarks": WdBuildingBlockTypesFromString = wdTypeWatermarks
        Case "wdTypeAutoText": WdBuildingBlockTypesFromString = wdTypeAutoText
        Case "wdTypeTextBox": WdBuildingBlockTypesFromString = wdTypeTextBox
        Case "wdTypePageNumberTop": WdBuildingBlockTypesFromString = wdTypePageNumberTop
        Case "wdTypePageNumberBottom": WdBuildingBlockTypesFromString = wdTypePageNumberBottom
        Case "wdTypePageNumberPage": WdBuildingBlockTypesFromString = wdTypePageNumberPage
        Case "wdTypeTableOfContents": WdBuildingBlockTypesFromString = wdTypeTableOfContents
        Case "wdTypeCustomQuickParts": WdBuildingBlockTypesFromString = wdTypeCustomQuickParts
        Case "wdTypeCustomCoverPage": WdBuildingBlockTypesFromString = wdTypeCustomCoverPage
        Case "wdTypeCustomEquations": WdBuildingBlockTypesFromString = wdTypeCustomEquations
        Case "wdTypeCustomFooters": WdBuildingBlockTypesFromString = wdTypeCustomFooters
        Case "wdTypeCustomHeaders": WdBuildingBlockTypesFromString = wdTypeCustomHeaders
        Case "wdTypeCustomPageNumber": WdBuildingBlockTypesFromString = wdTypeCustomPageNumber
        Case "wdTypeCustomTables": WdBuildingBlockTypesFromString = wdTypeCustomTables
        Case "wdTypeCustomWatermarks": WdBuildingBlockTypesFromString = wdTypeCustomWatermarks
        Case "wdTypeCustomAutoText": WdBuildingBlockTypesFromString = wdTypeCustomAutoText
        Case "wdTypeCustomTextBox": WdBuildingBlockTypesFromString = wdTypeCustomTextBox
        Case "wdTypeCustomPageNumberTop": WdBuildingBlockTypesFromString = wdTypeCustomPageNumberTop
        Case "wdTypeCustomPageNumberBottom": WdBuildingBlockTypesFromString = wdTypeCustomPageNumberBottom
        Case "wdTypeCustomPageNumberPage": WdBuildingBlockTypesFromString = wdTypeCustomPageNumberPage
        Case "wdTypeCustomTableOfContents": WdBuildingBlockTypesFromString = wdTypeCustomTableOfContents
        Case "wdTypeCustom1": WdBuildingBlockTypesFromString = wdTypeCustom1
        Case "wdTypeCustom2": WdBuildingBlockTypesFromString = wdTypeCustom2
        Case "wdTypeCustom3": WdBuildingBlockTypesFromString = wdTypeCustom3
        Case "wdTypeCustom4": WdBuildingBlockTypesFromString = wdTypeCustom4
        Case "wdTypeCustom5": WdBuildingBlockTypesFromString = wdTypeCustom5
        Case "wdTypeBibliography": WdBuildingBlockTypesFromString = wdTypeBibliography
        Case "wdTypeCustomBibliography": WdBuildingBlockTypesFromString = wdTypeCustomBibliography
    End Select
End Function

Function WdBuildingBlockTypesToString(value As WdBuildingBlockTypes) As String
    Select Case value
        Case wdTypeQuickParts: WdBuildingBlockTypesToString = "wdTypeQuickParts"
        Case wdTypeCoverPage: WdBuildingBlockTypesToString = "wdTypeCoverPage"
        Case wdTypeEquations: WdBuildingBlockTypesToString = "wdTypeEquations"
        Case wdTypeFooters: WdBuildingBlockTypesToString = "wdTypeFooters"
        Case wdTypeHeaders: WdBuildingBlockTypesToString = "wdTypeHeaders"
        Case wdTypePageNumber: WdBuildingBlockTypesToString = "wdTypePageNumber"
        Case wdTypeTables: WdBuildingBlockTypesToString = "wdTypeTables"
        Case wdTypeWatermarks: WdBuildingBlockTypesToString = "wdTypeWatermarks"
        Case wdTypeAutoText: WdBuildingBlockTypesToString = "wdTypeAutoText"
        Case wdTypeTextBox: WdBuildingBlockTypesToString = "wdTypeTextBox"
        Case wdTypePageNumberTop: WdBuildingBlockTypesToString = "wdTypePageNumberTop"
        Case wdTypePageNumberBottom: WdBuildingBlockTypesToString = "wdTypePageNumberBottom"
        Case wdTypePageNumberPage: WdBuildingBlockTypesToString = "wdTypePageNumberPage"
        Case wdTypeTableOfContents: WdBuildingBlockTypesToString = "wdTypeTableOfContents"
        Case wdTypeCustomQuickParts: WdBuildingBlockTypesToString = "wdTypeCustomQuickParts"
        Case wdTypeCustomCoverPage: WdBuildingBlockTypesToString = "wdTypeCustomCoverPage"
        Case wdTypeCustomEquations: WdBuildingBlockTypesToString = "wdTypeCustomEquations"
        Case wdTypeCustomFooters: WdBuildingBlockTypesToString = "wdTypeCustomFooters"
        Case wdTypeCustomHeaders: WdBuildingBlockTypesToString = "wdTypeCustomHeaders"
        Case wdTypeCustomPageNumber: WdBuildingBlockTypesToString = "wdTypeCustomPageNumber"
        Case wdTypeCustomTables: WdBuildingBlockTypesToString = "wdTypeCustomTables"
        Case wdTypeCustomWatermarks: WdBuildingBlockTypesToString = "wdTypeCustomWatermarks"
        Case wdTypeCustomAutoText: WdBuildingBlockTypesToString = "wdTypeCustomAutoText"
        Case wdTypeCustomTextBox: WdBuildingBlockTypesToString = "wdTypeCustomTextBox"
        Case wdTypeCustomPageNumberTop: WdBuildingBlockTypesToString = "wdTypeCustomPageNumberTop"
        Case wdTypeCustomPageNumberBottom: WdBuildingBlockTypesToString = "wdTypeCustomPageNumberBottom"
        Case wdTypeCustomPageNumberPage: WdBuildingBlockTypesToString = "wdTypeCustomPageNumberPage"
        Case wdTypeCustomTableOfContents: WdBuildingBlockTypesToString = "wdTypeCustomTableOfContents"
        Case wdTypeCustom1: WdBuildingBlockTypesToString = "wdTypeCustom1"
        Case wdTypeCustom2: WdBuildingBlockTypesToString = "wdTypeCustom2"
        Case wdTypeCustom3: WdBuildingBlockTypesToString = "wdTypeCustom3"
        Case wdTypeCustom4: WdBuildingBlockTypesToString = "wdTypeCustom4"
        Case wdTypeCustom5: WdBuildingBlockTypesToString = "wdTypeCustom5"
        Case wdTypeBibliography: WdBuildingBlockTypesToString = "wdTypeBibliography"
        Case wdTypeCustomBibliography: WdBuildingBlockTypesToString = "wdTypeCustomBibliography"
    End Select
End Function
