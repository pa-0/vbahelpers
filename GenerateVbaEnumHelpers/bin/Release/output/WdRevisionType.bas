Attribute VB_Name = "wWdRevisionType"
Function WdRevisionTypeFromString(value As String) As WdRevisionType
    If IsNumeric(value) Then
        WdRevisionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNoRevision": WdRevisionTypeFromString = wdNoRevision
        Case "wdRevisionInsert": WdRevisionTypeFromString = wdRevisionInsert
        Case "wdRevisionDelete": WdRevisionTypeFromString = wdRevisionDelete
        Case "wdRevisionProperty": WdRevisionTypeFromString = wdRevisionProperty
        Case "wdRevisionParagraphNumber": WdRevisionTypeFromString = wdRevisionParagraphNumber
        Case "wdRevisionDisplayField": WdRevisionTypeFromString = wdRevisionDisplayField
        Case "wdRevisionReconcile": WdRevisionTypeFromString = wdRevisionReconcile
        Case "wdRevisionConflict": WdRevisionTypeFromString = wdRevisionConflict
        Case "wdRevisionStyle": WdRevisionTypeFromString = wdRevisionStyle
        Case "wdRevisionReplace": WdRevisionTypeFromString = wdRevisionReplace
        Case "wdRevisionParagraphProperty": WdRevisionTypeFromString = wdRevisionParagraphProperty
        Case "wdRevisionTableProperty": WdRevisionTypeFromString = wdRevisionTableProperty
        Case "wdRevisionSectionProperty": WdRevisionTypeFromString = wdRevisionSectionProperty
        Case "wdRevisionStyleDefinition": WdRevisionTypeFromString = wdRevisionStyleDefinition
        Case "wdRevisionMovedFrom": WdRevisionTypeFromString = wdRevisionMovedFrom
        Case "wdRevisionMovedTo": WdRevisionTypeFromString = wdRevisionMovedTo
        Case "wdRevisionCellInsertion": WdRevisionTypeFromString = wdRevisionCellInsertion
        Case "wdRevisionCellDeletion": WdRevisionTypeFromString = wdRevisionCellDeletion
        Case "wdRevisionCellMerge": WdRevisionTypeFromString = wdRevisionCellMerge
        Case "wdRevisionCellSplit": WdRevisionTypeFromString = wdRevisionCellSplit
        Case "wdRevisionConflictInsert": WdRevisionTypeFromString = wdRevisionConflictInsert
        Case "wdRevisionConflictDelete": WdRevisionTypeFromString = wdRevisionConflictDelete
    End Select
End Function

Function WdRevisionTypeToString(value As WdRevisionType) As String
    Select Case value
        Case wdNoRevision: WdRevisionTypeToString = "wdNoRevision"
        Case wdRevisionInsert: WdRevisionTypeToString = "wdRevisionInsert"
        Case wdRevisionDelete: WdRevisionTypeToString = "wdRevisionDelete"
        Case wdRevisionProperty: WdRevisionTypeToString = "wdRevisionProperty"
        Case wdRevisionParagraphNumber: WdRevisionTypeToString = "wdRevisionParagraphNumber"
        Case wdRevisionDisplayField: WdRevisionTypeToString = "wdRevisionDisplayField"
        Case wdRevisionReconcile: WdRevisionTypeToString = "wdRevisionReconcile"
        Case wdRevisionConflict: WdRevisionTypeToString = "wdRevisionConflict"
        Case wdRevisionStyle: WdRevisionTypeToString = "wdRevisionStyle"
        Case wdRevisionReplace: WdRevisionTypeToString = "wdRevisionReplace"
        Case wdRevisionParagraphProperty: WdRevisionTypeToString = "wdRevisionParagraphProperty"
        Case wdRevisionTableProperty: WdRevisionTypeToString = "wdRevisionTableProperty"
        Case wdRevisionSectionProperty: WdRevisionTypeToString = "wdRevisionSectionProperty"
        Case wdRevisionStyleDefinition: WdRevisionTypeToString = "wdRevisionStyleDefinition"
        Case wdRevisionMovedFrom: WdRevisionTypeToString = "wdRevisionMovedFrom"
        Case wdRevisionMovedTo: WdRevisionTypeToString = "wdRevisionMovedTo"
        Case wdRevisionCellInsertion: WdRevisionTypeToString = "wdRevisionCellInsertion"
        Case wdRevisionCellDeletion: WdRevisionTypeToString = "wdRevisionCellDeletion"
        Case wdRevisionCellMerge: WdRevisionTypeToString = "wdRevisionCellMerge"
        Case wdRevisionCellSplit: WdRevisionTypeToString = "wdRevisionCellSplit"
        Case wdRevisionConflictInsert: WdRevisionTypeToString = "wdRevisionConflictInsert"
        Case wdRevisionConflictDelete: WdRevisionTypeToString = "wdRevisionConflictDelete"
    End Select
End Function
