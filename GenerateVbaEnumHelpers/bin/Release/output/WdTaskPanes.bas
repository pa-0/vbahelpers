Attribute VB_Name = "wWdTaskPanes"
Function WdTaskPanesFromString(value As String) As WdTaskPanes
    If IsNumeric(value) Then
        WdTaskPanesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTaskPaneFormatting": WdTaskPanesFromString = wdTaskPaneFormatting
        Case "wdTaskPaneRevealFormatting": WdTaskPanesFromString = wdTaskPaneRevealFormatting
        Case "wdTaskPaneMailMerge": WdTaskPanesFromString = wdTaskPaneMailMerge
        Case "wdTaskPaneTranslate": WdTaskPanesFromString = wdTaskPaneTranslate
        Case "wdTaskPaneSearch": WdTaskPanesFromString = wdTaskPaneSearch
        Case "wdTaskPaneXMLStructure": WdTaskPanesFromString = wdTaskPaneXMLStructure
        Case "wdTaskPaneDocumentProtection": WdTaskPanesFromString = wdTaskPaneDocumentProtection
        Case "wdTaskPaneDocumentActions": WdTaskPanesFromString = wdTaskPaneDocumentActions
        Case "wdTaskPaneSharedWorkspace": WdTaskPanesFromString = wdTaskPaneSharedWorkspace
        Case "wdTaskPaneHelp": WdTaskPanesFromString = wdTaskPaneHelp
        Case "wdTaskPaneResearch": WdTaskPanesFromString = wdTaskPaneResearch
        Case "wdTaskPaneFaxService": WdTaskPanesFromString = wdTaskPaneFaxService
        Case "wdTaskPaneXMLDocument": WdTaskPanesFromString = wdTaskPaneXMLDocument
        Case "wdTaskPaneDocumentUpdates": WdTaskPanesFromString = wdTaskPaneDocumentUpdates
        Case "wdTaskPaneSignature": WdTaskPanesFromString = wdTaskPaneSignature
        Case "wdTaskPaneStyleInspector": WdTaskPanesFromString = wdTaskPaneStyleInspector
        Case "wdTaskPaneDocumentManagement": WdTaskPanesFromString = wdTaskPaneDocumentManagement
        Case "wdTaskPaneApplyStyles": WdTaskPanesFromString = wdTaskPaneApplyStyles
        Case "wdTaskPaneNav": WdTaskPanesFromString = wdTaskPaneNav
        Case "wdTaskPaneSelection": WdTaskPanesFromString = wdTaskPaneSelection
    End Select
End Function

Function WdTaskPanesToString(value As WdTaskPanes) As String
    Select Case value
        Case wdTaskPaneFormatting: WdTaskPanesToString = "wdTaskPaneFormatting"
        Case wdTaskPaneRevealFormatting: WdTaskPanesToString = "wdTaskPaneRevealFormatting"
        Case wdTaskPaneMailMerge: WdTaskPanesToString = "wdTaskPaneMailMerge"
        Case wdTaskPaneTranslate: WdTaskPanesToString = "wdTaskPaneTranslate"
        Case wdTaskPaneSearch: WdTaskPanesToString = "wdTaskPaneSearch"
        Case wdTaskPaneXMLStructure: WdTaskPanesToString = "wdTaskPaneXMLStructure"
        Case wdTaskPaneDocumentProtection: WdTaskPanesToString = "wdTaskPaneDocumentProtection"
        Case wdTaskPaneDocumentActions: WdTaskPanesToString = "wdTaskPaneDocumentActions"
        Case wdTaskPaneSharedWorkspace: WdTaskPanesToString = "wdTaskPaneSharedWorkspace"
        Case wdTaskPaneHelp: WdTaskPanesToString = "wdTaskPaneHelp"
        Case wdTaskPaneResearch: WdTaskPanesToString = "wdTaskPaneResearch"
        Case wdTaskPaneFaxService: WdTaskPanesToString = "wdTaskPaneFaxService"
        Case wdTaskPaneXMLDocument: WdTaskPanesToString = "wdTaskPaneXMLDocument"
        Case wdTaskPaneDocumentUpdates: WdTaskPanesToString = "wdTaskPaneDocumentUpdates"
        Case wdTaskPaneSignature: WdTaskPanesToString = "wdTaskPaneSignature"
        Case wdTaskPaneStyleInspector: WdTaskPanesToString = "wdTaskPaneStyleInspector"
        Case wdTaskPaneDocumentManagement: WdTaskPanesToString = "wdTaskPaneDocumentManagement"
        Case wdTaskPaneApplyStyles: WdTaskPanesToString = "wdTaskPaneApplyStyles"
        Case wdTaskPaneNav: WdTaskPanesToString = "wdTaskPaneNav"
        Case wdTaskPaneSelection: WdTaskPanesToString = "wdTaskPaneSelection"
    End Select
End Function
