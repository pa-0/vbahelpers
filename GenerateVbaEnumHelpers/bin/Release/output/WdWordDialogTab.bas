Attribute VB_Name = "wWdWordDialogTab"
Function WdWordDialogTabFromString(value As String) As WdWordDialogTab
    If IsNumeric(value) Then
        WdWordDialogTabFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDialogToolsOptionsTabGeneral": WdWordDialogTabFromString = wdDialogToolsOptionsTabGeneral
        Case "wdDialogToolsOptionsTabView": WdWordDialogTabFromString = wdDialogToolsOptionsTabView
        Case "wdDialogToolsOptionsTabPrint": WdWordDialogTabFromString = wdDialogToolsOptionsTabPrint
        Case "wdDialogToolsOptionsTabSave": WdWordDialogTabFromString = wdDialogToolsOptionsTabSave
        Case "wdDialogToolsOptionsTabProofread": WdWordDialogTabFromString = wdDialogToolsOptionsTabProofread
        Case "wdDialogToolsOptionsTabUserInfo": WdWordDialogTabFromString = wdDialogToolsOptionsTabUserInfo
        Case "wdDialogToolsOptionsTabEdit": WdWordDialogTabFromString = wdDialogToolsOptionsTabEdit
        Case "wdDialogToolsOptionsTabFileLocations": WdWordDialogTabFromString = wdDialogToolsOptionsTabFileLocations
        Case "wdDialogToolsOptionsTabTrackChanges": WdWordDialogTabFromString = wdDialogToolsOptionsTabTrackChanges
        Case "wdDialogToolsOptionsTabCompatibility": WdWordDialogTabFromString = wdDialogToolsOptionsTabCompatibility
        Case "wdDialogToolsOptionsTabTypography": WdWordDialogTabFromString = wdDialogToolsOptionsTabTypography
        Case "wdDialogToolsOptionsTabHangulHanjaConversion": WdWordDialogTabFromString = wdDialogToolsOptionsTabHangulHanjaConversion
        Case "wdDialogToolsOptionsTabFuzzy": WdWordDialogTabFromString = wdDialogToolsOptionsTabFuzzy
        Case "wdDialogToolsOptionsTabBidi": WdWordDialogTabFromString = wdDialogToolsOptionsTabBidi
        Case "wdDialogToolsOptionsTabAcetate": WdWordDialogTabFromString = wdDialogToolsOptionsTabAcetate
        Case "wdDialogToolsOptionsTabSecurity": WdWordDialogTabFromString = wdDialogToolsOptionsTabSecurity
        Case "wdDialogFilePageSetupTabMargins": WdWordDialogTabFromString = wdDialogFilePageSetupTabMargins
        Case "wdDialogFilePageSetupTabPaper": WdWordDialogTabFromString = wdDialogFilePageSetupTabPaper
        Case "wdDialogFilePageSetupTabLayout": WdWordDialogTabFromString = wdDialogFilePageSetupTabLayout
        Case "wdDialogFilePageSetupTabCharsLines": WdWordDialogTabFromString = wdDialogFilePageSetupTabCharsLines
        Case "wdDialogInsertSymbolTabSymbols": WdWordDialogTabFromString = wdDialogInsertSymbolTabSymbols
        Case "wdDialogInsertSymbolTabSpecialCharacters": WdWordDialogTabFromString = wdDialogInsertSymbolTabSpecialCharacters
        Case "wdDialogNoteOptionsTabAllFootnotes": WdWordDialogTabFromString = wdDialogNoteOptionsTabAllFootnotes
        Case "wdDialogNoteOptionsTabAllEndnotes": WdWordDialogTabFromString = wdDialogNoteOptionsTabAllEndnotes
        Case "wdDialogInsertIndexAndTablesTabIndex": WdWordDialogTabFromString = wdDialogInsertIndexAndTablesTabIndex
        Case "wdDialogInsertIndexAndTablesTabTableOfContents": WdWordDialogTabFromString = wdDialogInsertIndexAndTablesTabTableOfContents
        Case "wdDialogInsertIndexAndTablesTabTableOfFigures": WdWordDialogTabFromString = wdDialogInsertIndexAndTablesTabTableOfFigures
        Case "wdDialogInsertIndexAndTablesTabTableOfAuthorities": WdWordDialogTabFromString = wdDialogInsertIndexAndTablesTabTableOfAuthorities
        Case "wdDialogOrganizerTabStyles": WdWordDialogTabFromString = wdDialogOrganizerTabStyles
        Case "wdDialogOrganizerTabAutoText": WdWordDialogTabFromString = wdDialogOrganizerTabAutoText
        Case "wdDialogOrganizerTabCommandBars": WdWordDialogTabFromString = wdDialogOrganizerTabCommandBars
        Case "wdDialogOrganizerTabMacros": WdWordDialogTabFromString = wdDialogOrganizerTabMacros
        Case "wdDialogFormatFontTabFont": WdWordDialogTabFromString = wdDialogFormatFontTabFont
        Case "wdDialogFormatFontTabCharacterSpacing": WdWordDialogTabFromString = wdDialogFormatFontTabCharacterSpacing
        Case "wdDialogFormatFontTabAnimation": WdWordDialogTabFromString = wdDialogFormatFontTabAnimation
        Case "wdDialogFormatBordersAndShadingTabBorders": WdWordDialogTabFromString = wdDialogFormatBordersAndShadingTabBorders
        Case "wdDialogFormatBordersAndShadingTabPageBorder": WdWordDialogTabFromString = wdDialogFormatBordersAndShadingTabPageBorder
        Case "wdDialogFormatBordersAndShadingTabShading": WdWordDialogTabFromString = wdDialogFormatBordersAndShadingTabShading
        Case "wdDialogToolsEnvelopesAndLabelsTabEnvelopes": WdWordDialogTabFromString = wdDialogToolsEnvelopesAndLabelsTabEnvelopes
        Case "wdDialogToolsEnvelopesAndLabelsTabLabels": WdWordDialogTabFromString = wdDialogToolsEnvelopesAndLabelsTabLabels
        Case "wdDialogFormatParagraphTabIndentsAndSpacing": WdWordDialogTabFromString = wdDialogFormatParagraphTabIndentsAndSpacing
        Case "wdDialogFormatParagraphTabTextFlow": WdWordDialogTabFromString = wdDialogFormatParagraphTabTextFlow
        Case "wdDialogFormatParagraphTabTeisai": WdWordDialogTabFromString = wdDialogFormatParagraphTabTeisai
        Case "wdDialogFormatDrawingObjectTabColorsAndLines": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabColorsAndLines
        Case "wdDialogFormatDrawingObjectTabSize": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabSize
        Case "wdDialogFormatDrawingObjectTabPosition": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabPosition
        Case "wdDialogFormatDrawingObjectTabWrapping": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabWrapping
        Case "wdDialogFormatDrawingObjectTabPicture": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabPicture
        Case "wdDialogFormatDrawingObjectTabTextbox": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabTextbox
        Case "wdDialogFormatDrawingObjectTabWeb": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabWeb
        Case "wdDialogFormatDrawingObjectTabHR": WdWordDialogTabFromString = wdDialogFormatDrawingObjectTabHR
        Case "wdDialogToolsAutoCorrectExceptionsTabFirstLetter": WdWordDialogTabFromString = wdDialogToolsAutoCorrectExceptionsTabFirstLetter
        Case "wdDialogToolsAutoCorrectExceptionsTabInitialCaps": WdWordDialogTabFromString = wdDialogToolsAutoCorrectExceptionsTabInitialCaps
        Case "wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet": WdWordDialogTabFromString = wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet
        Case "wdDialogToolsAutoCorrectExceptionsTabIac": WdWordDialogTabFromString = wdDialogToolsAutoCorrectExceptionsTabIac
        Case "wdDialogFormatBulletsAndNumberingTabBulleted": WdWordDialogTabFromString = wdDialogFormatBulletsAndNumberingTabBulleted
        Case "wdDialogFormatBulletsAndNumberingTabNumbered": WdWordDialogTabFromString = wdDialogFormatBulletsAndNumberingTabNumbered
        Case "wdDialogFormatBulletsAndNumberingTabOutlineNumbered": WdWordDialogTabFromString = wdDialogFormatBulletsAndNumberingTabOutlineNumbered
        Case "wdDialogLetterWizardTabLetterFormat": WdWordDialogTabFromString = wdDialogLetterWizardTabLetterFormat
        Case "wdDialogLetterWizardTabRecipientInfo": WdWordDialogTabFromString = wdDialogLetterWizardTabRecipientInfo
        Case "wdDialogLetterWizardTabOtherElements": WdWordDialogTabFromString = wdDialogLetterWizardTabOtherElements
        Case "wdDialogLetterWizardTabSenderInfo": WdWordDialogTabFromString = wdDialogLetterWizardTabSenderInfo
        Case "wdDialogToolsAutoManagerTabAutoCorrect": WdWordDialogTabFromString = wdDialogToolsAutoManagerTabAutoCorrect
        Case "wdDialogToolsAutoManagerTabAutoFormatAsYouType": WdWordDialogTabFromString = wdDialogToolsAutoManagerTabAutoFormatAsYouType
        Case "wdDialogToolsAutoManagerTabAutoText": WdWordDialogTabFromString = wdDialogToolsAutoManagerTabAutoText
        Case "wdDialogToolsAutoManagerTabAutoFormat": WdWordDialogTabFromString = wdDialogToolsAutoManagerTabAutoFormat
        Case "wdDialogToolsAutoManagerTabSmartTags": WdWordDialogTabFromString = wdDialogToolsAutoManagerTabSmartTags
        Case "wdDialogTablePropertiesTabTable": WdWordDialogTabFromString = wdDialogTablePropertiesTabTable
        Case "wdDialogTablePropertiesTabRow": WdWordDialogTabFromString = wdDialogTablePropertiesTabRow
        Case "wdDialogTablePropertiesTabColumn": WdWordDialogTabFromString = wdDialogTablePropertiesTabColumn
        Case "wdDialogTablePropertiesTabCell": WdWordDialogTabFromString = wdDialogTablePropertiesTabCell
        Case "wdDialogEmailOptionsTabSignature": WdWordDialogTabFromString = wdDialogEmailOptionsTabSignature
        Case "wdDialogEmailOptionsTabStationary": WdWordDialogTabFromString = wdDialogEmailOptionsTabStationary
        Case "wdDialogEmailOptionsTabQuoting": WdWordDialogTabFromString = wdDialogEmailOptionsTabQuoting
        Case "wdDialogWebOptionsGeneral": WdWordDialogTabFromString = wdDialogWebOptionsGeneral
        Case "wdDialogWebOptionsBrowsers": WdWordDialogTabFromString = wdDialogWebOptionsBrowsers
        Case "wdDialogWebOptionsFiles": WdWordDialogTabFromString = wdDialogWebOptionsFiles
        Case "wdDialogWebOptionsPictures": WdWordDialogTabFromString = wdDialogWebOptionsPictures
        Case "wdDialogWebOptionsEncoding": WdWordDialogTabFromString = wdDialogWebOptionsEncoding
        Case "wdDialogWebOptionsFonts": WdWordDialogTabFromString = wdDialogWebOptionsFonts
        Case "wdDialogTemplates": WdWordDialogTabFromString = wdDialogTemplates
        Case "wdDialogTemplatesXMLSchema": WdWordDialogTabFromString = wdDialogTemplatesXMLSchema
        Case "wdDialogTemplatesXMLExpansionPacks": WdWordDialogTabFromString = wdDialogTemplatesXMLExpansionPacks
        Case "wdDialogTemplatesLinkedCSS": WdWordDialogTabFromString = wdDialogTemplatesLinkedCSS
        Case "wdDialogStyleManagementTabEdit": WdWordDialogTabFromString = wdDialogStyleManagementTabEdit
        Case "wdDialogStyleManagementTabRecommend": WdWordDialogTabFromString = wdDialogStyleManagementTabRecommend
        Case "wdDialogStyleManagementTabRestrict": WdWordDialogTabFromString = wdDialogStyleManagementTabRestrict
    End Select
End Function

Function WdWordDialogTabToString(value As WdWordDialogTab) As String
    Select Case value
        Case wdDialogToolsOptionsTabGeneral: WdWordDialogTabToString = "wdDialogToolsOptionsTabGeneral"
        Case wdDialogToolsOptionsTabView: WdWordDialogTabToString = "wdDialogToolsOptionsTabView"
        Case wdDialogToolsOptionsTabPrint: WdWordDialogTabToString = "wdDialogToolsOptionsTabPrint"
        Case wdDialogToolsOptionsTabSave: WdWordDialogTabToString = "wdDialogToolsOptionsTabSave"
        Case wdDialogToolsOptionsTabProofread: WdWordDialogTabToString = "wdDialogToolsOptionsTabProofread"
        Case wdDialogToolsOptionsTabUserInfo: WdWordDialogTabToString = "wdDialogToolsOptionsTabUserInfo"
        Case wdDialogToolsOptionsTabEdit: WdWordDialogTabToString = "wdDialogToolsOptionsTabEdit"
        Case wdDialogToolsOptionsTabFileLocations: WdWordDialogTabToString = "wdDialogToolsOptionsTabFileLocations"
        Case wdDialogToolsOptionsTabTrackChanges: WdWordDialogTabToString = "wdDialogToolsOptionsTabTrackChanges"
        Case wdDialogToolsOptionsTabCompatibility: WdWordDialogTabToString = "wdDialogToolsOptionsTabCompatibility"
        Case wdDialogToolsOptionsTabTypography: WdWordDialogTabToString = "wdDialogToolsOptionsTabTypography"
        Case wdDialogToolsOptionsTabHangulHanjaConversion: WdWordDialogTabToString = "wdDialogToolsOptionsTabHangulHanjaConversion"
        Case wdDialogToolsOptionsTabFuzzy: WdWordDialogTabToString = "wdDialogToolsOptionsTabFuzzy"
        Case wdDialogToolsOptionsTabBidi: WdWordDialogTabToString = "wdDialogToolsOptionsTabBidi"
        Case wdDialogToolsOptionsTabAcetate: WdWordDialogTabToString = "wdDialogToolsOptionsTabAcetate"
        Case wdDialogToolsOptionsTabSecurity: WdWordDialogTabToString = "wdDialogToolsOptionsTabSecurity"
        Case wdDialogFilePageSetupTabMargins: WdWordDialogTabToString = "wdDialogFilePageSetupTabMargins"
        Case wdDialogFilePageSetupTabPaper: WdWordDialogTabToString = "wdDialogFilePageSetupTabPaper"
        Case wdDialogFilePageSetupTabLayout: WdWordDialogTabToString = "wdDialogFilePageSetupTabLayout"
        Case wdDialogFilePageSetupTabCharsLines: WdWordDialogTabToString = "wdDialogFilePageSetupTabCharsLines"
        Case wdDialogInsertSymbolTabSymbols: WdWordDialogTabToString = "wdDialogInsertSymbolTabSymbols"
        Case wdDialogInsertSymbolTabSpecialCharacters: WdWordDialogTabToString = "wdDialogInsertSymbolTabSpecialCharacters"
        Case wdDialogNoteOptionsTabAllFootnotes: WdWordDialogTabToString = "wdDialogNoteOptionsTabAllFootnotes"
        Case wdDialogNoteOptionsTabAllEndnotes: WdWordDialogTabToString = "wdDialogNoteOptionsTabAllEndnotes"
        Case wdDialogInsertIndexAndTablesTabIndex: WdWordDialogTabToString = "wdDialogInsertIndexAndTablesTabIndex"
        Case wdDialogInsertIndexAndTablesTabTableOfContents: WdWordDialogTabToString = "wdDialogInsertIndexAndTablesTabTableOfContents"
        Case wdDialogInsertIndexAndTablesTabTableOfFigures: WdWordDialogTabToString = "wdDialogInsertIndexAndTablesTabTableOfFigures"
        Case wdDialogInsertIndexAndTablesTabTableOfAuthorities: WdWordDialogTabToString = "wdDialogInsertIndexAndTablesTabTableOfAuthorities"
        Case wdDialogOrganizerTabStyles: WdWordDialogTabToString = "wdDialogOrganizerTabStyles"
        Case wdDialogOrganizerTabAutoText: WdWordDialogTabToString = "wdDialogOrganizerTabAutoText"
        Case wdDialogOrganizerTabCommandBars: WdWordDialogTabToString = "wdDialogOrganizerTabCommandBars"
        Case wdDialogOrganizerTabMacros: WdWordDialogTabToString = "wdDialogOrganizerTabMacros"
        Case wdDialogFormatFontTabFont: WdWordDialogTabToString = "wdDialogFormatFontTabFont"
        Case wdDialogFormatFontTabCharacterSpacing: WdWordDialogTabToString = "wdDialogFormatFontTabCharacterSpacing"
        Case wdDialogFormatFontTabAnimation: WdWordDialogTabToString = "wdDialogFormatFontTabAnimation"
        Case wdDialogFormatBordersAndShadingTabBorders: WdWordDialogTabToString = "wdDialogFormatBordersAndShadingTabBorders"
        Case wdDialogFormatBordersAndShadingTabPageBorder: WdWordDialogTabToString = "wdDialogFormatBordersAndShadingTabPageBorder"
        Case wdDialogFormatBordersAndShadingTabShading: WdWordDialogTabToString = "wdDialogFormatBordersAndShadingTabShading"
        Case wdDialogToolsEnvelopesAndLabelsTabEnvelopes: WdWordDialogTabToString = "wdDialogToolsEnvelopesAndLabelsTabEnvelopes"
        Case wdDialogToolsEnvelopesAndLabelsTabLabels: WdWordDialogTabToString = "wdDialogToolsEnvelopesAndLabelsTabLabels"
        Case wdDialogFormatParagraphTabIndentsAndSpacing: WdWordDialogTabToString = "wdDialogFormatParagraphTabIndentsAndSpacing"
        Case wdDialogFormatParagraphTabTextFlow: WdWordDialogTabToString = "wdDialogFormatParagraphTabTextFlow"
        Case wdDialogFormatParagraphTabTeisai: WdWordDialogTabToString = "wdDialogFormatParagraphTabTeisai"
        Case wdDialogFormatDrawingObjectTabColorsAndLines: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabColorsAndLines"
        Case wdDialogFormatDrawingObjectTabSize: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabSize"
        Case wdDialogFormatDrawingObjectTabPosition: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabPosition"
        Case wdDialogFormatDrawingObjectTabWrapping: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabWrapping"
        Case wdDialogFormatDrawingObjectTabPicture: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabPicture"
        Case wdDialogFormatDrawingObjectTabTextbox: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabTextbox"
        Case wdDialogFormatDrawingObjectTabWeb: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabWeb"
        Case wdDialogFormatDrawingObjectTabHR: WdWordDialogTabToString = "wdDialogFormatDrawingObjectTabHR"
        Case wdDialogToolsAutoCorrectExceptionsTabFirstLetter: WdWordDialogTabToString = "wdDialogToolsAutoCorrectExceptionsTabFirstLetter"
        Case wdDialogToolsAutoCorrectExceptionsTabInitialCaps: WdWordDialogTabToString = "wdDialogToolsAutoCorrectExceptionsTabInitialCaps"
        Case wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet: WdWordDialogTabToString = "wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet"
        Case wdDialogToolsAutoCorrectExceptionsTabIac: WdWordDialogTabToString = "wdDialogToolsAutoCorrectExceptionsTabIac"
        Case wdDialogFormatBulletsAndNumberingTabBulleted: WdWordDialogTabToString = "wdDialogFormatBulletsAndNumberingTabBulleted"
        Case wdDialogFormatBulletsAndNumberingTabNumbered: WdWordDialogTabToString = "wdDialogFormatBulletsAndNumberingTabNumbered"
        Case wdDialogFormatBulletsAndNumberingTabOutlineNumbered: WdWordDialogTabToString = "wdDialogFormatBulletsAndNumberingTabOutlineNumbered"
        Case wdDialogLetterWizardTabLetterFormat: WdWordDialogTabToString = "wdDialogLetterWizardTabLetterFormat"
        Case wdDialogLetterWizardTabRecipientInfo: WdWordDialogTabToString = "wdDialogLetterWizardTabRecipientInfo"
        Case wdDialogLetterWizardTabOtherElements: WdWordDialogTabToString = "wdDialogLetterWizardTabOtherElements"
        Case wdDialogLetterWizardTabSenderInfo: WdWordDialogTabToString = "wdDialogLetterWizardTabSenderInfo"
        Case wdDialogToolsAutoManagerTabAutoCorrect: WdWordDialogTabToString = "wdDialogToolsAutoManagerTabAutoCorrect"
        Case wdDialogToolsAutoManagerTabAutoFormatAsYouType: WdWordDialogTabToString = "wdDialogToolsAutoManagerTabAutoFormatAsYouType"
        Case wdDialogToolsAutoManagerTabAutoText: WdWordDialogTabToString = "wdDialogToolsAutoManagerTabAutoText"
        Case wdDialogToolsAutoManagerTabAutoFormat: WdWordDialogTabToString = "wdDialogToolsAutoManagerTabAutoFormat"
        Case wdDialogToolsAutoManagerTabSmartTags: WdWordDialogTabToString = "wdDialogToolsAutoManagerTabSmartTags"
        Case wdDialogTablePropertiesTabTable: WdWordDialogTabToString = "wdDialogTablePropertiesTabTable"
        Case wdDialogTablePropertiesTabRow: WdWordDialogTabToString = "wdDialogTablePropertiesTabRow"
        Case wdDialogTablePropertiesTabColumn: WdWordDialogTabToString = "wdDialogTablePropertiesTabColumn"
        Case wdDialogTablePropertiesTabCell: WdWordDialogTabToString = "wdDialogTablePropertiesTabCell"
        Case wdDialogEmailOptionsTabSignature: WdWordDialogTabToString = "wdDialogEmailOptionsTabSignature"
        Case wdDialogEmailOptionsTabStationary: WdWordDialogTabToString = "wdDialogEmailOptionsTabStationary"
        Case wdDialogEmailOptionsTabQuoting: WdWordDialogTabToString = "wdDialogEmailOptionsTabQuoting"
        Case wdDialogWebOptionsGeneral: WdWordDialogTabToString = "wdDialogWebOptionsGeneral"
        Case wdDialogWebOptionsBrowsers: WdWordDialogTabToString = "wdDialogWebOptionsBrowsers"
        Case wdDialogWebOptionsFiles: WdWordDialogTabToString = "wdDialogWebOptionsFiles"
        Case wdDialogWebOptionsPictures: WdWordDialogTabToString = "wdDialogWebOptionsPictures"
        Case wdDialogWebOptionsEncoding: WdWordDialogTabToString = "wdDialogWebOptionsEncoding"
        Case wdDialogWebOptionsFonts: WdWordDialogTabToString = "wdDialogWebOptionsFonts"
        Case wdDialogTemplates: WdWordDialogTabToString = "wdDialogTemplates"
        Case wdDialogTemplatesXMLSchema: WdWordDialogTabToString = "wdDialogTemplatesXMLSchema"
        Case wdDialogTemplatesXMLExpansionPacks: WdWordDialogTabToString = "wdDialogTemplatesXMLExpansionPacks"
        Case wdDialogTemplatesLinkedCSS: WdWordDialogTabToString = "wdDialogTemplatesLinkedCSS"
        Case wdDialogStyleManagementTabEdit: WdWordDialogTabToString = "wdDialogStyleManagementTabEdit"
        Case wdDialogStyleManagementTabRecommend: WdWordDialogTabToString = "wdDialogStyleManagementTabRecommend"
        Case wdDialogStyleManagementTabRestrict: WdWordDialogTabToString = "wdDialogStyleManagementTabRestrict"
    End Select
End Function