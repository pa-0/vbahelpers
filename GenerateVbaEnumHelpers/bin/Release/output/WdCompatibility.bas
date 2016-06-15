Attribute VB_Name = "wWdCompatibility"
Function WdCompatibilityFromString(value As String) As WdCompatibility
    If IsNumeric(value) Then
        WdCompatibilityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNoTabHangIndent": WdCompatibilityFromString = wdNoTabHangIndent
        Case "wdNoSpaceRaiseLower": WdCompatibilityFromString = wdNoSpaceRaiseLower
        Case "wdPrintColBlack": WdCompatibilityFromString = wdPrintColBlack
        Case "wdWrapTrailSpaces": WdCompatibilityFromString = wdWrapTrailSpaces
        Case "wdNoColumnBalance": WdCompatibilityFromString = wdNoColumnBalance
        Case "wdConvMailMergeEsc": WdCompatibilityFromString = wdConvMailMergeEsc
        Case "wdSuppressSpBfAfterPgBrk": WdCompatibilityFromString = wdSuppressSpBfAfterPgBrk
        Case "wdSuppressTopSpacing": WdCompatibilityFromString = wdSuppressTopSpacing
        Case "wdOrigWordTableRules": WdCompatibilityFromString = wdOrigWordTableRules
        Case "wdTransparentMetafiles": WdCompatibilityFromString = wdTransparentMetafiles
        Case "wdShowBreaksInFrames": WdCompatibilityFromString = wdShowBreaksInFrames
        Case "wdSwapBordersFacingPages": WdCompatibilityFromString = wdSwapBordersFacingPages
        Case "wdLeaveBackslashAlone": WdCompatibilityFromString = wdLeaveBackslashAlone
        Case "wdExpandShiftReturn": WdCompatibilityFromString = wdExpandShiftReturn
        Case "wdDontULTrailSpace": WdCompatibilityFromString = wdDontULTrailSpace
        Case "wdDontBalanceSingleByteDoubleByteWidth": WdCompatibilityFromString = wdDontBalanceSingleByteDoubleByteWidth
        Case "wdSuppressTopSpacingMac5": WdCompatibilityFromString = wdSuppressTopSpacingMac5
        Case "wdSpacingInWholePoints": WdCompatibilityFromString = wdSpacingInWholePoints
        Case "wdPrintBodyTextBeforeHeader": WdCompatibilityFromString = wdPrintBodyTextBeforeHeader
        Case "wdNoLeading": WdCompatibilityFromString = wdNoLeading
        Case "wdNoSpaceForUL": WdCompatibilityFromString = wdNoSpaceForUL
        Case "wdMWSmallCaps": WdCompatibilityFromString = wdMWSmallCaps
        Case "wdNoExtraLineSpacing": WdCompatibilityFromString = wdNoExtraLineSpacing
        Case "wdTruncateFontHeight": WdCompatibilityFromString = wdTruncateFontHeight
        Case "wdSubFontBySize": WdCompatibilityFromString = wdSubFontBySize
        Case "wdUsePrinterMetrics": WdCompatibilityFromString = wdUsePrinterMetrics
        Case "wdWW6BorderRules": WdCompatibilityFromString = wdWW6BorderRules
        Case "wdExactOnTop": WdCompatibilityFromString = wdExactOnTop
        Case "wdSuppressBottomSpacing": WdCompatibilityFromString = wdSuppressBottomSpacing
        Case "wdWPSpaceWidth": WdCompatibilityFromString = wdWPSpaceWidth
        Case "wdWPJustification": WdCompatibilityFromString = wdWPJustification
        Case "wdLineWrapLikeWord6": WdCompatibilityFromString = wdLineWrapLikeWord6
        Case "wdShapeLayoutLikeWW8": WdCompatibilityFromString = wdShapeLayoutLikeWW8
        Case "wdFootnoteLayoutLikeWW8": WdCompatibilityFromString = wdFootnoteLayoutLikeWW8
        Case "wdDontUseHTMLParagraphAutoSpacing": WdCompatibilityFromString = wdDontUseHTMLParagraphAutoSpacing
        Case "wdDontAdjustLineHeightInTable": WdCompatibilityFromString = wdDontAdjustLineHeightInTable
        Case "wdForgetLastTabAlignment": WdCompatibilityFromString = wdForgetLastTabAlignment
        Case "wdAutospaceLikeWW7": WdCompatibilityFromString = wdAutospaceLikeWW7
        Case "wdAlignTablesRowByRow": WdCompatibilityFromString = wdAlignTablesRowByRow
        Case "wdLayoutRawTableWidth": WdCompatibilityFromString = wdLayoutRawTableWidth
        Case "wdLayoutTableRowsApart": WdCompatibilityFromString = wdLayoutTableRowsApart
        Case "wdUseWord97LineBreakingRules": WdCompatibilityFromString = wdUseWord97LineBreakingRules
        Case "wdDontBreakWrappedTables": WdCompatibilityFromString = wdDontBreakWrappedTables
        Case "wdDontSnapTextToGridInTableWithObjects": WdCompatibilityFromString = wdDontSnapTextToGridInTableWithObjects
        Case "wdSelectFieldWithFirstOrLastCharacter": WdCompatibilityFromString = wdSelectFieldWithFirstOrLastCharacter
        Case "wdApplyBreakingRules": WdCompatibilityFromString = wdApplyBreakingRules
        Case "wdDontWrapTextWithPunctuation": WdCompatibilityFromString = wdDontWrapTextWithPunctuation
        Case "wdDontUseAsianBreakRulesInGrid": WdCompatibilityFromString = wdDontUseAsianBreakRulesInGrid
        Case "wdUseWord2002TableStyleRules": WdCompatibilityFromString = wdUseWord2002TableStyleRules
        Case "wdGrowAutofit": WdCompatibilityFromString = wdGrowAutofit
        Case "wdUseNormalStyleForList": WdCompatibilityFromString = wdUseNormalStyleForList
        Case "wdDontUseIndentAsNumberingTabStop": WdCompatibilityFromString = wdDontUseIndentAsNumberingTabStop
        Case "wdFELineBreak11": WdCompatibilityFromString = wdFELineBreak11
        Case "wdAllowSpaceOfSameStyleInTable": WdCompatibilityFromString = wdAllowSpaceOfSameStyleInTable
        Case "wdWW11IndentRules": WdCompatibilityFromString = wdWW11IndentRules
        Case "wdDontAutofitConstrainedTables": WdCompatibilityFromString = wdDontAutofitConstrainedTables
        Case "wdAutofitLikeWW11": WdCompatibilityFromString = wdAutofitLikeWW11
        Case "wdUnderlineTabInNumList": WdCompatibilityFromString = wdUnderlineTabInNumList
        Case "wdHangulWidthLikeWW11": WdCompatibilityFromString = wdHangulWidthLikeWW11
        Case "wdSplitPgBreakAndParaMark": WdCompatibilityFromString = wdSplitPgBreakAndParaMark
        Case "wdDontVertAlignCellWithShape": WdCompatibilityFromString = wdDontVertAlignCellWithShape
        Case "wdDontBreakConstrainedForcedTables": WdCompatibilityFromString = wdDontBreakConstrainedForcedTables
        Case "wdDontVertAlignInTextbox": WdCompatibilityFromString = wdDontVertAlignInTextbox
        Case "wdWord11KerningPairs": WdCompatibilityFromString = wdWord11KerningPairs
        Case "wdCachedColBalance": WdCompatibilityFromString = wdCachedColBalance
        Case "wdDisableOTKerning": WdCompatibilityFromString = wdDisableOTKerning
        Case "wdFlipMirrorIndents": WdCompatibilityFromString = wdFlipMirrorIndents
        Case "wdDontOverrideTableStyleFontSzAndJustification": WdCompatibilityFromString = wdDontOverrideTableStyleFontSzAndJustification
    End Select
End Function

Function WdCompatibilityToString(value As WdCompatibility) As String
    Select Case value
        Case wdNoTabHangIndent: WdCompatibilityToString = "wdNoTabHangIndent"
        Case wdNoSpaceRaiseLower: WdCompatibilityToString = "wdNoSpaceRaiseLower"
        Case wdPrintColBlack: WdCompatibilityToString = "wdPrintColBlack"
        Case wdWrapTrailSpaces: WdCompatibilityToString = "wdWrapTrailSpaces"
        Case wdNoColumnBalance: WdCompatibilityToString = "wdNoColumnBalance"
        Case wdConvMailMergeEsc: WdCompatibilityToString = "wdConvMailMergeEsc"
        Case wdSuppressSpBfAfterPgBrk: WdCompatibilityToString = "wdSuppressSpBfAfterPgBrk"
        Case wdSuppressTopSpacing: WdCompatibilityToString = "wdSuppressTopSpacing"
        Case wdOrigWordTableRules: WdCompatibilityToString = "wdOrigWordTableRules"
        Case wdTransparentMetafiles: WdCompatibilityToString = "wdTransparentMetafiles"
        Case wdShowBreaksInFrames: WdCompatibilityToString = "wdShowBreaksInFrames"
        Case wdSwapBordersFacingPages: WdCompatibilityToString = "wdSwapBordersFacingPages"
        Case wdLeaveBackslashAlone: WdCompatibilityToString = "wdLeaveBackslashAlone"
        Case wdExpandShiftReturn: WdCompatibilityToString = "wdExpandShiftReturn"
        Case wdDontULTrailSpace: WdCompatibilityToString = "wdDontULTrailSpace"
        Case wdDontBalanceSingleByteDoubleByteWidth: WdCompatibilityToString = "wdDontBalanceSingleByteDoubleByteWidth"
        Case wdSuppressTopSpacingMac5: WdCompatibilityToString = "wdSuppressTopSpacingMac5"
        Case wdSpacingInWholePoints: WdCompatibilityToString = "wdSpacingInWholePoints"
        Case wdPrintBodyTextBeforeHeader: WdCompatibilityToString = "wdPrintBodyTextBeforeHeader"
        Case wdNoLeading: WdCompatibilityToString = "wdNoLeading"
        Case wdNoSpaceForUL: WdCompatibilityToString = "wdNoSpaceForUL"
        Case wdMWSmallCaps: WdCompatibilityToString = "wdMWSmallCaps"
        Case wdNoExtraLineSpacing: WdCompatibilityToString = "wdNoExtraLineSpacing"
        Case wdTruncateFontHeight: WdCompatibilityToString = "wdTruncateFontHeight"
        Case wdSubFontBySize: WdCompatibilityToString = "wdSubFontBySize"
        Case wdUsePrinterMetrics: WdCompatibilityToString = "wdUsePrinterMetrics"
        Case wdWW6BorderRules: WdCompatibilityToString = "wdWW6BorderRules"
        Case wdExactOnTop: WdCompatibilityToString = "wdExactOnTop"
        Case wdSuppressBottomSpacing: WdCompatibilityToString = "wdSuppressBottomSpacing"
        Case wdWPSpaceWidth: WdCompatibilityToString = "wdWPSpaceWidth"
        Case wdWPJustification: WdCompatibilityToString = "wdWPJustification"
        Case wdLineWrapLikeWord6: WdCompatibilityToString = "wdLineWrapLikeWord6"
        Case wdShapeLayoutLikeWW8: WdCompatibilityToString = "wdShapeLayoutLikeWW8"
        Case wdFootnoteLayoutLikeWW8: WdCompatibilityToString = "wdFootnoteLayoutLikeWW8"
        Case wdDontUseHTMLParagraphAutoSpacing: WdCompatibilityToString = "wdDontUseHTMLParagraphAutoSpacing"
        Case wdDontAdjustLineHeightInTable: WdCompatibilityToString = "wdDontAdjustLineHeightInTable"
        Case wdForgetLastTabAlignment: WdCompatibilityToString = "wdForgetLastTabAlignment"
        Case wdAutospaceLikeWW7: WdCompatibilityToString = "wdAutospaceLikeWW7"
        Case wdAlignTablesRowByRow: WdCompatibilityToString = "wdAlignTablesRowByRow"
        Case wdLayoutRawTableWidth: WdCompatibilityToString = "wdLayoutRawTableWidth"
        Case wdLayoutTableRowsApart: WdCompatibilityToString = "wdLayoutTableRowsApart"
        Case wdUseWord97LineBreakingRules: WdCompatibilityToString = "wdUseWord97LineBreakingRules"
        Case wdDontBreakWrappedTables: WdCompatibilityToString = "wdDontBreakWrappedTables"
        Case wdDontSnapTextToGridInTableWithObjects: WdCompatibilityToString = "wdDontSnapTextToGridInTableWithObjects"
        Case wdSelectFieldWithFirstOrLastCharacter: WdCompatibilityToString = "wdSelectFieldWithFirstOrLastCharacter"
        Case wdApplyBreakingRules: WdCompatibilityToString = "wdApplyBreakingRules"
        Case wdDontWrapTextWithPunctuation: WdCompatibilityToString = "wdDontWrapTextWithPunctuation"
        Case wdDontUseAsianBreakRulesInGrid: WdCompatibilityToString = "wdDontUseAsianBreakRulesInGrid"
        Case wdUseWord2002TableStyleRules: WdCompatibilityToString = "wdUseWord2002TableStyleRules"
        Case wdGrowAutofit: WdCompatibilityToString = "wdGrowAutofit"
        Case wdUseNormalStyleForList: WdCompatibilityToString = "wdUseNormalStyleForList"
        Case wdDontUseIndentAsNumberingTabStop: WdCompatibilityToString = "wdDontUseIndentAsNumberingTabStop"
        Case wdFELineBreak11: WdCompatibilityToString = "wdFELineBreak11"
        Case wdAllowSpaceOfSameStyleInTable: WdCompatibilityToString = "wdAllowSpaceOfSameStyleInTable"
        Case wdWW11IndentRules: WdCompatibilityToString = "wdWW11IndentRules"
        Case wdDontAutofitConstrainedTables: WdCompatibilityToString = "wdDontAutofitConstrainedTables"
        Case wdAutofitLikeWW11: WdCompatibilityToString = "wdAutofitLikeWW11"
        Case wdUnderlineTabInNumList: WdCompatibilityToString = "wdUnderlineTabInNumList"
        Case wdHangulWidthLikeWW11: WdCompatibilityToString = "wdHangulWidthLikeWW11"
        Case wdSplitPgBreakAndParaMark: WdCompatibilityToString = "wdSplitPgBreakAndParaMark"
        Case wdDontVertAlignCellWithShape: WdCompatibilityToString = "wdDontVertAlignCellWithShape"
        Case wdDontBreakConstrainedForcedTables: WdCompatibilityToString = "wdDontBreakConstrainedForcedTables"
        Case wdDontVertAlignInTextbox: WdCompatibilityToString = "wdDontVertAlignInTextbox"
        Case wdWord11KerningPairs: WdCompatibilityToString = "wdWord11KerningPairs"
        Case wdCachedColBalance: WdCompatibilityToString = "wdCachedColBalance"
        Case wdDisableOTKerning: WdCompatibilityToString = "wdDisableOTKerning"
        Case wdFlipMirrorIndents: WdCompatibilityToString = "wdFlipMirrorIndents"
        Case wdDontOverrideTableStyleFontSzAndJustification: WdCompatibilityToString = "wdDontOverrideTableStyleFontSzAndJustification"
    End Select
End Function