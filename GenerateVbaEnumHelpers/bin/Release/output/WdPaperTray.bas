Attribute VB_Name = "wWdPaperTray"
Function WdPaperTrayFromString(value As String) As WdPaperTray
    If IsNumeric(value) Then
        WdPaperTrayFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPrinterDefaultBin": WdPaperTrayFromString = wdPrinterDefaultBin
        Case "wdPrinterUpperBin": WdPaperTrayFromString = wdPrinterUpperBin
        Case "wdPrinterOnlyBin": WdPaperTrayFromString = wdPrinterOnlyBin
        Case "wdPrinterLowerBin": WdPaperTrayFromString = wdPrinterLowerBin
        Case "wdPrinterMiddleBin": WdPaperTrayFromString = wdPrinterMiddleBin
        Case "wdPrinterManualFeed": WdPaperTrayFromString = wdPrinterManualFeed
        Case "wdPrinterEnvelopeFeed": WdPaperTrayFromString = wdPrinterEnvelopeFeed
        Case "wdPrinterManualEnvelopeFeed": WdPaperTrayFromString = wdPrinterManualEnvelopeFeed
        Case "wdPrinterAutomaticSheetFeed": WdPaperTrayFromString = wdPrinterAutomaticSheetFeed
        Case "wdPrinterTractorFeed": WdPaperTrayFromString = wdPrinterTractorFeed
        Case "wdPrinterSmallFormatBin": WdPaperTrayFromString = wdPrinterSmallFormatBin
        Case "wdPrinterLargeFormatBin": WdPaperTrayFromString = wdPrinterLargeFormatBin
        Case "wdPrinterLargeCapacityBin": WdPaperTrayFromString = wdPrinterLargeCapacityBin
        Case "wdPrinterPaperCassette": WdPaperTrayFromString = wdPrinterPaperCassette
        Case "wdPrinterFormSource": WdPaperTrayFromString = wdPrinterFormSource
    End Select
End Function

Function WdPaperTrayToString(value As WdPaperTray) As String
    Select Case value
        Case wdPrinterDefaultBin: WdPaperTrayToString = "wdPrinterDefaultBin"
        Case wdPrinterUpperBin: WdPaperTrayToString = "wdPrinterUpperBin"
        Case wdPrinterOnlyBin: WdPaperTrayToString = "wdPrinterOnlyBin"
        Case wdPrinterLowerBin: WdPaperTrayToString = "wdPrinterLowerBin"
        Case wdPrinterMiddleBin: WdPaperTrayToString = "wdPrinterMiddleBin"
        Case wdPrinterManualFeed: WdPaperTrayToString = "wdPrinterManualFeed"
        Case wdPrinterEnvelopeFeed: WdPaperTrayToString = "wdPrinterEnvelopeFeed"
        Case wdPrinterManualEnvelopeFeed: WdPaperTrayToString = "wdPrinterManualEnvelopeFeed"
        Case wdPrinterAutomaticSheetFeed: WdPaperTrayToString = "wdPrinterAutomaticSheetFeed"
        Case wdPrinterTractorFeed: WdPaperTrayToString = "wdPrinterTractorFeed"
        Case wdPrinterSmallFormatBin: WdPaperTrayToString = "wdPrinterSmallFormatBin"
        Case wdPrinterLargeFormatBin: WdPaperTrayToString = "wdPrinterLargeFormatBin"
        Case wdPrinterLargeCapacityBin: WdPaperTrayToString = "wdPrinterLargeCapacityBin"
        Case wdPrinterPaperCassette: WdPaperTrayToString = "wdPrinterPaperCassette"
        Case wdPrinterFormSource: WdPaperTrayToString = "wdPrinterFormSource"
    End Select
End Function
