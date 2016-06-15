Attribute VB_Name = "wWdWordDialogTabHID"
Function WdWordDialogTabHIDFromString(value As String) As WdWordDialogTabHID
    If IsNumeric(value) Then
        WdWordDialogTabHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDialogFilePageSetupTabPaperSize": WdWordDialogTabHIDFromString = wdDialogFilePageSetupTabPaperSize
        Case "wdDialogFilePageSetupTabPaperSource": WdWordDialogTabHIDFromString = wdDialogFilePageSetupTabPaperSource
    End Select
End Function

Function WdWordDialogTabHIDToString(value As WdWordDialogTabHID) As String
    Select Case value
        Case wdDialogFilePageSetupTabPaperSize: WdWordDialogTabHIDToString = "wdDialogFilePageSetupTabPaperSize"
        Case wdDialogFilePageSetupTabPaperSource: WdWordDialogTabHIDToString = "wdDialogFilePageSetupTabPaperSource"
    End Select
End Function
