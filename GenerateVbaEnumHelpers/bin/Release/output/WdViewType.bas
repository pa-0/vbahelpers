Attribute VB_Name = "wWdViewType"
Function WdViewTypeFromString(value As String) As WdViewType
    If IsNumeric(value) Then
        WdViewTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNormalView": WdViewTypeFromString = wdNormalView
        Case "wdOutlineView": WdViewTypeFromString = wdOutlineView
        Case "wdPrintView": WdViewTypeFromString = wdPrintView
        Case "wdPrintPreview": WdViewTypeFromString = wdPrintPreview
        Case "wdMasterView": WdViewTypeFromString = wdMasterView
        Case "wdWebView": WdViewTypeFromString = wdWebView
        Case "wdReadingView": WdViewTypeFromString = wdReadingView
        Case "wdConflictView": WdViewTypeFromString = wdConflictView
    End Select
End Function

Function WdViewTypeToString(value As WdViewType) As String
    Select Case value
        Case wdNormalView: WdViewTypeToString = "wdNormalView"
        Case wdOutlineView: WdViewTypeToString = "wdOutlineView"
        Case wdPrintView: WdViewTypeToString = "wdPrintView"
        Case wdPrintPreview: WdViewTypeToString = "wdPrintPreview"
        Case wdMasterView: WdViewTypeToString = "wdMasterView"
        Case wdWebView: WdViewTypeToString = "wdWebView"
        Case wdReadingView: WdViewTypeToString = "wdReadingView"
        Case wdConflictView: WdViewTypeToString = "wdConflictView"
    End Select
End Function
