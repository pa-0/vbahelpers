Attribute VB_Name = "wWdExportOptimizeFor"
Function WdExportOptimizeForFromString(value As String) As WdExportOptimizeFor
    If IsNumeric(value) Then
        WdExportOptimizeForFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdExportOptimizeForPrint": WdExportOptimizeForFromString = wdExportOptimizeForPrint
        Case "wdExportOptimizeForOnScreen": WdExportOptimizeForFromString = wdExportOptimizeForOnScreen
    End Select
End Function

Function WdExportOptimizeForToString(value As WdExportOptimizeFor) As String
    Select Case value
        Case wdExportOptimizeForPrint: WdExportOptimizeForToString = "wdExportOptimizeForPrint"
        Case wdExportOptimizeForOnScreen: WdExportOptimizeForToString = "wdExportOptimizeForOnScreen"
    End Select
End Function
