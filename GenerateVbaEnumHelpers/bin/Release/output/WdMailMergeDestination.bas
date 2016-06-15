Attribute VB_Name = "wWdMailMergeDestination"
Function WdMailMergeDestinationFromString(value As String) As WdMailMergeDestination
    If IsNumeric(value) Then
        WdMailMergeDestinationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSendToNewDocument": WdMailMergeDestinationFromString = wdSendToNewDocument
        Case "wdSendToPrinter": WdMailMergeDestinationFromString = wdSendToPrinter
        Case "wdSendToEmail": WdMailMergeDestinationFromString = wdSendToEmail
        Case "wdSendToFax": WdMailMergeDestinationFromString = wdSendToFax
    End Select
End Function

Function WdMailMergeDestinationToString(value As WdMailMergeDestination) As String
    Select Case value
        Case wdSendToNewDocument: WdMailMergeDestinationToString = "wdSendToNewDocument"
        Case wdSendToPrinter: WdMailMergeDestinationToString = "wdSendToPrinter"
        Case wdSendToEmail: WdMailMergeDestinationToString = "wdSendToEmail"
        Case wdSendToFax: WdMailMergeDestinationToString = "wdSendToFax"
    End Select
End Function
