Attribute VB_Name = "wPbMailMergeDestination"
Function PbMailMergeDestinationFromString(value As String) As PbMailMergeDestination
    If IsNumeric(value) Then
        PbMailMergeDestinationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbSendToPrinter": PbMailMergeDestinationFromString = pbSendToPrinter
        Case "pbMergeToNewPublication": PbMailMergeDestinationFromString = pbMergeToNewPublication
        Case "pbMergeToExistingPublication": PbMailMergeDestinationFromString = pbMergeToExistingPublication
        Case "pbSendEmail": PbMailMergeDestinationFromString = pbSendEmail
    End Select
End Function

Function PbMailMergeDestinationToString(value As PbMailMergeDestination) As String
    Select Case value
        Case pbSendToPrinter: PbMailMergeDestinationToString = "pbSendToPrinter"
        Case pbMergeToNewPublication: PbMailMergeDestinationToString = "pbMergeToNewPublication"
        Case pbMergeToExistingPublication: PbMailMergeDestinationToString = "pbMergeToExistingPublication"
        Case pbSendEmail: PbMailMergeDestinationToString = "pbSendEmail"
    End Select
End Function
