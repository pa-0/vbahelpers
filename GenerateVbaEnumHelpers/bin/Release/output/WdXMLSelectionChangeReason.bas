Attribute VB_Name = "wWdXMLSelectionChangeReason"
Function WdXMLSelectionChangeReasonFromString(value As String) As WdXMLSelectionChangeReason
    If IsNumeric(value) Then
        WdXMLSelectionChangeReasonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdXMLSelectionChangeReasonMove": WdXMLSelectionChangeReasonFromString = wdXMLSelectionChangeReasonMove
        Case "wdXMLSelectionChangeReasonInsert": WdXMLSelectionChangeReasonFromString = wdXMLSelectionChangeReasonInsert
        Case "wdXMLSelectionChangeReasonDelete": WdXMLSelectionChangeReasonFromString = wdXMLSelectionChangeReasonDelete
    End Select
End Function

Function WdXMLSelectionChangeReasonToString(value As WdXMLSelectionChangeReason) As String
    Select Case value
        Case wdXMLSelectionChangeReasonMove: WdXMLSelectionChangeReasonToString = "wdXMLSelectionChangeReasonMove"
        Case wdXMLSelectionChangeReasonInsert: WdXMLSelectionChangeReasonToString = "wdXMLSelectionChangeReasonInsert"
        Case wdXMLSelectionChangeReasonDelete: WdXMLSelectionChangeReasonToString = "wdXMLSelectionChangeReasonDelete"
    End Select
End Function
