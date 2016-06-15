Attribute VB_Name = "wWdPrintOutItem"
Function WdPrintOutItemFromString(value As String) As WdPrintOutItem
    If IsNumeric(value) Then
        WdPrintOutItemFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPrintDocumentContent": WdPrintOutItemFromString = wdPrintDocumentContent
        Case "wdPrintProperties": WdPrintOutItemFromString = wdPrintProperties
        Case "wdPrintComments": WdPrintOutItemFromString = wdPrintComments
        Case "wdPrintMarkup": WdPrintOutItemFromString = wdPrintMarkup
        Case "wdPrintStyles": WdPrintOutItemFromString = wdPrintStyles
        Case "wdPrintAutoTextEntries": WdPrintOutItemFromString = wdPrintAutoTextEntries
        Case "wdPrintKeyAssignments": WdPrintOutItemFromString = wdPrintKeyAssignments
        Case "wdPrintEnvelope": WdPrintOutItemFromString = wdPrintEnvelope
        Case "wdPrintDocumentWithMarkup": WdPrintOutItemFromString = wdPrintDocumentWithMarkup
    End Select
End Function

Function WdPrintOutItemToString(value As WdPrintOutItem) As String
    Select Case value
        Case wdPrintDocumentContent: WdPrintOutItemToString = "wdPrintDocumentContent"
        Case wdPrintProperties: WdPrintOutItemToString = "wdPrintProperties"
        Case wdPrintComments: WdPrintOutItemToString = "wdPrintComments"
        Case wdPrintMarkup: WdPrintOutItemToString = "wdPrintMarkup"
        Case wdPrintStyles: WdPrintOutItemToString = "wdPrintStyles"
        Case wdPrintAutoTextEntries: WdPrintOutItemToString = "wdPrintAutoTextEntries"
        Case wdPrintKeyAssignments: WdPrintOutItemToString = "wdPrintKeyAssignments"
        Case wdPrintEnvelope: WdPrintOutItemToString = "wdPrintEnvelope"
        Case wdPrintDocumentWithMarkup: WdPrintOutItemToString = "wdPrintDocumentWithMarkup"
    End Select
End Function
