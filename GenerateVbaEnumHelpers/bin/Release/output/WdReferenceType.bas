Attribute VB_Name = "wWdReferenceType"
Function WdReferenceTypeFromString(value As String) As WdReferenceType
    If IsNumeric(value) Then
        WdReferenceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRefTypeNumberedItem": WdReferenceTypeFromString = wdRefTypeNumberedItem
        Case "wdRefTypeHeading": WdReferenceTypeFromString = wdRefTypeHeading
        Case "wdRefTypeBookmark": WdReferenceTypeFromString = wdRefTypeBookmark
        Case "wdRefTypeFootnote": WdReferenceTypeFromString = wdRefTypeFootnote
        Case "wdRefTypeEndnote": WdReferenceTypeFromString = wdRefTypeEndnote
    End Select
End Function

Function WdReferenceTypeToString(value As WdReferenceType) As String
    Select Case value
        Case wdRefTypeNumberedItem: WdReferenceTypeToString = "wdRefTypeNumberedItem"
        Case wdRefTypeHeading: WdReferenceTypeToString = "wdRefTypeHeading"
        Case wdRefTypeBookmark: WdReferenceTypeToString = "wdRefTypeBookmark"
        Case wdRefTypeFootnote: WdReferenceTypeToString = "wdRefTypeFootnote"
        Case wdRefTypeEndnote: WdReferenceTypeToString = "wdRefTypeEndnote"
    End Select
End Function
