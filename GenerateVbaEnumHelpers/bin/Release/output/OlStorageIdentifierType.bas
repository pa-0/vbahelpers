Attribute VB_Name = "wOlStorageIdentifierType"
Function OlStorageIdentifierTypeFromString(value As String) As OlStorageIdentifierType
    If IsNumeric(value) Then
        OlStorageIdentifierTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olIdentifyBySubject": OlStorageIdentifierTypeFromString = olIdentifyBySubject
        Case "olIdentifyByEntryID": OlStorageIdentifierTypeFromString = olIdentifyByEntryID
        Case "olIdentifyByMessageClass": OlStorageIdentifierTypeFromString = olIdentifyByMessageClass
    End Select
End Function

Function OlStorageIdentifierTypeToString(value As OlStorageIdentifierType) As String
    Select Case value
        Case olIdentifyBySubject: OlStorageIdentifierTypeToString = "olIdentifyBySubject"
        Case olIdentifyByEntryID: OlStorageIdentifierTypeToString = "olIdentifyByEntryID"
        Case olIdentifyByMessageClass: OlStorageIdentifierTypeToString = "olIdentifyByMessageClass"
    End Select
End Function
