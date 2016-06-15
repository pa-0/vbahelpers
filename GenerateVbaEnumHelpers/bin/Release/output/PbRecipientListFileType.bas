Attribute VB_Name = "wPbRecipientListFileType"
Function PbRecipientListFileTypeFromString(value As String) As PbRecipientListFileType
    If IsNumeric(value) Then
        PbRecipientListFileTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbAsMdbFile": PbRecipientListFileTypeFromString = pbAsMdbFile
        Case "pbAsCsvFile": PbRecipientListFileTypeFromString = pbAsCsvFile
    End Select
End Function

Function PbRecipientListFileTypeToString(value As PbRecipientListFileType) As String
    Select Case value
        Case pbAsMdbFile: PbRecipientListFileTypeToString = "pbAsMdbFile"
        Case pbAsCsvFile: PbRecipientListFileTypeToString = "pbAsCsvFile"
    End Select
End Function
