Attribute VB_Name = "wPbLinkedFileStatus"
Function PbLinkedFileStatusFromString(value As String) As PbLinkedFileStatus
    If IsNumeric(value) Then
        PbLinkedFileStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbLinkedFileOK": PbLinkedFileStatusFromString = pbLinkedFileOK
        Case "pbLinkedFileMissing": PbLinkedFileStatusFromString = pbLinkedFileMissing
        Case "pbLinkedFileModified": PbLinkedFileStatusFromString = pbLinkedFileModified
    End Select
End Function

Function PbLinkedFileStatusToString(value As PbLinkedFileStatus) As String
    Select Case value
        Case pbLinkedFileOK: PbLinkedFileStatusToString = "pbLinkedFileOK"
        Case pbLinkedFileMissing: PbLinkedFileStatusToString = "pbLinkedFileMissing"
        Case pbLinkedFileModified: PbLinkedFileStatusToString = "pbLinkedFileModified"
    End Select
End Function
