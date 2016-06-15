Attribute VB_Name = "wWdMailMergeMainDocType"
Function WdMailMergeMainDocTypeFromString(value As String) As WdMailMergeMainDocType
    If IsNumeric(value) Then
        WdMailMergeMainDocTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFormLetters": WdMailMergeMainDocTypeFromString = wdFormLetters
        Case "wdMailingLabels": WdMailMergeMainDocTypeFromString = wdMailingLabels
        Case "wdEnvelopes": WdMailMergeMainDocTypeFromString = wdEnvelopes
        Case "wdDirectory": WdMailMergeMainDocTypeFromString = wdDirectory
        Case "wdCatalog": WdMailMergeMainDocTypeFromString = wdCatalog
        Case "wdEMail": WdMailMergeMainDocTypeFromString = wdEMail
        Case "wdFax": WdMailMergeMainDocTypeFromString = wdFax
        Case "wdNotAMergeDocument": WdMailMergeMainDocTypeFromString = wdNotAMergeDocument
    End Select
End Function

Function WdMailMergeMainDocTypeToString(value As WdMailMergeMainDocType) As String
    Select Case value
        Case wdFormLetters: WdMailMergeMainDocTypeToString = "wdFormLetters"
        Case wdMailingLabels: WdMailMergeMainDocTypeToString = "wdMailingLabels"
        Case wdEnvelopes: WdMailMergeMainDocTypeToString = "wdEnvelopes"
        Case wdDirectory: WdMailMergeMainDocTypeToString = "wdDirectory"
        Case wdCatalog: WdMailMergeMainDocTypeToString = "wdCatalog"
        Case wdEMail: WdMailMergeMainDocTypeToString = "wdEMail"
        Case wdFax: WdMailMergeMainDocTypeToString = "wdFax"
        Case wdNotAMergeDocument: WdMailMergeMainDocTypeToString = "wdNotAMergeDocument"
    End Select
End Function
