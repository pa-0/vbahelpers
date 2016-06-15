Attribute VB_Name = "wMsoDocInspectorStatus"
Function MsoDocInspectorStatusFromString(value As String) As MsoDocInspectorStatus
    If IsNumeric(value) Then
        MsoDocInspectorStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoDocInspectorStatusDocOk": MsoDocInspectorStatusFromString = msoDocInspectorStatusDocOk
        Case "msoDocInspectorStatusIssueFound": MsoDocInspectorStatusFromString = msoDocInspectorStatusIssueFound
        Case "msoDocInspectorStatusError": MsoDocInspectorStatusFromString = msoDocInspectorStatusError
    End Select
End Function

Function MsoDocInspectorStatusToString(value As MsoDocInspectorStatus) As String
    Select Case value
        Case msoDocInspectorStatusDocOk: MsoDocInspectorStatusToString = "msoDocInspectorStatusDocOk"
        Case msoDocInspectorStatusIssueFound: MsoDocInspectorStatusToString = "msoDocInspectorStatusIssueFound"
        Case msoDocInspectorStatusError: MsoDocInspectorStatusToString = "msoDocInspectorStatusError"
    End Select
End Function
