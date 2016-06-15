Attribute VB_Name = "wXlLinkStatus"
Function XlLinkStatusFromString(value As String) As XlLinkStatus
    If IsNumeric(value) Then
        XlLinkStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlLinkStatusOK": XlLinkStatusFromString = xlLinkStatusOK
        Case "xlLinkStatusMissingFile": XlLinkStatusFromString = xlLinkStatusMissingFile
        Case "xlLinkStatusMissingSheet": XlLinkStatusFromString = xlLinkStatusMissingSheet
        Case "xlLinkStatusOld": XlLinkStatusFromString = xlLinkStatusOld
        Case "xlLinkStatusSourceNotCalculated": XlLinkStatusFromString = xlLinkStatusSourceNotCalculated
        Case "xlLinkStatusIndeterminate": XlLinkStatusFromString = xlLinkStatusIndeterminate
        Case "xlLinkStatusNotStarted": XlLinkStatusFromString = xlLinkStatusNotStarted
        Case "xlLinkStatusInvalidName": XlLinkStatusFromString = xlLinkStatusInvalidName
        Case "xlLinkStatusSourceNotOpen": XlLinkStatusFromString = xlLinkStatusSourceNotOpen
        Case "xlLinkStatusSourceOpen": XlLinkStatusFromString = xlLinkStatusSourceOpen
        Case "xlLinkStatusCopiedValues": XlLinkStatusFromString = xlLinkStatusCopiedValues
    End Select
End Function

Function XlLinkStatusToString(value As XlLinkStatus) As String
    Select Case value
        Case xlLinkStatusOK: XlLinkStatusToString = "xlLinkStatusOK"
        Case xlLinkStatusMissingFile: XlLinkStatusToString = "xlLinkStatusMissingFile"
        Case xlLinkStatusMissingSheet: XlLinkStatusToString = "xlLinkStatusMissingSheet"
        Case xlLinkStatusOld: XlLinkStatusToString = "xlLinkStatusOld"
        Case xlLinkStatusSourceNotCalculated: XlLinkStatusToString = "xlLinkStatusSourceNotCalculated"
        Case xlLinkStatusIndeterminate: XlLinkStatusToString = "xlLinkStatusIndeterminate"
        Case xlLinkStatusNotStarted: XlLinkStatusToString = "xlLinkStatusNotStarted"
        Case xlLinkStatusInvalidName: XlLinkStatusToString = "xlLinkStatusInvalidName"
        Case xlLinkStatusSourceNotOpen: XlLinkStatusToString = "xlLinkStatusSourceNotOpen"
        Case xlLinkStatusSourceOpen: XlLinkStatusToString = "xlLinkStatusSourceOpen"
        Case xlLinkStatusCopiedValues: XlLinkStatusToString = "xlLinkStatusCopiedValues"
    End Select
End Function
