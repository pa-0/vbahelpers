Attribute VB_Name = "wWdXMLValidationStatus"
Function WdXMLValidationStatusFromString(value As String) As WdXMLValidationStatus
    If IsNumeric(value) Then
        WdXMLValidationStatusFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdXMLValidationStatusOK": WdXMLValidationStatusFromString = wdXMLValidationStatusOK
        Case "wdXMLValidationStatusCustom": WdXMLValidationStatusFromString = wdXMLValidationStatusCustom
    End Select
End Function

Function WdXMLValidationStatusToString(value As WdXMLValidationStatus) As String
    Select Case value
        Case wdXMLValidationStatusOK: WdXMLValidationStatusToString = "wdXMLValidationStatusOK"
        Case wdXMLValidationStatusCustom: WdXMLValidationStatusToString = "wdXMLValidationStatusCustom"
    End Select
End Function
