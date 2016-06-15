Attribute VB_Name = "wMsoAlertButtonType"
Function MsoAlertButtonTypeFromString(value As String) As MsoAlertButtonType
    If IsNumeric(value) Then
        MsoAlertButtonTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlertButtonOK": MsoAlertButtonTypeFromString = msoAlertButtonOK
        Case "msoAlertButtonOKCancel": MsoAlertButtonTypeFromString = msoAlertButtonOKCancel
        Case "msoAlertButtonAbortRetryIgnore": MsoAlertButtonTypeFromString = msoAlertButtonAbortRetryIgnore
        Case "msoAlertButtonYesNoCancel": MsoAlertButtonTypeFromString = msoAlertButtonYesNoCancel
        Case "msoAlertButtonYesNo": MsoAlertButtonTypeFromString = msoAlertButtonYesNo
        Case "msoAlertButtonRetryCancel": MsoAlertButtonTypeFromString = msoAlertButtonRetryCancel
        Case "msoAlertButtonYesAllNoCancel": MsoAlertButtonTypeFromString = msoAlertButtonYesAllNoCancel
    End Select
End Function

Function MsoAlertButtonTypeToString(value As MsoAlertButtonType) As String
    Select Case value
        Case msoAlertButtonOK: MsoAlertButtonTypeToString = "msoAlertButtonOK"
        Case msoAlertButtonOKCancel: MsoAlertButtonTypeToString = "msoAlertButtonOKCancel"
        Case msoAlertButtonAbortRetryIgnore: MsoAlertButtonTypeToString = "msoAlertButtonAbortRetryIgnore"
        Case msoAlertButtonYesNoCancel: MsoAlertButtonTypeToString = "msoAlertButtonYesNoCancel"
        Case msoAlertButtonYesNo: MsoAlertButtonTypeToString = "msoAlertButtonYesNo"
        Case msoAlertButtonRetryCancel: MsoAlertButtonTypeToString = "msoAlertButtonRetryCancel"
        Case msoAlertButtonYesAllNoCancel: MsoAlertButtonTypeToString = "msoAlertButtonYesAllNoCancel"
    End Select
End Function
