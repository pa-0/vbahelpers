Attribute VB_Name = "wMsoBalloonButtonType"
Function MsoBalloonButtonTypeFromString(value As String) As MsoBalloonButtonType
    If IsNumeric(value) Then
        MsoBalloonButtonTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBalloonButtonNull": MsoBalloonButtonTypeFromString = msoBalloonButtonNull
        Case "msoBalloonButtonYesToAll": MsoBalloonButtonTypeFromString = msoBalloonButtonYesToAll
        Case "msoBalloonButtonOptions": MsoBalloonButtonTypeFromString = msoBalloonButtonOptions
        Case "msoBalloonButtonTips": MsoBalloonButtonTypeFromString = msoBalloonButtonTips
        Case "msoBalloonButtonClose": MsoBalloonButtonTypeFromString = msoBalloonButtonClose
        Case "msoBalloonButtonSnooze": MsoBalloonButtonTypeFromString = msoBalloonButtonSnooze
        Case "msoBalloonButtonSearch": MsoBalloonButtonTypeFromString = msoBalloonButtonSearch
        Case "msoBalloonButtonIgnore": MsoBalloonButtonTypeFromString = msoBalloonButtonIgnore
        Case "msoBalloonButtonAbort": MsoBalloonButtonTypeFromString = msoBalloonButtonAbort
        Case "msoBalloonButtonRetry": MsoBalloonButtonTypeFromString = msoBalloonButtonRetry
        Case "msoBalloonButtonNext": MsoBalloonButtonTypeFromString = msoBalloonButtonNext
        Case "msoBalloonButtonBack": MsoBalloonButtonTypeFromString = msoBalloonButtonBack
        Case "msoBalloonButtonNo": MsoBalloonButtonTypeFromString = msoBalloonButtonNo
        Case "msoBalloonButtonYes": MsoBalloonButtonTypeFromString = msoBalloonButtonYes
        Case "msoBalloonButtonCancel": MsoBalloonButtonTypeFromString = msoBalloonButtonCancel
        Case "msoBalloonButtonOK": MsoBalloonButtonTypeFromString = msoBalloonButtonOK
    End Select
End Function

Function MsoBalloonButtonTypeToString(value As MsoBalloonButtonType) As String
    Select Case value
        Case msoBalloonButtonNull: MsoBalloonButtonTypeToString = "msoBalloonButtonNull"
        Case msoBalloonButtonYesToAll: MsoBalloonButtonTypeToString = "msoBalloonButtonYesToAll"
        Case msoBalloonButtonOptions: MsoBalloonButtonTypeToString = "msoBalloonButtonOptions"
        Case msoBalloonButtonTips: MsoBalloonButtonTypeToString = "msoBalloonButtonTips"
        Case msoBalloonButtonClose: MsoBalloonButtonTypeToString = "msoBalloonButtonClose"
        Case msoBalloonButtonSnooze: MsoBalloonButtonTypeToString = "msoBalloonButtonSnooze"
        Case msoBalloonButtonSearch: MsoBalloonButtonTypeToString = "msoBalloonButtonSearch"
        Case msoBalloonButtonIgnore: MsoBalloonButtonTypeToString = "msoBalloonButtonIgnore"
        Case msoBalloonButtonAbort: MsoBalloonButtonTypeToString = "msoBalloonButtonAbort"
        Case msoBalloonButtonRetry: MsoBalloonButtonTypeToString = "msoBalloonButtonRetry"
        Case msoBalloonButtonNext: MsoBalloonButtonTypeToString = "msoBalloonButtonNext"
        Case msoBalloonButtonBack: MsoBalloonButtonTypeToString = "msoBalloonButtonBack"
        Case msoBalloonButtonNo: MsoBalloonButtonTypeToString = "msoBalloonButtonNo"
        Case msoBalloonButtonYes: MsoBalloonButtonTypeToString = "msoBalloonButtonYes"
        Case msoBalloonButtonCancel: MsoBalloonButtonTypeToString = "msoBalloonButtonCancel"
        Case msoBalloonButtonOK: MsoBalloonButtonTypeToString = "msoBalloonButtonOK"
    End Select
End Function
