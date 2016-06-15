Attribute VB_Name = "wMsoButtonSetType"
Function MsoButtonSetTypeFromString(value As String) As MsoButtonSetType
    If IsNumeric(value) Then
        MsoButtonSetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoButtonSetNone": MsoButtonSetTypeFromString = msoButtonSetNone
        Case "msoButtonSetOK": MsoButtonSetTypeFromString = msoButtonSetOK
        Case "msoButtonSetCancel": MsoButtonSetTypeFromString = msoButtonSetCancel
        Case "msoButtonSetOkCancel": MsoButtonSetTypeFromString = msoButtonSetOkCancel
        Case "msoButtonSetYesNo": MsoButtonSetTypeFromString = msoButtonSetYesNo
        Case "msoButtonSetYesNoCancel": MsoButtonSetTypeFromString = msoButtonSetYesNoCancel
        Case "msoButtonSetBackClose": MsoButtonSetTypeFromString = msoButtonSetBackClose
        Case "msoButtonSetNextClose": MsoButtonSetTypeFromString = msoButtonSetNextClose
        Case "msoButtonSetBackNextClose": MsoButtonSetTypeFromString = msoButtonSetBackNextClose
        Case "msoButtonSetRetryCancel": MsoButtonSetTypeFromString = msoButtonSetRetryCancel
        Case "msoButtonSetAbortRetryIgnore": MsoButtonSetTypeFromString = msoButtonSetAbortRetryIgnore
        Case "msoButtonSetSearchClose": MsoButtonSetTypeFromString = msoButtonSetSearchClose
        Case "msoButtonSetBackNextSnooze": MsoButtonSetTypeFromString = msoButtonSetBackNextSnooze
        Case "msoButtonSetTipsOptionsClose": MsoButtonSetTypeFromString = msoButtonSetTipsOptionsClose
        Case "msoButtonSetYesAllNoCancel": MsoButtonSetTypeFromString = msoButtonSetYesAllNoCancel
    End Select
End Function

Function MsoButtonSetTypeToString(value As MsoButtonSetType) As String
    Select Case value
        Case msoButtonSetNone: MsoButtonSetTypeToString = "msoButtonSetNone"
        Case msoButtonSetOK: MsoButtonSetTypeToString = "msoButtonSetOK"
        Case msoButtonSetCancel: MsoButtonSetTypeToString = "msoButtonSetCancel"
        Case msoButtonSetOkCancel: MsoButtonSetTypeToString = "msoButtonSetOkCancel"
        Case msoButtonSetYesNo: MsoButtonSetTypeToString = "msoButtonSetYesNo"
        Case msoButtonSetYesNoCancel: MsoButtonSetTypeToString = "msoButtonSetYesNoCancel"
        Case msoButtonSetBackClose: MsoButtonSetTypeToString = "msoButtonSetBackClose"
        Case msoButtonSetNextClose: MsoButtonSetTypeToString = "msoButtonSetNextClose"
        Case msoButtonSetBackNextClose: MsoButtonSetTypeToString = "msoButtonSetBackNextClose"
        Case msoButtonSetRetryCancel: MsoButtonSetTypeToString = "msoButtonSetRetryCancel"
        Case msoButtonSetAbortRetryIgnore: MsoButtonSetTypeToString = "msoButtonSetAbortRetryIgnore"
        Case msoButtonSetSearchClose: MsoButtonSetTypeToString = "msoButtonSetSearchClose"
        Case msoButtonSetBackNextSnooze: MsoButtonSetTypeToString = "msoButtonSetBackNextSnooze"
        Case msoButtonSetTipsOptionsClose: MsoButtonSetTypeToString = "msoButtonSetTipsOptionsClose"
        Case msoButtonSetYesAllNoCancel: MsoButtonSetTypeToString = "msoButtonSetYesAllNoCancel"
    End Select
End Function
