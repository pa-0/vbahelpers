Attribute VB_Name = "wMsoWizardMsgType"
Function MsoWizardMsgTypeFromString(value As String) As MsoWizardMsgType
    If IsNumeric(value) Then
        MsoWizardMsgTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoWizardMsgLocalStateOn": MsoWizardMsgTypeFromString = msoWizardMsgLocalStateOn
        Case "msoWizardMsgLocalStateOff": MsoWizardMsgTypeFromString = msoWizardMsgLocalStateOff
        Case "msoWizardMsgShowHelp": MsoWizardMsgTypeFromString = msoWizardMsgShowHelp
        Case "msoWizardMsgSuspending": MsoWizardMsgTypeFromString = msoWizardMsgSuspending
        Case "msoWizardMsgResuming": MsoWizardMsgTypeFromString = msoWizardMsgResuming
    End Select
End Function

Function MsoWizardMsgTypeToString(value As MsoWizardMsgType) As String
    Select Case value
        Case msoWizardMsgLocalStateOn: MsoWizardMsgTypeToString = "msoWizardMsgLocalStateOn"
        Case msoWizardMsgLocalStateOff: MsoWizardMsgTypeToString = "msoWizardMsgLocalStateOff"
        Case msoWizardMsgShowHelp: MsoWizardMsgTypeToString = "msoWizardMsgShowHelp"
        Case msoWizardMsgSuspending: MsoWizardMsgTypeToString = "msoWizardMsgSuspending"
        Case msoWizardMsgResuming: MsoWizardMsgTypeToString = "msoWizardMsgResuming"
    End Select
End Function
