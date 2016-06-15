Attribute VB_Name = "wMsoWizardActType"
Function MsoWizardActTypeFromString(value As String) As MsoWizardActType
    If IsNumeric(value) Then
        MsoWizardActTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoWizardActInactive": MsoWizardActTypeFromString = msoWizardActInactive
        Case "msoWizardActActive": MsoWizardActTypeFromString = msoWizardActActive
        Case "msoWizardActSuspend": MsoWizardActTypeFromString = msoWizardActSuspend
        Case "msoWizardActResume": MsoWizardActTypeFromString = msoWizardActResume
    End Select
End Function

Function MsoWizardActTypeToString(value As MsoWizardActType) As String
    Select Case value
        Case msoWizardActInactive: MsoWizardActTypeToString = "msoWizardActInactive"
        Case msoWizardActActive: MsoWizardActTypeToString = "msoWizardActActive"
        Case msoWizardActSuspend: MsoWizardActTypeToString = "msoWizardActSuspend"
        Case msoWizardActResume: MsoWizardActTypeToString = "msoWizardActResume"
    End Select
End Function
