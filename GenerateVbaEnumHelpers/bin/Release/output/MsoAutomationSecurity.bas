Attribute VB_Name = "wMsoAutomationSecurity"
Function MsoAutomationSecurityFromString(value As String) As MsoAutomationSecurity
    If IsNumeric(value) Then
        MsoAutomationSecurityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAutomationSecurityLow": MsoAutomationSecurityFromString = msoAutomationSecurityLow
        Case "msoAutomationSecurityByUI": MsoAutomationSecurityFromString = msoAutomationSecurityByUI
        Case "msoAutomationSecurityForceDisable": MsoAutomationSecurityFromString = msoAutomationSecurityForceDisable
    End Select
End Function

Function MsoAutomationSecurityToString(value As MsoAutomationSecurity) As String
    Select Case value
        Case msoAutomationSecurityLow: MsoAutomationSecurityToString = "msoAutomationSecurityLow"
        Case msoAutomationSecurityByUI: MsoAutomationSecurityToString = "msoAutomationSecurityByUI"
        Case msoAutomationSecurityForceDisable: MsoAutomationSecurityToString = "msoAutomationSecurityForceDisable"
    End Select
End Function
