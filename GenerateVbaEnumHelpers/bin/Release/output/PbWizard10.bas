Attribute VB_Name = "wPbWizard10"
Function PbWizard10FromString(value As String) As PbWizard10
    If IsNumeric(value) Then
        PbWizard10FromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWizardWebSites": PbWizard10FromString = pbWizardWebSites
        Case "pbWizardGreetingCards": PbWizard10FromString = pbWizardGreetingCards
        Case "pbWizardInvitations": PbWizard10FromString = pbWizardInvitations
    End Select
End Function

Function PbWizard10ToString(value As PbWizard10) As String
    Select Case value
        Case pbWizardWebSites: PbWizard10ToString = "pbWizardWebSites"
        Case pbWizardGreetingCards: PbWizard10ToString = "pbWizardGreetingCards"
        Case pbWizardInvitations: PbWizard10ToString = "pbWizardInvitations"
    End Select
End Function
